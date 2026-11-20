"""pbip_writer.py - Power BI 프로젝트(PBIP) 폴더를 코드로 만드는 작은 라이브러리 (교안 제작용)

Power BI Desktop 은 .pbix 말고도 "Power BI 프로젝트(.pbip)" 라는 글자 파일 형식을 열 수 있다.
  이름.pbip                       ← Desktop 으로 여는 파일
  이름.SemanticModel/definition   ← 의미 체계 모델 (테이블·열·측정값·관계) : TMDL 글자 파일
  이름.Report/definition          ← 보고서 (페이지·시각적 개체) : PBIR JSON 파일
교안의 장별 완성 파일(.pbix)은 이 라이브러리로 PBIP 를 만든 뒤, Desktop 으로 열어
새로 고침하고 [다른 이름으로 저장 > .pbix] 해서 만든다 (tools/make_pbix.py).
"""
import json
import shutil
from pathlib import Path

THEME_SRC = Path(r"C:\Program Files\WindowsApps\Microsoft.MicrosoftPowerBIDesktop_2.159.1231.0_x64__8wekyb3d8bbwe"
                 r"\bin\WebView2Resources\minerva\sharedresources\BaseThemes\CY26SU10.json")
SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition"
VER = {"visual": "2.13.0", "report": "3.3.0", "page": "2.1.0"}


def _w(path: Path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    if isinstance(data, str):
        path.write_text(data, encoding="utf-8")
    else:
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


# ====================================================================== 의미 체계 모델 (TMDL)
class Column:
    def __init__(self, name, dtype, fmt=None, summarize=None, expr=None, sort_by=None, is_key=False, hidden=False):
        self.name, self.dtype, self.fmt, self.summarize = name, dtype, fmt, summarize
        self.expr, self.sort_by, self.is_key, self.hidden = expr, sort_by, is_key, hidden

    def tmdl(self):
        head = f"\tcolumn {self.name}" + (f" = {self.expr}" if self.expr else "")
        lines = [head, f"\t\tdataType: {self.dtype}"]
        if self.is_key:
            lines.append("\t\tisKey")
        if self.fmt:
            lines.append(f"\t\tformatString: {self.fmt}")
        if self.summarize:
            lines.append(f"\t\tsummarizeBy: {self.summarize}")
        if self.hidden:
            lines.append("\t\tisHidden")
        if not self.expr:
            lines.append(f"\t\tsourceColumn: {'[' + self.name + ']' if getattr(self, 'calc_source', False) else self.name}")
        if self.sort_by:
            lines.append(f"\t\tsortByColumn: {self.sort_by}")
        return "\n".join(lines)


class Measure:
    def __init__(self, name, expr, fmt="#,0", folder=None):
        self.name, self.expr, self.fmt, self.folder = name, expr, fmt, folder

    def tmdl(self):
        lines = [f"\tmeasure {self.name} = {self.expr}", f"\t\tformatString: {self.fmt}"]
        if self.folder:
            lines.append(f"\t\tdisplayFolder: {self.folder}")
        return "\n".join(lines)


class Table:
    def __init__(self, name, columns, m=None, dax=None, measures=(), data_category=None):
        self.name, self.columns, self.m, self.dax = name, columns, m, dax
        self.measures, self.data_category = list(measures), data_category

    def tmdl(self):
        out = [f"table {self.name}"]
        if self.data_category:
            out.append(f"\tdataCategory: {self.data_category}")
        out.append("")
        for ms in self.measures:
            out += [ms.tmdl(), ""]
        for c in self.columns:
            out += [c.tmdl(), ""]
        if self.m:
            body = "\n".join("\t\t\t\t" + line for line in self.m.strip("\n").split("\n"))
            out += [f"\tpartition {self.name} = m", "\t\tmode: import", "\t\tsource =", body, ""]
        else:
            out += [f"\tpartition {self.name} = calculated", "\t\tmode: import", f"\t\tsource = {self.dax}", ""]
        return "\n".join(out)


def write_model(folder: Path, tables, relationships=()):
    """relationships: [(from_table, from_col, to_table, to_col), ...]  (다 쪽 → 일 쪽)"""
    _w(folder / "definition.pbism", {"version": "4.2", "settings": {}})
    d = folder / "definition"
    _w(d / "database.tmdl", "database\n\tcompatibilityLevel: 1600\n")
    model = ["model Model", "\tculture: ko-KR", "\tdefaultPowerBIDataSourceVersion: powerBI_V3",
             "\tsourceQueryCulture: ko-KR", ""] + [f"ref table {t.name}" for t in tables]
    _w(d / "model.tmdl", "\n".join(model) + "\n")
    for t in tables:
        _w(d / "tables" / f"{t.name}.tmdl", t.tmdl())
    if relationships:
        rel = []
        for i, (ft, fc, tt, tc) in enumerate(relationships, 1):
            rel += [f"relationship r{i}", f"\tfromColumn: {ft}.{fc}", f"\ttoColumn: {tt}.{tc}", ""]
        _w(d / "relationships.tmdl", "\n".join(rel))


# ====================================================================== 보고서 (PBIR)
def col(table, column):
    return {"Column": {"Expression": {"SourceRef": {"Entity": table}}, "Property": column}}


def mea(table, measure):
    return {"Measure": {"Expression": {"SourceRef": {"Entity": table}}, "Property": measure}}


def _proj(field):
    if "Aggregation" in field:  # 열을 그대로 끌어다 놓은 암시적 합계
        c = field["Aggregation"]["Expression"]["Column"]
        t, p = c["Expression"]["SourceRef"]["Entity"], c["Property"]
        return {"field": field, "queryRef": f"Sum({t}.{p})", "nativeQueryRef": f"{p}의 합계"}
    kind = "Column" if "Column" in field else "Measure"
    t = field[kind]["Expression"]["SourceRef"]["Entity"]
    p = field[kind]["Property"]
    return {"field": field, "queryRef": f"{t}.{p}", "nativeQueryRef": p}


def lit(v):
    """PBIR 서식 값(Literal) - 글자는 '...' 로 감싸고, 숫자는 D/L 접미사"""
    if isinstance(v, bool):
        return {"expr": {"Literal": {"Value": "true" if v else "false"}}}
    if isinstance(v, (int, float)):
        return {"expr": {"Literal": {"Value": f"{v}D"}}}
    return {"expr": {"Literal": {"Value": f"'{v}'"}}}


class Visual:
    """시각적 개체 하나. roles = {"Category": [field...], "Y": [field...]} 같은 모양"""

    def __init__(self, vtype, x, y, w, h, roles=None, title=None, sort=None, objects=None, container=None, name=None):
        self.vtype, self.pos = vtype, (x, y, w, h)
        self.roles, self.title, self.sort = roles or {}, title, sort
        self.objects, self.container, self.name = objects or {}, container or {}, name

    def json(self, name, z):
        x, y, w, h = self.pos
        visual = {"visualType": self.vtype}
        if self.roles:
            qs = {role: {"projections": [_proj(f) for f in fields]} for role, fields in self.roles.items()}
            visual["query"] = {"queryState": qs}
            if self.sort:
                field, direction = self.sort
                visual["query"]["sortDefinition"] = {"sort": [{"field": field, "direction": direction}],
                                                     "isDefaultSort": True}
        if self.objects:
            visual["objects"] = {k: [{"properties": {pk: (pv if isinstance(pv, dict) else lit(pv)) for pk, pv in v.items()}}]
                                 for k, v in self.objects.items()}
        vco = dict(self.container)
        if self.title is not None:
            vco["title"] = {"show": True, "text": self.title}
        if vco:
            visual["visualContainerObjects"] = {k: [{"properties": {pk: (pv if isinstance(pv, dict) else lit(pv)) for pk, pv in v.items()}}]
                                                for k, v in vco.items()}
        visual["drillFilterOtherVisuals"] = True
        return {"$schema": f"{SCHEMA}/visualContainer/{VER['visual']}/schema.json", "name": name,
                "position": {"x": x, "y": y, "z": z, "width": w, "height": h, "tabOrder": z}, "visual": visual}


class TextBox(Visual):
    def __init__(self, text, x, y, w, h, size=20, bold=True, color="#3E2723"):
        super().__init__("textbox", x, y, w, h)
        self.text, self.size, self.bold, self.color = text, size, bold, color

    def json(self, name, z):
        x, y, w, h = self.pos
        run = {"value": self.text, "textStyle": {"fontSize": f"{self.size}pt", "color": self.color}}
        if self.bold:
            run["textStyle"]["fontWeight"] = "bold"
        para = {"textRuns": [run]}
        return {"$schema": f"{SCHEMA}/visualContainer/{VER['visual']}/schema.json", "name": name,
                "position": {"x": x, "y": y, "z": z, "width": w, "height": h, "tabOrder": z},
                "visual": {"visualType": "textbox",
                           "objects": {"general": [{"properties": {"paragraphs": [para]}}]},
                           "drillFilterOtherVisuals": True}}


class Page:
    def __init__(self, name, display, visuals, width=1280, height=720, background=None):
        self.name, self.display, self.visuals = name, display, visuals
        self.width, self.height, self.background = width, height, background


def write_report(folder: Path, model_folder_name: str, pages, custom_theme: Path = None):
    _w(folder / "definition.pbir", {"$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json",
                                    "version": "4.0", "datasetReference": {"byPath": {"path": f"../{model_folder_name}"}}})
    d = folder / "definition"
    _w(d / "version.json", {"$schema": f"{SCHEMA}/versionMetadata/1.0.0/schema.json", "version": "2.0.0"})
    report = {"$schema": f"{SCHEMA}/report/{VER['report']}/schema.json",
              "themeCollection": {"baseTheme": {"name": "CY26SU10", "reportVersionAtImport": VER, "type": "SharedResources"}},
              "resourcePackages": [{"name": "SharedResources", "type": "SharedResources",
                                    "items": [{"name": "CY26SU10", "path": "BaseThemes/CY26SU10.json", "type": "BaseTheme"}]}]}
    base = folder / "StaticResources" / "SharedResources" / "BaseThemes"
    base.mkdir(parents=True, exist_ok=True)
    shutil.copy(THEME_SRC, base / "CY26SU10.json")
    if custom_theme:
        # 보고서 안의 리소스 파일 이름은 영문으로 둔다 (한글 이름이면 .pbix 로 저장할 때 "올바른 파일 경로가 아닙니다" 오류)
        tname = "HaruCoffeeTheme"
        report["themeCollection"]["customTheme"] = {"name": tname + ".json", "reportVersionAtImport": VER, "type": "RegisteredResources"}
        report["resourcePackages"].append({"name": "RegisteredResources", "type": "RegisteredResources",
                                           "items": [{"name": tname + ".json", "path": tname + ".json", "type": "CustomTheme"}]})
        reg = folder / "StaticResources" / "RegisteredResources"
        reg.mkdir(parents=True, exist_ok=True)
        shutil.copy(custom_theme, reg / (tname + ".json"))
    _w(d / "report.json", report)
    _w(d / "pages" / "pages.json", {"$schema": f"{SCHEMA}/pagesMetadata/1.0.0/schema.json",
                                    "pageOrder": [p.name for p in pages], "activePageName": pages[0].name})
    for p in pages:
        pj = {"$schema": f"{SCHEMA}/page/{VER['page']}/schema.json", "name": p.name, "displayName": p.display,
              "displayOption": "FitToPage", "height": p.height, "width": p.width}
        if p.background:
            pj["objects"] = {"background": [{"properties": {"color": {"solid": {"color": lit(p.background)}},
                                                            "transparency": lit(0)}}]}
        _w(d / "pages" / p.name / "page.json", pj)
        for i, v in enumerate(p.visuals):
            vname = v.name or f"{p.name}_v{i + 1:02d}"
            _w(d / "pages" / p.name / "visuals" / vname / "visual.json", v.json(vname, i))


def write_project(out: Path, name: str, tables, relationships, pages, custom_theme=None):
    shutil.rmtree(out, ignore_errors=True)
    out.mkdir(parents=True)
    _w(out / f"{name}.pbip", {"version": "1.0", "artifacts": [{"report": {"path": f"{name}.Report"}}],
                              "settings": {"enableAutoRecovery": True}})
    write_model(out / f"{name}.SemanticModel", tables, relationships)
    write_report(out / f"{name}.Report", f"{name}.SemanticModel", pages, custom_theme)
    return out / f"{name}.pbip"
