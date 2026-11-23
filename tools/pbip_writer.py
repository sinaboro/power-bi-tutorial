"""pbip_writer.py - Power BI 프로젝트(PBIP) 폴더를 코드로 만드는 작은 라이브러리 (교안 제작용)

Power BI Desktop 은 .pbix 말고도 "Power BI 프로젝트(.pbip)" 라는 글자 파일 형식을 열 수 있다.
  이름.pbip                       ← Desktop 으로 여는 파일
  이름.SemanticModel/definition   ← 의미 체계 모델 (테이블·열·측정값·관계) : TMDL 글자 파일
  이름.Report/definition          ← 보고서 (페이지·시각적 개체) : PBIR JSON 파일
교안의 장별 완성 파일(.pbix)은 이 라이브러리로 PBIP 를 만든 뒤, Desktop 으로 열어
새로 고침하고 [다른 이름으로 저장 > .pbix] 해서 만든다 (tools/make_pbix.py).
"""
import json
import re
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
def q(name):
    """TMDL 이름: 글자·숫자·_ 만 있으면 그대로, 빈칸이나 기호가 있으면 '이름' 으로 감싼다"""
    return name if re.fullmatch(r"\w+", name) else "'" + name.replace("'", "''") + "'"


class Column:
    def __init__(self, name, dtype, fmt=None, summarize=None, expr=None, sort_by=None, is_key=False, hidden=False):
        self.name, self.dtype, self.fmt, self.summarize = name, dtype, fmt, summarize
        self.expr, self.sort_by, self.is_key, self.hidden = expr, sort_by, is_key, hidden

    def tmdl(self):
        head = f"\tcolumn {q(self.name)}" + (f" = {self.expr}" if self.expr else "")
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
            lines.append(f"\t\tsourceColumn: {'[' + self.name + ']' if getattr(self, 'calc_source', False) else q(self.name)}")
        if self.sort_by:
            lines.append(f"\t\tsortByColumn: {q(self.sort_by)}")
        return "\n".join(lines)


class Measure:
    def __init__(self, name, expr, fmt="#,0", folder=None):
        self.name, self.expr, self.fmt, self.folder = name, expr, fmt, folder

    def tmdl(self):
        expr = self.expr.strip("\n")
        if "\n" in expr:      # 여러 줄 식(VAR … RETURN)은 다음 줄부터 들여 쓴다
            expr = "\n" + "\n".join("\t\t\t" + line for line in expr.split("\n"))
        lines = [f"\tmeasure {q(self.name)} = {expr}", f"\t\tformatString: {self.fmt}"]
        if self.folder:
            lines.append(f"\t\tdisplayFolder: {self.folder}")
        return "\n".join(lines)


class Table:
    def __init__(self, name, columns, m=None, dax=None, measures=(), data_category=None, hierarchies=None):
        self.name, self.columns, self.m, self.dax = name, columns, m, dax
        self.measures, self.data_category = list(measures), data_category
        self.hierarchies = hierarchies or {}          # {"계층 이름": ["위 수준 열", "아래 수준 열", ...]}

    def tmdl(self):
        out = [f"table {q(self.name)}"]
        if self.data_category:
            out.append(f"\tdataCategory: {self.data_category}")
        out.append("")
        for ms in self.measures:
            out += [ms.tmdl(), ""]
        for c in self.columns:
            out += [c.tmdl(), ""]
        for hname, levels in self.hierarchies.items():       # 계층 (드릴다운용)
            out.append(f"\thierarchy {q(hname)}")
            for lv in levels:
                out += [f"\t\tlevel {q(lv)}", f"\t\t\tcolumn: {q(lv)}"]
            out.append("")
        if self.m:
            body = "\n".join("\t\t\t\t" + line for line in self.m.strip("\n").split("\n"))
            out += [f"\tpartition {q(self.name)} = m", "\t\tmode: import", "\t\tsource =", body, ""]
        else:
            out += [f"\tpartition {q(self.name)} = calculated", "\t\tmode: import", f"\t\tsource = {self.dax}", ""]
        return "\n".join(out)


class Expression:
    """모델에 테이블로 로드하지 않는 Power Query 쿼리, 또는 매개 변수 (expressions.tmdl 에 들어간다)"""

    def __init__(self, name, m, parameter=False):
        self.name, self.m, self.parameter = name, m, parameter

    def tmdl(self):
        if self.parameter:   # 매개 변수 = 값 하나 + "이것은 매개 변수" 표시(meta)
            head = (f"expression {q(self.name)} = {self.m} meta [IsParameterQuery=true, Type=\"Text\", "
                    f"IsParameterQueryRequired=true]")
            return "\n".join([head, "\tannotation PBI_ResultType = Text", ""])
        body = "\n".join("\t\t" + line for line in self.m.strip("\n").split("\n"))
        return "\n".join([f"expression {q(self.name)} =", body, "", "\tannotation PBI_ResultType = Table", ""])


def write_model(folder: Path, tables, relationships=(), expressions=(), roles=None):
    """relationships: [(from_table, from_col, to_table, to_col), ...]  (다 쪽 → 일 쪽)
       다섯째 값으로 {"active": False}(비활성 관계) 나 {"both": True}(양방향) 를 줄 수 있다.
       expressions: 로드하지 않는 쿼리·매개 변수 (Expression)
       roles: {"역할 이름": {"테이블": "DAX 필터 식"}}  (행 수준 보안 RLS)"""
    _w(folder / "definition.pbism", {"version": "4.2", "settings": {}})
    d = folder / "definition"
    _w(d / "database.tmdl", "database\n\tcompatibilityLevel: 1600\n")
    model = ["model Model", "\tculture: ko-KR", "\tdefaultPowerBIDataSourceVersion: powerBI_V3",
             "\tsourceQueryCulture: ko-KR", ""] + [f"ref table {q(t.name)}" for t in tables]
    _w(d / "model.tmdl", "\n".join(model) + "\n")
    for t in tables:
        _w(d / "tables" / f"{t.name}.tmdl", t.tmdl())
    if expressions:
        _w(d / "expressions.tmdl", "\n".join(e.tmdl() for e in expressions))
    if relationships:
        rel = []
        for i, r in enumerate(relationships, 1):
            ft, fc, tt, tc = r[:4]
            opt = r[4] if len(r) > 4 else {}
            rel.append(f"relationship r{i}")
            if opt.get("active") is False:
                rel.append("\tisActive: false")
            if opt.get("both"):
                rel.append("\tcrossFilteringBehavior: bothDirections")
            rel += [f"\tfromColumn: {q(ft)}.{q(fc)}", f"\ttoColumn: {q(tt)}.{q(tc)}", ""]
        _w(d / "relationships.tmdl", "\n".join(rel))
    for rname, perms in (roles or {}).items():
        lines = [f"role {q(rname)}", "\tmodelPermission: read", ""]
        for tname, expr in perms.items():
            lines += [f"\ttablePermission {q(tname)} = {expr}", ""]
        _w(d / "roles" / f"{rname}.tmdl", "\n".join(lines))


# ====================================================================== 보고서 (PBIR)
def col(table, column):
    return {"Column": {"Expression": {"SourceRef": {"Entity": table}}, "Property": column}}


def mea(table, measure):
    return {"Measure": {"Expression": {"SourceRef": {"Entity": table}}, "Property": measure}}


# 암시적 집계 함수 번호 (Power BI 내부 번호) → (queryRef 이름, 한글 이름)
AGG = {0: ("Sum", "합계"), 1: ("Avg", "평균"), 2: ("Count", "개수"), 3: ("Min", "최소값"), 4: ("Max", "최대값"),
       5: ("CountNonNull", "개수")}


def _proj(field):
    if "Aggregation" in field:  # 열을 그대로 끌어다 놓은 암시적 집계 (합계, 개수 …)
        c = field["Aggregation"]["Expression"]["Column"]
        t, p = c["Expression"]["SourceRef"]["Entity"], c["Property"]
        fn, ko = AGG[field["Aggregation"]["Function"]]
        return {"field": field, "queryRef": f"{fn}({t}.{p})", "nativeQueryRef": f"{p}의 {ko}"}
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

    def __init__(self, vtype, x, y, w, h, roles=None, title=None, sort=None, objects=None, container=None, name=None,
                 cf=None):
        self.vtype, self.pos = vtype, (x, y, w, h)
        self.roles, self.title, self.sort = roles or {}, title, sort
        self.objects, self.container, self.name = objects or {}, container or {}, name
        self.cf = cf or {}          # 조건부 서식: {"측정값 이름": 배경색 규칙}  (표·행렬의 값 셀)
        self.filters = []           # 이 시각적 개체에만 거는 필터 (예: 상위 N)

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
        for measure, rule in self.cf.items():      # 값 셀 배경색을 측정값 크기에 따라 칠한다
            ref = next(_proj(f)["queryRef"] for fs in self.roles.values() for f in fs
                       if "Measure" in f and f["Measure"]["Property"] == measure)
            visual.setdefault("objects", {}).setdefault("values", []).append({
                "properties": {"backColor": rule},
                "selector": {"data": [{"dataViewWildcard": {"matchingOption": 1}}], "metadata": ref}})
        vco = dict(self.container)
        if self.title is not None:
            vco["title"] = {"show": True, "text": self.title}
        if vco:
            visual["visualContainerObjects"] = {k: [{"properties": {pk: (pv if isinstance(pv, dict) else lit(pv)) for pk, pv in v.items()}}]
                                                for k, v in vco.items()}
        visual["drillFilterOtherVisuals"] = True
        out = {"$schema": f"{SCHEMA}/visualContainer/{VER['visual']}/schema.json", "name": name,
               "position": {"x": x, "y": y, "z": z, "width": w, "height": h, "tabOrder": z}, "visual": visual}
        if self.filters:
            out["filterConfig"] = {"filters": self.filters}
        return out


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
        self.binding = None     # 드릴스루 / 도구 설명 페이지 설정 (pageBinding)
        self.filters = []       # 페이지 필터 (드릴스루 필터 등)
        self.hidden = False     # True 면 보기 모드에서 페이지 탭을 숨긴다


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
        tname = "OlistTheme" if "olist" in str(custom_theme).lower() else "HaruCoffeeTheme"
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
        if p.binding:
            pj["pageBinding"] = p.binding
            pj["type"] = p.binding["type"]          # 페이지 용도도 따로 적어야 도구 설명·드릴스루 페이지로 동작한다
        if p.filters:
            pj["filterConfig"] = {"filters": p.filters}
        if p.hidden:
            pj["visibility"] = "HiddenInViewMode"
        if p.width != 1280 or p.height != 720:
            pj["displayOption"] = "ActualSize"
        _w(d / "pages" / p.name / "page.json", pj)
        for i, v in enumerate(p.visuals):
            vname = v.name or f"{p.name}_v{i + 1:02d}"
            _w(d / "pages" / p.name / "visuals" / vname / "visual.json", v.json(vname, i))


def write_project(out: Path, name: str, tables, relationships, pages, custom_theme=None, expressions=(), roles=None):
    shutil.rmtree(out, ignore_errors=True)
    out.mkdir(parents=True)
    _w(out / f"{name}.pbip", {"version": "1.0", "artifacts": [{"report": {"path": f"{name}.Report"}}],
                              "settings": {"enableAutoRecovery": True}})
    write_model(out / f"{name}.SemanticModel", tables, relationships, expressions, roles)
    write_report(out / f"{name}.Report", f"{name}.SemanticModel", pages, custom_theme)
    return out / f"{name}.pbip"
