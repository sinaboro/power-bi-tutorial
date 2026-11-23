"""make_pbix.py - 장별 완성 파일(.pbix) 만들기 (교안 제작용, Power BI Desktop 필요)

    python tools/make_pbix.py 7            # 7장 완성 파일만
    python tools/make_pbix.py all          # 1~7장 + 과제 풀이 전부
    python tools/make_pbix.py 7 --keep     # 저장 후 Desktop 을 닫지 않는다 (화면 확인용)

하는 일 (장마다)
  1. course.py 의 정의로 Power BI 프로젝트(PBIP)를 build/ 에 만든다.
  2. Power BI Desktop 으로 연다 → [홈 > 새로 고침] 으로 데이터를 불러온다.
  3. 모델에 DAX 질의를 보내 행 수·측정값이 기대한 값과 같은지 확인한다 (tools/dax_query.ps1).
  4. 페이지마다 화면을 캡처한다 → docs/images/pbi_<장>_<페이지>.png
  5. [파일 > 다른 이름으로 저장 > 이 장치 찾아보기] 로 pbix/<장>.pbix 를 저장하고 Desktop 을 닫는다.
"""
import os
import subprocess
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import course
import pbi_ui
from pbip_writer import write_project

ROOT = Path(__file__).resolve().parent.parent
BUILD = ROOT / "build"
EXE = (r"C:\Program Files\WindowsApps\Microsoft.MicrosoftPowerBIDesktop_2.159.1231.0_x64__8wekyb3d8bbwe"
       r"\bin\PBIDesktop.exe")
QFILE = Path(os.environ["TEMP"]) / "adomd" / "q.dax"


def dax(query):
    """열린 모델에 DAX 를 보내 결과 표(글자)를 돌려준다"""
    QFILE.write_text(query, encoding="utf-8")
    r = subprocess.run(["powershell", "-NoProfile", "-File", str(ROOT / "tools" / "dax_query.ps1"), "-QueryFile", str(QFILE)],
                       capture_output=True, text=True, encoding="utf-8", errors="replace")
    return (r.stdout + r.stderr).strip()


def close_desktop():
    subprocess.run(["powershell", "-NoProfile", "-Command",
                    "Get-Process PBIDesktop,msmdsrv -ErrorAction SilentlyContinue | Stop-Process -Force"], capture_output=True)
    time.sleep(3)


def dismiss_popups(w):
    """처음 열 때 뜨는 안내 풍선(Direct Lake 등)의 닫기 단추를 누른다"""
    for b in w.descendants(control_type="Button"):
        try:
            if b.window_text() in ("닫기", "Close") and b.is_visible():
                r, wr = b.rectangle(), w.rectangle()
                if r.top - wr.top > 30:          # 창 오른쪽 위의 [닫기] 는 건드리지 않는다
                    b.click_input()
        except Exception:
            pass


def refresh_and_wait(w, table="판매", timeout=240):
    import pbi_refresh
    pbi_refresh.refresh(w)
    end = time.time() + timeout
    while time.time() < end:
        time.sleep(5)
        out = dax(f"EVALUATE ROW(\"n\", COUNTROWS('{table}'))")
        nums = [s for s in out.split() if s.isdigit()]
        if nums and int(nums[-1]) > 0:
            time.sleep(8)                         # 나머지 테이블·시각적 개체가 그려질 시간
            return int(nums[-1])
    raise RuntimeError("새로 고침이 끝나지 않았습니다: " + out)


def goto_page(w, display):
    tabs = [t for t in w.descendants(control_type="TabItem") if t.window_text() == display and t.is_visible()]
    if tabs:
        tabs[-1].click_input()
        time.sleep(4)


def save_as(w, target: Path):
    """[파일] → [다른 이름으로 저장] → [이 장치 찾아보기] → 파일 형식 .pbix → 저장"""
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists():
        target.unlink()
    w.set_focus()
    [e for e in w.descendants() if e.window_text() == "파일" and e.is_visible()][0].click_input()
    time.sleep(2)
    [e for e in w.descendants(control_type="TabItem") if e.window_text() == "다른 이름으로 저장" and e.is_visible()][0].click_input()
    time.sleep(2)
    [e for e in w.descendants(control_type="Button") if e.window_text() == "이 장치 찾아보기" and e.is_visible()][0].click_input()
    for _ in range(20):
        time.sleep(1)
        dlg = [c for c in w.children() if c.element_info.class_name == "#32770"]
        if dlg:
            break
    d = dlg[0]
    ft = [e for e in d.descendants(control_type="ComboBox") if e.element_info.automation_id == "FileTypeControlHost"][0]
    ft.expand()
    time.sleep(1)
    [i for i in ft.descendants(control_type="ListItem") if "pbix" in i.window_text()][0].select()
    time.sleep(1)
    [e for e in d.descendants(control_type="Edit") if e.element_info.automation_id == "1001"][0].set_edit_text(str(target))
    [e for e in d.descendants(control_type="Button") if e.element_info.automation_id == "1"][0].click_input()
    for _ in range(120):
        time.sleep(2)
        if target.exists() and target.stat().st_size > 0 and "pbip" not in w.window_text().lower():
            break
    time.sleep(6)
    return target.stat().st_size


def build(key, keep=False, shots=True):
    if key == "과제":
        import course_cvs as mod
        name, proj = "과제_편의점_풀이", mod.project()
    elif key.startswith("a"):                      # 심화 과정: a1, a2, … (course_adv.py)
        import course_adv as mod
        stage = int(key[1:])
        name, proj = mod.STAGES[stage], mod.project(stage)
    else:
        stage = int(key)
        name, proj = course.STAGES[stage], course.project(stage)
    close_desktop()
    pbip = write_project(BUILD / name, name, **proj)
    subprocess.Popen([EXE, str(pbip)])
    w = pbi_ui.main_window(timeout=120)
    time.sleep(15)
    w = pbi_ui.main_window()
    w.set_focus()
    dismiss_popups(w)
    names_ = [t.name for t in proj["tables"]]
    first = "주문상품" if "주문상품" in names_ else names_[0]      # 심화: 가장 큰 테이블이 다 불러와졌는지 본다
    n = refresh_and_wait(w, first)
    print(f"[{name}] 불러온 행 수: {first} = {n:,}")
    dismiss_popups(w)
    # 측정값이 있는 장은 핵심 숫자를 출력해 pandas 로 계산한 기대값과 비교한다
    names = [m.name for m in proj["tables"][0].measures]
    if names:
        q = "EVALUATE ROW(" + ", ".join(f'"{n}", [{n}]' for n in names) + ")"
        print("  측정값:", " ".join(dax(q).split()))
    if shots:
        for i, p in enumerate(proj["pages"]):
            goto_page(w, p.display)
            out = ROOT / "docs" / "images" / f"pbi_{name}_p{i + 1}.png"
            pbi_ui.shot(str(out), w)
            print("  캡처:", out.name)
        if len(proj["pages"]) > 1:                   # 학생이 파일을 열면 첫 페이지가 보이게 돌아가서 저장
            goto_page(w, proj["pages"][0].display)
    if "--nosave" in sys.argv:
        return w
    tmp = ROOT / "pbix" / f"stage_{key}.pbix"          # 저장 대화 상자에는 영문 이름으로 저장한 뒤
    size = save_as(w, tmp)
    final = ROOT / "pbix" / f"{name}.pbix"              # Desktop 을 닫은 뒤 한글 이름으로 바꾼다
    if keep:
        print("  --keep: Desktop 을 열어 둔다. 파일 이름은", tmp.name)
        return w
    close_desktop()
    if final.exists():
        final.unlink()
    tmp.rename(final)
    print(f"  저장: pbix/{name}.pbix ({size / 1024 / 1024:.1f} MB)")
    return w


if __name__ == "__main__":
    keep = "--keep" in sys.argv
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if args == ["all"]:
        keys = [str(k) for k in course.STAGES] + ["과제"]
    elif args == ["adv"]:                          # 심화 과정 전체
        import course_adv
        keys = [f"a{k}" for k in course_adv.STAGES]
    else:
        keys = args
    for k in keys:
        build(k, keep=keep)
