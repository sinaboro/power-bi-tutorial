"""capture_ui.py - 교안에 넣을 Power BI Desktop 조작 화면을 자동으로 캡처한다 (교안 제작용)

    python tools/capture_ui.py            # 모든 화면
    python tools/capture_ui.py pq model   # 일부만 (start, getdata, pq, tableview, model, measure, fields, format)

장별 완성 파일(pbix/*.pbix)이 먼저 있어야 한다 (tools/make_pbix.py).
캡처는 docs/images/ui_*.png 로 저장한다. 창 영역만 캡처하므로 다른 프로그램은 찍히지 않는다.
"""
import subprocess
import sys
import time
from pathlib import Path

import pywinauto.mouse as mouse
from pywinauto import Desktop

sys.path.insert(0, str(Path(__file__).parent))
import pbi_ui
from make_pbix import EXE, close_desktop, dismiss_popups

ROOT = Path(__file__).resolve().parent.parent
IMG = ROOT / "docs" / "images"
DATA = Path(r"C:\study\power-bi-tutorial\data")


def visible(w, text, ctype=None, exact=True):
    out = []
    for e in w.descendants(control_type=ctype) if ctype else w.descendants():
        try:
            t = e.window_text()
            if (t == text if exact else text in t) and e.is_visible() and e.rectangle().width() > 0:
                out.append(e)
        except Exception:
            pass
    return out


def click(w, text, ctype=None, exact=True, index=0, wait=2.0):
    found = visible(w, text, ctype, exact)
    if not found:
        raise RuntimeError(f"'{text}' 을(를) 찾지 못했습니다")
    found[index].click_input()
    time.sleep(wait)


def select_table(w, table):
    """데이터 창에서 테이블 이름(TreeItem)을 눌러 그 테이블을 고른다 (펼치기 단추가 아니라)"""
    items = [e for e in w.descendants(control_type="TreeItem")
             if e.is_visible() and e.window_text() in (f"테이블 {table}", f"계산된 테이블 {table}")]
    items[0].click_input()
    time.sleep(4)


def shot(name, win):
    r = win.rectangle()
    mouse.move(coords=(r.left + 700, r.bottom - 15))     # 마우스를 치워 도구 설명 풍선이 안 찍히게
    time.sleep(1)
    path = IMG / f"ui_{name}.png"
    pbi_ui.shot(str(path), win)
    print("  캡처:", path.name)


def open_pbix(name):
    close_desktop()
    subprocess.Popen([EXE, str(ROOT / "pbix" / f"{name}.pbix")])
    w = pbi_ui.main_window(timeout=120)
    time.sleep(20)
    w = pbi_ui.main_window()
    w.set_focus()
    dismiss_popups(w)
    time.sleep(2)
    return w


def child_window(title, exact=True):
    """PBI 본 창 안에 뜨는 대화 상자(탐색 창, CSV 미리 보기 등)를 제목으로 찾는다"""
    for c in pbi_ui.main_window().children():
        t = c.window_text()
        if c.element_info.control_type == "Window" and (t == title if exact else title in t):
            return c
    raise RuntimeError(f"창 '{title}' 을(를) 찾지 못했습니다")


def top_windows():
    return [d for d in Desktop(backend="uia").windows() if pbi_ui._proc_name(d.element_info.process_id) == "PBIDesktop.exe"]


# ---------------------------------------------------------------- 1장: 시작 화면, 데이터 가져오기
def cap_start():
    close_desktop()
    subprocess.Popen([EXE])
    w = pbi_ui.main_window(timeout=120)
    time.sleep(20)
    w = pbi_ui.main_window()
    path = IMG / "ui_start.png"                           # 시작(홈) 화면 - 아래쪽 최근 파일 목록은 잘라 낸다
    pbi_ui.shot(str(path), w, crop=(0, 0, w.rectangle().width(), 620))
    print("  캡처:", path.name)
    r = w.rectangle()                                     # 시작 화면은 웹 화면이라 이름으로 못 찾는다 → [빈 보고서] 타일 위치를 누른다
    mouse.click(coords=(r.left + 336, r.top + 195))
    time.sleep(10)
    w = pbi_ui.main_window()
    dismiss_popups(w)
    shot("blank", w)


def open_file(dlg, path):
    """Windows 열기 대화 상자의 파일 이름 칸에 경로를 넣고 Enter ([열기]는 분할 단추라 Enter 가 확실하다)"""
    edit = [e for e in dlg.descendants(control_type="Edit") if e.element_info.automation_id == "1148"][0]
    edit.set_edit_text(str(path))
    edit.type_keys("{ENTER}")


def cap_getdata():
    w = pbi_ui.main_window()
    # [Excel 통합 문서] → 파일 선택 → 탐색 창
    click(w, "Excel 통합 문서", "Button", wait=3)
    dlg = [c for c in w.children() if c.element_info.class_name == "#32770"][0]
    open_file(dlg, DATA / "카페_메뉴.xlsx")
    time.sleep(10)
    nav = [child_window("탐색 창")]                      # 탐색 창은 본 창 안의 자식 창이다
    try:
        cb = visible(nav[0], "메뉴", "CheckBox", exact=False)[0].rectangle()
        mouse.click(coords=(cb.left + 23, (cb.top + cb.bottom) // 2))   # 줄 왼쪽의 체크 상자를 켠다
        time.sleep(4)
    except Exception as ex:
        print("  (메뉴 시트 선택 실패)", ex)
    shot("excel_navigator", nav[0])
    click(nav[0], "취소", "Button", wait=3)
    # [데이터 가져오기] 대화 상자 → 텍스트/CSV
    w = pbi_ui.main_window()
    btn = visible(w, "데이터 가져오기", "Button")
    big = max(b.rectangle().height() for b in btn)
    lower = [b for b in btn if b.rectangle().height() < big and b.rectangle().top > min(x.rectangle().top for x in btn)]
    lower[0].click_input()                                # 큰 단추의 아래쪽 절반(▾) = 자주 쓰는 원본 목록
    time.sleep(3)
    w = pbi_ui.main_window()
    shot("getdata_menu", w)                               # 드롭다운 목록
    click(w, "텍스트/CSV", exact=False, wait=3)
    dlg = [c for c in w.children() if c.element_info.class_name == "#32770"][0]
    open_file(dlg, DATA / "카페_판매.csv")
    time.sleep(12)
    prev = [child_window("카페_판매", exact=False)]
    shot("csv_preview", prev[0])
    click(prev[0], "취소", "Button", wait=2)


# ---------------------------------------------------------------- 1장: 테이블 보기
def cap_tableview():
    w = open_pbix("01_데이터불러오기")
    click(w, "테이블 보기", "TabItem", wait=4)
    select_table(w, "판매")
    shot("tableview_raw", w)


# ---------------------------------------------------------------- 2장: Power Query 편집기
def cap_pq():
    w = open_pbix("02_데이터정리")
    click(w, "홈", "TabItem", wait=2)                    # 리본이 다른 탭이면 [데이터 변환] 단추가 숨어 있다
    btn = [b for b in visible(w, "데이터 변환", "Button") if b.rectangle().height() < 50]
    btn[0].click_input()
    time.sleep(15)
    # Power Query 편집기는 제목이 보고서와 같은 별도 창이다 → 본 창이 아닌 PBIDesktop 창을 찾는다
    main = pbi_ui.main_window()
    pq = [d for d in top_windows() if d.handle != main.handle][0]
    pq.maximize()
    time.sleep(3)
    pq.set_focus()

    def query(name):                                      # 왼쪽 쿼리 목록(화면 왼쪽 끝의 글자)을 누른다
        [e for e in visible(pq, name, "Text") if e.rectangle().left < 150][0].click_input()
        time.sleep(6)

    def step(name):                                       # 오른쪽 적용된 단계 목록의 단계를 누른다
        visible(pq, name, "DataItem")[0].click_input()
        time.sleep(6)

    for q, name in [("판매", "pq_sales"), ("메뉴", "pq_menu"), ("매장", "pq_store")]:
        query(q)
        shot(name, pq)
    query("판매")                                         # 처음 단계(변경된 유형)를 눌러 정리 전 모습
    step("변경된 유형")
    shot("pq_sales_before", pq)
    step("제거된 열 수")
    pq.close()
    time.sleep(3)


# ---------------------------------------------------------------- 3장: 모델 보기
def cap_model():
    w = open_pbix("03_관계와날짜")
    click(w, "모델 보기", "TabItem", wait=6)
    shot("model", w)
    click(w, "테이블 보기", "TabItem", wait=4)
    select_table(w, "날짜")
    shot("tableview_date", w)


# ---------------------------------------------------------------- 4장: 측정값 수식 입력줄
def cap_measure():
    w = open_pbix("04_측정값")
    # 데이터 창에서 판매 테이블을 펼치고 [객단가] 를 선택하면 수식 입력줄에 식이 보인다
    r = w.rectangle()
    mouse.click(coords=(r.left + 900, r.top + 800))       # 캔버스 빈 곳 - 시각적 개체 선택 풀기
    time.sleep(2)
    try:
        click(w, "판매", "Button", wait=2)          # 판매 테이블 펼치기
        items = [e for e in w.descendants(control_type="TreeItem") if e.is_visible() and "객단가" in e.window_text()]
        ir = items[0].rectangle()
        mouse.click(coords=(ir.left + 75, (ir.top + ir.bottom) // 2))   # 체크 상자가 아니라 이름 글자를 누른다
        time.sleep(4)
    except Exception as ex:
        print("  (측정값 선택 실패)", ex)
    shot("measure", w)


# ---------------------------------------------------------------- 5·6장: 시각적 개체 빌드 / 서식 창
def cap_fields():
    w = open_pbix("05_차트")
    r = w.rectangle()
    mouse.click(coords=(r.left + 1000, r.top + 300))      # 매장별 매출 막대 차트 클릭 (선택)
    time.sleep(3)
    shot("fields", w)


def cap_format():
    w = open_pbix("06_슬라이서와서식")
    r = w.rectangle()
    mouse.click(coords=(r.left + 1250, r.top + 440))      # 매장별 매출 차트의 제목 옆 빈 곳 (막대를 누르면 교차 필터가 걸린다)
    time.sleep(3)
    try:
        click(w, "시각적 개체 서식 지정", exact=False, wait=4)
    except Exception as ex:
        print("  (서식 창 열기 실패)", ex)
    shot("format", w)


ALL = {"start": cap_start, "getdata": cap_getdata, "tableview": cap_tableview, "pq": cap_pq, "model": cap_model,
       "measure": cap_measure, "fields": cap_fields, "format": cap_format}

if __name__ == "__main__":
    keys = sys.argv[1:] or list(ALL)
    for k in keys:
        print("[", k, "]")
        try:
            ALL[k]()
        except Exception as ex:
            print("  실패:", ex)
    close_desktop()
