"""capture_adv.py - 심화 교안에 넣을 Power BI Desktop 조작 화면 캡처 (교안 제작용)

    python tools/capture_adv.py              # 모든 화면
    python tools/capture_adv.py pq model     # 일부만 (pq, model, measure, report, roles, perf)

장별 완성 파일(pbix/심화*.pbix)이 먼저 있어야 한다 (python tools/make_pbix.py adv).
캡처는 docs/images/ui_a_*.png 로 저장한다. Power BI 창 영역만 찍는다.
"""
import sys
import time
from pathlib import Path

import pywinauto.mouse as mouse

sys.path.insert(0, str(Path(__file__).parent))
import pbi_ui
from capture_ui import IMG, click, open_pbix, top_windows, visible
from make_pbix import close_desktop


def shot(name, win, crop=None, park=True):
    if park:                                   # 마우스를 치워 도구 설명 풍선이 안 찍히게
        r = win.rectangle()
        mouse.move(coords=(r.left + 700, r.bottom - 15))
        time.sleep(1)
    path = IMG / f"ui_a_{name}.png"
    pbi_ui.shot(str(path), win, crop=crop)
    print("  캡처:", path.name)


def open_pq(w):
    """[홈 → 데이터 변환] 으로 Power Query 편집기를 열고 최대화해서 돌려준다"""
    click(w, "홈", "TabItem", wait=2)
    [b for b in visible(w, "데이터 변환", "Button") if b.rectangle().height() < 50][0].click_input()
    time.sleep(15)
    # 편집기 창은 제목이 보고서와 같으므로 "적용된 단계" 가 있는 창을 찾는다
    pq = [d for d in top_windows() if visible(d, "적용된 단계", exact=False)][0]
    pq.maximize()
    time.sleep(3)
    pq.set_focus()
    return pq


def pq_query(pq, name):
    """Power Query 편집기 왼쪽 쿼리 목록에서 쿼리 하나를 누른다"""
    # 매개 변수는 "DataPath (C:\...)" 처럼 값이 붙어 보이므로 이름으로 시작하는지로 찾는다
    [e for e in pq.descendants(control_type="Text") if e.is_visible() and e.rectangle().left < 150
     and (e.window_text() == name or e.window_text().startswith(name + " ("))][0].click_input()
    time.sleep(6)


def goto(w, page):
    """아래쪽 페이지 탭을 눌러 그 페이지로 간다"""
    tabs = [t for t in w.descendants(control_type="TabItem") if t.window_text() == page and t.is_visible()]
    tabs[-1].click_input()
    time.sleep(5)


def child(win, title_part):
    for c in win.children():
        if c.element_info.control_type == "Window" and title_part in c.window_text():
            return c
    return None


# ---------------------------------------------------------------- 1·2장: 매개 변수, 병합, 조건 열
def cap_pq():
    w = open_pbix("심화02_병합과조건열")
    pq = open_pq(w)
    pq_query(pq, "DataPath")
    shot("param", pq)
    pq_query(pq, "주문")
    shot("pq_orders", pq)
    # 쿼리 병합 대화 상자
    try:
        click(pq, "쿼리 병합", exact=False, wait=4)
        dlg = child(pq, "병합") or pq
        shot("merge", dlg)
        click(dlg, "취소", "Button", wait=2)
    except Exception as ex:
        print("  (병합 창 실패)", ex)
    # 조건 열 추가 대화 상자
    try:
        click(pq, "열 추가", "TabItem", wait=2)
        click(pq, "조건 열", "Button", wait=4)
        dlg = child(pq, "조건 열") or pq
        shot("conditional", dlg)
        click(dlg, "취소", "Button", wait=2)
    except Exception as ex:
        print("  (조건 열 창 실패)", ex)
    pq.close()


# ---------------------------------------------------------------- 3장: 모델 보기
def cap_model():
    w = open_pbix("심화03_모델링")
    click(w, "모델 보기", "TabItem", wait=8)
    shot("model", w)


# ---------------------------------------------------------------- 4장: VAR 측정값과 수식 입력줄
def first_item(w, text, exact=False):
    """데이터 창(보고서 보기)의 트리 항목 - 같은 트리가 여러 벌 있어 가장 오른쪽 위의 보이는 것을 고른다"""
    items = [e for e in w.descendants(control_type="TreeItem") if e.is_visible()
             and (e.window_text() == text if exact else text in e.window_text())]
    return sorted(items, key=lambda e: (-e.rectangle().left, e.rectangle().top))[0]


def expand(w, text):
    """트리 항목을 누르고 → 키로 펼친다"""
    it = first_item(w, text)
    r = it.rectangle()
    mouse.click(coords=(r.left + 60, (r.top + r.bottom) // 2))
    time.sleep(1)
    it.type_keys("{RIGHT}")
    time.sleep(2)


def cap_measure():
    w = open_pbix("심화04_CALCULATE")
    r = w.rectangle()
    mouse.click(coords=(r.left + 900, r.top + 940))          # 캔버스 빈 곳 - 선택 풀기
    time.sleep(2)
    expand(w, "_측정값")
    expand(w, "2 CALCULATE")
    it = first_item(w, "지연율")
    ir = it.rectangle()
    mouse.click(coords=(ir.right - 60, (ir.top + ir.bottom) // 2))  # 체크 상자가 아니라 이름 글자 쪽을 눌러 측정값 선택
    time.sleep(3)
    shot("measure_var", w)


# ---------------------------------------------------------------- 7장: 상위 N, 도구 설명, 드릴스루
def cap_report():
    w = open_pbix("심화07_고급보고서기능")
    goto(w, "판매 개요")
    r = w.rectangle()
    bar_x, bar_y = r.left + 330, r.top + 508                 # 상위 10 막대 차트의 첫 막대
    mouse.click(coords=(r.left + 300, r.top + 425))          # 차트 제목 부근을 눌러 선택
    time.sleep(3)
    shot("topn", w)
    mouse.click(coords=(r.left + 700, r.top + 960))          # 선택 풀기
    time.sleep(2)
    for dx in range(0, 60, 6):                               # 막대 위로 천천히 움직여야 도구 설명이 뜬다
        mouse.move(coords=(bar_x - 60 + dx, bar_y))
        time.sleep(0.2)
    time.sleep(4)
    shot("tooltip", w, park=False)
    mouse.move(coords=(r.left + 700, r.top + 960))
    time.sleep(2)
    mouse.click(coords=(bar_x, bar_y), button="right")       # 오른쪽 클릭 → 드릴스루 (하위 메뉴는 마우스를 올리면 열린다)
    time.sleep(3)
    try:
        dt = [e for e in w.descendants(control_type="MenuItem") if e.is_visible() and "드릴스루" in e.window_text()]
        dr = dt[0].rectangle()
        mouse.move(coords=((dr.left + dr.right) // 2, (dr.top + dr.bottom) // 2))
        time.sleep(3)
    except Exception as ex:
        print("  (드릴스루 메뉴 실패)", ex)
    shot("drillthrough", w, park=False)
    w.type_keys("{ESC}")


# ---------------------------------------------------------------- 8장: 역할, 역할로 보기, 성능 분석기
def cap_roles():
    w = open_pbix("심화08_행수준보안")
    click(w, "모델링", "TabItem", wait=2)
    click(w, "역할 관리", exact=False, wait=8)
    try:                                                      # 역할 하나와 테이블 '주' 를 골라 규칙(DAX 식)이 보이게
        mw = pbi_ui.main_window()
        click(mw, "남동부 담당", exact=False, wait=2)
        # 테이블 목록의 '주' 는 이름으로 잡히지 않아 대화 상자 제목 위치에서 상대 좌표로 누른다
        wr = mw.rectangle()
        mouse.click(coords=(wr.left + 718, wr.top + 568))   # 테이블 목록의 '주' (창 기준 위치)
        time.sleep(3)
    except Exception as ex:
        print("  (역할 고르기 실패)", ex)
    shot("roles", pbi_ui.main_window())
    for name in ("닫기", "취소"):
        try:
            click(pbi_ui.main_window(), name, "Button", wait=3)
            break
        except Exception:
            pass
    w = pbi_ui.main_window()
    click(w, "모델링", "TabItem", wait=2)
    try:
        click(w, "역할로 보기", exact=False, wait=5)
    except Exception:
        click(w, "표시 방법", exact=False, wait=5)          # 2.159 의 리본 단추 이름은 "표시 방법:"
    try:
        cb = visible(pbi_ui.main_window(), "남동부 담당", exact=False)
        cb[0].click_input()
        time.sleep(1)
        click(pbi_ui.main_window(), "확인", "Button", wait=10)
    except Exception as ex:
        print("  (역할로 보기 실패)", ex)
    shot("viewas", pbi_ui.main_window())


def cap_perf():
    w = open_pbix("심화07_고급보고서기능")
    goto(w, "판매 개요")
    click(w, "보기", "TabItem", wait=2)
    click(w, "성능 분석기", exact=False, wait=4)
    click(w, "기록 시작", exact=False, wait=2)
    click(w, "시각적 개체 새로 고침", exact=False, wait=15)
    shot("perf", w)


ALL = {"pq": cap_pq, "model": cap_model, "measure": cap_measure, "report": cap_report, "roles": cap_roles,
       "perf": cap_perf}

if __name__ == "__main__":
    keys = sys.argv[1:] or list(ALL)
    for k in keys:
        print("[", k, "]")
        try:
            ALL[k]()
        except Exception as ex:
            print("  실패:", ex)
    close_desktop()
