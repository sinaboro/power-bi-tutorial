"""교안에 들어가는 개념 그림을 만든다 (docs/images/fig_*.png).

    python tools/make_figures.py

matplotlib 으로 상자와 화살표를 그린다. 글꼴은 맑은 고딕(Windows 기본).
Power BI 화면 캡처(pbi_*.png)는 이 파일이 아니라 tools/make_pbix.py, tools/capture_ui.py 가 만든다.
"""
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch

plt.rcParams["font.family"] = "Malgun Gothic"
plt.rcParams["axes.unicode_minus"] = False

OUT = Path(__file__).resolve().parent.parent / "docs" / "images"
OUT.mkdir(parents=True, exist_ok=True)

NAVY, BROWN, GREEN, ORANGE, RED, GREY, GOLD = "#1f3a5f", "#6F4E37", "#2e7d32", "#e8890c", "#c0392b", "#6b7280", "#B8860B"
FILL = {"navy": "#e3ebf5", "brown": "#f3e9e2", "green": "#e6f4ea", "orange": "#fff1dc", "red": "#fdecea",
        "grey": "#f1f3f5", "gold": "#fff8d6"}
EDGE = {"navy": NAVY, "brown": BROWN, "green": GREEN, "orange": ORANGE, "red": RED, "grey": GREY, "gold": GOLD}


def canvas(w, h):
    fig, ax = plt.subplots(figsize=(w, h), dpi=150)
    ax.set_xlim(0, w * 10)
    ax.set_ylim(0, h * 10)
    ax.axis("off")
    return fig, ax


def box(ax, x, y, w, h, text, color="navy", size=10, bold=False, align="center"):
    """(x, y) 는 왼쪽 아래. text 는 여러 줄 가능."""
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.3,rounding_size=1.2",
                                fc=FILL[color], ec=EDGE[color], lw=1.4, clip_on=False))
    tx = x + w / 2 if align == "center" else x + 1.2
    ax.text(tx, y + h / 2, text, ha=align, va="center", fontsize=size,
            fontweight="bold" if bold else "normal", color="#1f2937", linespacing=1.5)


def arrow(ax, x1, y1, x2, y2, label="", color=NAVY, rad=0.0, size=8.5, lx=0, ly=0):
    ax.add_patch(FancyArrowPatch((x1, y1), (x2, y2), arrowstyle="-|>", mutation_scale=13, color=color, lw=1.5,
                                 connectionstyle=f"arc3,rad={rad}", clip_on=False))
    if label:
        ax.text((x1 + x2) / 2 + lx, (y1 + y2) / 2 + ly, label, ha="center", va="center", fontsize=size,
                color=color, bbox=dict(fc="white", ec="none", pad=1.5))


def title(ax, text):
    xl, yl = ax.get_xlim(), ax.get_ylim()
    ax.text(xl[1] / 2, yl[1] - 2.5, text, ha="center", va="top", fontsize=11.5, fontweight="bold", color=NAVY)


def save(fig, name):
    fig.savefig(OUT / name, bbox_inches="tight", pad_inches=0.12, facecolor="white")
    plt.close(fig)
    print("그림:", name)


# ------------------------------------------------------------------ 1장: Power BI 작업 흐름
def fig_bi_flow():
    fig, ax = canvas(10, 3.4)
    title(ax, "Power BI 로 보고서를 만드는 순서")
    steps = [("① 데이터 원본\nCSV · 엑셀\nDB · 웹", "grey"), ("② Power Query\n가져오기\n정리하기 (2장)", "green"),
             ("③ 데이터 모델\n관계 (3장)\n측정값 (4장)", "gold"), ("④ 보고서\n차트 (5장)\n슬라이서 (6장)", "brown"),
             ("⑤ 공유\nPDF · 서비스\n(7장)", "navy")]
    for i, (t, c) in enumerate(steps):
        box(ax, 1 + i * 20, 6, 16, 16, t, c, 9.3)
        if i:
            arrow(ax, 1 + (i - 1) * 20 + 16.4, 14, 1 + i * 20 - 0.4, 14)
    ax.text(50, 2, "Power BI Desktop 한 프로그램 안에서 ②~④ 를 모두 한다. 결과는 .pbix 파일 하나에 저장된다.",
            ha="center", fontsize=8.8, color=GREY)
    save(fig, "fig_bi_flow.png")


# ------------------------------------------------------------------ 2장: 적용된 단계
def fig_query_steps():
    fig, ax = canvas(10, 3.6)
    title(ax, "Power Query 는 정리 작업을 \"단계\" 로 기록한다 (판매 쿼리)")
    steps = ["원본", "승격된\n헤더", "변경된\n유형", "제거된\n중복 항목 수", "필터링된\n행", "바꾼 값", "삽입된\n날짜·시간", "추가된\n사용자 지정", "제거된\n열 수"]
    for i, t in enumerate(steps):
        x = 1 + i * 11
        box(ax, x, 12, 9, 10, t, "green" if i else "grey", 8)
        if i:
            arrow(ax, x - 1.6, 17, x - 0.4, 17)
    ax.text(50, 6, "새로 고침할 때마다 이 단계들이 처음부터 다시 실행된다 → 원본 파일이 바뀌어도 같은 정리가 자동으로 된다",
            ha="center", fontsize=8.8, color=GREEN)
    ax.text(50, 2, "원본 CSV · 엑셀 파일은 절대 바뀌지 않는다. 정리된 결과만 Power BI 모델에 들어간다.", ha="center",
            fontsize=8.8, color=GREY)
    save(fig, "fig_query_steps.png")


# ------------------------------------------------------------------ 3장: 별 모양 모델
def fig_star():
    fig, ax = canvas(9, 5)
    title(ax, "별 모양(스타) 스키마 - 가운데 사실 테이블, 바깥에 차원 테이블")
    box(ax, 33, 18, 24, 14, "판매  (사실 테이블)\n주문번호 · 수량 · 매출액\n매장코드 · 메뉴코드 · 주문일자", "brown", 9, True)
    box(ax, 2, 33, 22, 11, "메뉴  (차원)\n메뉴코드 · 메뉴명 · 분류", "gold", 9)
    box(ax, 66, 33, 22, 11, "매장  (차원)\n매장코드 · 매장명 · 지역", "gold", 9)
    box(ax, 34, 1, 22, 11, "날짜  (차원)\nDate · 연도 · 월 · 요일", "gold", 9)
    arrow(ax, 24.5, 38, 32.5, 29, "1 → *", BROWN, size=8.5, ly=2)
    arrow(ax, 65.5, 38, 57.5, 29, "1 → *", BROWN, size=8.5, ly=2)
    arrow(ax, 45, 12.5, 45, 17.5, "1 → *", BROWN, size=8.5, lx=5)
    ax.text(80, 8, "차원 = \"무엇으로 나눠 볼까\"\n(메뉴별, 매장별, 월별)\n사실 = \"무엇을 셀까\"\n(매출, 수량, 주문)", ha="center",
            fontsize=8.5, color=GREY, va="center")
    save(fig, "fig_star.png")


# ------------------------------------------------------------------ 3장: 1 대 다 관계
def fig_one_many():
    fig, ax = canvas(9.5, 3.8)
    title(ax, "1 : 다(*) 관계 - 메뉴코드가 두 테이블을 잇는다")
    box(ax, 2, 4, 30, 24, "메뉴 (일 쪽)\n───────────────\nM01  아메리카노  커피\nM02  카페라떼    커피\nM16  치즈케이크  디저트\n\n메뉴코드가 한 번씩만 나온다", "gold", 8.8)
    box(ax, 50, 4, 42, 24, "판매 (다 쪽) - 2023-01-01 실제 데이터 일부\n─────────────────────────\nS01-20230101-0023  M02  수량 1  단가 4,700\nS02-20230101-0011  M02  수량 1  단가 4,700\nS02-20230101-0039  M01  수량 1  단가 4,200\nS02-20230101-0094  M16  수량 1  단가 6,500\n같은 메뉴코드가 여러 번 나온다", "brown", 8.8)
    arrow(ax, 32.5, 16, 49.5, 16, "메뉴코드 = 메뉴코드", BROWN, ly=2.5)
    ax.text(41, 10, "1", fontsize=12, fontweight="bold", color=ORANGE)
    ax.text(47, 10, "*", fontsize=14, fontweight="bold", color=ORANGE)
    save(fig, "fig_one_many.png")


# ------------------------------------------------------------------ 4장: 계산 열 vs 측정값
def fig_column_measure():
    fig, ax = canvas(9.5, 3.8)
    title(ax, "계산 열과 측정값의 차이")
    box(ax, 2, 3, 43, 26, "계산 열 (또는 Power Query 사용자 지정 열)\n\n행마다 하나씩 계산해서 표에 저장\n매출액 = 수량 × 단가\n\n→ 행 36만 개에 값이 하나씩 생긴다\n→ 슬라이서·축으로 쓸 수 있다", "green", 8.8)
    box(ax, 50, 3, 43, 26, "측정값\n\n보고서에서 보는 순간 계산 (저장 X)\n총매출 = SUM('판매'[매출액])\n\n→ 매장별·월별로 놓으면 그 조건으로\n   다시 계산된다 (필터 컨텍스트)", "gold", 8.8)
    save(fig, "fig_column_measure.png")


# ------------------------------------------------------------------ 4장: 필터 컨텍스트
def fig_filter_context():
    fig, ax = canvas(10, 4.2)
    title(ax, "같은 측정값 [총매출] 이 행마다 다른 값이 되는 이유")
    box(ax, 1, 4, 26, 28, "행렬 시각적 개체\n─────────────\n매장명     총매출\n강남역점   865,535,200\n판교점     430,592,300\n홍대점     820,379,200", "brown", 8.8)
    box(ax, 37, 22, 28, 10, "① \"강남역점\" 행을 그릴 때\n매장명 = 강남역점 필터가 걸린다", "navy", 8.6)
    box(ax, 37, 6, 28, 10, "② 관계를 따라 판매 표에도\n강남역점 행만 남는다", "navy", 8.6)
    box(ax, 74, 6, 25, 26, "③ 남은 행으로\nSUM('판매'[매출액])\n\n= 865,535,200\n\n행마다 ①~③ 반복", "gold", 8.8)
    arrow(ax, 27.5, 24, 36.5, 27)
    arrow(ax, 51, 21.5, 51, 16.5)
    arrow(ax, 65.5, 11, 73.5, 15)
    save(fig, "fig_filter_context.png")


# ------------------------------------------------------------------ 7장: 대시보드 배치
def fig_layout():
    fig, ax = canvas(9, 5.6)
    title(ax, "한눈에 읽히는 대시보드 배치 (Z 순서로 읽는다)")
    box(ax, 2, 41, 52, 6, "제목 - 무엇에 대한 보고서인가", "grey", 8.6)
    box(ax, 58, 41, 30, 6, "슬라이서 (연도 · 매장)", "navy", 8.6)
    for i, t in enumerate(["총매출", "주문수", "객단가", "판매수량"]):
        box(ax, 2 + i * 22, 31, 20, 7, f"카드\n{t}", "gold", 8.4)
    box(ax, 2, 15, 52, 13, "추세 (꺾은선)\n시간에 따라 어떻게 변했나", "brown", 8.6)
    box(ax, 58, 15, 30, 13, "비교 (막대)\n어디가 큰가", "brown", 8.6)
    box(ax, 2, 1, 30, 11, "구성 (도넛)\n무엇이 차지하나", "green", 8.6)
    box(ax, 36, 1, 52, 11, "자세한 숫자 (행렬)", "green", 8.6)
    save(fig, "fig_layout.png")


if __name__ == "__main__":
    fig_bi_flow()
    fig_query_steps()
    fig_star()
    fig_one_many()
    fig_column_measure()
    fig_filter_context()
    fig_layout()
