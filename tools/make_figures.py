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


# ================================================================== 심화 과정 그림
# ------------------------------------------------------------------ 심화 0장: Olist 파일과 연결 열
def fig_olist_tables():
    fig, ax = canvas(10, 5.2)
    title(ax, "Olist 데이터 - 파일 7개와 연결 열")
    box(ax, 38, 22, 24, 13, "orders (주문)\norder_id · customer_id\n주문상태 · 날짜 5개", "navy", 9.5, True)
    box(ax, 2, 36, 26, 11, "customers (고객)\ncustomer_id\ncustomer_unique_id · state", "grey", 9)
    box(ax, 72, 36, 26, 11, "order_reviews (리뷰)\nreview_id · order_id\nreview_score", "orange", 9)
    box(ax, 72, 20, 26, 11, "order_payments (결제)\norder_id · payment_type\npayment_value", "orange", 9)
    box(ax, 38, 3, 24, 11, "order_items (주문상품)\norder_id · product_id\nseller_id · price", "orange", 9, True)
    box(ax, 2, 3, 26, 11, "products (상품)\nproduct_id\ncategory_name (포르투갈어)", "green", 9)
    box(ax, 72, 3, 26, 11, "sellers (판매자)\nseller_id\nseller_city · seller_state", "green", 9)
    arrow(ax, 28, 41, 38, 32, "customer_id", GREY, size=8)
    arrow(ax, 72, 41, 62, 32, "order_id", ORANGE, size=8)
    arrow(ax, 72, 25.5, 62, 28, "order_id", ORANGE, size=8, ly=1.5)
    arrow(ax, 50, 14, 50, 22, "order_id", ORANGE, size=8)
    arrow(ax, 38, 8.5, 28, 8.5, "product_id", GREEN, size=8, ly=2)
    arrow(ax, 62, 8.5, 72, 8.5, "seller_id", GREEN, size=8, ly=2)
    save(fig, "fig_olist_tables.png")


# ------------------------------------------------------------------ 심화 3장: Olist 모델
def fig_olist_model():
    fig, ax = canvas(10, 5.6)
    title(ax, "Olist 모델 - 필터는 화살표 방향(일 → 다)으로 흐른다")
    box(ax, 4, 38, 18, 9, "날짜\n(Date)", "green", 10, True)
    box(ax, 4, 22, 18, 9, "주\n(주코드·지역)", "green", 10, True)
    box(ax, 38, 28, 24, 12, "주문\n주문ID · 고객ID · 고객주\n구매일 · 배송완료일 · 배송상태", "navy", 9.5, True)
    box(ax, 76, 40, 20, 9, "리뷰\n평점", "orange", 9.5, True)
    box(ax, 76, 26, 20, 9, "결제\n결제금액 · 결제수단", "orange", 9.5, True)
    box(ax, 38, 6, 24, 11, "주문상품\n가격 · 배송비", "orange", 9.5, True)
    box(ax, 4, 4, 18, 9, "상품\n카테고리 · 대분류", "green", 9.5, True)
    box(ax, 78, 4, 18, 9, "판매자\n판매자주", "green", 9.5, True)
    arrow(ax, 22, 43, 38, 37, "1 → *  구매일 (활성)", GREEN, size=8, ly=2.5)
    ax.add_patch(FancyArrowPatch((22, 40), (38, 34), arrowstyle="-|>", mutation_scale=13, color=GREY, lw=1.3,
                                 linestyle="--", clip_on=False))
    ax.text(13, 34.2, "점선: 배송완료일 (비활성)", ha="center", fontsize=7.6, color=GREY)
    arrow(ax, 22, 26.5, 38, 30, "1 → *", GREEN, size=8, ly=-2)
    arrow(ax, 62, 37, 76, 44, "1 → *", NAVY, size=8)
    arrow(ax, 62, 33, 76, 30.5, "1 → *", NAVY, size=8)
    arrow(ax, 50, 28, 50, 17, "1 → *", NAVY, size=8, lx=4)
    arrow(ax, 22, 8.5, 38, 10, "1 → *", GREEN, size=8, ly=2)
    arrow(ax, 78, 8.5, 62, 10, "1 → *", GREEN, size=8, ly=2)
    save(fig, "fig_olist_model.png")


# ------------------------------------------------------------------ 심화 3장: 필터 방향 실험
def fig_filter_direction():
    fig, ax = canvas(10, 3.6)
    title(ax, "행렬의 '가구·인테리어' 행 - 필터가 어디까지 흐를까?")
    box(ax, 2, 10, 18, 12, "상품\n대분류 = 가구·인테리어", "green", 9.5, True)
    box(ax, 38, 10, 22, 12, "주문상품\n→ 상품매출 2,239,686\n→ 상품주문수 18,221", "orange", 9.2, True)
    box(ax, 76, 10, 22, 12, "주문\n걸러지지 않음\n→ 주문수 99,441 (전체)", "red", 9.2, True)
    arrow(ax, 21, 16, 38, 16, "1 → * 흐른다 (O)", GREEN, size=8.5, ly=3)
    ax.add_patch(FancyArrowPatch((60, 16), (76, 16), arrowstyle="-|>", mutation_scale=13, color=RED, lw=1.5,
                                 linestyle="--", clip_on=False))
    ax.text(68, 19.5, "* → 1 안 흐른다 (X)", ha="center", fontsize=8.5, color=RED)
    ax.text(50, 4, "해결: 주문상품에서 세는 측정값(상품주문수) 또는 CROSSFILTER(…, Both) 로 그 측정값에서만 방향을 바꾼다",
            ha="center", fontsize=8.8, color="#1f2937")
    save(fig, "fig_filter_direction.png")


# ------------------------------------------------------------------ 심화 4장: CALCULATE
def fig_calculate():
    fig, ax = canvas(10, 3.6)
    title(ax, "CALCULATE([주문수], '주문'[주문상태] = \"배송완료\")  -  '남동부' 행에서")
    box(ax, 2, 8, 24, 15, "① 원래 필터\n(행렬의 행)\n\n지역 = 남동부", "grey", 9.5)
    box(ax, 36, 8, 26, 15, "② 조건을 바꾼 필터\n\n지역 = 남동부\n그리고 주문상태 = 배송완료", "navy", 9.5, True)
    box(ax, 72, 8, 26, 15, "③ 바뀐 필터로 계산\n\n[주문수]\n= 남동부의 배송완료 주문 수", "green", 9.5)
    arrow(ax, 26, 15.5, 36, 15.5, "조건 더하기", NAVY, size=8.5, ly=3)
    arrow(ax, 62, 15.5, 72, 15.5, "계산", NAVY, size=8.5, ly=3)
    ax.text(50, 2.5, "REMOVEFILTERS 는 ②에서 필터를 없애고, USERELATIONSHIP·CROSSFILTER 는 ②에서 관계를 바꾼다",
            ha="center", fontsize=8.8, color="#1f2937")
    save(fig, "fig_calculate.png")


# ------------------------------------------------------------------ 심화 5장: 시간 인텔리전스
def fig_time_intel():
    import matplotlib.patches as mp
    fig, ax = canvas(10, 4.2)
    title(ax, "2018년 3월 행에서 각 함수가 고르는 기간")
    months = ["17-01", "17-02", "17-03", "17-04", "…", "17-12", "18-01", "18-02", "18-03"]
    xs = [16 + i * 9 for i in range(len(months))]
    for x, m in zip(xs, months):
        ax.text(x + 4, 30.5, m, ha="center", fontsize=8, color=GREY)
    rows = [("TOTALYTD", [6, 7, 8], NAVY), ("SAMEPERIODLASTYEAR", [2], ORANGE),
            ("DATEADD -1 MONTH", [7], GREEN), ("DATESINPERIOD -3 MONTH", [6, 7, 8], RED)]
    for r, (name, idx, colr) in enumerate(rows):
        y = 23 - r * 6.2
        ax.text(14, y + 2, name, ha="right", va="center", fontsize=8.6, color="#1f2937")
        for i, x in enumerate(xs):
            on = i in idx
            ax.add_patch(mp.Rectangle((x, y), 8, 4, fc=colr if on else "#eef0f3", ec="white", alpha=0.85 if on else 1))
    ax.add_patch(mp.Rectangle((xs[8] - 0.4, 2.5), 8.8, 26, fill=False, ec=NAVY, lw=1.2, ls="--"))
    ax.text(xs[8] + 4, 0.6, "지금 행", ha="center", fontsize=8, color=NAVY)
    save(fig, "fig_time_intel.png")


if __name__ == "__main__":
    fig_olist_tables()
    fig_olist_model()
    fig_filter_direction()
    fig_calculate()
    fig_time_intel()
    fig_bi_flow()
    fig_query_steps()
    fig_star()
    fig_one_many()
    fig_column_measure()
    fig_filter_context()
    fig_layout()
