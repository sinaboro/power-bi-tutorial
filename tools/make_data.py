"""make_data.py - 실습 데이터 만들기 (가상의 카페 '하루커피' 3년치 + 편의점 과제 1년치)

    python tools/make_data.py          # 전부
    python tools/make_data.py cvs      # 편의점(과제) 데이터만

만드는 파일 (data 폴더)
  카페_판매.csv     2023-01-01 ~ 2025-12-31 판매 내역 (주문 상세 한 줄 = 메뉴 하나)
  카페_메뉴.xlsx    메뉴 22개 (메뉴코드, 메뉴명, 분류, 판매가, 원가)
  카페_매장.xlsx    매장 3곳 (매장코드, 매장명, 지역, 오픈일, 좌석수)
  편의점_판매.csv / 편의점_상품.xlsx / 편의점_점포.xlsx   (과제용, 2025년 1년치)

데이터에 일부러 넣은 "지저분한 값" (2장 Power Query 실습용)
  - 판매: 똑같은 줄이 두 번 들어간 중복 행, 수량이 빈칸인 행
  - 판매: 결제수단 표기가 섞여 있음 ("카드" / "신용카드" / "CARD")
  - 메뉴: 메뉴명 앞뒤에 빈칸, 같은 메뉴가 두 번 적힌 행

숫자는 난수로 만들지만 seed 를 고정했기 때문에 몇 번을 실행해도 같은 파일이 나온다.
"""
from datetime import date, datetime, timedelta
from pathlib import Path

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "data"
OUT.mkdir(exist_ok=True)
rng = np.random.default_rng(2026)  # 고정 seed → 항상 같은 데이터

# ====================================================================== 카페 - 매장
STORES = pd.DataFrame([
    # 매장코드, 매장명, 지역, 오픈일, 좌석수
    ("S01", "강남역점", "서울", date(2021, 5, 1), 48),
    ("S02", "홍대점", "서울", date(2022, 3, 12), 40),
    ("S03", "판교점", "경기", date(2024, 3, 15), 56),
], columns=["매장코드", "매장명", "지역", "오픈일", "좌석수"])

# 매장별 하루 평균 주문 수 (평일, 주말) - 상권 특징: 강남·판교는 직장인(평일), 홍대는 주말
STORE_BASE = {"S01": (95, 55), "S02": (70, 105), "S03": (85, 35)}
# 매장별 시간대 분포 (8시~21시, 14칸)
HOURS = np.arange(8, 22)
HOUR_W = {
    "S01": [9, 12, 8, 6, 10, 11, 7, 6, 5, 5, 4, 3, 2, 1],   # 출근·점심
    "S02": [2, 3, 4, 6, 9, 10, 11, 11, 10, 9, 8, 7, 6, 4],  # 오후·저녁
    "S03": [8, 11, 9, 7, 11, 12, 8, 7, 6, 5, 3, 2, 1, 1],   # 출근·점심
}

# ====================================================================== 카페 - 메뉴
MENU = pd.DataFrame([
    # 메뉴코드, 메뉴명, 분류, 판매가(2025년 현재), 원가, 인기도, 계절성(여름+, 겨울-)
    ("M01", "아메리카노", "커피", 4500, 900, 30, 0.2),
    ("M02", "카페라떼", "커피", 5000, 1300, 16, 0.0),
    ("M03", "바닐라라떼", "커피", 5500, 1500, 9, 0.0),
    ("M04", "카라멜마키아토", "커피", 5800, 1600, 6, -0.1),
    ("M05", "콜드브루", "커피", 5000, 1100, 7, 0.8),
    ("M06", "카푸치노", "커피", 5000, 1300, 4, -0.3),
    ("M07", "에스프레소", "커피", 4000, 800, 2, 0.0),
    ("M08", "녹차라떼", "논커피", 5500, 1600, 4, 0.0),
    ("M09", "초코라떼", "논커피", 5500, 1500, 4, -0.6),
    ("M10", "딸기라떼", "논커피", 6000, 2000, 3, -0.4),
    ("M11", "자몽에이드", "논커피", 6000, 1800, 4, 1.0),
    ("M12", "레몬에이드", "논커피", 5800, 1700, 3, 1.0),
    ("M13", "캐모마일", "티", 4800, 700, 2, -0.5),
    ("M14", "얼그레이", "티", 4800, 700, 2, -0.3),
    ("M15", "유자차", "티", 5500, 1200, 3, -1.0),
    ("M16", "치즈케이크", "디저트", 6500, 2600, 4, 0.0),
    ("M17", "티라미수", "디저트", 6800, 2800, 3, 0.0),
    ("M18", "마카롱", "디저트", 3000, 1100, 4, 0.0),
    ("M19", "크루아상", "베이커리", 4200, 1500, 4, 0.0),
    ("M20", "소금빵", "베이커리", 3800, 1200, 5, 0.0),
    ("M21", "베이글", "베이커리", 4000, 1300, 3, 0.0),
    ("M22", "샌드위치", "베이커리", 7500, 3200, 4, 0.0),
], columns=["메뉴코드", "메뉴명", "분류", "판매가", "원가", "인기도", "계절성"])


def price_on(menu_row, day):
    """판매 당시의 단가. 커피는 2025-01-01 부터 300원 올랐다 (그 전에는 300원 쌌다)."""
    if menu_row["분류"] == "커피" and day < date(2025, 1, 1):
        return int(menu_row["판매가"] - 300)
    return int(menu_row["판매가"])


def season(day):
    """여름(7월)에 +1, 겨울(1월)에 -1 이 되는 계절 값"""
    return -np.cos((day.timetuple().tm_yday - 15) / 365 * 2 * np.pi)


def make_cafe():
    start, end = date(2023, 1, 1), date(2025, 12, 31)
    menu_codes = MENU["메뉴코드"].to_numpy()
    rows = []
    order_line = 0
    day = start
    while day <= end:
        weekday = day.weekday() < 5
        years = (day - start).days / 365
        growth = 1.06 ** years                               # 해마다 약 6% 성장
        s = season(day)
        # 메뉴 선택 확률: 인기도 × (1 + 계절성 × 계절 값)
        weight = MENU["인기도"].to_numpy() * np.clip(1 + MENU["계절성"].to_numpy() * s * 0.6, 0.15, None)
        prob = weight / weight.sum()
        for _, st in STORES.iterrows():
            code = st["매장코드"]
            if day < max(st["오픈일"], start):
                continue
            wk, we = STORE_BASE[code]
            base = (wk if weekday else we) * growth
            if code == "S03":                                 # 새 매장은 처음 몇 달 동안 손님이 늘어난다
                months = (day - st["오픈일"]).days / 30
                base *= min(1.0, 0.55 + months * 0.08)
            if day.month == 12 and day.day >= 20:            # 연말 특수
                base *= 1.15
            n_orders = max(5, int(rng.normal(base, base * 0.12)))
            hours = rng.choice(HOURS, size=n_orders, p=np.array(HOUR_W[code]) / sum(HOUR_W[code]))
            for k in range(n_orders):
                t = datetime(day.year, day.month, day.day, int(hours[k]), int(rng.integers(0, 60)), int(rng.integers(0, 60)))
                order_no = f"{code}-{day:%Y%m%d}-{k + 1:04d}"
                pay = rng.choice(["카드", "간편결제", "현금"], p=[0.62, 0.30, 0.08] if day.year > 2023 else [0.68, 0.22, 0.10])
                kind = rng.choice(["매장", "포장"], p=[0.55, 0.45] if code != "S02" else [0.7, 0.3])
                n_items = rng.choice([1, 2, 3], p=[0.62, 0.28, 0.10])
                for m in rng.choice(menu_codes, size=n_items, replace=False, p=prob):
                    mrow = MENU.loc[MENU["메뉴코드"] == m].iloc[0]
                    qty = int(rng.choice([1, 2, 3], p=[0.85, 0.12, 0.03]))
                    order_line += 1
                    rows.append((order_line, order_no, t, code, m, qty, price_on(mrow, day), pay, kind))
        day += timedelta(days=1)

    sales = pd.DataFrame(rows, columns=["주문상세번호", "주문번호", "주문일시", "매장코드", "메뉴코드", "수량", "단가", "결제수단", "주문유형"])
    sales = sales.sort_values(["주문일시", "주문상세번호"]).reset_index(drop=True)
    sales["주문일시"] = sales["주문일시"].dt.strftime("%Y-%m-%d %H:%M:%S")

    # ---- 지저분한 값 넣기 (2장 실습용) ----
    card = sales.index[sales["결제수단"] == "카드"]
    sales.loc[rng.choice(card, size=int(len(card) * 0.04), replace=False), "결제수단"] = "신용카드"
    card = sales.index[sales["결제수단"] == "카드"]
    sales.loc[rng.choice(card, size=int(len(card) * 0.01), replace=False), "결제수단"] = "CARD"
    sales["수량"] = sales["수량"].astype("Int64")
    sales.loc[rng.choice(sales.index, size=30, replace=False), "수량"] = pd.NA       # 수량 빈칸 30줄
    dup = sales.loc[rng.choice(sales.index, size=40, replace=False)]                # 중복 40줄
    sales = pd.concat([sales, dup]).sort_values(["주문일시", "주문상세번호"]).reset_index(drop=True)
    # utf-8-sig : 맨 앞에 BOM 을 붙여 Power BI·엑셀이 한글을 UTF-8 로 바로 알아보게 한다
    sales.to_csv(OUT / "카페_판매.csv", index=False, encoding="utf-8-sig")

    menu_out = MENU[["메뉴코드", "메뉴명", "분류", "판매가", "원가"]].copy()
    menu_out.loc[menu_out["메뉴코드"] == "M03", "메뉴명"] = " 바닐라라떼"        # 앞에 빈칸
    menu_out.loc[menu_out["메뉴코드"] == "M16", "메뉴명"] = "치즈케이크  "       # 뒤에 빈칸
    menu_out.loc[menu_out["메뉴코드"] == "M20", "메뉴명"] = "소금빵 "
    menu_out = pd.concat([menu_out, menu_out[menu_out["메뉴코드"] == "M05"]]).sort_values("메뉴코드")  # 중복 행
    with pd.ExcelWriter(OUT / "카페_메뉴.xlsx") as w:
        menu_out.to_excel(w, sheet_name="메뉴", index=False)
    with pd.ExcelWriter(OUT / "카페_매장.xlsx") as w:
        STORES.to_excel(w, sheet_name="매장", index=False)
    return sales


# ====================================================================== 편의점 (과제)
CVS_STORES = pd.DataFrame([
    ("C01", "역삼중앙점", "오피스"), ("C02", "신촌대학점", "대학가"), ("C03", "일산주택점", "주거지"),
], columns=["점포코드", "점포명", "상권"])
CVS_ITEMS = pd.DataFrame([
    # 상품코드, 상품명, 카테고리, 가격, 인기도, 계절성
    ("P01", "생수 500ml", "음료", 1000, 12, 0.8), ("P02", "콜라 500ml", "음료", 2200, 8, 0.6),
    ("P03", "커피우유", "음료", 1700, 7, 0.0), ("P04", "에너지드링크", "음료", 2500, 5, 0.2),
    ("P05", "아이스커피컵", "음료", 2000, 6, 1.2), ("P06", "새우깡", "과자", 1600, 5, 0.0),
    ("P07", "초코파이", "과자", 4800, 3, 0.0), ("P08", "감자칩", "과자", 1800, 5, 0.0),
    ("P09", "참치마요 삼각김밥", "간편식사", 1300, 10, 0.0), ("P10", "제육 도시락", "간편식사", 5500, 6, 0.0),
    ("P11", "샌드위치", "간편식사", 3200, 5, 0.0), ("P12", "컵라면 큰사발", "라면", 1600, 8, -0.6),
    ("P13", "봉지라면", "라면", 1100, 3, -0.4), ("P14", "아이스크림 바", "아이스크림", 1500, 6, 1.5),
    ("P15", "아이스크림 콘", "아이스크림", 2200, 4, 1.5), ("P16", "핫팩", "생활용품", 1200, 3, -2.5),
    ("P17", "물티슈", "생활용품", 1500, 3, 0.0), ("P18", "우산", "생활용품", 7000, 1, 0.3),
], columns=["상품코드", "상품명", "카테고리", "가격", "인기도", "계절성"])
CVS_BASE = {"C01": (260, 90), "C02": (230, 170), "C03": (180, 220)}
# 상권별 시간대 분포 (0시~23시, 24칸): 오피스는 출근·점심·퇴근, 대학가는 오후·밤늦게, 주거지는 저녁
CVS_HOUR_W = {
    "C01": [1, 1, 0, 0, 0, 1, 3, 9, 14, 8, 5, 9, 14, 8, 5, 4, 5, 8, 10, 6, 3, 2, 1, 1],
    "C02": [5, 4, 3, 1, 1, 1, 1, 2, 4, 5, 5, 7, 9, 8, 8, 8, 8, 9, 10, 10, 10, 9, 8, 7],
    "C03": [2, 1, 1, 0, 0, 1, 2, 4, 5, 4, 4, 5, 6, 5, 5, 6, 7, 10, 12, 12, 10, 8, 5, 3],
}


def make_cvs():
    rng = np.random.default_rng(2027)  # 편의점은 따로 seed → 카페 데이터와 상관없이 다시 만들 수 있다
    rows = []
    no = 0
    day = date(2025, 1, 1)
    while day <= date(2025, 12, 31):
        weekday = day.weekday() < 5
        s = season(day)
        w = CVS_ITEMS["인기도"].to_numpy() * np.clip(1 + CVS_ITEMS["계절성"].to_numpy() * s * 0.5, 0.05, None)
        p = w / w.sum()
        for code in CVS_STORES["점포코드"]:
            base = CVS_BASE[code][0 if weekday else 1]
            n = max(10, int(rng.normal(base, base * 0.1)))
            items = rng.choice(CVS_ITEMS["상품코드"], size=n, p=p)
            hw = np.array(CVS_HOUR_W[code], dtype=float)
            hours = rng.choice(np.arange(24), size=n, p=hw / hw.sum())
            for k in range(n):
                no += 1
                rows.append((no, day.isoformat(), int(hours[k]), code, items[k], int(rng.choice([1, 2], p=[0.8, 0.2]))))
        day += timedelta(days=1)
    sales = pd.DataFrame(rows, columns=["판매번호", "판매일자", "판매시", "점포코드", "상품코드", "수량"])
    price = dict(zip(CVS_ITEMS["상품코드"], CVS_ITEMS["가격"]))
    sales["판매금액"] = sales["상품코드"].map(price) * sales["수량"]
    sales.to_csv(OUT / "편의점_판매.csv", index=False, encoding="utf-8-sig")
    with pd.ExcelWriter(OUT / "편의점_상품.xlsx") as w:
        CVS_ITEMS[["상품코드", "상품명", "카테고리", "가격"]].to_excel(w, sheet_name="상품", index=False)
    with pd.ExcelWriter(OUT / "편의점_점포.xlsx") as w:
        CVS_STORES.to_excel(w, sheet_name="점포", index=False)
    return sales


if __name__ == "__main__":
    import sys
    # python tools/make_data.py cvs  → 편의점(과제) 데이터만 다시 만든다
    cafe = make_cafe() if "cvs" not in sys.argv else []
    cvs = make_cvs()
    for f in sorted(OUT.iterdir()):
        print(f"{f.name:16s} {f.stat().st_size / 1024 / 1024:6.2f} MB")
    print("카페 판매 줄 수:", len(cafe), " 편의점 판매 줄 수:", len(cvs))
