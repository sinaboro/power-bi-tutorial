"""course.py - 장별 완성 파일의 내용 (모델 + 보고서) 정의 (교안 제작용)

stage(n) 이 n 장이 끝났을 때의 상태를 돌려준다.
  1장 데이터 불러오기 → 2장 Power Query 정리 → 3장 관계·날짜 테이블 → 4장 측정값
  → 5장 차트 → 6장 슬라이서·서식 → 7장 대시보드 완성,  그리고 과제(편의점) 풀이
Power Query 단계 이름은 한국어 Power BI Desktop 이 자동으로 붙이는 이름과 같게 썼다.
"""
from pathlib import Path

from pbip_writer import Column as C, Measure as M, Page, Table, TextBox, Visual, col, mea, lit

DATA = r"C:\study\power-bi-tutorial\data"
ROOT = Path(__file__).resolve().parent.parent
THEME = ROOT / "data" / "하루커피_테마.json"


def agg(table, column, fn=0):
    """열을 그대로 끌어다 놓았을 때 생기는 암시적 집계 (0 = 합계)"""
    return {"Aggregation": {"Expression": col(table, column), "Function": fn}}


# ============================================================================ Power Query (M)
def m_sales(stage):
    m = f'''let
    원본 = Csv.Document(File.Contents("{DATA}\\카페_판매.csv"),[Delimiter=",", Columns=9, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"승격된 헤더" = Table.PromoteHeaders(원본, [PromoteAllScalars=true]),
    #"변경된 유형" = Table.TransformColumnTypes(#"승격된 헤더",{{{{"주문상세번호", Int64.Type}}, {{"주문번호", type text}}, {{"주문일시", type datetime}}, {{"매장코드", type text}}, {{"메뉴코드", type text}}, {{"수량", Int64.Type}}, {{"단가", Int64.Type}}, {{"결제수단", type text}}, {{"주문유형", type text}}}})'''
    if stage < 2:
        return m + '\nin\n    #"변경된 유형"'
    return m + ''',
    #"제거된 중복 항목 수" = Table.Distinct(#"변경된 유형"),
    #"필터링된 행" = Table.SelectRows(#"제거된 중복 항목 수", each [수량] <> null),
    #"바꾼 값" = Table.ReplaceValue(#"필터링된 행","신용카드","카드",Replacer.ReplaceText,{"결제수단"}),
    #"바꾼 값1" = Table.ReplaceValue(#"바꾼 값","CARD","카드",Replacer.ReplaceText,{"결제수단"}),
    #"삽입된 날짜" = Table.AddColumn(#"바꾼 값1", "날짜", each DateTime.Date([주문일시]), type date),
    #"삽입된 시간" = Table.AddColumn(#"삽입된 날짜", "시간", each Time.Hour([주문일시]), Int64.Type),
    #"이름을 바꾼 열 수" = Table.RenameColumns(#"삽입된 시간",{{"날짜", "주문일자"}, {"시간", "주문시"}}),
    #"추가된 사용자 지정 항목" = Table.AddColumn(#"이름을 바꾼 열 수", "매출액", each [수량] * [단가]),
    #"변경된 유형1" = Table.TransformColumnTypes(#"추가된 사용자 지정 항목",{{"매출액", Int64.Type}}),
    #"제거된 열 수" = Table.RemoveColumns(#"변경된 유형1",{"주문상세번호", "주문일시"})
in
    #"제거된 열 수"'''


def m_menu(stage):
    m = f'''let
    원본 = Excel.Workbook(File.Contents("{DATA}\\카페_메뉴.xlsx"), null, true),
    메뉴_Sheet = 원본{{[Item="메뉴",Kind="Sheet"]}}[Data],
    #"승격된 헤더" = Table.PromoteHeaders(메뉴_Sheet, [PromoteAllScalars=true]),
    #"변경된 유형" = Table.TransformColumnTypes(#"승격된 헤더",{{{{"메뉴코드", type text}}, {{"메뉴명", type text}}, {{"분류", type text}}, {{"판매가", Int64.Type}}, {{"원가", Int64.Type}}}})'''
    if stage < 2:
        return m + '\nin\n    #"변경된 유형"'
    return m + ''',
    #"잘린 텍스트" = Table.TransformColumns(#"변경된 유형",{{"메뉴명", Text.Trim, type text}}),
    #"제거된 중복 항목 수" = Table.Distinct(#"잘린 텍스트")
in
    #"제거된 중복 항목 수"'''


def m_store(stage):
    m = f'''let
    원본 = Excel.Workbook(File.Contents("{DATA}\\카페_매장.xlsx"), null, true),
    매장_Sheet = 원본{{[Item="매장",Kind="Sheet"]}}[Data],
    #"승격된 헤더" = Table.PromoteHeaders(매장_Sheet, [PromoteAllScalars=true]),
    #"변경된 유형" = Table.TransformColumnTypes(#"승격된 헤더",{{{{"매장코드", type text}}, {{"매장명", type text}}, {{"지역", type text}}, {{"오픈일", type datetime}}, {{"좌석수", Int64.Type}}}})'''
    if stage < 2:
        return m + '\nin\n    #"변경된 유형"'
    return m + ''',
    #"변경된 유형1" = Table.TransformColumnTypes(#"변경된 유형",{{"오픈일", type date}})
in
    #"변경된 유형1"'''


# ============================================================================ 테이블
def tables(stage):
    if stage < 2:
        sales_cols = [C("주문상세번호", "int64", "0", "sum"), C("주문번호", "string"), C("주문일시", "dateTime", "General Date"),
                      C("매장코드", "string"), C("메뉴코드", "string"), C("수량", "int64", "0", "sum"),
                      C("단가", "int64", "0", "sum"), C("결제수단", "string"), C("주문유형", "string")]
    else:
        sales_cols = [C("주문번호", "string"), C("매장코드", "string"), C("메뉴코드", "string"),
                      C("수량", "int64", "#,0", "sum"), C("단가", "int64", "#,0", "sum"), C("결제수단", "string"),
                      C("주문유형", "string"), C("주문일자", "dateTime", "yyyy-mm-dd"), C("주문시", "int64", "0", "none"),
                      C("매출액", "int64", "#,0", "sum")]
    measures = []
    if stage >= 4:
        measures = [
            M("총매출", "SUM('판매'[매출액])"),
            M("판매수량", "SUM('판매'[수량])"),
            M("주문수", "DISTINCTCOUNT('판매'[주문번호])"),
            M("객단가", "DIVIDE([총매출], [주문수])"),
        ]
    store_open_type = "dateTime"
    t = [
        Table("판매", sales_cols, m=m_sales(stage), measures=measures),
        Table("메뉴", [C("메뉴코드", "string"), C("메뉴명", "string"), C("분류", "string"),
                     C("판매가", "int64", "#,0", "sum"), C("원가", "int64", "#,0", "sum")], m=m_menu(stage)),
        Table("매장", [C("매장코드", "string"), C("매장명", "string"), C("지역", "string"),
                     C("오픈일", store_open_type, "yyyy-mm-dd" if stage >= 2 else "General Date"),
                     C("좌석수", "int64", "0", "sum")], m=m_store(stage)),
    ]
    if stage >= 3:
        d = C("Date", "dateTime", "yyyy-mm-dd", is_key=True)
        d.calc_source = True
        t.append(Table("날짜", [
            d,
            C("연도", "int64", "0", "none", expr="YEAR('날짜'[Date])"),
            C("월", "int64", "0", "none", expr="MONTH('날짜'[Date])"),
            C("연월", "string", expr="FORMAT('날짜'[Date], \"yyyy-mm\")"),
            C("월시작일", "dateTime", "yyyy-mm", "none", expr="DATE(YEAR('날짜'[Date]), MONTH('날짜'[Date]), 1)"),
            C("요일", "string", expr="FORMAT('날짜'[Date], \"aaa\")", sort_by="요일번호"),
            C("요일번호", "int64", "0", "none", expr="WEEKDAY('날짜'[Date], 2)"),
        ], dax="CALENDAR(DATE(2023, 1, 1), DATE(2025, 12, 31))", data_category="Time"))
    return t


def relationships(stage):
    if stage < 3:
        return []
    return [("판매", "메뉴코드", "메뉴", "메뉴코드"), ("판매", "매장코드", "매장", "매장코드"),
            ("판매", "주문일자", "날짜", "Date")]


# ============================================================================ 보고서
S, MN, ST, D = "판매", "메뉴", "매장", "날짜"
LABELS = {"labels": {"show": True}}


def card(measure, x, y, w=290, h=110, plain=True):
    """새 카드(cardVisual). plain=True 면 표시 단위를 '없음' 으로 → 2,116,506,700 처럼 전체 숫자"""
    obj = {"value": {"labelDisplayUnits": 1}} if plain else {}
    return Visual("cardVisual", x, y, w, h, {"Data": [mea(S, measure)]}, objects=obj)


MILLION = {"labels": {"show": True, "labelDisplayUnits": 1000000}, "valueAxis": {"labelDisplayUnits": 1000000}}
PERCENT = {"labels": {"labelStyle": "Percent of total"}, "centerValue": {"show": False}}   # 비율만, 가운데 합계 숨김
COUNT_LABELS = {"labels": {"show": True, "labelDisplayUnits": 1}}                          # 93,842 처럼 전체 숫자
AXIS_M = {"valueAxis": {"labelDisplayUnits": 1000000}}


def page_stage1():
    return [Page("p1", "1장 불러온 데이터", [
        TextBox("1장 - 데이터 불러오기", 30, 20, 600, 50),
        Visual("tableEx", 30, 90, 700, 200, {"Values": [col(ST, "매장코드"), col(ST, "매장명"), col(ST, "지역"),
                                                         col(ST, "오픈일"), col(ST, "좌석수")]}, title="매장 목록"),
    ])]


def page_stage2():
    return [Page("p1", "2장 정리한 데이터", [
        TextBox("2장 - Power Query 로 정리한 메뉴", 30, 20, 700, 50),
        Visual("tableEx", 30, 90, 760, 600, {"Values": [col(MN, "메뉴코드"), col(MN, "메뉴명"), col(MN, "분류"),
                                                         col(MN, "판매가"), col(MN, "원가")]}, title="메뉴 (빈칸·중복 정리 후)"),
    ])]


def page_stage3():
    return [Page("p1", "3장 관계 확인", [
        TextBox("3장 - 관계를 만들면 표가 함께 움직인다", 30, 20, 900, 50),
        Visual("pivotTable", 30, 90, 900, 220, {"Rows": [col(ST, "매장명")], "Columns": [col(D, "연도")],
                                                 "Values": [agg(S, "매출액")]}, title="매장별 · 연도별 매출액 합계"),
    ])]


def page_stage4():
    v = [TextBox("4장 - 측정값 만들기", 30, 20, 600, 50)]
    for i, m in enumerate(["총매출", "주문수", "객단가", "판매수량"]):
        v.append(card(m, 30 + i * 305, 90, plain=False))   # 4장은 기본 표시 단위(자동) 그대로 - 6장에서 바꾼다
    v.append(Visual("pivotTable", 30, 230, 1220, 220, {"Rows": [col(ST, "매장명")],
                    "Values": [mea(S, "총매출"), mea(S, "주문수"), mea(S, "객단가"), mea(S, "판매수량")]},
                    title="매장별 측정값 (같은 측정값이 행마다 다르게 계산된다)"))
    return [Page("p1", "4장 측정값", v)]


def analysis_visuals(styled):
    """5장 차트들. styled=True 면 6장의 서식(데이터 레이블 등)을 더한다"""
    lab = LABELS if styled else {}
    return [
        Visual("lineChart", 20, 90, 760, 300, {"Category": [col(D, "월시작일")], "Y": [mea(S, "총매출")]},
               title="월별 매출 추이"),
        Visual("clusteredColumnChart", 800, 90, 460, 300, {"Category": [col(ST, "매장명")], "Y": [mea(S, "총매출")]},
               title="매장별 매출", objects=lab),
        Visual("donutChart", 20, 405, 400, 300, {"Category": [col(MN, "분류")], "Y": [mea(S, "총매출")]},
               title="분류별 매출 비중"),
        Visual("pivotTable", 440, 405, 820, 300, {"Rows": [col(ST, "매장명")], "Columns": [col(D, "연도")],
                                                   "Values": [mea(S, "총매출")]}, title="매장 × 연도 매출"),
    ]


def page_stage5():
    v = [TextBox("5장 - 차트 만들기", 30, 20, 600, 50)] + analysis_visuals(False)
    v2 = [TextBox("5장 - 메뉴와 시간대", 30, 20, 600, 50),
          Visual("clusteredBarChart", 20, 90, 600, 615, {"Category": [col(MN, "메뉴명")], "Y": [mea(S, "판매수량")]},
                 title="메뉴별 판매수량", sort=(mea(S, "판매수량"), "Descending")),
          Visual("clusteredColumnChart", 640, 90, 620, 300, {"Category": [col(S, "주문시")], "Y": [mea(S, "주문수")]},
                 title="시간대별 주문수", sort=(col(S, "주문시"), "Ascending"))]
    return [Page("p1", "매출 분석", v), Page("p2", "메뉴·시간", v2)]


def slicers(y=20):
    return [Visual("slicer", 820, y, 200, 60, {"Values": [col(D, "연도")]},
                   objects={"data": {"mode": "Dropdown"}}),
            Visual("slicer", 1040, y, 220, 60, {"Values": [col(ST, "매장명")]},
                   objects={"data": {"mode": "Dropdown"}})]


def page_summary(styled=True):
    v = [TextBox("하루커피 매출 대시보드", 20, 15, 700, 55, size=22)] + slicers()
    for i, m in enumerate(["총매출", "주문수", "객단가", "판매수량"]):
        v.append(card(m, 20 + i * 312, 85, w=300, h=100))
    v += [
        Visual("lineChart", 20, 200, 760, 250, {"Category": [col(D, "월시작일")], "Y": [mea(S, "총매출")]}, title="월별 매출 추이",
               objects={"valueAxis": {"labelDisplayUnits": 1000000}}),
        Visual("clusteredColumnChart", 800, 200, 460, 250, {"Category": [col(ST, "매장명")], "Y": [mea(S, "총매출")]},
               title="매장별 매출 (백만 원)", objects=MILLION),
        Visual("donutChart", 20, 465, 400, 240, {"Category": [col(MN, "분류")], "Y": [mea(S, "총매출")]}, title="분류별 매출 비중",
               objects=PERCENT),
        Visual("pivotTable", 440, 465, 820, 240, {"Rows": [col(ST, "매장명")], "Columns": [col(D, "연도")],
                                                   "Values": [mea(S, "총매출")]}, title="매장 × 연도 매출"),
    ]
    return Page("p1", "매출 요약", v)


def page_menu_time():
    v = [TextBox("메뉴 · 시간 분석", 20, 15, 700, 55, size=22)] + slicers()
    v += [
        Visual("slicer", 600, 20, 200, 60, {"Values": [col(MN, "분류")]}, objects={"data": {"mode": "Dropdown"}}),
        Visual("clusteredBarChart", 20, 85, 600, 620, {"Category": [col(MN, "메뉴명")], "Y": [mea(S, "판매수량")]},
               title="메뉴별 판매수량", sort=(mea(S, "판매수량"), "Descending"), objects=COUNT_LABELS),
        Visual("clusteredColumnChart", 640, 85, 620, 300, {"Category": [col(S, "주문시")], "Y": [mea(S, "주문수")]},
               title="시간대별 주문수", sort=(col(S, "주문시"), "Ascending")),
        Visual("clusteredColumnChart", 640, 400, 300, 305, {"Category": [col(D, "요일")], "Y": [mea(S, "총매출")]},
               title="요일별 매출 (백만 원)", sort=(col(D, "요일"), "Ascending"), objects=AXIS_M),
        Visual("donutChart", 960, 400, 300, 305, {"Category": [col(S, "결제수단")], "Y": [mea(S, "주문수")]},
               title="결제수단별 주문 비중", objects=PERCENT),
    ]
    return Page("p2", "메뉴·시간 분석", v)


def pages(stage):
    if stage == 1:
        return page_stage1()
    if stage == 2:
        return page_stage2()
    if stage == 3:
        return page_stage3()
    if stage == 4:
        return page_stage4()
    if stage == 5:
        return page_stage5()
    if stage == 6:
        p = page_summary()
        p.display = "매출 분석"
        return [p]
    return [page_summary(), page_menu_time()]


STAGES = {
    1: "01_데이터불러오기", 2: "02_데이터정리", 3: "03_관계와날짜", 4: "04_측정값",
    5: "05_차트", 6: "06_슬라이서와서식", 7: "07_대시보드완성",
}


def project(stage):
    return dict(tables=tables(stage), relationships=relationships(stage), pages=pages(stage),
                custom_theme=THEME if stage >= 6 else None)
