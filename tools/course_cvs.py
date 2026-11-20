"""course_cvs.py - 과제(편의점 매출 대시보드) 풀이 파일의 내용 (교안 제작용)

카페 과정에서 배운 순서를 편의점 데이터에 그대로 적용한 모범 답안이다.
  데이터 불러오기 → 형식 확인 → 날짜 테이블과 관계 → 측정값 → 대시보드 한 페이지
"""
from pathlib import Path

from course import DATA, THEME, MILLION, PERCENT, COUNT_LABELS, card as _card
from pbip_writer import Column as C, Measure as M, Page, Table, TextBox, Visual, col, mea

S, P, ST, D = "판매", "상품", "점포", "날짜"


def m_csv():
    return f'''let
    원본 = Csv.Document(File.Contents("{DATA}\\편의점_판매.csv"),[Delimiter=",", Columns=7, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"승격된 헤더" = Table.PromoteHeaders(원본, [PromoteAllScalars=true]),
    #"변경된 유형" = Table.TransformColumnTypes(#"승격된 헤더",{{{{"판매번호", Int64.Type}}, {{"판매일자", type date}}, {{"판매시", Int64.Type}}, {{"점포코드", type text}}, {{"상품코드", type text}}, {{"수량", Int64.Type}}, {{"판매금액", Int64.Type}}}})
in
    #"변경된 유형"'''


def m_xlsx(file, sheet, types):
    t = ", ".join(f'{{"{c}", {ty}}}' for c, ty in types)
    return f'''let
    원본 = Excel.Workbook(File.Contents("{DATA}\\{file}"), null, true),
    {sheet}_Sheet = 원본{{[Item="{sheet}",Kind="Sheet"]}}[Data],
    #"승격된 헤더" = Table.PromoteHeaders({sheet}_Sheet, [PromoteAllScalars=true]),
    #"변경된 유형" = Table.TransformColumnTypes(#"승격된 헤더",{{{t}}})
in
    #"변경된 유형"'''


def card(measure, x, y, w=300, h=100):
    v = _card(measure, x, y, w, h)
    v.roles = {"Data": [mea(S, measure)]}
    return v


def project():
    d = C("Date", "dateTime", "yyyy-mm-dd", is_key=True)
    d.calc_source = True
    tables = [
        Table(S, [C("판매번호", "int64", "0", "none"), C("판매일자", "dateTime", "yyyy-mm-dd"), C("판매시", "int64", "0", "none"),
                  C("점포코드", "string"), C("상품코드", "string"), C("수량", "int64", "#,0", "sum"),
                  C("판매금액", "int64", "#,0", "sum")], m=m_csv(),
              measures=[M("총매출", "SUM('판매'[판매금액])"), M("판매수량", "SUM('판매'[수량])"),
                        M("판매건수", "COUNTROWS('판매')"), M("건당매출", "DIVIDE([총매출], [판매건수])")]),
        Table(P, [C("상품코드", "string"), C("상품명", "string"), C("카테고리", "string"), C("가격", "int64", "#,0", "sum")],
              m=m_xlsx("편의점_상품.xlsx", "상품", [("상품코드", "type text"), ("상품명", "type text"),
                                                   ("카테고리", "type text"), ("가격", "Int64.Type")])),
        Table(ST, [C("점포코드", "string"), C("점포명", "string"), C("상권", "string")],
              m=m_xlsx("편의점_점포.xlsx", "점포", [("점포코드", "type text"), ("점포명", "type text"), ("상권", "type text")])),
        Table(D, [d, C("월", "int64", "0", "none", expr="MONTH('날짜'[Date])"),
                  C("월시작일", "dateTime", "yyyy-mm", "none", expr="DATE(YEAR('날짜'[Date]), MONTH('날짜'[Date]), 1)"),
                  C("요일", "string", expr="FORMAT('날짜'[Date], \"aaa\")", sort_by="요일번호"),
                  C("요일번호", "int64", "0", "none", expr="WEEKDAY('날짜'[Date], 2)")],
              dax="CALENDAR(DATE(2025, 1, 1), DATE(2025, 12, 31))", data_category="Time"),
    ]
    rels = [(S, "상품코드", P, "상품코드"), (S, "점포코드", ST, "점포코드"), (S, "판매일자", D, "Date")]
    v = [TextBox("편의점 매출 대시보드 (2025년)", 20, 15, 700, 55, size=22),
         Visual("slicer", 820, 20, 200, 60, {"Values": [col(ST, "점포명")]}, objects={"data": {"mode": "Dropdown"}}),
         Visual("slicer", 1040, 20, 220, 60, {"Values": [col(P, "카테고리")]}, objects={"data": {"mode": "Dropdown"}})]
    for i, m in enumerate(["총매출", "판매건수", "건당매출", "판매수량"]):
        v.append(card(m, 20 + i * 312, 85))
    v += [
        Visual("lineChart", 20, 200, 620, 250, {"Category": [col(D, "월시작일")], "Y": [mea(S, "총매출")]}, title="월별 매출 추이",
               objects={"valueAxis": {"labelDisplayUnits": 1000000}}),
        Visual("clusteredColumnChart", 660, 200, 600, 250, {"Category": [col(S, "판매시")], "Y": [mea(S, "판매건수")]},
               title="시간대별 판매건수", sort=(col(S, "판매시"), "Ascending")),
        Visual("clusteredBarChart", 20, 465, 620, 240, {"Category": [col(P, "카테고리")], "Y": [mea(S, "총매출")]},
               title="카테고리별 매출 (백만 원)", sort=(mea(S, "총매출"), "Descending"), objects=MILLION),
        Visual("donutChart", 660, 465, 600, 240, {"Category": [col(ST, "점포명")], "Y": [mea(S, "총매출")]},
               title="점포별 매출 비중", objects=PERCENT),
    ]
    return dict(tables=tables, relationships=rels, pages=[Page("p1", "편의점 대시보드", v)], custom_theme=THEME)
