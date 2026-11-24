"""course_adv.py - 심화 과정(Olist) 장별 완성 파일의 내용 (모델 + 보고서) 정의 (교안 제작용)

stage(n) 이 심화 n 장이 끝났을 때의 상태를 돌려준다.
  1장 여러 파일과 매개 변수 → 2장 Power Query 병합·조건 열 → 3장 모델링(여러 사실 테이블, 비활성 관계)
  → 4장 CALCULATE → 5장 반복 함수·시간 인텔리전스 → 6장 배송·만족도 분석
  → 7장 고급 보고서 기능 → 8장 행 수준 보안(RLS) → 9장 프로젝트(판매자 성과) 풀이
Power Query 단계 이름은 한국어 Power BI Desktop 이 자동으로 붙이는 이름과 같게 썼다.
M 식 안의 데이터 위치는 매개 변수 DataPath 하나로 모았다 (학생 PC 의 폴더가 달라도 한 곳만 고치면 된다).
"""
from pathlib import Path

from pbip_writer import Column as C, Expression, Measure as M, Page, Table, TextBox, Visual, col, mea

ROOT = Path(__file__).resolve().parent.parent
DATA = r"C:\study\power-bi-tutorial\data\olist"
THEME = ROOT / "data" / "olist" / "Olist_테마.json"


def csv(file, columns, quote="None"):
    """CSV 원본 단계. quote="Csv" 면 따옴표 안의 줄 바꿈을 값의 일부로 읽는다 (리뷰 파일)"""
    return (f'원본 = Csv.Document(File.Contents(DataPath & "\\{file}"),[Delimiter=",", Columns={columns}, '
            f'Encoding=65001, QuoteStyle=QuoteStyle.{quote}])')


def types(prev, pairs):
    """Table.TransformColumnTypes 의 {{"열", 형식}, …} 부분"""
    body = ", ".join("{" + f'"{c}", {t}' + "}" for c, t in pairs)
    return f"Table.TransformColumnTypes({prev},{{{body}}})"


def steps(*lines):
    """let … in 식 만들기: lines = [(단계 이름, 식), …]. 마지막 단계를 돌려준다"""
    out = []
    for name, expr in lines:
        n = name if name.isidentifier() else f'#"{name}"'
        out.append(f"    {n} = {expr}")
    last = lines[-1][0]
    last = last if last.isidentifier() else f'#"{last}"'
    return "let\n" + ",\n".join(out) + f"\nin\n    {last}"


def ref(name):
    return name if name.isidentifier() else f'#"{name}"'


# ============================================================================ Power Query (M)
TXT, DTM, INT, NUM, DT = "type text", "type datetime", "Int64.Type", "type number", "type date"

ORDER_TYPES = [("order_id", TXT), ("customer_id", TXT), ("order_status", TXT), ("order_purchase_timestamp", DTM),
               ("order_approved_at", DTM), ("order_delivered_carrier_date", DTM),
               ("order_delivered_customer_date", DTM), ("order_estimated_delivery_date", DTM)]


def m_orders(stage):
    base = [("원본", csv("olist_orders_dataset.csv", 8)[5:]),
            ("승격된 헤더", "Table.PromoteHeaders(원본, [PromoteAllScalars=true])"),
            ("변경된 유형", types(ref("승격된 헤더"), ORDER_TYPES))]
    if stage < 2:
        return steps(*base)
    return steps(*base,
        ("병합된 쿼리", 'Table.NestedJoin(#"변경된 유형", {"customer_id"}, 고객, {"customer_id"}, "고객", JoinKind.LeftOuter)'),
        ("확장된 고객", 'Table.ExpandTableColumn(#"병합된 쿼리", "고객", {"customer_unique_id", "customer_state"}, '
                     '{"customer_unique_id", "customer_state"})'),
        ("추가된 사용자 지정 항목", 'Table.AddColumn(#"확장된 고객", "구매일", each DateTime.Date([order_purchase_timestamp]))'),
        ("추가된 사용자 지정 항목1", 'Table.AddColumn(#"추가된 사용자 지정 항목", "배송완료일", '
                               'each DateTime.Date([order_delivered_customer_date]))'),
        ("추가된 사용자 지정 항목2", 'Table.AddColumn(#"추가된 사용자 지정 항목1", "예상배송일", '
                               'each DateTime.Date([order_estimated_delivery_date]))'),
        ("추가된 사용자 지정 항목3", 'Table.AddColumn(#"추가된 사용자 지정 항목2", "배송일수", '
                               'each if [배송완료일] = null then null else Duration.Days([배송완료일] - [구매일]))'),
        # 배송상태는 빈 값(null) 검사가 필요해 사용자 지정 열(식)로, 주문상태는 조건 열 대화 상자로 만든다
        ("추가된 사용자 지정 항목4", 'Table.AddColumn(#"추가된 사용자 지정 항목3", "배송상태", each if [배송완료일] = null then "미배송" '
                               'else if [배송완료일] > [예상배송일] then "지연" else "정시")'),
        ("추가된 조건 열", 'Table.AddColumn(#"추가된 사용자 지정 항목4", "주문상태", each if [order_status] = "delivered" then "배송완료" '
                       'else if [order_status] = "shipped" then "배송중" else if [order_status] = "canceled" then "취소" '
                       'else "기타")'),
        # 사용자 지정 열·조건 열 대화 상자에는 데이터 형식 칸이 없다 → 만든 뒤 한 번에 형식을 바꾼다
        ("변경된 유형1", 'Table.TransformColumnTypes(#"추가된 조건 열",{{"구매일", type date}, {"배송완료일", type date}, '
                      '{"예상배송일", type date}, {"배송일수", Int64.Type}, {"배송상태", type text}, {"주문상태", type text}})'),
        ("제거된 열 수", 'Table.RemoveColumns(#"변경된 유형1",{"customer_id", "order_status", "order_purchase_timestamp", '
                     '"order_approved_at", "order_delivered_carrier_date", "order_delivered_customer_date", '
                     '"order_estimated_delivery_date"})'),
        ("이름을 바꾼 열 수", 'Table.RenameColumns(#"제거된 열 수",{{"order_id", "주문ID"}, {"customer_unique_id", "고객ID"}, '
                         '{"customer_state", "고객주"}})'))


def m_items(stage):
    base = [("원본", csv("olist_order_items_dataset.csv", 7)[5:]),
            ("승격된 헤더", "Table.PromoteHeaders(원본, [PromoteAllScalars=true])"),
            ("변경된 유형", types(ref("승격된 헤더"), [("order_id", TXT), ("order_item_id", INT), ("product_id", TXT),
                                                 ("seller_id", TXT), ("shipping_limit_date", DTM), ("price", NUM),
                                                 ("freight_value", NUM)]))]
    if stage < 2:
        return steps(*base)
    return steps(*base,
        ("제거된 열 수", 'Table.RemoveColumns(#"변경된 유형",{"shipping_limit_date"})'),
        ("이름을 바꾼 열 수", 'Table.RenameColumns(#"제거된 열 수",{{"order_id", "주문ID"}, {"order_item_id", "순번"}, '
                         '{"product_id", "상품ID"}, {"seller_id", "판매자ID"}, {"price", "가격"}, {"freight_value", "배송비"}})'))


def m_payments(stage):
    base = [("원본", csv("olist_order_payments_dataset.csv", 5)[5:]),
            ("승격된 헤더", "Table.PromoteHeaders(원본, [PromoteAllScalars=true])"),
            ("변경된 유형", types(ref("승격된 헤더"), [("order_id", TXT), ("payment_sequential", INT), ("payment_type", TXT),
                                                 ("payment_installments", INT), ("payment_value", NUM)]))]
    if stage < 2:
        return steps(*base)
    return steps(*base,
        ("추가된 조건 열", 'Table.AddColumn(#"변경된 유형", "결제수단", each if [payment_type] = "credit_card" then "신용카드" '
                       'else if [payment_type] = "boleto" then "볼레토" else if [payment_type] = "voucher" then "바우처" '
                       'else if [payment_type] = "debit_card" then "체크카드" else "기타")'),
        ("변경된 유형1", 'Table.TransformColumnTypes(#"추가된 조건 열",{{"결제수단", type text}})'),
        ("제거된 열 수", 'Table.RemoveColumns(#"변경된 유형1",{"payment_type"})'),
        ("이름을 바꾼 열 수", 'Table.RenameColumns(#"제거된 열 수",{{"order_id", "주문ID"}, {"payment_sequential", "결제순번"}, '
                         '{"payment_installments", "할부개월"}, {"payment_value", "결제금액"}})'))


def m_reviews(stage):
    # 1장: 기본값(QuoteStyle.None)으로 불러오면 리뷰 글 안의 줄 바꿈 때문에 행이 쪼개진다 → 2장에서 고친다
    base = [("원본", csv("olist_order_reviews_dataset.csv", 7, "None" if stage < 2 else "Csv")[5:]),
            ("승격된 헤더", "Table.PromoteHeaders(원본, [PromoteAllScalars=true])"),
            ("변경된 유형", types(ref("승격된 헤더"), [("review_id", TXT), ("order_id", TXT), ("review_score", INT),
                                                 ("review_comment_title", TXT), ("review_comment_message", TXT),
                                                 ("review_creation_date", DTM), ("review_answer_timestamp", DTM)]))]
    if stage < 2:
        return steps(*base)
    return steps(*base,
        ("제거된 다른 열 수", 'Table.SelectColumns(#"변경된 유형",{"review_id", "order_id", "review_score", "review_creation_date"})'),
        ("이름을 바꾼 열 수", 'Table.RenameColumns(#"제거된 다른 열 수",{{"review_id", "리뷰ID"}, {"order_id", "주문ID"}, '
                         '{"review_score", "평점"}, {"review_creation_date", "리뷰작성일"}})'),
        ("변경된 유형1", 'Table.TransformColumnTypes(#"이름을 바꾼 열 수",{{"리뷰작성일", type date}})'))


def m_products(stage):
    base = [("원본", csv("olist_products_dataset.csv", 9)[5:]),
            ("승격된 헤더", "Table.PromoteHeaders(원본, [PromoteAllScalars=true])"),
            ("변경된 유형", types(ref("승격된 헤더"), [("product_id", TXT), ("product_category_name", TXT),
                                                 ("product_name_lenght", INT), ("product_description_lenght", INT),
                                                 ("product_photos_qty", INT), ("product_weight_g", INT),
                                                 ("product_length_cm", INT), ("product_height_cm", INT),
                                                 ("product_width_cm", INT)]))]
    if stage < 2:
        return steps(*base)
    return steps(*base,
        ("병합된 쿼리", 'Table.NestedJoin(#"변경된 유형", {"product_category_name"}, 카테고리번역, {"product_category_name"}, '
                     '"카테고리번역", JoinKind.LeftOuter)'),
        ("확장된 카테고리번역", 'Table.ExpandTableColumn(#"병합된 쿼리", "카테고리번역", {"product_category_name_english"}, '
                        '{"product_category_name_english"})'),
        ("병합된 쿼리1", 'Table.NestedJoin(#"확장된 카테고리번역", {"product_category_name_english"}, 카테고리한글, '
                      '{"category_english"}, "카테고리한글", JoinKind.LeftOuter)'),
        ("확장된 카테고리한글", 'Table.ExpandTableColumn(#"병합된 쿼리1", "카테고리한글", {"카테고리", "대분류"}, {"카테고리", "대분류"})'),
        ("바꾼 값", 'Table.ReplaceValue(#"확장된 카테고리한글",null,"(분류 없음)",Replacer.ReplaceValue,{"카테고리"})'),
        ("바꾼 값1", 'Table.ReplaceValue(#"바꾼 값",null,"기타",Replacer.ReplaceValue,{"대분류"})'),
        ("제거된 다른 열 수", 'Table.SelectColumns(#"바꾼 값1",{"product_id", "카테고리", "대분류"})'),
        ("이름을 바꾼 열 수", 'Table.RenameColumns(#"제거된 다른 열 수",{{"product_id", "상품ID"}})'))


def m_sellers(stage):
    base = [("원본", csv("olist_sellers_dataset.csv", 4)[5:]),
            ("승격된 헤더", "Table.PromoteHeaders(원본, [PromoteAllScalars=true])"),
            ("변경된 유형", types(ref("승격된 헤더"), [("seller_id", TXT), ("seller_zip_code_prefix", TXT),
                                                 ("seller_city", TXT), ("seller_state", TXT)]))]
    if stage < 2:
        return steps(*base)
    return steps(*base,
        ("제거된 열 수", 'Table.RemoveColumns(#"변경된 유형",{"seller_zip_code_prefix"})'),
        ("이름을 바꾼 열 수", 'Table.RenameColumns(#"제거된 열 수",{{"seller_id", "판매자ID"}, {"seller_city", "판매자도시"}, '
                         '{"seller_state", "판매자주"}})'))


def m_states(stage):
    base = [("원본", csv("브라질_주.csv", 3)[5:]),
            ("승격된 헤더", "Table.PromoteHeaders(원본, [PromoteAllScalars=true])"),
            ("변경된 유형", types(ref("승격된 헤더"), [("state", TXT), ("주이름", TXT), ("지역", TXT)]))]
    if stage < 2:
        return steps(*base)
    return steps(*base, ("이름을 바꾼 열 수", 'Table.RenameColumns(#"변경된 유형",{{"state", "주코드"}})'))


def m_customers():
    return steps(("원본", csv("olist_customers_dataset.csv", 5)[5:]),
                 ("승격된 헤더", "Table.PromoteHeaders(원본, [PromoteAllScalars=true])"),
                 ("변경된 유형", types(ref("승격된 헤더"), [("customer_id", TXT), ("customer_unique_id", TXT),
                                                      ("customer_zip_code_prefix", TXT), ("customer_city", TXT),
                                                      ("customer_state", TXT)])))


def m_translation():
    return steps(("원본", csv("product_category_name_translation.csv", 2)[5:]),
                 ("승격된 헤더", "Table.PromoteHeaders(원본, [PromoteAllScalars=true])"),
                 ("변경된 유형", types(ref("승격된 헤더"), [("product_category_name", TXT),
                                                      ("product_category_name_english", TXT)])))


def m_korean():
    return steps(("원본", csv("카테고리_한글.csv", 3)[5:]),
                 ("승격된 헤더", "Table.PromoteHeaders(원본, [PromoteAllScalars=true])"),
                 ("변경된 유형", types(ref("승격된 헤더"), [("category_english", TXT), ("카테고리", TXT), ("대분류", TXT)])))


# ============================================================================ 테이블
def raw_cols(pairs):
    kind = {TXT: "string", DTM: "dateTime", INT: "int64", NUM: "double", DT: "dateTime"}
    out = []
    for name, t in pairs:
        c = C(name, kind[t], fmt={DTM: "General Date", DT: "yyyy-mm-dd", NUM: "#,0.00", INT: "0"}.get(t),
              summarize="none" if t in (TXT, DTM, DT) or name.endswith("_id") else "sum")
        out.append(c)
    return out


def tables(stage):
    if stage < 2:   # 1장: 불러오기만 한 원래 모양 (열 이름도 영어 그대로)
        return [
            Table("주문", raw_cols(ORDER_TYPES), m=m_orders(1)),
            Table("주문상품", raw_cols([("order_id", TXT), ("order_item_id", INT), ("product_id", TXT), ("seller_id", TXT),
                                     ("shipping_limit_date", DTM), ("price", NUM), ("freight_value", NUM)]), m=m_items(1)),
            Table("결제", raw_cols([("order_id", TXT), ("payment_sequential", INT), ("payment_type", TXT),
                                  ("payment_installments", INT), ("payment_value", NUM)]), m=m_payments(1)),
            Table("리뷰", raw_cols([("review_id", TXT), ("order_id", TXT), ("review_score", INT),
                                  ("review_comment_title", TXT), ("review_comment_message", TXT),
                                  ("review_creation_date", DTM), ("review_answer_timestamp", DTM)]), m=m_reviews(1)),
            Table("상품", raw_cols([("product_id", TXT), ("product_category_name", TXT), ("product_name_lenght", INT),
                                  ("product_description_lenght", INT), ("product_photos_qty", INT),
                                  ("product_weight_g", INT), ("product_length_cm", INT), ("product_height_cm", INT),
                                  ("product_width_cm", INT)]), m=m_products(1)),
            Table("판매자", raw_cols([("seller_id", TXT), ("seller_zip_code_prefix", TXT), ("seller_city", TXT),
                                   ("seller_state", TXT)]), m=m_sellers(1)),
            Table("주", raw_cols([("state", TXT), ("주이름", TXT), ("지역", TXT)]), m=m_states(1)),
        ]
    t = [
        Table("주문", [C("주문ID", "string"), C("고객ID", "string"), C("고객주", "string"),
                     C("구매일", "dateTime", "yyyy-mm-dd"), C("배송완료일", "dateTime", "yyyy-mm-dd"),
                     C("예상배송일", "dateTime", "yyyy-mm-dd"), C("배송일수", "int64", "0", "none"),
                     C("배송상태", "string"), C("주문상태", "string")], m=m_orders(stage)),
        Table("주문상품", [C("주문ID", "string"), C("순번", "int64", "0", "none"), C("상품ID", "string"),
                       C("판매자ID", "string"), C("가격", "double", "#,0.00", "sum"), C("배송비", "double", "#,0.00", "sum")],
              m=m_items(stage)),
        Table("결제", [C("주문ID", "string"), C("결제순번", "int64", "0", "none"), C("할부개월", "int64", "0", "none"),
                     C("결제금액", "double", "#,0.00", "sum"), C("결제수단", "string")], m=m_payments(stage)),
        Table("리뷰", [C("리뷰ID", "string"), C("주문ID", "string"), C("평점", "int64", "0", "none"),
                     C("리뷰작성일", "dateTime", "yyyy-mm-dd")], m=m_reviews(stage)),
        Table("상품", [C("상품ID", "string"), C("카테고리", "string"), C("대분류", "string")], m=m_products(stage),
              hierarchies={"대분류 계층 구조": ["대분류", "카테고리"]} if stage >= 3 else None),
        Table("판매자", [C("판매자ID", "string"), C("판매자도시", "string"), C("판매자주", "string")], m=m_sellers(stage)),
        Table("주", [C("주코드", "string"), C("주이름", "string"), C("지역", "string")], m=m_states(stage),
              hierarchies={"지역 계층 구조": ["지역", "주이름"]} if stage >= 3 else None),
    ]
    if stage >= 3:
        d = C("Date", "dateTime", "yyyy-mm-dd", is_key=True)
        d.calc_source = True
        t.append(Table("날짜", [
            d,
            C("연도", "int64", "0", "none", expr="YEAR('날짜'[Date])"),
            C("분기", "string", expr="\"Q\" & QUARTER('날짜'[Date])"),
            C("월", "int64", "0", "none", expr="MONTH('날짜'[Date])"),
            C("연월", "string", expr="FORMAT('날짜'[Date], \"yyyy-mm\")"),
            C("월시작일", "dateTime", "yyyy-mm", "none", expr="DATE(YEAR('날짜'[Date]), MONTH('날짜'[Date]), 1)"),
            C("요일", "string", expr="FORMAT('날짜'[Date], \"aaa\")", sort_by="요일번호"),
            C("요일번호", "int64", "0", "none", expr="WEEKDAY('날짜'[Date], 2)"),
        ], dax="CALENDAR(DATE(2016, 1, 1), DATE(2018, 12, 31))", data_category="Time",
            hierarchies={"연도 계층 구조": ["연도", "분기", "월"]}))
        t.insert(0, Table("_측정값", [C("임시", "int64", hidden=True, expr=None)], dax='ROW("임시", 0)',
                          measures=measures(stage)))
        t[0].columns[0].calc_source = True
    return t


def expressions(stage):
    ex = [Expression("DataPath", f'"{DATA}"', parameter=True)]
    if stage >= 2:      # 병합에만 쓰고 모델에는 로드하지 않는 쿼리 (로드 사용 안 함)
        ex += [Expression("고객", m_customers()), Expression("카테고리번역", m_translation()),
               Expression("카테고리한글", m_korean())]
    return ex


def relationships(stage):
    if stage < 3:
        return []
    return [("주문상품", "주문ID", "주문", "주문ID"), ("결제", "주문ID", "주문", "주문ID"), ("리뷰", "주문ID", "주문", "주문ID"),
            ("주문상품", "상품ID", "상품", "상품ID"), ("주문상품", "판매자ID", "판매자", "판매자ID"),
            ("주문", "고객주", "주", "주코드"), ("주문", "구매일", "날짜", "Date"),
            ("주문", "배송완료일", "날짜", "Date", {"active": False})]


# ============================================================================ 측정값
MS = "_측정값"
F0, F2, PCT = "#,0", "#,0.00", "0.0%"


def measures(stage):
    ms = [
        M("상품매출", "SUM('주문상품'[가격])", F0, "1 기본"),
        M("배송비", "SUM('주문상품'[배송비])", F0, "1 기본"),
        M("주문수", "COUNTROWS('주문')", F0, "1 기본"),
        M("상품주문수", "DISTINCTCOUNT('주문상품'[주문ID])", F0, "1 기본"),
        M("객단가", "DIVIDE([상품매출], [상품주문수])", F2, "1 기본"),
        M("고객수", "DISTINCTCOUNT('주문'[고객ID])", F0, "1 기본"),
        M("평균 평점", "AVERAGE('리뷰'[평점])", F2, "1 기본"),
    ]
    if stage >= 4:
        ms += [
            M("배송완료 주문수", "CALCULATE([주문수], '주문'[주문상태] = \"배송완료\")", F0, "2 CALCULATE"),
            M("지연 주문수", "CALCULATE([주문수], '주문'[배송상태] = \"지연\")", F0, "2 CALCULATE"),
            # DAX 변수 이름은 영문자·숫자·_ 만 쓸 수 있다 (한글 변수 이름은 오류)
            M("지연율", "VAR LateOrders = [지연 주문수]\nVAR ArrivedOrders = CALCULATE([주문수], '주문'[배송상태] <> \"미배송\")\n"
                      "RETURN DIVIDE(LateOrders, ArrivedOrders)", "0.00%", "2 CALCULATE"),
            M("매출 비중", "DIVIDE([상품매출], CALCULATE([상품매출], REMOVEFILTERS('상품')))", "0.00%", "2 CALCULATE"),
            M("카테고리 평균 평점", "CALCULATE([평균 평점], CROSSFILTER('주문상품'[주문ID], '주문'[주문ID], Both))", F2,
              "2 CALCULATE"),
            M("배송완료 기준 주문수", "CALCULATE([주문수], USERELATIONSHIP('주문'[배송완료일], '날짜'[Date]))", F0, "2 CALCULATE"),
        ]
    if stage >= 5:
        ms += [
            M("총매출(배송비 포함)", "SUMX('주문상품', '주문상품'[가격] + '주문상품'[배송비])", F0, "3 반복 함수"),
            M("평균 배송일수", "AVERAGE('주문'[배송일수])", "0.0", "3 반복 함수"),
            M("재구매 고객수", "COUNTROWS(FILTER(VALUES('주문'[고객ID]), CALCULATE([주문수]) > 1))", F0, "3 반복 함수"),
            M("재구매율", "DIVIDE([재구매 고객수], [고객수])", "0.00%", "3 반복 함수"),
            M("매출 YTD", "TOTALYTD([상품매출], '날짜'[Date])", F0, "4 시간 인텔리전스"),
            M("전년 동기 매출", "CALCULATE([상품매출], SAMEPERIODLASTYEAR('날짜'[Date]))", F0, "4 시간 인텔리전스"),
            M("전년 대비 성장률", "DIVIDE([상품매출] - [전년 동기 매출], [전년 동기 매출])", "0.0%", "4 시간 인텔리전스"),
            M("전월 매출", "CALCULATE([상품매출], DATEADD('날짜'[Date], -1, MONTH))", F0, "4 시간 인텔리전스"),
            M("전월 대비 증감률", "DIVIDE([상품매출] - [전월 매출], [전월 매출])", "0.0%", "4 시간 인텔리전스"),
            M("최근 3개월 월평균 매출", "CALCULATE([상품매출], DATESINPERIOD('날짜'[Date], MAX('날짜'[Date]), -3, MONTH)) / 3",
              F0, "4 시간 인텔리전스"),
        ]
    if stage >= 7:
        ms += [M("카테고리 지연율", "CALCULATE([지연율], CROSSFILTER('주문상품'[주문ID], '주문'[주문ID], Both))", "0.00%",
                 "2 CALCULATE")]
    return ms


# ============================================================================ 보고서 (장별 화면은 아래에서 채운다)
def agg(table, column, fn=0):
    return {"Aggregation": {"Expression": col(table, column), "Function": fn}}


def card(measure, x, y, w=240, h=100, units=1):
    return Visual("cardVisual", x, y, w, h, {"Data": [mea(MS, measure)]}, objects={"value": {"labelDisplayUnits": units}})


def count_card(table, column, label, x, y, w=170, h=120):
    """행 수 카드 - 열의 '개수'(빈칸 아닌 값의 수, 함수 번호 5). 2번은 '개수(고유)' 라서 행 수가 아니다"""
    return Visual("cardVisual", x, y, w, h, {"Data": [agg(table, column, 5)]}, title=label,
                  objects={"value": {"labelDisplayUnits": 1}})


def page_stage1():
    v = [TextBox("심화 1장 - Olist 데이터 7개 불러오기", 30, 20, 900, 50)]
    for i, (t, c) in enumerate([("주문", "order_id"), ("주문상품", "order_id"), ("결제", "order_id"), ("리뷰", "review_id"),
                                ("상품", "product_id"), ("판매자", "seller_id"), ("주", "state")]):
        v.append(count_card(t, c, f"{t} 행 수", 30 + i * 175, 90))
    v.append(Visual("tableEx", 30, 210, 1220, 480, {"Values": [col("주문", c) for c, _ in ORDER_TYPES]},
                    title="주문 테이블 (불러온 그대로)"))
    return [Page("p1", "1장 불러온 데이터", v)]


def page_stage2():
    return [Page("p1", "2장 정리한 데이터", [
        TextBox("심화 2장 - 병합과 조건 열로 정리한 데이터", 30, 20, 900, 50),
        Visual("tableEx", 30, 90, 600, 600, {"Values": [col("상품", "카테고리"), col("상품", "대분류"), agg("상품", "상품ID", 2)]},
               title="상품 - 병합으로 붙인 한글 카테고리"),
        Visual("tableEx", 650, 90, 600, 220, {"Values": [col("주문", "배송상태"), agg("주문", "주문ID", 2)]},
               title="주문 - 조건 열로 만든 배송상태"),
        Visual("tableEx", 650, 330, 600, 220, {"Values": [col("결제", "결제수단"), agg("결제", "결제금액", 0)]},
               title="결제 - 조건 열로 만든 결제수단"),
    ])]


def page_stage3():
    return [Page("p1", "3장 모델", [
        TextBox("심화 3장 - 필터는 관계를 따라 한 방향으로 흐른다", 30, 20, 1000, 50),
        Visual("pivotTable", 30, 90, 760, 420, {"Rows": [col("상품", "대분류")],
                                                "Values": [mea(MS, "상품매출"), mea(MS, "주문수"), mea(MS, "상품주문수")]},
               title="대분류별 - 주문수(주문 테이블)는 왜 모두 같을까?"),
        Visual("pivotTable", 810, 90, 440, 420, {"Rows": [col("주", "지역")],
                                                 "Values": [mea(MS, "주문수"), mea(MS, "고객수"), mea(MS, "평균 평점")]},
               title="지역별 주문수 · 고객수 · 평점"),
    ])]


def page_stage4():
    v = [TextBox("심화 4장 - CALCULATE 로 조건 바꾸기", 30, 20, 900, 50)]
    for i, m in enumerate(["주문수", "배송완료 주문수", "지연 주문수", "지연율"]):
        v.append(card(m, 30 + i * 305, 85, w=290))
    v += [
        with_filters(Visual("lineChart", 30, 200, 760, 250, {"Category": [col("날짜", "월시작일")],
                                                             "Y": [mea(MS, "주문수"), mea(MS, "배송완료 기준 주문수")]},
                            title="구매일 기준 vs 배송완료일 기준 주문수 (2017-01 ~ 2018-08)"), period_filter()),
        Visual("pivotTable", 810, 200, 440, 500, {"Rows": [col("상품", "대분류")],
                                                  "Values": [mea(MS, "상품매출"), mea(MS, "매출 비중"), mea(MS, "카테고리 평균 평점")]},
               title="대분류별 매출 비중과 평점"),
        Visual("clusteredColumnChart", 30, 465, 760, 240, {"Category": [col("주문", "배송상태")], "Y": [mea(MS, "평균 평점")]},
               title="배송상태별 평균 평점", objects={"labels": {"show": True}}),
    ]
    return [Page("p1", "4장 CALCULATE", v)]


def _where(table, column, condition):
    return {"Version": 2, "From": [{"Name": "d", "Entity": table, "Type": 0}], "Where": [{"Condition": condition}]}


def _c(table, column):
    return {"Column": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": column}}


def period_filter(start="2017-01-01", end="2018-08-31"):
    """시각적 개체 필터 - 날짜[Date] 가 start ~ end 인 데이터만 (분석 기간)
    Olist 데이터는 2016년(주문 수백 건)과 2018년 9~10월(수십 건)이 너무 적어 추세·증감률을 흐린다"""
    def cmp(kind, day):
        return {"Comparison": {"ComparisonKind": kind, "Left": _c("날짜", "Date"),
                               "Right": {"Literal": {"Value": f"datetime'{day}T00:00:00'"}}}}
    return {"name": "period", "field": col("날짜", "Date"), "type": "Advanced",
            "filter": _where("날짜", "Date", {"And": {"Left": cmp(2, start), "Right": cmp(4, end)}})}


def year_filter(year):
    """시각적 개체 필터 - 날짜[연도] = year"""
    return {"name": "year", "field": col("날짜", "연도"), "type": "Categorical",
            "filter": _where("날짜", "연도", {"In": {"Expressions": [_c("날짜", "연도")],
                                                     "Values": [[{"Literal": {"Value": f"{year}L"}}]]}})}


def with_filters(visual, *filters):
    visual.filters = list(filters)
    return visual


def page_stage5():
    v = [TextBox("심화 5장 - 반복 함수와 시간 인텔리전스", 30, 20, 900, 50)]
    for i, m in enumerate(["상품매출", "총매출(배송비 포함)", "재구매 고객수", "재구매율"]):
        v.append(card(m, 30 + i * 305, 85, w=290))
    v += [
        with_filters(Visual("lineChart", 30, 200, 760, 250, {"Category": [col("날짜", "월시작일")],
                                                             "Y": [mea(MS, "상품매출"), mea(MS, "전년 동기 매출")]},
                            title="월별 매출과 전년 동기 매출 (2017-01 ~ 2018-08)",
                            objects={"valueAxis": {"labelDisplayUnits": 1000}}), period_filter()),
        with_filters(Visual("pivotTable", 810, 200, 440, 505, {"Rows": [col("날짜", "연월")],
                                                               "Values": [mea(MS, "상품매출"), mea(MS, "전년 동기 매출"),
                                                                          mea(MS, "전년 대비 성장률"), mea(MS, "전월 대비 증감률"),
                                                                          mea(MS, "매출 YTD")]},
                            title="2018년 월별 성장률과 누계"), period_filter(), year_filter(2018)),
        with_filters(Visual("lineChart", 30, 465, 760, 240, {"Category": [col("날짜", "월시작일")],
                                                             "Y": [mea(MS, "상품매출"), mea(MS, "최근 3개월 월평균 매출")]},
                            title="월 매출과 최근 3개월 월평균 (2017-01 ~ 2018-08)",
                            objects={"valueAxis": {"labelDisplayUnits": 1000}}), period_filter()),
    ]
    return [Page("p1", "5장 시간 분석", v)]


LABELS = {"labels": {"show": True}}
DROP = {"data": {"mode": "Dropdown"}}


def slicer(table, column, x, y, w=200, h=60):
    return Visual("slicer", x, y, w, h, {"Values": [col(table, column)]}, objects=DROP)


def page_stage6():
    """6장: 배송과 고객 만족도 분석 페이지"""
    v = [TextBox("배송과 고객 만족도", 20, 15, 600, 55, size=22),
         slicer("날짜", "연도", 840, 15), slicer("주", "지역", 1060, 15)]
    for i, m in enumerate(["평균 배송일수", "지연율", "평균 평점", "재구매율"]):
        v.append(card(m, 20 + i * 312, 85, w=300, h=95))
    v += [
        Visual("clusteredColumnChart", 20, 195, 400, 250, {"Category": [col("주문", "배송상태")], "Y": [mea(MS, "평균 평점")]},
               title="배송상태별 평균 평점", objects=LABELS),
        with_filters(Visual("lineChart", 440, 195, 820, 250, {"Category": [col("날짜", "월시작일")], "Y": [mea(MS, "지연율")]},
                            title="월별 지연율 (2017-01 ~ 2018-08)"), period_filter()),
        Visual("clusteredBarChart", 20, 460, 600, 245, {"Category": [col("주", "주이름")], "Y": [mea(MS, "평균 배송일수")]},
               title="고객 주별 평균 배송일수 (긴 순)", sort=(mea(MS, "평균 배송일수"), "Descending")),
        Visual("scatterChart", 640, 460, 620, 245, {"Category": [col("주", "주이름")], "X": [mea(MS, "평균 배송일수")],
                                                    "Y": [mea(MS, "평균 평점")]},
               title="주별 배송일수와 평점 - 오래 걸릴수록 평점이 낮다"),
    ]
    return Page("p1", "배송과 만족도", v)


def topn_filter(name, table, column, mtable, measure, n):
    """시각적 개체 필터 '상위 N' (measure 기준 위에서 n 개)"""
    return {"name": name, "field": col(table, column), "type": "TopN", "filter": {
        "Version": 2,
        "From": [{"Name": "subquery", "Expression": {"Subquery": {"Query": {
            "Version": 2,
            "From": [{"Name": "s", "Entity": table, "Type": 0}, {"Name": "m", "Entity": mtable, "Type": 0}],
            "Select": [{"Column": {"Expression": {"SourceRef": {"Source": "s"}}, "Property": column}, "Name": "field"}],
            "OrderBy": [{"Direction": 2, "Expression": {"Measure": {"Expression": {"SourceRef": {"Source": "m"}},
                                                                    "Property": measure}}}],
            "Top": n}}}, "Type": 2},
                 {"Name": "s", "Entity": table, "Type": 0}],
        "Where": [{"Condition": {"In": {"Expressions": [{"Column": {"Expression": {"SourceRef": {"Source": "s"}},
                                                                     "Property": column}}],
                                        "Table": {"SourceRef": {"Source": "subquery"}}}}}]}}


def gradient(measure, low="#FFFFFF", high="#E74C3C"):
    """조건부 서식 - 값이 클수록 진한 색 (배경색 그라데이션)"""
    return {"solid": {"color": {"expr": {"FillRule": {
        "Input": mea(MS, measure),
        "FillRule": {"linearGradient2": {"min": {"color": {"Literal": {"Value": f"'{low}'"}}},
                                         "max": {"color": {"Literal": {"Value": f"'{high}'"}}},
                                         "nullColoringStrategy": {"strategy": {"Literal": {"Value": "'asZero'"}}}}}}}}}}


def page_overview():
    """7장: 드릴다운·상위 N·조건부 서식·도구 설명 페이지가 들어간 개요 페이지"""
    v = [TextBox("Olist 판매 개요", 20, 15, 600, 55, size=22),
         slicer("날짜", "연도", 840, 15), slicer("주", "지역", 1060, 15)]
    for i, m in enumerate(["상품매출", "상품주문수", "객단가", "평균 평점"]):
        v.append(card(m, 20 + i * 312, 85, w=300, h=95))
    top = Visual("clusteredBarChart", 20, 195, 600, 270, {"Category": [col("상품", "카테고리")], "Y": [mea(MS, "상품매출")]},
                 title="매출 상위 10 카테고리 (마우스 올리기: 요약 / 오른쪽 클릭: 드릴스루)", sort=(mea(MS, "상품매출"), "Descending"),
                 objects={"labels": {"show": True, "labelDisplayUnits": 1000}},
                 container={"visualTooltip": {"show": True, "type": "Canvas", "section": "p3"}})  # Canvas = 보고서 페이지 도구 설명
    top.filters = [topn_filter("top10", "상품", "카테고리", MS, "상품매출", 10)]
    v += [
        top,
        Visual("clusteredColumnChart", 640, 195, 620, 270, {"Category": [col("주", "지역"), col("주", "주이름")],
                                                            "Y": [mea(MS, "상품매출")]},
               title="지역별 매출 (↓ 로 주까지 드릴다운)", sort=(mea(MS, "상품매출"), "Descending"),
               objects={"valueAxis": {"labelDisplayUnits": 1000}}),
        Visual("pivotTable", 20, 478, 1240, 235, {"Rows": [col("상품", "대분류")],
                                                  "Values": [mea(MS, "상품매출"), mea(MS, "매출 비중"), mea(MS, "카테고리 지연율"),
                                                             mea(MS, "카테고리 평균 평점")]},
               title="대분류별 성과 (지연율이 높을수록 빨간색)",
               sort=(mea(MS, "상품매출"), "Descending"),
               cf={"카테고리 지연율": gradient("카테고리 지연율", "#FFFFFF", "#E74C3C")}),
    ]
    return Page("p1", "판매 개요", v)


def page_drillthrough():
    """7장: 카테고리를 오른쪽 클릭 → 드릴스루 → 이 페이지가 그 카테고리로 걸러져 열린다"""
    v = [TextBox("카테고리 상세", 20, 15, 600, 55, size=22)]
    for i, m in enumerate(["상품매출", "상품주문수", "카테고리 평균 평점", "매출 비중"]):
        v.append(card(m, 20 + i * 312, 85, w=300, h=95))
    v += [
        with_filters(Visual("lineChart", 20, 195, 760, 250, {"Category": [col("날짜", "월시작일")], "Y": [mea(MS, "상품매출")]},
                            title="월별 매출 (2017-01 ~ 2018-08)"), period_filter()),
        Visual("donutChart", 800, 195, 460, 250, {"Category": [col("주", "지역")], "Y": [mea(MS, "상품매출")]},
               title="고객 지역별 매출 비중", objects={"labels": {"labelStyle": "Percent of total"}, "centerValue": {"show": False}}),
        Visual("tableEx", 20, 460, 1240, 245, {"Values": [col("판매자", "판매자ID"), col("판매자", "판매자주"),
                                                         mea(MS, "상품매출"), mea(MS, "상품주문수")]},
               title="이 카테고리를 많이 판 판매자", sort=(mea(MS, "상품매출"), "Descending")),
    ]
    p = Page("p2", "카테고리 상세", v)
    p.binding = {"name": "dt_category", "type": "Drillthrough",
                 "parameters": [{"name": "dt_param", "boundFilter": "dt_filter", "fieldExpr": col("상품", "카테고리")}]}
    p.filters = [{"name": "dt_filter", "field": col("상품", "카테고리"), "type": "Categorical", "howCreated": "Drillthrough"}]
    return p


def page_tooltip():
    """7장: 막대에 마우스를 올리면 뜨는 작은 '도구 설명 페이지'"""
    p = Page("p3", "카테고리 툴팁", [
        Visual("cardVisual", 10, 5, 145, 90, {"Data": [mea(MS, "상품매출")]},
               objects={"value": {"labelDisplayUnits": 1000, "fontSize": 16}}),   # 작은 페이지라 값 글자를 줄인다
        Visual("cardVisual", 165, 5, 145, 90, {"Data": [mea(MS, "카테고리 평균 평점")]},
               objects={"value": {"fontSize": 16}}),
        with_filters(Visual("lineChart", 10, 100, 300, 135, {"Category": [col("날짜", "월시작일")], "Y": [mea(MS, "상품매출")]},
                            title="월별 매출"), period_filter()),
    ], width=320, height=240)
    p.binding = {"name": "tt_category", "type": "Tooltip"}
    p.hidden = True
    return p


ROLES = {"남동부 담당": {"주": "[지역] = \"남동부\""}, "상파울루 담당": {"주": "[주코드] = \"SP\""}}


def page_project():
    """9장 프로젝트 풀이: 판매자 성과 대시보드"""
    v = [TextBox("판매자 성과 대시보드", 20, 15, 600, 55, size=22),
         slicer("날짜", "연도", 840, 15), slicer("판매자", "판매자주", 1060, 15)]
    for i, m in enumerate(["판매자수", "상품매출", "판매자 평균 평점", "판매자 지연율"]):
        v.append(card(m, 20 + i * 312, 85, w=300, h=95))
    top = Visual("clusteredBarChart", 20, 195, 600, 510, {"Category": [col("판매자", "판매자명")], "Y": [mea(MS, "상품매출")]},
                 title="매출 상위 15 판매자", sort=(mea(MS, "상품매출"), "Descending"),
                 objects={"labels": {"show": True, "labelDisplayUnits": 1000}})
    top.filters = [topn_filter("top15", "판매자", "판매자명", MS, "상품매출", 15)]
    v += [
        top,
        Visual("clusteredColumnChart", 640, 195, 620, 240, {"Category": [col("판매자", "판매자주")], "Y": [mea(MS, "판매자수")]},
               title="주별 판매자 수", sort=(mea(MS, "판매자수"), "Descending"),
               objects={"labels": {"show": True, "labelDisplayUnits": 1}, "valueAxis": {"labelDisplayUnits": 1}}),
        Visual("pivotTable", 640, 450, 620, 255, {"Rows": [col("판매자", "판매자명")],
                                                  "Values": [mea(MS, "상품매출"), mea(MS, "상품주문수"), mea(MS, "판매자 평균 평점"),
                                                             mea(MS, "판매자 지연율")]},
               title="판매자별 성과 (지연율이 높을수록 빨간색)", sort=(mea(MS, "상품매출"), "Descending"),
               cf={"판매자 지연율": gradient("판매자 지연율")}),
    ]
    return Page("p1", "판매자 성과", v)


def project_measures():
    return [
        M("판매자수", "DISTINCTCOUNT('주문상품'[판매자ID])", F0, "5 프로젝트"),
        M("판매자 평균 평점", "CALCULATE([평균 평점], CROSSFILTER('주문상품'[주문ID], '주문'[주문ID], Both))", F2, "5 프로젝트"),
        M("판매자 지연율", "CALCULATE([지연율], CROSSFILTER('주문상품'[주문ID], '주문'[주문ID], Both))", "0.00%", "5 프로젝트"),
    ]


def pages(stage):
    if stage <= 5:
        return {1: page_stage1, 2: page_stage2, 3: page_stage3, 4: page_stage4, 5: page_stage5}[stage]()
    if stage == 6:
        return [page_stage6()]
    if stage in (7, 8):
        return [page_overview(), page_drillthrough(), page_tooltip(), page_stage6_as("p4")]
    return [page_project()]


def page_stage6_as(name):
    p = page_stage6()
    p.name = name
    for i, v in enumerate(p.visuals):
        v.name = f"{name}_v{i + 1:02d}"
    return p


STAGES = {1: "심화01_여러파일과매개변수", 2: "심화02_병합과조건열", 3: "심화03_모델링", 4: "심화04_CALCULATE",
          5: "심화05_시간인텔리전스", 6: "심화06_배송과만족도분석", 7: "심화07_고급보고서기능", 8: "심화08_행수준보안",
          9: "심화09_프로젝트_판매자성과_풀이"}


def project(stage):
    t = tables(stage)
    if stage >= 9:   # 프로젝트: 판매자 이름 열(계산 열)과 측정값 3개를 더한다
        t[0].measures += project_measures()
        seller = [x for x in t if x.name == "판매자"][0]
        seller.columns.append(C("판매자명", "string", expr="LEFT('판매자'[판매자ID], 6) & \" · \" & '판매자'[판매자도시]"))
    return dict(tables=t, relationships=relationships(stage), pages=pages(stage),
                expressions=expressions(stage), custom_theme=THEME if stage >= 3 else None,
                roles=ROLES if stage >= 8 else None)
