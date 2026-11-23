r"""olist_reference.py - 심화 과정 정답 숫자를 pandas 로 계산한다 (교안 제작용)

Power Query 단계와 같은 규칙으로 데이터를 정리한 뒤 측정값을 계산해,
Power BI 가 낸 값(make_pbix.py 가 DAX 로 질의한 값)과 교안의 숫자가 맞는지 확인하는 데 쓴다.

    python tools\olist_reference.py
"""
from pathlib import Path

import pandas as pd

D = Path(__file__).resolve().parent.parent / "data" / "olist"


def load():
    o = pd.read_csv(D / "olist_orders_dataset.csv", parse_dates=["order_purchase_timestamp", "order_delivered_customer_date",
                                                                  "order_estimated_delivery_date"])
    c = pd.read_csv(D / "olist_customers_dataset.csv", dtype={"customer_zip_code_prefix": str})
    o = o.merge(c[["customer_id", "customer_unique_id", "customer_state"]], on="customer_id", how="left")
    o["구매일"] = o.order_purchase_timestamp.dt.normalize()
    o["배송완료일"] = o.order_delivered_customer_date.dt.normalize()
    o["예상배송일"] = o.order_estimated_delivery_date.dt.normalize()
    o["배송일수"] = (o["배송완료일"] - o["구매일"]).dt.days
    o["배송상태"] = "정시"
    o.loc[o["배송완료일"] > o["예상배송일"], "배송상태"] = "지연"
    o.loc[o["배송완료일"].isna(), "배송상태"] = "미배송"
    st = {"delivered": "배송완료", "shipped": "배송중", "canceled": "취소"}
    o["주문상태"] = o.order_status.map(st).fillna("기타")
    i = pd.read_csv(D / "olist_order_items_dataset.csv")
    p = pd.read_csv(D / "olist_order_payments_dataset.csv")
    r = pd.read_csv(D / "olist_order_reviews_dataset.csv")
    pr = pd.read_csv(D / "olist_products_dataset.csv")
    tr = pd.read_csv(D / "product_category_name_translation.csv", encoding="utf-8-sig")
    ko = pd.read_csv(D / "카테고리_한글.csv", encoding="utf-8-sig")
    pr = pr.merge(tr, on="product_category_name", how="left").merge(
        ko, left_on="product_category_name_english", right_on="category_english", how="left")
    pr["카테고리"] = pr["카테고리"].fillna("(분류 없음)")
    pr["대분류"] = pr["대분류"].fillna("기타")
    s = pd.read_csv(D / "olist_sellers_dataset.csv")
    states = pd.read_csv(D / "브라질_주.csv", encoding="utf-8-sig")
    return o, i, p, r, pr, s, states


def main():
    o, i, p, r, pr, s, states = load()
    i2 = i.merge(o[["order_id", "구매일", "customer_unique_id", "customer_state"]], on="order_id").merge(
        pr[["product_id", "카테고리", "대분류"]], on="product_id")
    print("행 수: 주문", len(o), "주문상품", len(i), "결제", len(p), "리뷰", len(r), "상품", len(pr), "판매자", len(s))
    print("상품매출", round(i.price.sum(), 2), "배송비", round(i.freight_value.sum(), 2), "결제금액", round(p.payment_value.sum(), 2))
    n_orders = len(o)
    print("주문수", n_orders, "주문수(상품기준)", i.order_id.nunique(), "고객수", o.customer_unique_id.nunique())
    print("객단가(상품기준)", round(i.price.sum() / i.order_id.nunique(), 2))
    dv = o[o["배송상태"] != "미배송"]
    print("배송상태", o["배송상태"].value_counts().to_dict(), "지연율", round((dv["배송상태"] == "지연").mean() * 100, 2))
    print("평균 배송일수", round(dv["배송일수"].mean(), 2), "평균 평점", round(r.review_score.mean(), 4))
    rr = r.merge(o[["order_id", "배송상태"]], on="order_id")
    print("배송상태별 평점", rr.groupby("배송상태").review_score.mean().round(2).to_dict())
    print("주문상태", o["주문상태"].value_counts().to_dict())
    # 연도별 매출
    i2["연도"] = i2["구매일"].dt.year
    print("연도별 상품매출", i2.groupby("연도").price.sum().round(2).to_dict())
    print("대분류 매출 상위", i2.groupby("대분류").price.sum().sort_values(ascending=False).round(2).head(5).to_dict())
    print("카테고리 매출 상위", i2.groupby("카테고리").price.sum().sort_values(ascending=False).round(2).head(5).to_dict())
    rep = o.groupby("customer_unique_id").order_id.nunique()
    print("재구매 고객수", int((rep > 1).sum()), "재구매율", round((rep > 1).mean() * 100, 2))
    ost = o.merge(states, left_on="customer_state", right_on="state")
    print("지역별 주문수", ost.groupby("지역").order_id.count().to_dict())
    print("결제수단", p.payment_type.value_counts().to_dict())


if __name__ == "__main__":
    main()
