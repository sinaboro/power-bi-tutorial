-- ============================================================================
-- 심화 10장 (선택) - Olist CSV 를 PostgreSQL 테이블로 불러오기
-- 컨테이너를 처음 시작할 때 한 번만 자동으로 실행된다 (docker-entrypoint-initdb.d).
-- CSV 는 docker-compose.yml 에서 /olist 폴더로 연결해 두었다.
-- ============================================================================

-- 주문 (주문 하나 = 한 줄)
CREATE TABLE orders (
    order_id                      text PRIMARY KEY,
    customer_id                   text,
    order_status                  text,
    order_purchase_timestamp      timestamp,
    order_approved_at             timestamp,
    order_delivered_carrier_date  timestamp,
    order_delivered_customer_date timestamp,
    order_estimated_delivery_date timestamp
);

-- 주문상품 (주문 안의 상품 하나 = 한 줄)
CREATE TABLE order_items (
    order_id            text,
    order_item_id       int,
    product_id          text,
    seller_id           text,
    shipping_limit_date timestamp,
    price               numeric(10, 2),
    freight_value       numeric(10, 2),
    PRIMARY KEY (order_id, order_item_id)
);

-- 결제
CREATE TABLE order_payments (
    order_id             text,
    payment_sequential   int,
    payment_type         text,
    payment_installments int,
    payment_value        numeric(10, 2)
);

-- 리뷰 (리뷰 글 안에 줄 바꿈이 있어도 CSV 형식이면 한 값으로 읽는다)
CREATE TABLE order_reviews (
    review_id               text,
    order_id                text,
    review_score            int,
    review_comment_title    text,
    review_comment_message  text,
    review_creation_date    timestamp,
    review_answer_timestamp timestamp
);

-- 고객 (주문할 때의 고객 정보)
CREATE TABLE customers (
    customer_id              text PRIMARY KEY,
    customer_unique_id       text,
    customer_zip_code_prefix text,
    customer_city            text,
    customer_state           text
);

-- 상품
CREATE TABLE products (
    product_id                 text PRIMARY KEY,
    product_category_name      text,
    product_name_lenght        int,
    product_description_lenght int,
    product_photos_qty         int,
    product_weight_g           int,
    product_length_cm          int,
    product_height_cm          int,
    product_width_cm           int
);

-- 판매자
CREATE TABLE sellers (
    seller_id              text PRIMARY KEY,
    seller_zip_code_prefix text,
    seller_city            text,
    seller_state           text
);

-- 카테고리 이름 번역 (포르투갈어 → 영어)
CREATE TABLE category_translation (
    product_category_name         text PRIMARY KEY,
    product_category_name_english text
);

-- CSV 불러오기: 첫 줄은 열 이름(HEADER), 따옴표로 감싼 값 안의 쉼표·줄 바꿈은 값의 일부
COPY orders               FROM '/olist/olist_orders_dataset.csv'               WITH (FORMAT csv, HEADER true);
COPY order_items          FROM '/olist/olist_order_items_dataset.csv'          WITH (FORMAT csv, HEADER true);
COPY order_payments       FROM '/olist/olist_order_payments_dataset.csv'       WITH (FORMAT csv, HEADER true);
COPY order_reviews        FROM '/olist/olist_order_reviews_dataset.csv'        WITH (FORMAT csv, HEADER true);
COPY customers            FROM '/olist/olist_customers_dataset.csv'            WITH (FORMAT csv, HEADER true);
COPY products             FROM '/olist/olist_products_dataset.csv'             WITH (FORMAT csv, HEADER true);
COPY sellers              FROM '/olist/olist_sellers_dataset.csv'              WITH (FORMAT csv, HEADER true);
COPY category_translation FROM '/olist/product_category_name_translation.csv'  WITH (FORMAT csv, HEADER true);

-- 보고서용 뷰: 월별 주문 요약 (Power BI 에서 "SQL 문" 대신 뷰를 골라 불러오는 연습용)
CREATE VIEW monthly_orders AS
SELECT date_trunc('month', o.order_purchase_timestamp)::date AS month_start,
       count(*)                                            AS orders,
       count(*) FILTER (WHERE o.order_status = 'delivered') AS delivered_orders
FROM orders o
GROUP BY 1
ORDER BY 1;
