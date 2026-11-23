r"""make_olist.py - 심화 과정 실습 데이터(data/olist) 준비 (교안 제작용)

    python tools\make_olist.py <Olist CSV 가 있는 폴더>
    예) python tools\make_olist.py C:\works\numpy\ecommerce\00_data

하는 일
  1. Kaggle "Brazilian E-Commerce Public Dataset by Olist" 의 CSV 7개와 카테고리 번역 파일을
     data/olist 로 그대로 복사한다. (위치 정보 olist_geolocation_dataset.csv 는 61MB 로 크고
     이 과정에서 쓰지 않아 뺀다)
  2. 학생이 Power Query 병합을 연습할 한글 매핑 파일 두 개를 만든다.
     - 카테고리_한글.csv : 영어 카테고리 이름 → 한글 카테고리, 대분류
     - 브라질_주.csv     : 주 약어(SP 등) → 한글 주 이름, 지역(남동부 등)
원본 데이터 출처: https://www.kaggle.com/datasets/olistbr/brazilian-ecommerce (CC BY-NC-SA 4.0)
"""
import csv
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "data" / "olist"

FILES = ["olist_orders_dataset.csv", "olist_order_items_dataset.csv", "olist_order_payments_dataset.csv",
         "olist_order_reviews_dataset.csv", "olist_customers_dataset.csv", "olist_products_dataset.csv",
         "olist_sellers_dataset.csv", "product_category_name_translation.csv"]

# 영어 카테고리 → (한글 카테고리, 대분류)
CATEGORY = {
    "health_beauty": ("건강·뷰티", "뷰티·건강"), "perfumery": ("향수", "뷰티·건강"),
    "diapers_and_hygiene": ("기저귀·위생용품", "뷰티·건강"),
    "computers_accessories": ("컴퓨터 주변기기", "전자·컴퓨터"), "computers": ("컴퓨터", "전자·컴퓨터"),
    "telephony": ("휴대폰", "전자·컴퓨터"), "fixed_telephony": ("유선 전화기", "전자·컴퓨터"),
    "electronics": ("전자제품", "전자·컴퓨터"), "tablets_printing_image": ("태블릿·프린터", "전자·컴퓨터"),
    "audio": ("오디오", "전자·컴퓨터"), "consoles_games": ("게임기·게임", "전자·컴퓨터"),
    "cine_photo": ("카메라·영상", "전자·컴퓨터"), "signaling_and_security": ("신호·보안장비", "전자·컴퓨터"),
    "security_and_services": ("보안 서비스", "전자·컴퓨터"),
    "bed_bath_table": ("침구·욕실·식탁", "가구·인테리어"), "furniture_decor": ("가구·장식", "가구·인테리어"),
    "office_furniture": ("사무용 가구", "가구·인테리어"), "furniture_living_room": ("거실 가구", "가구·인테리어"),
    "furniture_bedroom": ("침실 가구", "가구·인테리어"),
    "furniture_mattress_and_upholstery": ("매트리스·소파", "가구·인테리어"),
    "kitchen_dining_laundry_garden_furniture": ("주방·세탁·정원 가구", "가구·인테리어"),
    "home_confort": ("홈 편의용품", "가구·인테리어"), "home_comfort_2": ("홈 편의용품2", "가구·인테리어"),
    "housewares": ("생활용품", "생활·주방"), "small_appliances": ("소형 가전", "생활·주방"),
    "home_appliances": ("생활 가전", "생활·주방"), "home_appliances_2": ("생활 가전2", "생활·주방"),
    "small_appliances_home_oven_and_coffee": ("오븐·커피 가전", "생활·주방"), "air_conditioning": ("냉난방기", "생활·주방"),
    "la_cuisine": ("주방용품", "생활·주방"), "home_construction": ("집수리 자재", "생활·주방"),
    "watches_gifts": ("시계·선물", "패션"), "fashion_bags_accessories": ("가방·액세서리", "패션"),
    "fashion_shoes": ("신발", "패션"), "fashion_male_clothing": ("남성 의류", "패션"),
    "fashio_female_clothing": ("여성 의류", "패션"), "fashion_underwear_beach": ("속옷·수영복", "패션"),
    "fashion_sport": ("스포츠 의류", "패션"), "fashion_childrens_clothes": ("아동 의류", "패션"),
    "luggage_accessories": ("여행 가방", "패션"),
    "sports_leisure": ("스포츠·레저", "레저·취미"), "toys": ("장난감", "유아·완구"), "baby": ("유아용품", "유아·완구"),
    "cool_stuff": ("아이디어 상품", "레저·취미"), "musical_instruments": ("악기", "레저·취미"),
    "party_supplies": ("파티용품", "레저·취미"), "christmas_supplies": ("크리스마스 용품", "레저·취미"),
    "art": ("미술품", "레저·취미"), "arts_and_craftmanship": ("공예", "레저·취미"), "pet_shop": ("반려동물 용품", "레저·취미"),
    "flowers": ("꽃", "레저·취미"),
    "books_general_interest": ("일반 도서", "도서·음반·문구"), "books_technical": ("전문 도서", "도서·음반·문구"),
    "books_imported": ("수입 도서", "도서·음반·문구"), "stationery": ("문구", "도서·음반·문구"),
    "music": ("음악", "도서·음반·문구"), "cds_dvds_musicals": ("CD·DVD·뮤지컬", "도서·음반·문구"),
    "dvds_blu_ray": ("DVD·블루레이", "도서·음반·문구"),
    "food_drink": ("식품·음료", "식품"), "food": ("식품", "식품"), "drinks": ("음료", "식품"),
    "garden_tools": ("정원 도구", "공구·산업"), "construction_tools_construction": ("건설 공구", "공구·산업"),
    "costruction_tools_garden": ("정원 공구", "공구·산업"), "costruction_tools_tools": ("수공구", "공구·산업"),
    "construction_tools_lights": ("조명 공구", "공구·산업"), "construction_tools_safety": ("안전 장비", "공구·산업"),
    "agro_industry_and_commerce": ("농업·산업", "공구·산업"),
    "industry_commerce_and_business": ("산업·비즈니스", "공구·산업"),
    "auto": ("자동차용품", "생활·주방"), "market_place": ("마켓플레이스", "기타"),
}

# 주 약어 → (한글 주 이름, 지역)
STATE = {
    "AC": ("아크리", "북부"), "AM": ("아마조나스", "북부"), "AP": ("아마파", "북부"), "PA": ("파라", "북부"),
    "RO": ("혼도니아", "북부"), "RR": ("호라이마", "북부"), "TO": ("토칸칭스", "북부"),
    "AL": ("알라고아스", "북동부"), "BA": ("바이아", "북동부"), "CE": ("세아라", "북동부"), "MA": ("마라냥", "북동부"),
    "PB": ("파라이바", "북동부"), "PE": ("페르남부쿠", "북동부"), "PI": ("피아우이", "북동부"),
    "RN": ("히우그란지두노르치", "북동부"), "SE": ("세르지피", "북동부"),
    "DF": ("연방직할구(브라질리아)", "중서부"), "GO": ("고이아스", "중서부"), "MS": ("마투그로수두술", "중서부"),
    "MT": ("마투그로수", "중서부"),
    "ES": ("이스피리투산투", "남동부"), "MG": ("미나스제라이스", "남동부"), "RJ": ("리우데자네이루", "남동부"),
    "SP": ("상파울루", "남동부"),
    "PR": ("파라나", "남부"), "RS": ("히우그란지두술", "남부"), "SC": ("산타카타리나", "남부"),
}


def main(src):
    OUT.mkdir(parents=True, exist_ok=True)
    for f in FILES:
        shutil.copy(Path(src) / f, OUT / f)                # 원본 그대로 복사 (내용을 바꾸지 않는다)
        print("  복사:", f)
    # 번역 파일에 있는 영어 이름이 모두 한글 표에 있는지 확인
    with open(OUT / "product_category_name_translation.csv", encoding="utf-8-sig") as fp:
        eng = [r["product_category_name_english"] for r in csv.DictReader(fp)]
    missing = sorted(set(eng) - set(CATEGORY))
    assert not missing, f"한글 이름이 없는 카테고리: {missing}"
    with open(OUT / "카테고리_한글.csv", "w", encoding="utf-8-sig", newline="") as fp:
        w = csv.writer(fp)
        w.writerow(["category_english", "카테고리", "대분류"])
        for e in eng:
            w.writerow([e, *CATEGORY[e]])
    with open(OUT / "브라질_주.csv", "w", encoding="utf-8-sig", newline="") as fp:
        w = csv.writer(fp)
        w.writerow(["state", "주이름", "지역"])
        for k in sorted(STATE):
            w.writerow([k, *STATE[k]])
    print("  만듦: 카테고리_한글.csv (%d개), 브라질_주.csv (%d개)" % (len(eng), len(STATE)))


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else r"C:\works\numpy\ecommerce\00_data")
