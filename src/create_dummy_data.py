import os
import random
import pandas as pd
import numpy as np
from datetime import datetime, timedelta

# "data"라는 이름의 폴더 생성, 이미 존재한다면 에러 없이 그대로 진행
os.makedirs("excel_pipeline_project/data/inputs", exist_ok = True)

# 기본 설정
products = [
    ("모니터", 350000),
    ("키보드", 85000),
    ("마우스", 30000),
    ("노트북", 1200000),
    ("헤드셋", 150000)
]

start_date = datetime(2026, 9, 1)

# 최근 30일 내의 임의 날짜 생성
def generate_random_dates(count):
    return [(start_date + timedelta(days=random.randint(0, 30))).strftime('%Y-%m-%d') for _ in range(count)]

# (상품명, 단가) 튜플 쌍을 한 번씩만 추출
def get_random_products(rows):
    return [random.choice(products) for _ in range(rows)]

# 1. 서울점(100건): 공백이 들어간 컬럼명 + 정상 수치 데이터
# 판매일자: generate_random_dates 함수를 이용해 랜덤으로 생성
# 수량: 1~20 중 랜덤으로 생성
rows_seoul = 100
items_seoul = get_random_products(rows_seoul)

df_seoul = pd.DataFrame({
    ' 지점명 ' : ['서울점'] * rows_seoul,
    '판매일자' : generate_random_dates(rows_seoul),
    ' 상품명 ' : [item[0] for item in items_seoul],
    '수량' : [random.randint(1, 20) for _ in range(rows_seoul)],
    '단가' : [item[1] for item in items_seoul]
})

# 생성된 데이터프레임을 엑셀 파일로 추출
df_seoul.to_excel("excel_pipeline_project/data/inputs/branch_seoul.xlsx", index = False)


# 2. 부산점(80건): 수량/단가에 숫자, 쉼표, '개', '원' 섞임
rows_busan = 80
items_busan = get_random_products(rows_busan)
# 수량과 단가 형식을 세 가지로 나눠서 랜덤으로 한 가지 형식을 고르게 할 예정
qty_formats = [lambda q: f"{q:,}", lambda q: f"{q}개", lambda q: q]
price_formats = [lambda p: f"{p:,}원", lambda p: f"{p}원", lambda p: p]

df_busan = pd.DataFrame({
    '지점명' : ['부산점'] * rows_busan,
    '판매일자' : generate_random_dates(rows_busan),
    '상품명' : [item[0] for item in items_busan],
    '수량' : [random.choice(qty_formats)(random.randint(1, 15)) for _ in range(rows_busan)],
    '단가' : [random.choice(price_formats)(item[1]) for item in items_busan]
})

df_busan.to_excel("excel_pipeline_project/data/inputs/branch_busan.xlsx", index = False)


# 3. 인천점(70건): 무작위 결측값(NaN) 15% 비율로 발생
rows_incheon = 70
items_incheon = get_random_products(rows_incheon)
# 수량/단가 컬럼에 무작위 결측값 발생
quantities = [random.randint(1, 10) if random.random() > 0.15 else np.nan for _ in range(rows_incheon)]
prices = [item[1] if random.random() > 0.15 else np.nan for item in items_incheon]

df_incheon = pd.DataFrame({
    '지점명' : ['인천점'] * rows_incheon,
    '판매일자' : generate_random_dates(rows_incheon),
    '상품명' : [item[0] for item in items_incheon],
    '수량' : quantities,
    '단가' : prices
})

df_incheon.to_excel("excel_pipeline_project/data/inputs/branch_incheon.xlsx", index = False)


# 4. 대구점(60건): 다른 컬럼명 ('지점', '판매가') 사용
rows_daegu = 60
items_daegu = get_random_products(rows_daegu)

df_daegu = pd.DataFrame({
    '지점' : ['대구점'] * rows_daegu,
    '판매일자' : generate_random_dates(rows_daegu),
    '상품명' : [item[0] for item in items_daegu],
    '수량' : [random.randint(1, 12) for _ in range(rows_daegu)],
    '판매가' : [item[1] for item in items_daegu]
})

df_daegu.to_excel("excel_pipeline_project/data/inputs/branch_daegu.xlsx", index = False)


# 5. 울산점(50건): 필수 컬럼 누락 -> 에러 감지용
rows_ulsan = 50
items_ulsan = get_random_products(rows_ulsan)

df_ulsan = pd.DataFrame({
    '지점명' : ['울산점'] * rows_ulsan,
    '판매일자' : generate_random_dates(rows_ulsan),
    '상품명' : [item[0] for item in items_ulsan]
})

df_ulsan.to_excel("excel_pipeline_project/data/inputs/branch_ulsan.xlsx", index = False)

print("🎉 총 360여 건의 대용량 가상 엑셀 데이터 생성이 완료되었습니다!")
