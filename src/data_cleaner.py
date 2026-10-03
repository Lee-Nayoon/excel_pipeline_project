import glob
import os
import pandas as pd
import numpy as np

# 전국 지점별 엑셀 수집, 컬럼 표준화, 수량/단가 정제, 결측치 처리를 담당하는 클래스 생성
class DataCleaner:
    # [BEFORE] 하드코딩된 구분자 포함 방식
    # def __init__(self, input_dir: str = "data/inputs"):

    # [AFTER] OS별 표준 구분자를 준수하는 정석 방식
    def __init__(self, input_dir: str = os.path.join("data", "inputs")):
        self.input_dir = input_dir

        # 표준 컬럼명 매핑 (지점별로 상이한 컬럼을 대응시키는 과정)
        self.column_mapping = {
            '지점' : '지점명',
            '판매가' : '단가',
            '가격' : '단가'
        }

    # inputs/ 폴더 내의 모든 .xlsx 파일 경로를 패턴 검색으로 수집, 리스트 형태로 반환(명시)
    def collect_excel_files(self) -> list:
        pattern = os.path.join(self.input_dir, '*.xlsx')
        files = glob.glob(pattern)
        # 패턴과 일치하는 파일이 없을 시 FileNotFoundError 발생, 로그 기록
        if not files:
            raise FileNotFoundError(f"'{self.input_dir}' 폴더에 엑셀 파일이 존재하지 않습니다.")
        return files

    # 1. 컬럼명 양옆의 공백 제거 및 대표 컬럼명 표준화
    def sanitize_column_names(self, df: pd.DataFrame) -> pd.DataFrame:
        # 양옆 공백 제거
        df.columns = df.columns.str.strip()
        # 불일치 컬럼명 변환
        df = df.rename(columns = self.column_mapping)
        return df

    # 2. 수량/단가 컬럼의 텍스트(',', '개', '원') 제거 및 숫자형 변환
    def clean_numeric_column(self, series: pd.Series) -> pd.Series:
        cleaned = (
            series.astype(str)  # 데이터 정제 전 모든 요소를 문자형으로 변환
            .str.replace(',', '', regex = False)  # 천 단위 콤마 제거
            .str.replace('개', '', regex = False)  # 단위 '개' 제거
            .str.replace('원', '', regex =  False)  # 단위 '원' 제거
            .str.strip()  # 문자열 양옆의 불필요한 공백 제거
        )

        # 숫자형으로 변환하여 반환, 숫자로 변환이 불가한 문자('nan' 등)는 NaN으로 자동 처리
        return pd.to_numeric(cleaned, errors = 'coerce')

    # 개별 엑셀 파일 정제 파이프라인
    def process_single_file(self, file_path: str) -> pd.DataFrame:
        df = pd.read_excel(file_path)

        # 1. 컬럼명 표준화
        df = self.sanitize_column_names(df)

        # 2. 필수 컬럼 누락 검사 -> 누락 시 ValueError 발생, 로그 기록
        required_cols = ['지점명', '판매일자', '상품명', '수량', '단가']
        missing_cols = [col for col in required_cols if col not in df.columns]

        # 필수 컬럼이 없는 경우 예외 발생 -> 상위 실행부에서 해당 파일만 스킵
        if missing_cols:
            file_name = os.path.basename(file_path)
            raise ValueError(f"필수 컬럼 누락 {missing_cols} - 파일명: {file_name}")

        # 3. 수량 및 단가 정제
        df['수량'] = self.clean_numeric_column(df['수량'])
        df['단가'] = self.clean_numeric_column(df['단가'])

        # 단가 결측치는 동일 상품명의 단가 중 최빈값으로 보완, 동일 상품명의 모든 단가가 결측치일 경우 0으로 보완한 후 데이터 타입을 정수형으로 변환
        df['단가'] = (
            df.groupby('상품명')['단가']
            .transform(lambda x: x.fillna(x.mode()[0]) if not x.mode().empty else 0)
            .astype(int)
        )

        # 수량 결측치는 0으로 처리
        df['수량'] = df['수량'].fillna(0).astype(int)

        # 4. 파생 변수 계산: 총 매출액 (수량 * 단가)
        df['총매출액'] = df['수량'] * df['단가']

        return df[required_cols + ['총매출액']]

    # 전체 파일 수집 및 병합 파이프라인
    def run_pipeline(self, logger = None) -> pd.DataFrame:

        # 1. 파일 목록 수집 시 예외 처리
        try:
            file_paths = self.collect_excel_files()
        except FileNotFoundError as e:
            # 파일이 0개인 경우 에러 로그 기록 후 중단
            if logger:
                logger.error(f"[파이프라인 중단] {e}")
            # 에러를 그대로 상위(main.py)로 전달하여 중단시킴
            raise

        cleaned_dfs = []

        # 2. 개별 파일 정제
        for path in file_paths:
            try:
                cleaned_df = self.process_single_file(path)
                cleaned_dfs.append(cleaned_df)
            except ValueError as e:
                # 컬럼 누락 파일은 경고 로그 출력, 스킵한 후 다음 파일로 진행
                if logger:
                    logger.warning(f"[스킵] {e}")

        # 3. 정상 정제된 파일이 0개일 때
        if not cleaned_dfs:
            msg = "정제에 성공한 데이터가 단 1건도 없습니다. 원본 엑셀 파일을 확인해 주세요."
            if logger:
                logger.error(f"[파이프라인 중단] {msg}")
            raise ValueError(msg)

        # 모든 정상 정제된 데이터프레임을 하나로 합친 후 반환 
        return pd.concat(cleaned_dfs, ignore_index = True)

if __name__ == "__main__":
    # 단독 테스트 실행
    cleaner = DataCleaner(input_dir = os.path.join("data", "inputs"))
    result_df = cleaner.run_pipeline()
    print("전처리 완료 샘플 5건:")
    print(result_df.head())
    print(f"\n총 통합 행 수: {len(result_df)}개")