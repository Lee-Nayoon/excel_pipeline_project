import os
from src import DataCleaner, ExcelReporter
from utils import setup_logger

def main():

    # 1. 로거 초기화 (logs/app.log 저장 및 콘솔 출력 준비)
    logger = setup_logger()

    logger.info("==================================================")
    logger.info("전국 지점별 판매 데이터 자동 정제 파이프라인 시작")
    logger.info("==================================================")

    # 2. DataCleaner 인스턴스 생성 (실제 엑셀 위치: data/inputs)
    input_dir = os.path.join("data", "inputs")
    cleaner = DataCleaner(input_dir)

    try:
        # 3. 파이프라인 실행
        # 폴더 내 엑셀 미존재 시: FileNotFoundError 발생 후 중단
        # 필수 컬럼 누락 파일 존재 시: ValueError 발생 후 해당 파일만 스킵, 경고 로그 발생
        merged_df = cleaner.run_pipeline(logger = logger)

        logger.info(
            f"데이터 정제 성공! 총 {len(merged_df):,}건의 데이터가 통합되었습니다."
        )

        # 4. 정제된 데이터프레임 요약 정보 출력
        print("\n" + "=" * 50)
        print("[ 정제된 통합 데이터 미리보기 (상위 5건) ]")
        print("=" * 50)
        print(merged_df.head())
        print("\n[ 데이터 요약 정보 ]")
        print(merged_df.info())
        print("=" * 50 + "\n")

        # 5. 정제된 중간 결과물들을 CSV로 저장 (파일 위치: data/outputs)
        output_dir = os.path.join("data", "outputs")
        os.makedirs(output_dir, exist_ok=True)

        output_csv_path = os.path.join(output_dir, "cleaned_merged_data.csv")
        merged_df.to_csv(output_csv_path, index = False, encoding = "utf-8-sig")
        logger.info(f"정제 데이터 CSV 저장 완료: {output_csv_path}")

        # 6. 경영진 보고서용 엑셀 파일 및 차트 생성 (ExcelReporter 연결)
        logger.info("경영진 보고서용 엑셀 생성 프로세스 시작...")
        reporter = ExcelReporter(output_dir = output_dir)

        output_excel_path = reporter.generate_report(
            df = merged_df,
            file_name = "전국_지점별_판매_실적_보고서.xlsx",
            logger = logger
        )

        logger.info("==================================================")
        logger.info(f"파이프라인 모든 과정 성공적 완료!")
        logger.info(f"최종 엑셀 보고서 경로: {output_excel_path}")
        logger.info("==================================================")


    except FileNotFoundError as e:
        logger.error(f"[파이프라인 실행 중단] {e}")
    except ValueError as e:
        logger.error(f"[파이프라인 실행 중단] {e}")
    except Exception as e:
        logger.critical(f"[시스템 예외 발생] {e}")

# 다른 파일에서 import main 했을 시 전체 로직이 작동되지 않게 방지
if __name__ == "__main__":
    main()