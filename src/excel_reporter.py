import os
import logging
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList

class ExcelReporter:
    """
    정제된 Pandas DataFrame 데이터를 활용하여
    경영진 보고서용 엑셀 파일 서식 적용 및 시각화 차트를 자동 생성하는 클래스
    """

    def __init__(self, output_dir: str = os.path.join("data", "outputs")):
        self.output_dir = output_dir
        os.makedirs(self.output_dir, exist_ok = True)

    def generate_report(self, df: pd.DataFrame, file_name: str = "sales_report.xlsx", logger: logging.Logger = None) -> str:
        output_path = os.path.join(self.output_dir, file_name)

        if df.empty:
            if logger:
                logger.warning("보고서 생성을 위한 데이터가 비어 있습니다.")
            raise ValueError("데이터프레임이 비어 있어 보고서를 생성할 수 없습니다.")

        # 1. openpyxl Workbook 생성
        wb = Workbook()
        ws_detail = wb.active
        ws_detail.title = "상세 판매 내역"

        # 2. 스타일 정의 (색상 패러디, 폰트, 테두리 등) 
        # "FFFFFF": 블랙, "1F4E78": 네이비
        header_font = Font(name = "맑은 고딕", size = 11, bold = True, color = "FFFFFF")
        header_fill = PatternFill(start_color = "1F4E78", fill_type = "solid")

        data_font = Font(name = "맑은 고딕", size = 10)
        title_font = Font(name = "맑은 고딕", size = 16, bold = True, color = "1F4E78")

        thin_border = Border(
            left = Side(style = "thin", color = "D9D9D9"),
            right = Side(style = "thin", color = "D9D9D9"),
            top = Side(style = "thin", color = "D9D9D9"),
            bottom = Side(style = "thin", color = "D9D9D9")
        )

        align_center = Alignment(horizontal = "center", vertical = "center")
        align_right = Alignment(horizontal = "right", vertical = "center")
        align_left = Alignment(horizontal = "left", vertical = "center")

        # 3. 상세 내역 시트 구성 (제목 배치)
        ws_detail.cell(row = 1, column = 1, value = "전국 지점별 판매 실적 상세 보고서").font = title_font
        ws_detail.row_dimensions[1].height = 35

        # 데이터를 3번째 행부터 작성하기 위해 2번째 행에 빈 행 추가
        ws_detail.append([])

        # DataFrame 데이터를 3번째 행부터 작성
        for r_idx, row in enumerate(dataframe_to_rows(df, index = False, header = True), start = 3):
            ws_detail.append(row)

        # 4. 헤더 및 데이터 행 스타일링
        # 헤더 행 (3번째 행)
        ws_detail.row_dimensions[3].height = 35
        for col_idx in range(1, len(df.columns) + 1):
            cell = ws_detail.cell(row = 3, column = col_idx)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = align_center
            cell.border = thin_border

        # 데이터 행 (4번째 행부터)
        for row_idx in range(4, 4 + len(df)):
            ws_detail.row_dimensions[row_idx].height = 20
            for col_idx, col_name in enumerate(df.columns, start = 1):
                cell = ws_detail.cell(row = row_idx, column = col_idx)
                cell.font = data_font
                cell.border = thin_border

                # 컬럼 유형별 정렬 및 숫자 서식 적용
                if "가" in col_name or "액" in col_name or "금" in col_name:
                    cell.alignment = align_right
                    cell.number_format = "#,##0"  # 천 단위 콤마
                elif "수량" in col_name:
                    cell.alignment = align_right
                    cell.number_format = "#,##0"
                elif "일자" in col_name or "날짜" in col_name:
                    cell.alignment = align_center
                else:
                    cell.alignment = align_left

        # 컬럼 너비 자동 조절
        for col in ws_detail.columns:
            max_len = 0
            # col[0].column은 숫자이므로 이를 알파벳으로 변환
            col_letter = get_column_letter(col[0].column)

            for cell in col:
                # 셀이 빈 값인 경우 빈 문자열로 대체해 오류 발생 없앰, 문자열 길이 측정 가능하도록 타입 변환
                val_str = str(cell.value or '')

                # 한글 글자 수 판별을 고려한 너비 계산
                # ord(c) 문자 c의 유니코드 정수값 반환, ASCII 코드 범위를 벗어나는 문자의 개수 파악
                korean_len = sum(1 for c in val_str if ord(c) > 128)

                # 전체 글자 수에 ASCII 범위 밖의 글자 수만큼 더해 가중치 계산
                length = len(val_str) + korean_len
                if length > max_len:
                    max_len = length

            # 여유분 3만큼 추가, 최소 너비 12로 지정
            ws_detail.column_dimensions[col_letter].width = max(max_len + 3, 12)

        # 5. 요약 집계 시트 (지점별 매출 요약 & 막대차트 생성)
        ws_summary = wb.create_sheet(title = "지점별 매출 요약")
        ws_summary.cell(row = 1, column = 1, value = "지점별 총 매출액 요약").font = title_font
        ws_summary.row_dimensions[1].height = 35

        # 지점별 매출 집계 데이터 생성
        summary_df = df.groupby("지점명", as_index = False)['총매출액'].sum().sort_values(by = "총매출액", ascending = False)

        # 요약 헤더 작성
        ws_summary.cell(row = 3, column = 1, value = "지점명")
        ws_summary.cell(row = 3, column = 2, value = "총매출액")

        for col_idx in [1, 2]:
            cell = ws_summary.cell(row = 3, column = col_idx)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = align_center
            cell.border = thin_border

        # 요약 데이터 작성 (데이터 입력 시작 위치: 4번째 행)
        for r_idx, row in enumerate(summary_df.itertuples(index = False), start = 4):
            c1 = ws_summary.cell(row = r_idx, column = 1, value = row.지점명)
            c2 = ws_summary.cell(row = r_idx, column = 2, value = row.총매출액)

            c1.font = data_font
            c1.alignment = align_center
            c1.border = thin_border

            c2.font = data_font
            c2.alignment = align_center
            c2.number_format = "#,##0"
            c2.border = thin_border

        ws_summary.column_dimensions["A"].width = 15
        ws_summary.column_dimensions["B"].width = 20

        # 6. openpyxl BarChart 생성 및 시트 부착
        chart = BarChart()
        chart.type = "col"
        chart.style = 10
        chart.title = "지점별 총매출액 비교"
        chart.y_axis.title = "매출액 (원)"
        chart.x_axis.title = "지점명"

        data_ref = Reference(ws_summary, min_col = 2, min_row = 3, max_row = 3 + len(summary_df))
        cats_ref = Reference(ws_summary, min_col = 1, min_row = 4, max_row = 3 + len(summary_df))

        chart.add_data(data_ref, titles_from_data = True)
        chart.set_categories(cats_ref)
        chart.legend = None  # 범례 제거

        # 막대 위에 지점명 및 매출액 데이터 라벨 표시
        chart.dataLabels = DataLabelList()
        chart.dataLabels.showCatName = True
        chart.dataLabels.showVal = True

        # 차트 위치 설정
        ws_summary.add_chart(chart, "D3")

        # 엑셀 파일 저장
        wb.save(output_path)

        if logger:
            logger.info(f"엑셀 보고서 생성 완료: {output_path}")

        return output_path