import logging
import os
from datetime import datetime

# 로그 폴더 생성 및 콘솔/파일 로거 설정
def setup_logger(log_dir: str = "logs") -> logging.Logger:

    # 1. logs/ 폴더 없으면 자동 생성
    os.makedirs(log_dir, exist_ok = True)

    # 2. 로거 인스턴스 생성 (이름: 'DataPipeline')
    logger = logging.getLogger("DataPipeline")
    # INFO 레벨 이상 출력
    logger.setLevel(logging.INFO)

    # 이미 핸들러가 설정되어 있다면 중복 방지 (로그 2번 출력 현상 방지)
    if logger.handlers:
        return logger

    # 3. 로그 포맷 설정 (시간 [로그레벨] 메시지)
    formatter = logging.Formatter(
        "[%(asctime)s] [%(levelname)s] %(message)s", datefmt = "%Y-%m-%d %H:%M:%S"
    )

    # 4. 콘솔(터미널) 실시간 출력용 핸들러
    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)

    # 5. 파일 저장용 핸들러 (logs/app.log에 한글 깨짐 없이 저장)
    log_file_path = os.path.join(log_dir, "app.log")
    file_handler = logging.FileHandler(log_file_path, mode = "a", encoding = "utf-8")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    return logger