import argparse

from datetime import datetime,timedelta
import time

from to_pdf import ConvertPdf
from sftp_connect import Sftp
import sys

if __name__ == '__main__':
    args = argparse.ArgumentParser()
    args.add_argument("--env", type=str, default="prod",
                      help="sftp환경, 운영의 경우 prod")
    args.add_argument("--d", type=str, default=datetime.today().strftime("%Y%m%d"),
                      help="수행할 날짜 YYYYMMDD 형식")
    args.add_argument("--mode", type=str, default="all",
                      help="소스코드 모드, all:모든동작, conv_test:변환테스트")
    config = args.parse_args()

    if config.mode == "all":
        local_from_folder = "C:/convert_to_pdf_4/from_folder"
        local_to_folder = "C:/convert_to_pdf_4/PDF"
        error_log_path = "C:/convert_to_pdf_4/log_folder/log_" + config.d + ".csv"

        start = time.time()

        # [sftp] remote -> local 파일복사
        sftp = Sftp(config.env, config.d, local_from_folder, local_to_folder)
        sftp.get_file_from_sftp()

        # [변환]
        converter = ConvertPdf(local_from_folder, local_to_folder, config.d)
        converter.to_csv_error_file(error_log_path)

        # [sftp] local -> remote 파일복사
        sftp.put_file_to_sftp()

        print(str(timedelta(seconds = time.time()-start)).split(".")[0])
    elif config.mode == "conv_test":
        local_from_folder = "C:/convert_to_pdf_4/test_from_folder"
        local_to_folder = "C:/convert_to_pdf_4/test_PDF"
        error_log_path = "C:/convert_to_pdf_4/test_log_folder/log_" + config.d + ".csv"

        # [변환]
        converter = ConvertPdf(local_from_folder, local_to_folder, config.d)
        converter.to_csv_error_file(error_log_path)