import pandas as pd
import os
from ftplib import FTP, error_perm
from datetime import datetime

FTP_HOST = '27.0.15.132'
FTP_USER = 'served8af4'
FTP_PASS = 'RvyVKk9dA341'

EXCEL_FILE_PATH = r'E:\TestDownloadFile\FileDownload.xlsx'
BASE_DOWNLOAD_DIR = r'D:\DataDownload'
BASE_TARGET_DIR = '/domains/clouddrive15132ftp.superdata.vn/public_html/techzen/dhkt'

columns_to_process = [
    'materialVideoPath',
    'courseMaterial',
    'courseCoverImage',
    'courseDetailedImgPath1',
    'courseDetailedImgPath2',
    'courseDetailedImgPath3',
    'courseDetailedImgPath4',
    'courseDetailedImgPath5',
    '20240620',
    '20240822',
    '20240823'
]

if not os.path.exists(BASE_DOWNLOAD_DIR):
    os.makedirs(BASE_DOWNLOAD_DIR)

log_dir = os.path.join(BASE_DOWNLOAD_DIR, 'Log')
if not os.path.exists(log_dir):
    os.makedirs(log_dir)

log_file_path = os.path.join(log_dir, f'log_{datetime.now().strftime("%Y%m%d_%H%M%S")}.txt')

def log_message(message):
    with open(log_file_path, 'a', encoding='utf-8') as log_file:
        log_file.write(message + '\n')
    print(message)

log_message("Đang kết nối FTP...")

try:
    ftp = FTP(FTP_HOST)
    ftp.login(FTP_USER, FTP_PASS)
    log_message("Kết nối FTP thành công.")
except Exception as e:
    log_message(f"Lỗi kết nối FTP: {e}")
    exit(1)

def file_exists_on_ftp(ftp, path):
    try:
        ftp.size(path)
        return True
    except Exception:
        return False

df = pd.read_excel(EXCEL_FILE_PATH)

for index, row in df.iterrows():
    for col in columns_to_process:
        file_name = row.get(col)
        if pd.notna(file_name):
            file_name = str(file_name)

            if col.startswith('courseDetailedImgPath'):
                parent_folder = 'courseDetailedFiles'
                subfolder = col
                target_dir = os.path.join(BASE_TARGET_DIR, parent_folder, subfolder)
                download_subdir = os.path.join(BASE_DOWNLOAD_DIR, parent_folder, subfolder)

            elif col in ['20240620', '20240822', '20240823']:
                parent_folder = 'courseQuestionImage'
                subfolder = col
                target_dir = os.path.join(BASE_TARGET_DIR, parent_folder, subfolder)
                download_subdir = os.path.join(BASE_DOWNLOAD_DIR, parent_folder, subfolder)

            elif col in ['courseCoverImage']:
                parent_folder = 'courseFiles'
                subfolder = col
                target_dir = os.path.join(BASE_TARGET_DIR, parent_folder, subfolder)
                download_subdir = os.path.join(BASE_DOWNLOAD_DIR, parent_folder, subfolder)

            else:
                target_dir = os.path.join(BASE_TARGET_DIR, col)
                download_subdir = os.path.join(BASE_DOWNLOAD_DIR, col)

            if not os.path.exists(download_subdir):
                os.makedirs(download_subdir)

            source_file_path = os.path.join(target_dir, file_name).replace('\\', '/')
            local_file_path = os.path.join(download_subdir, file_name)

            log_message(f'Đang kiểm tra: {source_file_path}')

            if file_exists_on_ftp(ftp, source_file_path):
                log_message(f'Đang tải: {source_file_path}')
                try:
                    with open(local_file_path, 'wb') as local_file:
                        ftp.retrbinary(f'RETR {source_file_path}', local_file.write)
                    log_message(f'Tải về tệp: {local_file_path}')
                except error_perm as e:
                    log_message(f'Lỗi quyền truy cập tệp: {source_file_path} - {e}')
                except Exception as e:
                    log_message(f'Lỗi tải tệp: {source_file_path} - {e}')
            else:
                log_message(f'Tệp không tồn tại: {source_file_path}')

log_message("Hoàn thành tải xuống từ FTP.")

try:
    ftp.quit()
    log_message("Đóng kết nối FTP.")
except Exception as e:
    log_message(f'Lỗi khi đóng kết nối FTP: {e}')