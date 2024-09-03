import pandas as pd
import os

# Thông tin FTP (bị comment lại cho test dưới local)
# FTP_HOST = '27.0.15.132'
# FTP_USER = 'served8af4'
# FTP_PASS = 'RvyVKk9dA341'

EXCEL_FILE_PATH = r'E:\TestDownloadFile\FileDownload.xlsx'
BASE_DOWNLOAD_DIR = r'E:\TestDownloadFile\Local'
BASE_TARGET_DIR = r'E:\TestDownloadFile'

columns_to_process = [
    'materialVideoPath',
    'courseDetailed',
    'courseMaterial',
    'courseDetailedImgPath1',
    'courseDetailedImgPath2',
    'courseDetailedImgPath3',
    'courseDetailedImgPath4',
    'courseDetailedImgPath5',
    '20240821',
    '20240822',
    '20240823'
]

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

            elif col in ['20240821', '20240822', '20240823']:
                parent_folder = 'courseQuestionImage'
                subfolder = col
                target_dir = os.path.join(BASE_TARGET_DIR, parent_folder, subfolder)
                download_subdir = os.path.join(BASE_DOWNLOAD_DIR, parent_folder, subfolder)

            else:
                target_dir = os.path.join(BASE_TARGET_DIR, col)
                download_subdir = os.path.join(BASE_DOWNLOAD_DIR, col)

            if not os.path.exists(download_subdir):
                os.makedirs(download_subdir)

            source_file_path = os.path.join(target_dir, file_name)
            local_file_path = os.path.join(download_subdir, file_name)

            if os.path.exists(source_file_path):
                with open(source_file_path, 'rb') as source_file, open(local_file_path, 'wb') as local_file:
                    local_file.write(source_file.read())
                print(f'Tải về tệp: {local_file_path}')
            else:
                print(f'Không tìm thấy tệp: {source_file_path}')

print("Hoàn thành kiểm tra dưới local.")
