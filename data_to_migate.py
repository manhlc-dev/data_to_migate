import pandas as pd
from ftplib import FTP
import os

# Thông tin FTP
FTP_HOST = '27.0.15.132'
FTP_USER = 'served8af4'
FTP_PASS = 'RvyVKk9dA341'

# Đường dẫn đến tệp Excel chứa danh sách tên tệp cần tải
EXCEL_FILE_PATH = r'D:\StudyArt DHKT\Tài liệu\Quản trị học.xlsx'

# Thư mục tải về
DOWNLOAD_DIR = r'D:\StudyArt DHKT\Tài liệu\Backup File\courseMaterial'

# Thư mục gốc trên máy chủ FTP
BASE_TARGET_DIR = '/domains/clouddrive15132ftp.superdata.vn/public_html/techzen/dhkt'

# Tạo thư mục tải nếu chưa tồn tại
if not os.path.exists(DOWNLOAD_DIR):
    os.makedirs(DOWNLOAD_DIR)

# Kết nối đến máy chủ FTP
ftp = FTP(FTP_HOST)
ftp.login(FTP_USER, FTP_PASS)

try:
    # Đọc danh sách tệp từ tệp Excel
    df = pd.read_excel(EXCEL_FILE_PATH)

    # Tải từng tệp từ danh sách
    for index, row in df.iterrows():
        file_name = str(row['materialpath'])  # Chuyển đổi tên tệp thành chuỗi
        subdir = str(row['category'])  # Chuyển đổi tên thư mục thành chuỗi
        
        # Xác định thư mục mục tiêu dựa trên 'category'
        if subdir == 'materialVideoPath':
            target_dir = os.path.join(BASE_TARGET_DIR, 'materialVideoPath')
        elif subdir == 'courseMaterial':
            target_dir = os.path.join(BASE_TARGET_DIR, 'courseMaterial')
        elif subdir == 'courseQuestionImage':
            target_dir = os.path.join(BASE_TARGET_DIR, 'courseQuestionImage')
        elif subdir == 'courseDetailed':
            target_dir = os.path.join(BASE_TARGET_DIR, 'courseDetailed')
        else:
            print(f'Bỏ qua dòng không xác định: {subdir}')
            continue

        # Di chuyển đến thư mục mục tiêu trên máy chủ FTP
        ftp.cwd(target_dir)
        print(f'Đang ở thư mục: {ftp.pwd()}')

        local_file_path = os.path.join(DOWNLOAD_DIR, file_name)
        
        # Tải tệp từ máy chủ FTP
        with open(local_file_path, 'wb') as local_file:
            ftp.retrbinary(f'RETR {file_name}', local_file.write)
            print(f'Tải về tệp: {local_file_path}')

finally:
    # Đóng kết nối FTP
    ftp.quit()
