# db_config.py
import os
APP_ROOT = os.path.dirname(os.path.abspath(__file__)) # 获取 db_config.py 文件的目录

# 数据库连接配置
# db_config.py
DB_CONFIG = {
    'host': '43.143.60.119',
    'user': 'root',
    'password': 'Shenjiao666777;',
    'database': 'ShenJiao',
    'port': 3306,
    'auth_plugin': 'mysql_native_password' # ADD THIS LINE
}

# 获取原始 Word 文档的基础 URL
# 用于从Django服务器下载Word文件
WORD_FILE_BASE_URL = "http://43.143.60.119:7777"

EXTERNAL_HOSTNAME = "127.0.0.1"  # <<< YOUR ACTUAL PUBLIC IP or DOMAIN

# Flask 应用的目录配置 (这些目录将在运行Flask的Windows机器上)
# 确保这些目录存在或由应用创建

# 上传文件临时存储目录 (当从Django下载Word文件时，会先存到这里)
UPLOAD_FOLDER = os.path.join(APP_ROOT, 'uploads')

# 生成的审校清单 (.docx 文件) 的存放目录
GENERATED_DOCS_DIR = os.path.join(APP_ROOT, 'generated_proof_lists')

# 从Word文档解析出来的图片存放目录 (当前配置下，extractWordElement_web.py 不再提取图片，但保留此配置项以备将来使用)
IMAGE_OUTPUT_DIR_FLASK = os.path.join(APP_ROOT, 'parsed_word_images_flask')