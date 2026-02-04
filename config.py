# config.py
# 百度 AI 密钥配置
BAIDU_CONFIG = {
    'APP_ID': '122001991',
    'API_KEY': 'aE4bjyOR0B0JWQxhGvtMhTMh',
    'SECRET_KEY': 'PCMvnAFiL1gXg2y2Q3q2LYjSFu02fZwU'
}

# Poppler 路径配置
import os

# 动态获取项目根目录
PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
POPPLER_PATH = os.path.join(PROJECT_ROOT, 'poppler-25.12.0', 'Library', 'bin')