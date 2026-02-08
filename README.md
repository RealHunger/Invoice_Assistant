# 发票识别与管理系统

一个基于 Flask + 百度 OCR + SQLite 的发票识别和管理系统。支持批量上传、PDF 转图片、发票识别、数据导出。

## 本地开发快速开始

### 环境准备

- Python 3.8+
- Poppler（用于 PDF 转图片）：Windows 上项目已包含 `poppler-25.12.0`

### 安装依赖

```bash
pip install -r requirements.txt
```

### 配置百度 AI

编辑 `config.py`，填入你的百度 OCR 密钥：

```python
BAIDU_CONFIG = {
    'APP_ID': '你的APP_ID',
    'API_KEY': '你的API_KEY',
    'SECRET_KEY': '你的SECRET_KEY'
}
```

从 [百度智能云](https://cloud.baidu.com) 获取密钥。

### 运行应用

```bash
python app.py
```

访问 `http://localhost:5000`

## 功能

- **批量上传发票**：支持 PDF / JPG / PNG，自动识别
- **发票数据管理**：自动提取发票号、金额、商品名等信息
- **附件管理**：支持上传支付、订单截图等附件
- **数据导出**：导出为 Excel 和 ZIP 汇总包
- **撤销删除**：已删除的附件可恢复

## 目录结构

```
├── app.py                   # Flask 应用主文件
├── config.py                # 配置文件（密钥等）
├── requirements.txt         # Python 依赖
├── storage/                 # 上传文件存储
├── instance/                # SQLite 数据库（invoices_pro.db）
├── poppler-25.12.0/         # Poppler PDF 处理工具（Windows）
├── templates/               # HTML 模板
└── static/                  # 静态文件
```

## 笔记

- 数据库使用 SQLite，文件位置 `instance/invoices_pro.db`
- 所有上传的文件和图片保存在 `storage/` 目录
- Windows 环境已预装 Poppler；其它系统需独立安装

