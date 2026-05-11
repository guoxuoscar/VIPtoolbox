# 唯品上新工具箱

电商运营辅助工具集 — 基于 PySide6 的 Windows 桌面应用。

## 功能

| 功能 | 说明 |
|------|------|
| **商品图片下载** | 通过 Playwright 浏览器自动化，下载淘宝商品主图/SKU图/详情图/视频 |
| **一键抠图** | 基于 rembg + u2net.onnx 自动抠透明底图 |
| **图片压缩** | JPEG/PNG 批量压缩到指定大小 |
| **图片重命名** | 正则/序号批量重命名 |
| **PDF 编辑** | 拆分、合并、提取页面 |
| **ERP 商品资料导出** | 商品信息 → ERP条码对照表/批量新增商品资料/定价导入模板 |
| **OCR 识别** | 本地 PaddleOCR / 云端百度OCR / PaddleX OCR |
| **尺码映射** | 尺码标准化别名映射 |

## 安装

### 1. 环境要求

- Windows 10/11
- Python 3.9+
- Google Chrome（用于商品图片下载）

### 2. 安装依赖

```bash
cd 项目目录
python -m pip install -r requirements.txt
python -m playwright install chromium
```

### 3. 下载模型文件（必看）

以下文件**太大未包含在仓库**中，需手动下载放到对应位置：

#### 抠图模型（~170MB）
下载 [u2net.onnx](https://github.com/danielgatis/rembg/releases/download/v0.0.0/u2net.onnx) 放到项目根目录

#### OCR 引擎（可选）
如需要本地 OCR 功能，下载 PaddleOCR-json 放到 `toolbox/ocr_engine/`：
- [PaddleOCR-json v1.4.1](https://github.com/foyoux/paddleocr-json/releases)

### 4. 配置 OCR（可选）

如需使用云端 OCR（百度/PaddleX），创建 `toolbox/ocr_engine/baidu_api.json`：

```json
{
    "paddlex_api_url": "你的接口地址",
    "paddlex_token": "你的token",
    "cloud_provider": "paddlex"
}
```

> 该文件已在 `.gitignore` 中，不会提交到仓库。

## 运行

```bash
python main.py
```

## 目录结构

```
├── main.py                       # 主入口
├── requirements.txt              # Python 依赖
├── scripts/
│   └── init_github.py            # GitHub 仓库一键初始化
└── toolbox/
    ├── core/                     # 核心逻辑
    │   ├── browser.py            # Playwright 浏览器自动化
    │   ├── erp_product.py        # ERP 商品资料导出
    │   ├── ocr.py                # OCR 识别引擎
    │   ├── pdf_edit_core.py      # PDF 编辑
    │   ├── size_mapping.py       # 尺码映射
    │   └── utils.py              # 工具函数
    ├── ui/                       # 界面
    │   ├── main_window.pyw       # 主窗口
    │   ├── compress_page.py      # 压缩页面
    │   ├── rename_page.py        # 重命名页面
    │   ├── history_page.py       # 历史记录
    │   └── pages/                # 子页面
    ├── 表格模板/                  # Excel 模板文件
    ├── 尺码映射/                  # 尺码别名配置
    └── 参考数据/                  # 参考数据
```

## 提交到 GitHub

仓库已配置好，日常提交只需：

```bash
git add .
git commit -m "改了啥"
git push
```

如果是**全新项目**要推 GitHub：

```bash
python scripts/init_github.py --name 你的用户名/仓库名
```
