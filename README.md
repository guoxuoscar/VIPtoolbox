# 唯品上新工具箱

电商运营辅助工具集 — 基于 PySide6 的 Windows 桌面应用。专为唯品会/淘宝电商运营设计，涵盖商品图片采集、图片处理、PDF 编辑、ERP 资料导出、尺码表处理、文件管理等日常高频操作。

## 功能

| 功能 | 说明 |
|------|------|
| **唯品批量找图** | 按款号批量从唯品会商品页采集主图/SKU图/详情图 |
| **淘宝图片下载** | Playwright 浏览器自动化，下载淘宝商品主图/SKU图/详情图/视频，带防封策略 |
| **图片压缩** | JPEG/PNG 批量压缩到指定大小（按尺寸或文件大小） |
| **PDF 工具** | PDF 拆分、合并、提取页面 |
| **批量抠图** | 基于 rembg + u2net.onnx 自动抠透明底图，支持批量处理 |
| **批量尺码表录入** | 从图片/PDF 中识别尺码表并批量录入到 Excel 模板（支持 OCR） |
| **批量模板 & ERP** | 批量生成 QA/配件明细/试穿报告/属性模板；导出 ERP 条码对照表/商品资料/定价导入 |
| **批量重命名** | 正则匹配 / 序号编号 / 关键词批量重命名文件 |
| **文件工具** | 文件夹批量创建（按 Excel 列层级嵌套）、文件提取汇总、图片分类归文件夹、文件清单导出、文件分发、供应商编码生成、重复文件清理 |
| **操作历史** | 记录所有操作日志，方便回溯 |

## 安装

### 1. 环境要求

- Windows 10/11
- Python 3.9+
- Google Chrome（用于淘宝/唯品图片下载）

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
    │   ├── batch_template.py     # 批量模板生成（QA/配件/试穿/属性）
    │   ├── browser.py            # Playwright 浏览器自动化
    │   ├── erp_product.py        # ERP 商品资料导出
    │   ├── ocr.py                # OCR 识别引擎
    │   ├── pdf_edit_core.py      # PDF 编辑
    │   ├── size_mapping.py       # 尺码映射
    │   └── utils.py              # 工具函数
    ├── ui/                       # 界面
    │   ├── main_window.pyw       # 主窗口（侧边栏导航）
    │   ├── vip_image_finder_page.py  # 唯品批量找图
    │   ├── compress_page.py      # 图片压缩
    │   ├── rename_page.py        # 批量重命名
    │   ├── file_tools_page.py    # 文件工具（创建/提取/分类/导出/分发/清重）
    │   ├── history_page.py       # 操作历史
    │   ├── anti_ban_dialog.py    # 防封设置
    │   ├── path_drop.py          # 拖拽输入组件
    │   └── pages/
    │       └── batch_erp_page.py # 批量模板 & ERP 页面
    ├── ocr_engine/               # OCR 引擎（需自行下载）
    ├── 表格模板/                  # Excel 模板文件
    ├── 尺码映射/                  # 尺码别名配置
    └── 参考数据/                  # 参考数据
```

## Git 日常使用

```bash
git add -A
git commit -m "修改说明"
git push
```
