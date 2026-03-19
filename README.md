# PDF 转 Word 工具

将 PDF 文件转换为 Word 文档（.docx），**布局 100% 与原始 PDF 一致**。

## 功能特点

- **图片渲染模式**：将 PDF 每页渲染为高清图片（300 DPI）插入 Word
- **布局 100% 一致**：完美还原原始 PDF 的所有排版、图片、表格等
- **使用简单**：一行命令即可完成转换

## 安装

### 1. 创建虚拟环境（推荐）

```bash
py -m venv venv
```

### 2. 激活虚拟环境

**Windows (Git Bash):**
```bash
source venv/Scripts/activate
```

**Windows (PowerShell/CMD):**
```bash
venv\Scripts\activate
```

### 3. 安装依赖

```bash
pip install -r requirements.txt
```

## 使用方法

### 基本用法

```bash
py main.py input.pdf
```

会在同目录下生成 `input.docx`

### 指定输出文件

```bash
py main.py input.pdf -o output.docx
```

### 示例

```bash
# 转换当前目录的 PDF
py main.py 2026.pdf

# 指定输出位置
py main.py 2026.pdf -o D:/output/result.docx
```

## 项目结构

```
pdf2word/
├── main.py              # 主入口文件
├── converter/           # 转换模块
│   └── docx_writer.py   # Word 生成（图片渲染）
├── requirements.txt     # 依赖列表
└── README.md            # 说明文档
```

## 依赖说明

| 库 | 用途 |
|----|------|
| pymupdf (fitz) | PDF 解析和渲染 |
| python-docx | Word 文件生成 |
| Pillow | 图片处理 |

## 常见问题

### Q: 为什么转换后的 Word 文件很大？
A: 为了保持高质量，图片以 300 DPI 渲染。如需更小的文件，可修改代码中的缩放比例（默认 `fitz.Matrix(3, 3)`）。

### Q: 文字可以编辑吗？
A: 本工具采用图片渲染模式，文字不可编辑，但布局 100% 与原始 PDF 一致。如需编辑文字，建议使用其他 OCR 工具。

### Q: 如何在没有 Python 环境的电脑上使用？
A: 可使用 PyInstaller 打包为可执行文件：
```bash
pip install pyinstaller
pyinstaller --onefile --name pdf2word main.py
```

## 技术细节

- 使用 PyMuPDF 将每页渲染为 PNG 图片（3 倍缩放，约 300 DPI）
- 图片按 A4 页面比例插入 Word
- 自动添加分页符
- 布局 100% 与原始 PDF 一致

## License

MIT License
