"""PDF 转 Word 工具 - 主入口

支持将 PDF 页面渲染为图片插入 Word，布局 100% 与原始 PDF 一致。
"""

import argparse
import os
import sys
import fitz
from docx import Document
from docx.shared import Cm
import io


def convert_pdf_to_images_in_word(input_path: str, output_path: str = None) -> str:
    """
    将 PDF 每页渲染为图片插入 Word 文档

    Args:
        input_path: 输入 PDF 文件路径
        output_path: 输出 Word 文件路径（可选，默认为同目录下）

    Returns:
        输出文件路径
    """
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"文件不存在：{input_path}")

    if not input_path.lower().endswith('.pdf'):
        raise ValueError("输入文件必须是 PDF 格式")

    # 生成输出路径
    if output_path is None:
        base_name = os.path.splitext(input_path)[0]
        output_path = f"{base_name}.docx"

    print(f"正在处理：{os.path.basename(input_path)}")

    # 打开 PDF
    doc = fitz.open(input_path)
    word_doc = Document()

    # 设置 Word 页面为 A4
    section = word_doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(1.27)
    section.bottom_margin = Cm(1.27)
    section.left_margin = Cm(1.27)
    section.right_margin = Cm(1.27)

    for i in range(len(doc)):
        page = doc[i]

        # 渲染为图片（300 DPI）
        mat = fitz.Matrix(3, 3)  # 3 倍缩放
        pix = page.get_pixmap(matrix=mat)

        # 转换为 PNG 字节流
        img_bytes = io.BytesIO(pix.tobytes("png"))

        # 计算图片在 Word 中的尺寸
        pdf_width_cm = page.rect.width * 72 / 25.4
        pdf_height_cm = page.rect.height * 72 / 25.4

        # 缩放到 Word 页面宽度
        scale = (21 - 2.54) / pdf_width_cm
        img_width = pdf_width_cm * scale
        img_height = pdf_height_cm * scale

        # 插入图片
        if i > 0:
            word_doc.add_page_break()
        word_doc.add_picture(img_bytes, width=Cm(img_width))

        pix = None
        img_bytes = None

    doc.close()
    word_doc.save(output_path)

    print(f"转换完成！输出文件：{output_path}")
    return output_path


def main():
    """命令行入口"""
    parser = argparse.ArgumentParser(
        description='PDF 转 Word 工具 - 图片渲染模式（布局 100% 一致）'
    )
    parser.add_argument(
        'input',
        help='输入 PDF 文件路径'
    )
    parser.add_argument(
        '-o', '--output',
        help='输出 Word 文件路径（可选）'
    )

    args = parser.parse_args()

    try:
        convert_pdf_to_images_in_word(args.input, args.output)
    except Exception as e:
        print(f"错误：{e}", file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()
