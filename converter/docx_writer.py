"""Word 生成模块 - 将 PDF 页面渲染为图片插入 Word"""

from docx import Document
from docx.shared import Cm
import fitz
import io


def convert_pages_as_images(output_path: str, pdf_path: str) -> None:
    """
    将 PDF 每页渲染为图片插入 Word - 布局 100% 一致

    Args:
        output_path: 输出 Word 文件路径
        pdf_path: 输入 PDF 文件路径
    """
    # 打开 PDF
    doc = fitz.open(pdf_path)
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
