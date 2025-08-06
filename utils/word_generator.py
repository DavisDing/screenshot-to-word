import os
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

class WordGenerator:
    def __init__(self, base_dir, logger):
        self.base_dir = base_dir
        self.logger = logger

    def get_doc_path(self, case_name):
        folder = os.path.join(self.base_dir, case_name)
        if not os.path.exists(folder):
            os.makedirs(folder)
        doc_path = os.path.join(folder, f"{case_name}.docx")
        return doc_path

    def insert_case_image(self, case_name, case_desc, image_path):
        doc_path = self.get_doc_path(case_name)
        is_new_doc = not os.path.exists(doc_path)

        if is_new_doc:
            doc = Document()
            doc.add_heading(case_name, level=1)
            p = doc.add_paragraph(f"验证点：{case_desc}")
            p.space_after = Inches(0.1)
        else:
            try:
                doc = Document(doc_path)
                self.logger.write(f"加载已有 Word 文件：{doc_path}")
            except Exception as e:
                self.logger.write(f"[异常] 加载 Word 文件失败，创建新文档：{e}")
                doc = Document()
                doc.add_heading(case_name, level=1)
                p = doc.add_paragraph(f"验证点：{case_desc}")
                p.space_after = Inches(0.1)

        pic_count = sum(1 for p in doc.paragraphs if p.text.startswith("截图"))
        pic_number = pic_count + 1

        try:
            max_width = Inches(6.0)
            p_pic = doc.add_paragraph()
            run = p_pic.add_run()
            run.add_picture(image_path, width=max_width)
            p_pic.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

            caption = doc.add_paragraph(f"截图{pic_number}")
            caption.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

            doc.save(doc_path)
            self.logger.write(f"插入截图并保存 Word 文件：{doc_path}")
        except Exception as e:
            self.logger.write(f"[异常] 插入图片失败：{e}")