from __future__ import annotations

from pathlib import Path
from tempfile import TemporaryDirectory
import sys

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from openpyxl import Workbook, load_workbook
from PIL import Image, ImageDraw

from qt_app.services.excel_service import ExcelService
from qt_app.services.logger import AppLogger
from qt_app.services.word_service import WordService


def test_excel_mapping_and_word_generation() -> None:
    with TemporaryDirectory() as td:
        base = Path(td)
        logger = AppLogger(base / "logs" / "output.log")

        xlsx = base / "cases.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.append(["Case Name", "Check Point", "Step", "Desc", "Expected", "Status"])
        ws.append(["登录功能", "登录成功", "打开页面", "点击登录按钮", "显示首页", ""])
        ws.append(["已完成用例", "不应返回", "", "", "", "已执行"])
        wb.save(xlsx)

        excel = ExcelService(logger)
        preview = excel.preview(xlsx)
        assert preview.suggested_mapping["测试名称"] == "Case Name"
        assert preview.suggested_mapping["验证点"] == "Check Point"
        assert preview.suggested_mapping["测试结果"] == "Status"

        items = excel.load(xlsx, preview.suggested_mapping)
        assert len(items) == 1
        assert items[0].case_name == "登录功能"
        assert items[0].step_name == "打开页面"
        excel.mark_executed([items[0].row_index])
        excel.close()

        wb2 = load_workbook(xlsx)
        assert wb2.active["F2"].value == "已执行"
        wb2.close()

        image_path = base / "shot.png"
        img = Image.new("RGB", (400, 240), "white")
        ImageDraw.Draw(img).rectangle((20, 20, 180, 120), outline="red", width=4)
        img.save(image_path)

        word = WordService(base / "word_output", logger)
        doc_path = word.add_image("登录/功能:*?", "登录成功", image_path, "步骤名称：打开页面")
        assert doc_path.exists()
        assert doc_path.name == "登录_功能.docx"


if __name__ == "__main__":
    test_excel_mapping_and_word_generation()
    print("smoke tests passed")
