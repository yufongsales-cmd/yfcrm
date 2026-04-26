import os
import tempfile
import unittest
from types import SimpleNamespace
from unittest import mock

import paddleocr_gui as gui


class TestPaddleOCRHelpers(unittest.TestCase):
    def test_parse_result_extracts_lines(self):
        raw = [
            [
                [[[0, 0], [1, 0], [1, 1], [0, 1]], ("第一行", 0.98)],
                [[[0, 2], [1, 2], [1, 3], [0, 3]], ("第二行", 0.95)],
            ]
        ]

        lines = gui.PaddleOCRApp._parse_result(raw)
        self.assertEqual(len(lines), 2)
        self.assertEqual(lines[0].text, "第一行")
        self.assertAlmostEqual(lines[1].score, 0.95)

    def test_write_excel_uses_pandas_when_available(self):
        with tempfile.TemporaryDirectory() as d:
            out = os.path.join(d, "out.xlsx")
            called = {"ok": False}

            class FakeDataFrame:
                def __init__(self, data):
                    self.data = data

                def to_excel(self, save_path, index=False):
                    called["ok"] = True
                    self.data_used = self.data
                    with open(save_path, "wb") as f:
                        f.write(b"fake")

            fake_pandas = SimpleNamespace(DataFrame=FakeDataFrame)
            with mock.patch.dict("sys.modules", {"pandas": fake_pandas}):
                gui.PaddleOCRApp._write_excel(out, ["A", "B"])

            self.assertTrue(called["ok"])
            self.assertTrue(os.path.exists(out))

    def test_write_excel_falls_back_to_openpyxl_like_workbook(self):
        with tempfile.TemporaryDirectory() as d:
            out = os.path.join(d, "out.xlsx")

            class FakeSheet:
                def __init__(self):
                    self.title = ""
                    self.rows = []

                def append(self, row):
                    self.rows.append(row)

            class FakeWorkbook:
                def __init__(self):
                    self.active = FakeSheet()
                    self.saved_path = None

                def save(self, path):
                    self.saved_path = path
                    with open(path, "wb") as f:
                        f.write(b"fake-openpyxl")

            with mock.patch.dict("sys.modules", {}, clear=False):
                with mock.patch("builtins.__import__", side_effect=ImportError("no pandas")):
                    with mock.patch.object(gui, "Workbook", FakeWorkbook):
                        gui.PaddleOCRApp._write_excel(out, ["X", "Y"])

            self.assertTrue(os.path.exists(out))

    def test_parse_dnd_files_handles_braces_and_spaces(self):
        app = gui.PaddleOCRApp.__new__(gui.PaddleOCRApp)
        app.root = SimpleNamespace(tk=SimpleNamespace(splitlist=lambda s: ("{C:/a b/c.png}",)))

        files = gui.PaddleOCRApp._parse_dnd_files(app, "{C:/a b/c.png}")
        self.assertEqual(files, ["C:/a b/c.png"])

    def test_parse_dnd_files_handles_multiple_paths(self):
        app = gui.PaddleOCRApp.__new__(gui.PaddleOCRApp)
        app.root = SimpleNamespace(tk=SimpleNamespace(splitlist=lambda s: ("A.png", "B.png")))

        files = gui.PaddleOCRApp._parse_dnd_files(app, "A.png B.png")
        self.assertEqual(files, ["A.png", "B.png"])


if __name__ == "__main__":
    unittest.main()
