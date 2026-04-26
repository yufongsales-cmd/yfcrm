"""PaddleOCR GUI: drag image -> OCR -> review text -> export Excel.

Usage:
    python paddleocr_gui.py

Dependencies:
    pip install paddleocr paddlepaddle pandas openpyxl pillow tkinterdnd2
"""

from __future__ import annotations

import os
import threading
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional

from tkinter import filedialog, messagebox, ttk
import tkinter as tk

try:
    from PIL import Image, ImageTk
except Exception:  # pragma: no cover - optional dependency fallback
    Image = None
    ImageTk = None

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except Exception:  # pragma: no cover - optional dependency fallback
    DND_FILES = None
    TkinterDnD = None

try:
    from paddleocr import PaddleOCR
except Exception:  # pragma: no cover - optional dependency fallback
    PaddleOCR = None

try:
    from openpyxl import Workbook
except Exception:  # pragma: no cover - optional dependency fallback
    Workbook = None


@dataclass
class OcrLine:
    text: str
    score: float
    box: list


class PaddleOCRApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("PaddleOCR 圖片轉 Excel 工具")
        self.root.geometry("1100x760")

        self.current_image: Optional[Path] = None
        self.preview_photo: Optional[ImageTk.PhotoImage] = None
        self.ocr_engine: Optional[PaddleOCR] = None

        self._build_ui()

    def _build_ui(self) -> None:
        top = ttk.Frame(self.root, padding=10)
        top.pack(fill="x")

        self.drop_label = ttk.Label(
            top,
            text="把圖片拖拉到這裡，或按『選擇圖片』",
            relief="ridge",
            anchor="center",
            padding=12,
        )
        self.drop_label.pack(side="left", fill="x", expand=True, padx=(0, 8))

        if DND_FILES and hasattr(self.drop_label, "drop_target_register"):
            self.drop_label.drop_target_register(DND_FILES)
            self.drop_label.dnd_bind("<<Drop>>", self._on_drop)
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind("<<Drop>>", self._on_drop)

        ttk.Button(top, text="選擇圖片", command=self._choose_file).pack(side="left", padx=4)
        ttk.Button(top, text="執行 OCR", command=self._run_ocr).pack(side="left", padx=4)
        ttk.Button(top, text="匯出 Excel", command=self._export_excel).pack(side="left", padx=4)
        self.auto_ocr_on_drop = tk.BooleanVar(value=True)
        ttk.Checkbutton(top, text="拖拉後自動OCR", variable=self.auto_ocr_on_drop).pack(side="left", padx=8)

        main = ttk.Panedwindow(self.root, orient="horizontal")
        main.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        left_frame = ttk.Frame(main, padding=8)
        right_frame = ttk.Frame(main, padding=8)
        main.add(left_frame, weight=1)
        main.add(right_frame, weight=1)

        ttk.Label(left_frame, text="圖片預覽").pack(anchor="w")
        self.preview_canvas = tk.Canvas(left_frame, bg="#f5f5f5", width=480, height=620)
        self.preview_canvas.pack(fill="both", expand=True)

        ttk.Label(right_frame, text="OCR 結果（可手動修正）").pack(anchor="w")
        self.text_box = tk.Text(right_frame, wrap="word", font=("Arial", 12))
        self.text_box.pack(fill="both", expand=True)

        status = ttk.Frame(self.root, padding=(10, 0, 10, 10))
        status.pack(fill="x")
        self.status_var = tk.StringVar(value="就緒")
        ttk.Label(status, textvariable=self.status_var).pack(anchor="w")

    def _set_status(self, text: str) -> None:
        self.status_var.set(text)

    def _choose_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="選擇圖片",
            filetypes=[("Image", "*.png *.jpg *.jpeg *.bmp *.tif *.tiff")],
        )
        if file_path:
            self._load_image(Path(file_path))

    def _on_drop(self, event) -> None:
        files = self._parse_dnd_files(event.data)
        if not files:
            messagebox.showwarning("提醒", "未讀取到可用的檔案路徑。")
            return
        if len(files) > 1:
            messagebox.showinfo("提醒", "一次只處理一張圖片，已使用第一個檔案。")
        file_path = Path(files[0])
        self._load_image(file_path)
        if self.auto_ocr_on_drop.get():
            self._run_ocr()

    def _parse_dnd_files(self, raw_data: str) -> List[str]:
        """Normalize drag-and-drop payload to a clean file path list."""
        raw_data = (raw_data or "").strip()
        if not raw_data:
            return []
        try:
            parts = list(self.root.tk.splitlist(raw_data))
        except Exception:
            parts = [raw_data]
        normalized = []
        for item in parts:
            value = item.strip()
            if value.startswith("{") and value.endswith("}"):
                value = value[1:-1]
            if value:
                normalized.append(value)
        return normalized

    def _load_image(self, path: Path) -> None:
        if not path.exists():
            messagebox.showerror("錯誤", f"找不到檔案：{path}")
            return
        if Image is None or ImageTk is None:
            messagebox.showerror(
                "缺少套件",
                "載入圖片需要 Pillow。\n請安裝：pip install pillow",
            )
            return

        self.current_image = path
        self.drop_label.configure(text=f"目前圖片：{path.name}")

        img = Image.open(path)
        img.thumbnail((1000, 1000))
        self.preview_photo = ImageTk.PhotoImage(img)

        self.preview_canvas.delete("all")
        self.preview_canvas.create_image(10, 10, anchor="nw", image=self.preview_photo)
        self._set_status(f"已載入 {path.name}")

    def _ensure_engine(self) -> bool:
        if PaddleOCR is None:
            messagebox.showerror(
                "缺少套件",
                "找不到 paddleocr，請先安裝：\n"
                "pip install paddleocr paddlepaddle",
            )
            return False

        if self.ocr_engine is None:
            self._set_status("初始化 PaddleOCR（第一次可能較久）...")
            self.root.update_idletasks()
            self.ocr_engine = PaddleOCR(use_angle_cls=True, lang="ch")
        return True

    def _run_ocr(self) -> None:
        if not self.current_image:
            messagebox.showwarning("提醒", "請先拖拉或選擇圖片。")
            return
        if not self._ensure_engine():
            return

        def worker() -> None:
            try:
                self._set_status("OCR 辨識中...")
                result = self.ocr_engine.ocr(str(self.current_image), cls=True)
                lines = self._parse_result(result)

                def update_ui() -> None:
                    self.text_box.delete("1.0", "end")
                    self.text_box.insert("1.0", "\n".join([line.text for line in lines]))
                    self._set_status(f"完成，共辨識 {len(lines)} 行文字。")

                self.root.after(0, update_ui)
            except Exception as exc:
                self.root.after(0, lambda: messagebox.showerror("OCR 失敗", str(exc)))
                self.root.after(0, lambda: self._set_status("OCR 失敗"))

        threading.Thread(target=worker, daemon=True).start()

    @staticmethod
    def _parse_result(result) -> List[OcrLine]:
        lines: List[OcrLine] = []
        if not result:
            return lines
        for block in result:
            for entry in block:
                box = entry[0]
                text = entry[1][0]
                score = float(entry[1][1])
                lines.append(OcrLine(text=text, score=score, box=box))
        return lines

    def _export_excel(self) -> None:
        raw_text = self.text_box.get("1.0", "end").strip()
        if not raw_text:
            messagebox.showwarning("提醒", "目前沒有可匯出的文字。")
            return

        save_path = filedialog.asksaveasfilename(
            title="匯出 Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="ocr_result.xlsx",
        )
        if not save_path:
            return

        rows = [line.strip() for line in raw_text.splitlines() if line.strip()]
        self._write_excel(save_path, rows)
        self._set_status(f"已匯出：{os.path.basename(save_path)}")
        messagebox.showinfo("完成", f"已匯出到：\n{save_path}")

    @staticmethod
    def _write_excel(save_path: str, rows: List[str]) -> None:
        """Write OCR rows to .xlsx.

        Prefer pandas when available, otherwise fallback to openpyxl.
        """
        try:
            import pandas as pd  # local import to avoid hard dependency at startup

            df = pd.DataFrame({"row": list(range(1, len(rows) + 1)), "text": rows})
            df.to_excel(save_path, index=False)
            return
        except Exception:
            pass

        if Workbook is None:
            raise RuntimeError(
                "匯出 Excel 需要 pandas 或 openpyxl。\n"
                "請安裝：pip install pandas openpyxl"
            )

        wb = Workbook()
        ws = wb.active
        ws.title = "OCR"
        ws.append(["row", "text"])
        for idx, text in enumerate(rows, start=1):
            ws.append([idx, text])
        wb.save(save_path)


def create_root() -> tk.Tk:
    try:
        if TkinterDnD is not None:
            return TkinterDnD.Tk()
        return tk.Tk()
    except tk.TclError as exc:
        raise RuntimeError(
            "無法啟動 GUI：目前環境沒有可用的圖形介面（DISPLAY）。\n"
            "請在桌面環境中執行，或設定 X11/Wayland 轉發。"
        ) from exc


if __name__ == "__main__":
    try:
        app_root = create_root()
        app = PaddleOCRApp(app_root)
        app_root.mainloop()
    except RuntimeError as err:
        print(err)
