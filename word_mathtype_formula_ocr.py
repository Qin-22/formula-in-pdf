# -*- coding: utf-8 -*-
"""
Word / MathType / PDF 公式实时框选 -> LaTeX
================================================

功能：
1. 打开 DOCX / DOC
2. 自动调用 Microsoft Word 将 Word 文档渲染成临时 PDF
3. 也可以直接打开 PDF
4. 左侧显示页面
5. 支持：
   - 上一页 / 下一页
   - 页码跳转
   - 放大 / 缩小
   - 100%
   - 适应窗口
   - Ctrl + 鼠标滚轮缩放
   - 鼠标滚轮滚动
   - 中键拖动页面
6. 左键拖动框选公式
7. 框选过程中右侧实时显示裁剪预览
8. 松开鼠标后不会自动 OCR
9. 点击“确认框选并 OCR”后调用 pix2tex
10. LaTeX 显示在右侧，可一键复制
11. 每次 OCR 的公式图片自动保存到当前程序目录：
       formula_crops/
12. 支持重复框选、重复 OCR

适用：
- MathType 公式
- Word 公式
- 图片公式
- PDF 中的公式

说明：
本程序不直接读取 MathType OLE 内部数据。
它采用：
    Word -> PDF -> 框选公式 -> pix2tex -> LaTeX
因此不需要浏览器在线公式编辑器。

环境：
Windows
Microsoft Word
Python 3.8
PyMuPDF
Pillow
pywin32
pix2tex
"""

import os
import tempfile
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path

import fitz
from PIL import Image, ImageTk


# ============================================================
# Microsoft Word
# ============================================================

try:
    import win32com.client
except ImportError:
    win32com = None


# ============================================================
# pix2tex
# ============================================================

try:
    from pix2tex.cli import LatexOCR
except ImportError:
    LatexOCR = None


class WordFormulaOCR:

    def __init__(self, root):

        self.root = root

        self.root.title(
            "Word / MathType 公式实时框选 → LaTeX"
        )

        self.root.geometry("1500x900")
        self.root.minsize(1100, 700)

        # ====================================================
        # PDF
        # ====================================================

        self.doc = None
        self.pdf_path = None

        # ====================================================
        # 临时目录
        # ====================================================

        self.temp_dir = None

        # ====================================================
        # 页面
        # ====================================================

        self.page_index = 0

        # ====================================================
        # 缩放
        # ====================================================

        self.zoom = 1.20

        self.min_zoom = 0.30

        self.max_zoom = 5.00

        # ====================================================
        # 页面图片
        # ====================================================

        self.page_image = None

        self.tk_image = None

        # 页面在 Canvas 中的位置
        self.offset_x = 0
        self.offset_y = 0

        # ====================================================
        # 框选
        # ====================================================

        self.selecting = False

        self.start_x = None
        self.start_y = None

        self.current_rect = None

        # 最终确认区域
        self.confirmed_rect = None

        # ====================================================
        # 平移
        # ====================================================

        self.panning = False

        self.pan_last_x = None
        self.pan_last_y = None

        # ====================================================
        # OCR
        # ====================================================

        self.ocr_model = None

        # ====================================================
        # UI
        # ====================================================

        self.build_ui()

        self.bind_events()

        self.root.protocol(
            "WM_DELETE_WINDOW",
            self.on_close
        )

    # ========================================================
    # UI
    # ========================================================

    def build_ui(self):

        # ====================================================
        # 顶部工具栏
        # ====================================================

        top = tk.Frame(self.root)

        top.pack(
            fill="x",
            padx=8,
            pady=6
        )

        tk.Button(
            top,
            text="打开 Word",
            width=12,
            command=self.open_word
        ).pack(
            side="left",
            padx=3
        )

        tk.Button(
            top,
            text="打开 PDF",
            width=12,
            command=self.open_pdf
        ).pack(
            side="left",
            padx=3
        )

        tk.Button(
            top,
            text="上一页",
            width=9,
            command=self.prev_page
        ).pack(
            side="left",
            padx=3
        )

        tk.Button(
            top,
            text="下一页",
            width=9,
            command=self.next_page
        ).pack(
            side="left",
            padx=3
        )

        tk.Label(
            top,
            text="页码:"
        ).pack(
            side="left",
            padx=(12, 2)
        )

        self.page_var = tk.StringVar(
            value="0"
        )

        self.page_entry = tk.Entry(
            top,
            textvariable=self.page_var,
            width=6
        )

        self.page_entry.pack(
            side="left"
        )

        tk.Button(
            top,
            text="跳转",
            width=7,
            command=self.goto_page
        ).pack(
            side="left",
            padx=3
        )

        # 缩小
        tk.Button(
            top,
            text="−",
            width=4,
            command=lambda: self.change_zoom(0.85)
        ).pack(
            side="left",
            padx=(15, 2)
        )

        self.zoom_label = tk.Label(
            top,
            text="120%"
        )

        self.zoom_label.pack(
            side="left",
            padx=3
        )

        # 放大
        tk.Button(
            top,
            text="+",
            width=4,
            command=lambda: self.change_zoom(1.18)
        ).pack(
            side="left",
            padx=2
        )

        # 适应窗口
        tk.Button(
            top,
            text="适应窗口",
            width=10,
            command=self.fit_page
        ).pack(
            side="left",
            padx=3
        )

        # 100%
        tk.Button(
            top,
            text="100%",
            width=8,
            command=self.zoom_100
        ).pack(
            side="left",
            padx=3
        )

        # OCR
        tk.Button(
            top,
            text="确认框选并 OCR",
            width=18,
            command=self.confirm_ocr
        ).pack(
            side="left",
            padx=(18, 3)
        )

        # 取消
        tk.Button(
            top,
            text="取消框选",
            width=12,
            command=self.cancel_selection
        ).pack(
            side="left",
            padx=3
        )

        # 状态
        self.status_var = tk.StringVar(
            value="请打开 Word 文档或 PDF"
        )

        tk.Label(
            top,
            textvariable=self.status_var,
            anchor="w"
        ).pack(
            side="left",
            padx=15,
            fill="x",
            expand=True
        )

        # ====================================================
        # 主区域
        # ====================================================

        main = tk.PanedWindow(
            self.root,
            orient=tk.HORIZONTAL,
            sashrelief=tk.RAISED
        )

        main.pack(
            fill="both",
            expand=True,
            padx=8,
            pady=(0, 8)
        )

        # ====================================================
        # 左侧页面
        # ====================================================

        left_frame = tk.Frame(main)

        main.add(
            left_frame,
            stretch="always",
            minsize=650
        )

        self.canvas = tk.Canvas(
            left_frame,
            bg="#777777",
            highlightthickness=0
        )

        self.vbar = tk.Scrollbar(
            left_frame,
            orient="vertical",
            command=self.canvas.yview
        )

        self.hbar = tk.Scrollbar(
            left_frame,
            orient="horizontal",
            command=self.canvas.xview
        )

        self.canvas.configure(
            xscrollcommand=self.hbar.set,
            yscrollcommand=self.vbar.set
        )

        self.canvas.grid(
            row=0,
            column=0,
            sticky="nsew"
        )

        self.vbar.grid(
            row=0,
            column=1,
            sticky="ns"
        )

        self.hbar.grid(
            row=1,
            column=0,
            sticky="ew"
        )

        left_frame.rowconfigure(
            0,
            weight=1
        )

        left_frame.columnconfigure(
            0,
            weight=1
        )

        # ====================================================
        # 右侧
        # ====================================================

        right_frame = tk.Frame(
            main,
            width=500
        )

        main.add(
            right_frame,
            minsize=400
        )

        # ----------------------------------------------------
        # 公式预览
        # ----------------------------------------------------

        tk.Label(
            right_frame,
            text="公式实时预览",
            font=("Microsoft YaHei UI", 11, "bold")
        ).pack(
            anchor="w",
            padx=8,
            pady=(8, 4)
        )

        preview_box = tk.Frame(
            right_frame,
            bd=1,
            relief="sunken"
        )

        preview_box.pack(
            fill="x",
            padx=8,
            pady=4
        )

        self.preview_label = tk.Label(
            preview_box,
            text="框选公式后这里显示预览",
            bg="white",
            height=9,
            anchor="center"
        )

        self.preview_label.pack(
            fill="both",
            expand=True
        )

        # ----------------------------------------------------
        # LaTeX
        # ----------------------------------------------------

        tk.Label(
            right_frame,
            text="LaTeX",
            font=("Microsoft YaHei UI", 11, "bold")
        ).pack(
            anchor="w",
            padx=8,
            pady=(12, 4)
        )

        self.latex_text = tk.Text(
            right_frame,
            wrap="word",
            font=("Consolas", 12)
        )

        self.latex_text.pack(
            fill="both",
            expand=True,
            padx=8,
            pady=(0, 8)
        )

        # ----------------------------------------------------
        # 底部
        # ----------------------------------------------------

        bottom = tk.Frame(
            right_frame
        )

        bottom.pack(
            fill="x",
            padx=8,
            pady=(0, 8)
        )

        tk.Button(
            bottom,
            text="复制 LaTeX",
            width=14,
            command=self.copy_latex
        ).pack(
            side="left"
        )

        tk.Button(
            bottom,
            text="清空",
            width=10,
            command=self.clear_latex
        ).pack(
            side="left",
            padx=6
        )

        tk.Label(
            bottom,
            text="左键框选 | Ctrl+滚轮缩放 | 中键拖动",
            fg="#666666"
        ).pack(
            side="right"
        )

    # ========================================================
    # 事件
    # ========================================================

    def bind_events(self):

        # 左键
        self.canvas.bind(
            "<ButtonPress-1>",
            self.on_left_down
        )

        self.canvas.bind(
            "<B1-Motion>",
            self.on_left_drag
        )

        self.canvas.bind(
            "<ButtonRelease-1>",
            self.on_left_up
        )

        # 滚轮
        self.canvas.bind(
            "<MouseWheel>",
            self.on_mousewheel
        )

        self.canvas.bind(
            "<Shift-MouseWheel>",
            self.on_shift_wheel
        )

        # 中键
        self.canvas.bind(
            "<ButtonPress-2>",
            self.on_middle_down
        )

        self.canvas.bind(
            "<B2-Motion>",
            self.on_middle_drag
        )

        self.canvas.bind(
            "<ButtonRelease-2>",
            self.on_middle_up
        )

        # 快捷键
        self.root.bind(
            "<Control-o>",
            lambda e: self.open_word()
        )

        self.root.bind(
            "<Control-0>",
            lambda e: self.zoom_100()
        )

        self.root.bind(
            "<Control-plus>",
            lambda e: self.change_zoom(1.18)
        )

        self.root.bind(
            "<Control-minus>",
            lambda e: self.change_zoom(0.85)
        )

    # ========================================================
    # Word
    # ========================================================

    def open_word(self):

        path = filedialog.askopenfilename(
            title="选择 Word 文档",
            filetypes=[
                ("Word 文档", "*.docx"),
                ("Word 文档", "*.doc"),
                ("所有文件", "*.*")
            ]
        )

        if not path:
            return

        if win32com is None:

            messagebox.showerror(
                "缺少 pywin32",
                "当前环境没有 pywin32。\n\n"
                "请在 FilterNet 环境执行：\n\n"
                "python -m pip install pywin32 "
                "-i https://pypi.tuna.tsinghua.edu.cn/simple"
            )

            return

        try:

            self.status_var.set(
                "正在调用 Microsoft Word 转 PDF，请稍候..."
            )

            self.root.update_idletasks()

            pdf_path = self.word_to_pdf(
                path
            )

            self.load_pdf(
                pdf_path
            )

            self.status_var.set(
                f"已打开：{Path(path).name} | "
                "左键拖动框选公式"
            )

        except Exception as e:

            messagebox.showerror(
                "Word 转 PDF 失败",
                str(e)
            )

            self.status_var.set(
                "Word 转 PDF 失败"
            )

    def word_to_pdf(self, word_path):

        self.cleanup_temp()

        self.temp_dir = tempfile.TemporaryDirectory(
            prefix="word_formula_"
        )

        out_pdf = os.path.join(
            self.temp_dir.name,
            Path(word_path).stem + ".pdf"
        )

        word = None
        doc = None

        try:

            word = win32com.client.DispatchEx(
                "Word.Application"
            )

            word.Visible = False

            word.DisplayAlerts = 0

            doc = word.Documents.Open(
                os.path.abspath(word_path),
                ReadOnly=True,
                AddToRecentFiles=False
            )

            # 17 = PDF
            doc.ExportAsFixedFormat(
                OutputFileName=os.path.abspath(
                    out_pdf
                ),
                ExportFormat=17,
                OpenAfterExport=False,
                OptimizeFor=0,
                CreateBookmarks=0
            )

            doc.Close(False)

            doc = None

            return out_pdf

        finally:

            try:

                if doc is not None:
                    doc.Close(False)

            except Exception:
                pass

            try:

                if word is not None:
                    word.Quit()

            except Exception:
                pass

    # ========================================================
    # PDF
    # ========================================================

    def open_pdf(self):

        path = filedialog.askopenfilename(
            title="选择 PDF",
            filetypes=[
                ("PDF 文件", "*.pdf"),
                ("所有文件", "*.*")
            ]
        )

        if not path:
            return

        try:

            self.load_pdf(
                path
            )

            self.status_var.set(
                f"已打开：{Path(path).name}"
            )

        except Exception as e:

            messagebox.showerror(
                "打开 PDF 失败",
                str(e)
            )

    def load_pdf(self, pdf_path):

        if self.doc is not None:

            self.doc.close()

            self.doc = None

        self.doc = fitz.open(
            pdf_path
        )

        self.pdf_path = pdf_path

        self.page_index = 0

        self.zoom = 1.20

        self.confirmed_rect = None

        self.current_rect = None

        self.page_var.set(
            "1"
        )

        self.render_page()

        self.clear_preview()

    # ========================================================
    # 页面渲染
    # ========================================================

    def render_page(self):

        if self.doc is None:
            return

        page = self.doc[
            self.page_index
        ]

        matrix = fitz.Matrix(
            self.zoom,
            self.zoom
        )

        pix = page.get_pixmap(
            matrix=matrix,
            alpha=False
        )

        img = Image.frombytes(
            "RGB",
            [pix.width, pix.height],
            pix.samples
        )

        self.page_image = img

        self.tk_image = ImageTk.PhotoImage(
            img
        )

        self.canvas.delete(
            "all"
        )

        # 当前 Canvas 尺寸
        canvas_w = max(
            self.canvas.winfo_width(),
            1
        )

        canvas_h = max(
            self.canvas.winfo_height(),
            1
        )

        # 居中
        self.offset_x = max(
            (canvas_w - img.width) // 2,
            20
        )

        self.offset_y = max(
            (canvas_h - img.height) // 2,
            20
        )

        self.canvas.create_image(
            self.offset_x,
            self.offset_y,
            image=self.tk_image,
            anchor="nw",
            tags="page"
        )

        self.canvas.configure(
            scrollregion=(
                0,
                0,
                max(
                    img.width + self.offset_x * 2,
                    canvas_w
                ),
                max(
                    img.height + self.offset_y * 2,
                    canvas_h
                )
            )
        )

        self.zoom_label.config(
            text=f"{self.zoom * 100:.0f}%"
        )

        self.page_var.set(
            str(self.page_index + 1)
        )

    # ========================================================
    # 适应窗口
    # ========================================================

    def fit_page(self):

        if self.doc is None:
            return

        self.root.update_idletasks()

        page = self.doc[
            self.page_index
        ]

        rect = page.rect

        canvas_w = max(
            self.canvas.winfo_width() - 40,
            200
        )

        canvas_h = max(
            self.canvas.winfo_height() - 40,
            200
        )

        zx = canvas_w / rect.width

        zy = canvas_h / rect.height

        self.zoom = max(
            self.min_zoom,
            min(
                self.max_zoom,
                min(zx, zy)
            )
        )

        self.render_page()

    def zoom_100(self):

        self.zoom = 1.0

        self.render_page()

    def change_zoom(self, factor):

        if self.doc is None:
            return

        new_zoom = (
            self.zoom * factor
        )

        self.zoom = max(
            self.min_zoom,
            min(
                self.max_zoom,
                new_zoom
            )
        )

        self.render_page()

    # ========================================================
    # 翻页
    # ========================================================

    def prev_page(self):

        if self.doc is None:
            return

        if self.page_index > 0:

            self.page_index -= 1

            self.reset_selection()

            self.render_page()

            self.clear_preview()

    def next_page(self):

        if self.doc is None:
            return

        if self.page_index < len(self.doc) - 1:

            self.page_index += 1

            self.reset_selection()

            self.render_page()

            self.clear_preview()

    def goto_page(self):

        if self.doc is None:
            return

        try:

            n = int(
                self.page_var.get()
            ) - 1

            if 0 <= n < len(self.doc):

                self.page_index = n

                self.reset_selection()

                self.render_page()

                self.clear_preview()

            else:

                messagebox.showwarning(
                    "页码错误",
                    f"页码范围：1 - {len(self.doc)}"
                )

        except ValueError:

            messagebox.showwarning(
                "页码错误",
                "请输入整数页码"
            )

    # ========================================================
    # 坐标转换
    # ========================================================

    def canvas_to_image(self, x, y):

        return (
            x - self.offset_x,
            y - self.offset_y
        )

    def image_to_canvas(self, x, y):

        return (
            x + self.offset_x,
            y + self.offset_y
        )

    # ========================================================
    # 左键开始框选
    # ========================================================

    def on_left_down(self, event):

        if self.page_image is None:
            return

        self.selecting = True

        self.start_x = event.x

        self.start_y = event.y

        self.confirmed_rect = None

        self.canvas.delete(
            "selection"
        )

        self.current_rect = (
            self.canvas.create_rectangle(
                event.x,
                event.y,
                event.x,
                event.y,
                outline="red",
                width=3,
                tags="selection"
            )
        )

    # ========================================================
    # 框选过程中
    # ========================================================

    def on_left_drag(self, event):

        if not self.selecting:
            return

        x1 = self.start_x

        y1 = self.start_y

        x2 = event.x

        y2 = event.y

        self.canvas.coords(
            self.current_rect,
            x1,
            y1,
            x2,
            y2
        )

        # 实时预览
        self.update_selection_preview(
            x1,
            y1,
            x2,
            y2
        )

    # ========================================================
    # 松开鼠标
    # ========================================================

    def on_left_up(self, event):

        if not self.selecting:
            return

        self.selecting = False

        x1 = self.start_x
        y1 = self.start_y

        x2 = event.x
        y2 = event.y

        x1, x2 = sorted(
            (x1, x2)
        )

        y1, y2 = sorted(
            (y1, y2)
        )

        ix1, iy1 = self.canvas_to_image(
            x1,
            y1
        )

        ix2, iy2 = self.canvas_to_image(
            x2,
            y2
        )

        ix1 = max(
            0,
            min(
                self.page_image.width,
                int(ix1)
            )
        )

        iy1 = max(
            0,
            min(
                self.page_image.height,
                int(iy1)
            )
        )

        ix2 = max(
            0,
            min(
                self.page_image.width,
                int(ix2)
            )
        )

        iy2 = max(
            0,
            min(
                self.page_image.height,
                int(iy2)
            )
        )

        if (
            ix2 - ix1 < 10
            or
            iy2 - iy1 < 10
        ):

            self.cancel_selection()

            return

        self.confirmed_rect = (
            ix1,
            iy1,
            ix2,
            iy2
        )

        self.status_var.set(
            "框选完成，请检查右侧预览；"
            "确认无误后点击“确认框选并 OCR”"
        )

        self.update_selection_preview(
            x1,
            y1,
            x2,
            y2
        )

    # ========================================================
    # 实时预览
    # ========================================================

    def update_selection_preview(
        self,
        x1,
        y1,
        x2,
        y2
    ):

        if self.page_image is None:
            return

        x1, x2 = sorted(
            (x1, x2)
        )

        y1, y2 = sorted(
            (y1, y2)
        )

        ix1, iy1 = self.canvas_to_image(
            x1,
            y1
        )

        ix2, iy2 = self.canvas_to_image(
            x2,
            y2
        )

        ix1 = max(
            0,
            min(
                self.page_image.width,
                int(ix1)
            )
        )

        iy1 = max(
            0,
            min(
                self.page_image.height,
                int(iy1)
            )
        )

        ix2 = max(
            0,
            min(
                self.page_image.width,
                int(ix2)
            )
        )

        iy2 = max(
            0,
            min(
                self.page_image.height,
                int(iy2)
            )
        )

        if ix2 <= ix1 or iy2 <= iy1:
            return

        crop = self.page_image.crop(
            (
                ix1,
                iy1,
                ix2,
                iy2
            )
        )

        max_w = 450

        max_h = 220

        preview = crop.copy()

        preview.thumbnail(
            (
                max_w,
                max_h
            ),
            Image.Resampling.LANCZOS
        )

        self.preview_tk = ImageTk.PhotoImage(
            preview
        )

        self.preview_label.configure(
            image=self.preview_tk,
            text=""
        )

    # ========================================================
    # 取消框选
    # ========================================================

    def cancel_selection(self):

        self.reset_selection()

        self.clear_preview()

        self.status_var.set(
            "已取消框选，请重新拖动鼠标选择公式"
        )

    def reset_selection(self):

        self.selecting = False

        self.confirmed_rect = None

        self.current_rect = None

        self.canvas.delete(
            "selection"
        )

    # ========================================================
    # OCR
    # ========================================================

    def confirm_ocr(self):

        if self.page_image is None:

            messagebox.showwarning(
                "提示",
                "请先打开 Word 或 PDF"
            )

            return

        if self.confirmed_rect is None:

            messagebox.showwarning(
                "提示",
                "请先用鼠标框选一个公式"
            )

            return

        ix1, iy1, ix2, iy2 = (
            self.confirmed_rect
        )

        crop = self.page_image.crop(
            (
                ix1,
                iy1,
                ix2,
                iy2
            )
        )

        # ====================================================
        # 自动增加白边
        # ====================================================

        crop = self.add_margin(
            crop,
            25
        )

        # ====================================================
        # 保存公式图片
        # ====================================================

        # debug_dir = (
        #     Path.cwd()
        #     / "formula_crops"
        # )
        #
        # debug_dir.mkdir(
        #     exist_ok=True
        # )
        #
        # existing = list(
        #     debug_dir.glob("*.png")
        # )
        #
        # number = len(existing) + 1
        #
        # crop_path = (
        #     debug_dir
        #     /
        #     (
        #         f"page_{self.page_index + 1:03d}"
        #         f"_formula_{number:03d}.png"
        #     )
        # )
        #
        # crop.save(
        #     crop_path
        # )

        # ====================================================
        # 加载 OCR 模型
        # ====================================================

        try:

            if self.ocr_model is None:

                self.status_var.set(
                    "正在加载 LaTeX-OCR 模型..."
                )

                self.root.update_idletasks()

                if LatexOCR is None:

                    raise RuntimeError(
                        "pix2tex 没有正确安装。\n\n"
                        "请测试：\n"
                        "python -c "
                        "\"from pix2tex.cli import LatexOCR; "
                        "print('pix2tex OK')\""
                    )

                self.ocr_model = LatexOCR()

            # =================================================
            # OCR
            # =================================================

            self.status_var.set(
                "正在识别公式，请稍候..."
            )

            self.root.update_idletasks()

            latex = self.ocr_model(
                crop
            )

            # =================================================
            # 显示
            # =================================================

            self.latex_text.delete(
                "1.0",
                "end"
            )

            self.latex_text.insert(
                "1.0",
                str(latex)
            )

            self.status_var.set(
                "OCR 完成 | "
                # f"公式图片：{crop_path}"
            )

        except Exception as e:

            messagebox.showerror(
                "LaTeX-OCR 失败",
                str(e)
            )

            self.status_var.set(
                "OCR 失败"
            )

    # ========================================================
    # 增加白边
    # ========================================================

    @staticmethod
    def add_margin(
        img,
        margin=25
    ):

        result = Image.new(
            "RGB",
            (
                img.width + margin * 2,
                img.height + margin * 2
            ),
            "white"
        )

        result.paste(
            img,
            (
                margin,
                margin
            )
        )

        return result

    # ========================================================
    # 复制 LaTeX
    # ========================================================

    def copy_latex(self):

        text = self.latex_text.get(
            "1.0",
            "end"
        ).strip()

        if not text:
            return

        self.root.clipboard_clear()

        self.root.clipboard_append(
            text
        )

        self.root.update()

        self.status_var.set(
            "LaTeX 已复制到剪贴板"
        )

    # ========================================================
    # 清空 LaTeX
    # ========================================================

    def clear_latex(self):

        self.latex_text.delete(
            "1.0",
            "end"
        )

    # ========================================================
    # 滚轮
    # ========================================================

    def on_mousewheel(self, event):

        # Ctrl + 滚轮 = 缩放
        if event.state & 0x0004:

            if event.delta > 0:

                self.change_zoom(
                    1.12
                )

            else:

                self.change_zoom(
                    0.89
                )

            return

        # 普通滚轮 = 上下移动
        self.canvas.yview_scroll(
            -int(
                event.delta / 120
            ),
            "units"
        )

    def on_shift_wheel(self, event):

        self.canvas.xview_scroll(
            -int(
                event.delta / 120
            ),
            "units"
        )

    # ========================================================
    # 中键平移
    # ========================================================

    def on_middle_down(self, event):

        self.panning = True

        self.pan_last_x = event.x

        self.pan_last_y = event.y

    def on_middle_drag(self, event):

        if not self.panning:
            return

        dx = (
            event.x
            - self.pan_last_x
        )

        dy = (
            event.y
            - self.pan_last_y
        )

        self.canvas.xview_scroll(
            -dx,
            "units"
        )

        self.canvas.yview_scroll(
            -dy,
            "units"
        )

        self.pan_last_x = event.x

        self.pan_last_y = event.y

    def on_middle_up(self, event):

        self.panning = False

    # ========================================================
    # 清理临时文件
    # ========================================================

    def cleanup_temp(self):

        try:

            if self.doc is not None:
                self.doc.close()

        except Exception:
            pass

        self.doc = None

        try:

            if self.temp_dir is not None:
                self.temp_dir.cleanup()

        except Exception:
            pass

        self.temp_dir = None

    # ========================================================
    # 关闭
    # ========================================================

    def on_close(self):

        self.cleanup_temp()

        self.root.destroy()


# ============================================================
# 主程序
# ============================================================

def main():

    root = tk.Tk()

    app = WordFormulaOCR(
        root
    )

    root.mainloop()


if __name__ == "__main__":
    main()
