import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path

import fitz
from PIL import Image, ImageTk
from pix2tex.cli import LatexOCR


# ==============================
# 配置
# ==============================

OUTPUT_DIR = Path("formula_output")

# PDF页面渲染倍率
ZOOM = 2.0


# ==============================
# PDF → 图片
# ==============================

def render_page(page):
    """
    将 PDF 页面转换为 PIL Image
    """

    matrix = fitz.Matrix(ZOOM, ZOOM)

    pix = page.get_pixmap(
        matrix=matrix,
        alpha=False
    )

    image = Image.frombytes(
        "RGB",
        [pix.width, pix.height],
        pix.samples
    )

    return image


# ==============================
# 主程序
# ==============================

class FormulaOCRApp:

    def __init__(self, root):

        self.root = root

        self.root.title(
            "PDF Formula → LaTeX"
        )

        self.root.geometry(
            "1200x900"
        )

        # ----------------------
        # 状态
        # ----------------------

        self.pdf_path = None
        self.pdf = None

        self.page_index = 0

        self.page_image = None
        self.tk_image = None

        self.start_x = None
        self.start_y = None

        self.rect = None

        self.formula_count = 0

        # ----------------------
        # LaTeX OCR
        # ----------------------

        print("正在加载 LaTeX-OCR 模型...")

        self.latex_ocr = LatexOCR()

        print("模型加载完成")

        # ----------------------
        # UI
        # ----------------------

        self.create_ui()

    # ==========================
    # UI
    # ==========================

    def create_ui(self):

        top_frame = tk.Frame(
            self.root
        )

        top_frame.pack(
            fill=tk.X,
            padx=10,
            pady=10
        )

        open_button = tk.Button(
            top_frame,
            text="打开 PDF",
            command=self.open_pdf,
            width=12
        )

        open_button.pack(
            side=tk.LEFT,
            padx=5
        )

        previous_button = tk.Button(
            top_frame,
            text="上一页",
            command=self.previous_page,
            width=10
        )

        previous_button.pack(
            side=tk.LEFT,
            padx=5
        )

        next_button = tk.Button(
            top_frame,
            text="下一页",
            command=self.next_page,
            width=10
        )

        next_button.pack(
            side=tk.LEFT,
            padx=5
        )

        self.page_label = tk.Label(
            top_frame,
            text="未打开 PDF"
        )

        self.page_label.pack(
            side=tk.LEFT,
            padx=20
        )

        # ----------------------
        # Canvas
        # ----------------------

        canvas_frame = tk.Frame(
            self.root
        )

        canvas_frame.pack(
            fill=tk.BOTH,
            expand=True
        )

        self.canvas = tk.Canvas(
            canvas_frame,
            bg="gray"
        )

        self.canvas.pack(
            fill=tk.BOTH,
            expand=True
        )

        self.canvas.bind(
            "<ButtonPress-1>",
            self.mouse_down
        )

        self.canvas.bind(
            "<B1-Motion>",
            self.mouse_move
        )

        self.canvas.bind(
            "<ButtonRelease-1>",
            self.mouse_up
        )

        # ----------------------
        # LaTeX结果
        # ----------------------

        result_frame = tk.Frame(
            self.root
        )

        result_frame.pack(
            fill=tk.X,
            padx=10,
            pady=10
        )

        tk.Label(
            result_frame,
            text="LaTeX："
        ).pack(
            anchor=tk.W
        )

        self.result_text = tk.Text(
            result_frame,
            height=6,
            font=("Consolas", 12)
        )

        self.result_text.pack(
            fill=tk.X
        )

    # ==========================
    # 打开PDF
    # ==========================

    def open_pdf(self):

        file_path = filedialog.askopenfilename(
            title="选择 PDF",
            filetypes=[
                ("PDF files", "*.pdf")
            ]
        )

        if not file_path:
            return

        try:

            self.pdf_path = Path(
                file_path
            )

            self.pdf = fitz.open(
                file_path
            )

            self.page_index = 0

            self.show_page()

        except Exception as e:

            messagebox.showerror(
                "错误",
                str(e)
            )

    # ==========================
    # 显示页面
    # ==========================

    def show_page(self):

        if self.pdf is None:
            return

        page = self.pdf[
            self.page_index
        ]

        self.page_image = render_page(
            page
        )

        # ----------------------
        # 根据窗口大小缩放
        # ----------------------

        canvas_width = self.canvas.winfo_width()

        canvas_height = self.canvas.winfo_height()

        if canvas_width < 100:
            canvas_width = 1000

        if canvas_height < 100:
            canvas_height = 700

        image = self.page_image.copy()

        image.thumbnail(
            (
                canvas_width - 20,
                canvas_height - 20
            )
        )

        self.display_image = image

        self.tk_image = ImageTk.PhotoImage(
            image
        )

        self.canvas.delete(
            "all"
        )

        self.canvas.create_image(
            10,
            10,
            anchor=tk.NW,
            image=self.tk_image
        )

        self.page_label.config(
            text=(
                f"第 {self.page_index + 1} / "
                f"{len(self.pdf)} 页"
            )
        )

    # ==========================
    # 上一页
    # ==========================

    def previous_page(self):

        if self.pdf is None:
            return

        if self.page_index > 0:

            self.page_index -= 1

            self.show_page()

    # ==========================
    # 下一页
    # ==========================

    def next_page(self):

        if self.pdf is None:
            return

        if self.page_index < len(self.pdf) - 1:

            self.page_index += 1

            self.show_page()

    # ==========================
    # 鼠标按下
    # ==========================

    def mouse_down(self, event):

        self.start_x = event.x
        self.start_y = event.y

        if self.rect:
            self.canvas.delete(
                self.rect
            )

        self.rect = self.canvas.create_rectangle(
            self.start_x,
            self.start_y,
            event.x,
            event.y,
            outline="red",
            width=2
        )

    # ==========================
    # 鼠标拖动
    # ==========================

    def mouse_move(self, event):

        if self.rect:

            self.canvas.coords(
                self.rect,
                self.start_x,
                self.start_y,
                event.x,
                event.y
            )

    # ==========================
    # 鼠标释放
    # ==========================

    def mouse_up(self, event):

        if self.page_image is None:
            return

        end_x = event.x
        end_y = event.y

        x1 = min(
            self.start_x,
            end_x
        )

        y1 = min(
            self.start_y,
            end_y
        )

        x2 = max(
            self.start_x,
            end_x
        )

        y2 = max(
            self.start_y,
            end_y
        )

        # Canvas 显示尺寸
        display_width = self.display_image.width
        display_height = self.display_image.height

        # 原始 PDF 图片尺寸
        original_width = self.page_image.width
        original_height = self.page_image.height

        # 缩放比例
        scale_x = (
            original_width /
            display_width
        )

        scale_y = (
            original_height /
            display_height
        )

        # Canvas 左上角偏移
        x1 -= 10
        x2 -= 10

        y1 -= 10
        y2 -= 10

        # 转换到原始图片坐标
        x1 = int(x1 * scale_x)
        x2 = int(x2 * scale_x)

        y1 = int(y1 * scale_y)
        y2 = int(y2 * scale_y)

        # 防止越界
        x1 = max(0, x1)
        y1 = max(0, y1)

        x2 = min(
            original_width,
            x2
        )

        y2 = min(
            original_height,
            y2
        )

        if x2 - x1 < 20 or y2 - y1 < 20:
            return

        # 裁剪公式
        formula_image = self.page_image.crop(
            (
                x1,
                y1,
                x2,
                y2
            )
        )

        self.recognize_formula(
            formula_image
        )

    # ==========================
    # 公式识别
    # ==========================

    def recognize_formula(
        self,
        formula_image
    ):

        self.result_text.delete(
            "1.0",
            tk.END
        )

        self.result_text.insert(
            tk.END,
            "正在识别公式，请稍候...\n"
        )

        self.root.update()

        try:

            # ----------------------
            # 保存图片
            # ----------------------

            OUTPUT_DIR.mkdir(
                exist_ok=True
            )

            self.formula_count += 1

            page_number = (
                self.page_index + 1
            )

            image_name = (
                f"page_{page_number:03d}"
                f"_formula_{self.formula_count:03d}.png"
            )

            image_path = (
                OUTPUT_DIR /
                image_name
            )

            formula_image.save(
                image_path
            )

            # ----------------------
            # OCR
            # ----------------------

            latex = self.latex_ocr(
                formula_image
            )

            latex = latex.strip()

            # ----------------------
            # 保存 LaTeX
            # ----------------------

            tex_path = (
                image_path.with_suffix(
                    ".tex"
                )
            )

            tex_path.write_text(
                latex,
                encoding="utf-8"
            )

            # ----------------------
            # 显示
            # ----------------------

            self.result_text.delete(
                "1.0",
                tk.END
            )

            self.result_text.insert(
                tk.END,
                latex
            )

            # ----------------------
            # 添加到总文件
            # ----------------------

            all_tex = (
                OUTPUT_DIR /
                "formulas.tex"
            )

            with open(
                all_tex,
                "a",
                encoding="utf-8"
            ) as f:

                f.write(
                    "\n\n"
                    f"% Page {page_number}\n"
                )

                f.write(
                    latex
                )

                f.write(
                    "\n\n"
                )

            print(
                "\n识别结果："
            )

            print(
                latex
            )

            print(
                f"\n图片：{image_path}"
            )

            print(
                f"公式：{tex_path}"
            )

        except Exception as e:

            messagebox.showerror(
                "公式识别失败",
                str(e)
            )


# ==============================
# 程序入口
# ==============================

if __name__ == "__main__":

    root = tk.Tk()

    app = FormulaOCRApp(
        root
    )

    root.mainloop()