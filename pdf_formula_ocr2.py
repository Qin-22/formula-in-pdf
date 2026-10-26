import os
import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
from pix2tex.cli import LatexOCR


class PDFFormulaOCR:

    def __init__(self, root):

        self.root = root
        self.root.title("PDF 公式 OCR - LaTeX-OCR")
        self.root.geometry("1500x900")
        self.root.minsize(1100, 700)

        # =========================================================
        # PDF 状态
        # =========================================================

        self.doc = None
        self.pdf_path = None
        self.page_num = 0

        # PDF 缩放
        self.zoom = 1.5
        self.min_zoom = 0.5
        self.max_zoom = 5.0
        self.zoom_step = 0.2

        # 当前 PDF 页面图片
        self.page_image = None
        self.tk_image = None
        self.image_id = None

        # =========================================================
        # 框选状态
        # =========================================================

        # 鼠标按下位置
        self.select_start = None

        # 当前框选矩形
        self.select_rect = None

        # 当前已经确定的框选区域
        self.selected_box = None

        # 框选后的原始图片
        self.selected_image = None

        # 预览图片
        self.preview_tk_image = None

        # =========================================================
        # OCR
        # =========================================================

        self.last_latex = ""

        print("正在加载 LaTeX-OCR 模型...")

        self.ocr_model = LatexOCR()

        print("LaTeX-OCR 模型加载完成")

        # 创建界面
        self.create_ui()

        # 绑定鼠标事件
        self.bind_events()


    # =============================================================
    # 创建 UI
    # =============================================================

    def create_ui(self):

        # =========================================================
        # 顶部工具栏
        # =========================================================

        toolbar = tk.Frame(
            self.root,
            height=45
        )

        toolbar.pack(
            side=tk.TOP,
            fill=tk.X
        )

        # 打开 PDF
        tk.Button(
            toolbar,
            text="打开 PDF",
            command=self.open_pdf,
            width=10
        ).pack(
            side=tk.LEFT,
            padx=5,
            pady=5
        )

        # 上一页
        tk.Button(
            toolbar,
            text="上一页",
            command=self.prev_page,
            width=8
        ).pack(
            side=tk.LEFT,
            padx=2
        )

        # 下一页
        tk.Button(
            toolbar,
            text="下一页",
            command=self.next_page,
            width=8
        ).pack(
            side=tk.LEFT,
            padx=2
        )

        tk.Label(
            toolbar,
            text="页码:"
        ).pack(
            side=tk.LEFT,
            padx=(10, 2)
        )

        self.page_entry = tk.Entry(
            toolbar,
            width=6
        )

        self.page_entry.pack(
            side=tk.LEFT
        )

        tk.Button(
            toolbar,
            text="跳转",
            command=self.jump_page,
            width=6
        ).pack(
            side=tk.LEFT,
            padx=2
        )

        # ---------------------------------------------------------
        # 缩放
        # ---------------------------------------------------------

        tk.Frame(
            toolbar,
            width=15
        ).pack(
            side=tk.LEFT
        )

        tk.Button(
            toolbar,
            text="−",
            command=self.zoom_out,
            width=4
        ).pack(
            side=tk.LEFT
        )

        self.zoom_label = tk.Label(
            toolbar,
            text="150%",
            width=7
        )

        self.zoom_label.pack(
            side=tk.LEFT
        )

        tk.Button(
            toolbar,
            text="+",
            command=self.zoom_in,
            width=4
        ).pack(
            side=tk.LEFT
        )

        tk.Button(
            toolbar,
            text="适应窗口",
            command=self.fit_page,
            width=10
        ).pack(
            side=tk.LEFT,
            padx=5
        )

        tk.Button(
            toolbar,
            text="100%",
            command=self.actual_size,
            width=7
        ).pack(
            side=tk.LEFT
        )

        # 状态
        self.status_label = tk.Label(
            toolbar,
            text="请打开 PDF",
            anchor="w"
        )

        self.status_label.pack(
            side=tk.LEFT,
            padx=15
        )


        # =========================================================
        # 主区域
        # =========================================================

        main_frame = tk.Frame(
            self.root
        )

        main_frame.pack(
            side=tk.TOP,
            fill=tk.BOTH,
            expand=True
        )


        # =========================================================
        # 左侧 PDF
        # =========================================================

        pdf_frame = tk.Frame(
            main_frame,
            bg="#444444"
        )

        pdf_frame.pack(
            side=tk.LEFT,
            fill=tk.BOTH,
            expand=True
        )

        # PDF Canvas
        self.canvas = tk.Canvas(
            pdf_frame,
            bg="#444444",
            highlightthickness=0
        )

        self.canvas.pack(
            side=tk.LEFT,
            fill=tk.BOTH,
            expand=True
        )

        # 垂直滚动条
        v_scroll = tk.Scrollbar(
            pdf_frame,
            orient=tk.VERTICAL,
            command=self.canvas.yview
        )

        v_scroll.pack(
            side=tk.RIGHT,
            fill=tk.Y
        )

        # 水平滚动条
        h_scroll = tk.Scrollbar(
            pdf_frame,
            orient=tk.HORIZONTAL,
            command=self.canvas.xview
        )

        h_scroll.pack(
            side=tk.BOTTOM,
            fill=tk.X
        )

        self.canvas.configure(
            xscrollcommand=h_scroll.set,
            yscrollcommand=v_scroll.set
        )


        # =========================================================
        # 右侧区域
        # =========================================================

        right_frame = tk.Frame(
            main_frame,
            width=430
        )

        right_frame.pack(
            side=tk.RIGHT,
            fill=tk.Y
        )

        right_frame.pack_propagate(False)


        # =========================================================
        # 公式预览标题
        # =========================================================

        tk.Label(
            right_frame,
            text="公式实时预览",
            font=("Microsoft YaHei", 12, "bold")
        ).pack(
            side=tk.TOP,
            pady=(10, 5)
        )


        # =========================================================
        # 预览区域
        # =========================================================

        preview_container = tk.Frame(
            right_frame,
            bd=1,
            relief=tk.SOLID,
            bg="white"
        )

        preview_container.pack(
            side=tk.TOP,
            fill=tk.X,
            padx=10,
            pady=5
        )

        self.preview_canvas = tk.Canvas(
            preview_container,
            height=230,
            bg="white",
            highlightthickness=0
        )

        self.preview_canvas.pack(
            fill=tk.BOTH,
            expand=True
        )

        # 默认提示
        self.preview_text_id = self.preview_canvas.create_text(
            200,
            110,
            text="请在左侧 PDF 中\n拖动鼠标框选公式",
            font=("Microsoft YaHei", 11),
            fill="gray",
            justify=tk.CENTER
        )


        # =========================================================
        # 框选控制按钮
        # =========================================================

        button_frame = tk.Frame(
            right_frame
        )

        button_frame.pack(
            side=tk.TOP,
            fill=tk.X,
            padx=10,
            pady=5
        )

        self.confirm_button = tk.Button(
            button_frame,
            text="确认并 OCR",
            command=self.confirm_selection,
            state=tk.DISABLED,
            height=2
        )

        self.confirm_button.pack(
            side=tk.LEFT,
            fill=tk.X,
            expand=True,
            padx=(0, 4)
        )

        self.cancel_button = tk.Button(
            button_frame,
            text="取消",
            command=self.cancel_selection,
            state=tk.DISABLED,
            height=2
        )

        self.cancel_button.pack(
            side=tk.LEFT,
            fill=tk.X,
            expand=True,
            padx=(4, 0)
        )


        # =========================================================
        # OCR 结果
        # =========================================================

        tk.Label(
            right_frame,
            text="LaTeX OCR 结果",
            font=("Microsoft YaHei", 12, "bold")
        ).pack(
            side=tk.TOP,
            pady=(10, 5)
        )


        latex_frame = tk.Frame(
            right_frame
        )

        latex_frame.pack(
            side=tk.TOP,
            fill=tk.BOTH,
            expand=True,
            padx=10
        )


        self.latex_text = tk.Text(
            latex_frame,
            wrap=tk.WORD,
            font=("Consolas", 11)
        )

        self.latex_text.pack(
            side=tk.LEFT,
            fill=tk.BOTH,
            expand=True
        )


        latex_scroll = tk.Scrollbar(
            latex_frame,
            command=self.latex_text.yview
        )

        latex_scroll.pack(
            side=tk.RIGHT,
            fill=tk.Y
        )

        self.latex_text.configure(
            yscrollcommand=latex_scroll.set
        )


        # =========================================================
        # 复制按钮
        # =========================================================

        tk.Button(
            right_frame,
            text="复制 LaTeX",
            command=self.copy_latex,
            height=2
        ).pack(
            side=tk.BOTTOM,
            fill=tk.X,
            padx=10,
            pady=8
        )


        # =========================================================
        # 操作说明
        # =========================================================

        help_text = (
            "操作说明：\n"
            "• 左键拖动：框选公式\n"
            "• 框选过程中：实时预览\n"
            "• 确认后：点击「确认框选并 OCR」\n"
            "• 中键拖动：移动 PDF\n"
            "• 滚轮：上下移动\n"
            "• Ctrl + 滚轮：缩放 PDF"
        )

        tk.Label(
            right_frame,
            text=help_text,
            justify=tk.LEFT,
            anchor="w"
        ).pack(
            side=tk.BOTTOM,
            fill=tk.X,
            padx=10,
            pady=5
        )


    # =============================================================
    # 绑定事件
    # =============================================================

    def bind_events(self):

        # 左键框选
        self.canvas.bind(
            "<ButtonPress-1>",
            self.start_select
        )

        self.canvas.bind(
            "<B1-Motion>",
            self.update_select
        )

        self.canvas.bind(
            "<ButtonRelease-1>",
            self.end_select
        )

        # 中键拖动
        self.canvas.bind(
            "<ButtonPress-2>",
            self.start_pan
        )

        self.canvas.bind(
            "<B2-Motion>",
            self.do_pan
        )

        # 滚轮
        self.canvas.bind(
            "<MouseWheel>",
            self.mouse_wheel
        )

        # Ctrl + 滚轮
        self.canvas.bind(
            "<Control-MouseWheel>",
            self.ctrl_mouse_wheel
        )


    # =============================================================
    # 打开 PDF
    # =============================================================

    def open_pdf(self):

        path = filedialog.askopenfilename(
            title="选择 PDF",
            filetypes=[
                ("PDF files", "*.pdf"),
                ("All files", "*.*")
            ]
        )

        if not path:
            return

        try:

            if self.doc:
                self.doc.close()

            self.doc = fitz.open(path)

            self.pdf_path = path
            self.page_num = 0

            self.page_entry.delete(
                0,
                tk.END
            )

            self.page_entry.insert(
                0,
                "1"
            )

            self.selected_box = None
            self.selected_image = None

            self.clear_preview()

            self.fit_page()

            self.status_label.config(
                text=(
                    f"{os.path.basename(path)} | "
                    f"共 {len(self.doc)} 页"
                )
            )

        except Exception as e:

            messagebox.showerror(
                "打开 PDF 失败",
                str(e)
            )


    # =============================================================
    # 渲染 PDF 页面
    # =============================================================

    def render_page(self):

        if self.doc is None:
            return

        page = self.doc[self.page_num]

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

        self.image_id = self.canvas.create_image(
            20,
            20,
            image=self.tk_image,
            anchor=tk.NW
        )

        self.canvas.configure(
            scrollregion=(
                0,
                0,
                img.width + 40,
                img.height + 40
            )
        )

        self.zoom_label.config(
            text=f"{int(self.zoom * 100)}%"
        )

        self.page_entry.delete(
            0,
            tk.END
        )

        self.page_entry.insert(
            0,
            str(self.page_num + 1)
        )

        self.status_label.config(
            text=(
                f"{os.path.basename(self.pdf_path)} | "
                f"第 {self.page_num + 1}/{len(self.doc)} 页 | "
                f"缩放 {int(self.zoom * 100)}%"
            )
        )

        # 换页后清除之前框选
        self.selected_box = None
        self.selected_image = None
        self.clear_preview()

        self.confirm_button.config(
            state=tk.DISABLED
        )

        self.cancel_button.config(
            state=tk.DISABLED
        )


    # =============================================================
    # 缩放
    # =============================================================

    def zoom_in(self):

        if self.doc is None:
            return

        new_zoom = self.zoom + self.zoom_step

        if new_zoom <= self.max_zoom:

            self.zoom = new_zoom

            self.render_page()


    def zoom_out(self):

        if self.doc is None:
            return

        new_zoom = self.zoom - self.zoom_step

        if new_zoom >= self.min_zoom:

            self.zoom = new_zoom

            self.render_page()


    def actual_size(self):

        if self.doc is None:
            return

        self.zoom = 1.0

        self.render_page()


    def fit_page(self):

        if self.doc is None:
            return

        page = self.doc[self.page_num]

        rect = page.rect

        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()

        if canvas_width <= 1:
            canvas_width = 1000

        if canvas_height <= 1:
            canvas_height = 700

        available_width = canvas_width - 40
        available_height = canvas_height - 40

        zoom_x = available_width / rect.width
        zoom_y = available_height / rect.height

        self.zoom = min(
            zoom_x,
            zoom_y
        )

        self.zoom = max(
            self.min_zoom,
            min(
                self.zoom,
                self.max_zoom
            )
        )

        self.render_page()


    # =============================================================
    # 翻页
    # =============================================================

    def prev_page(self):

        if self.doc is None:
            return

        if self.page_num > 0:

            self.page_num -= 1

            self.render_page()


    def next_page(self):

        if self.doc is None:
            return

        if self.page_num < len(self.doc) - 1:

            self.page_num += 1

            self.render_page()


    def jump_page(self):

        if self.doc is None:
            return

        try:

            page = int(
                self.page_entry.get()
            )

            if 1 <= page <= len(self.doc):

                self.page_num = page - 1

                self.render_page()

            else:

                messagebox.showwarning(
                    "页码错误",
                    f"请输入 1 到 {len(self.doc)} 之间的页码"
                )

        except ValueError:

            messagebox.showwarning(
                "页码错误",
                "请输入正确的页码"
            )


    # =============================================================
    # 鼠标滚轮
    # =============================================================

    def mouse_wheel(self, event):

        self.canvas.yview_scroll(
            int(-event.delta / 120),
            "units"
        )


    def ctrl_mouse_wheel(self, event):

        if event.delta > 0:
            self.zoom_in()
        else:
            self.zoom_out()


    # =============================================================
    # PDF 平移
    # =============================================================

    def start_pan(self, event):

        self.pan_start = (
            event.x,
            event.y
        )


    def do_pan(self, event):

        if self.pan_start is None:
            return

        self.canvas.scan_dragto(
            event.x,
            event.y,
            gain=1
        )


    # =============================================================
    # 开始框选
    # =============================================================

    def start_select(self, event):

        if self.doc is None:
            return

        x = self.canvas.canvasx(
            event.x
        )

        y = self.canvas.canvasy(
            event.y
        )

        self.select_start = (
            x,
            y
        )

        # 删除之前矩形
        if self.select_rect:

            self.canvas.delete(
                self.select_rect
            )

        # 创建新的矩形
        self.select_rect = self.canvas.create_rectangle(
            x,
            y,
            x,
            y,
            outline="red",
            width=2
        )

        # 开始新的框选
        self.selected_box = None

        self.confirm_button.config(
            state=tk.DISABLED
        )

        self.cancel_button.config(
            state=tk.NORMAL
        )


    # =============================================================
    # 框选过程中
    # =============================================================

    def update_select(self, event):

        if self.select_start is None:
            return

        x = self.canvas.canvasx(
            event.x
        )

        y = self.canvas.canvasy(
            event.y
        )

        x0, y0 = self.select_start

        # 更新红色框
        self.canvas.coords(
            self.select_rect,
            x0,
            y0,
            x,
            y
        )

        # 实时预览
        self.update_preview(
            x0,
            y0,
            x,
            y
        )


    # =============================================================
    # 框选结束
    # =============================================================

    def end_select(self, event):

        if self.select_start is None:
            return

        x0, y0 = self.select_start

        x1 = self.canvas.canvasx(
            event.x
        )

        y1 = self.canvas.canvasy(
            event.y
        )

        left = min(
            x0,
            x1
        )

        top = min(
            y0,
            y1
        )

        right = max(
            x0,
            x1
        )

        bottom = max(
            y0,
            y1
        )

        image_x0 = left - 20
        image_y0 = top - 20

        image_x1 = right - 20
        image_y1 = bottom - 20

        # 防止框选区域太小
        if (
            image_x1 - image_x0 < 10
            or
            image_y1 - image_y0 < 10
        ):

            self.select_start = None

            return

        # 限制在页面图片范围内
        image_x0 = max(
            0,
            image_x0
        )

        image_y0 = max(
            0,
            image_y0
        )

        image_x1 = min(
            self.page_image.width,
            image_x1
        )

        image_y1 = min(
            self.page_image.height,
            image_y1
        )

        # 保存最终框选范围
        self.selected_box = (
            int(image_x0),
            int(image_y0),
            int(image_x1),
            int(image_y1)
        )

        # 保存最终图片
        self.selected_image = self.page_image.crop(
            self.selected_box
        )

        # 再更新一次预览
        self.show_preview(
            self.selected_image
        )

        # 开启确认按钮
        self.confirm_button.config(
            state=tk.NORMAL
        )

        self.cancel_button.config(
            state=tk.NORMAL
        )

        self.select_start = None

        self.status_label.config(
            text="已框选公式，请确认区域后点击「确认并 OCR」"
        )


    # =============================================================
    # 实时预览
    # =============================================================

    def update_preview(
        self,
        x0,
        y0,
        x1,
        y1
    ):

        if self.page_image is None:
            return

        left = min(
            x0,
            x1
        )

        top = min(
            y0,
            y1
        )

        right = max(
            x0,
            x1
        )

        bottom = max(
            y0,
            y1
        )

        image_x0 = int(
            max(
                0,
                left - 20
            )
        )

        image_y0 = int(
            max(
                0,
                top - 20
            )
        )

        image_x1 = int(
            min(
                self.page_image.width,
                right - 20
            )
        )

        image_y1 = int(
            min(
                self.page_image.height,
                bottom - 20
            )
        )

        if (
            image_x1 <= image_x0
            or
            image_y1 <= image_y0
        ):
            return

        try:

            crop = self.page_image.crop(
                (
                    image_x0,
                    image_y0,
                    image_x1,
                    image_y1
                )
            )

            self.show_preview(
                crop
            )

        except Exception:
            pass


    # =============================================================
    # 显示预览图片
    # =============================================================

    def show_preview(self, image):

        if image is None:
            return

        # 预览区域尺寸
        width = self.preview_canvas.winfo_width()
        height = self.preview_canvas.winfo_height()

        if width <= 1:
            width = 400

        if height <= 1:
            height = 230

        # ---------------------------------------------------------
        # 留一点边距
        # ---------------------------------------------------------

        available_width = width - 20
        available_height = height - 20

        # ---------------------------------------------------------
        # 计算等比例缩放
        # ---------------------------------------------------------

        scale_x = available_width / image.width
        scale_y = available_height / image.height

        scale = min(
            scale_x,
            scale_y
        )

        # 至少放大到能看清
        scale = max(
            scale,
            1.0
        )

        new_width = int(
            image.width * scale
        )

        new_height = int(
            image.height * scale
        )

        # 如果太大则限制
        if new_width > available_width:
            scale = available_width / image.width

            new_width = int(
                image.width * scale
            )

            new_height = int(
                image.height * scale
            )

        if new_height > available_height:
            scale = available_height / image.height

            new_width = int(
                image.width * scale
            )

            new_height = int(
                image.height * scale
            )

        # ---------------------------------------------------------
        # 高质量缩放
        # ---------------------------------------------------------

        preview = image.resize(
            (
                max(1, new_width),
                max(1, new_height)
            ),
            Image.Resampling.LANCZOS
        )

        self.preview_tk_image = ImageTk.PhotoImage(
            preview
        )

        # 清除 Canvas
        self.preview_canvas.delete(
            "all"
        )

        # 居中显示
        x = width // 2
        y = height // 2

        self.preview_canvas.create_image(
            x,
            y,
            image=self.preview_tk_image,
            anchor=tk.CENTER
        )


    # =============================================================
    # 清除预览
    # =============================================================

    def clear_preview(self):

        self.preview_canvas.delete(
            "all"
        )

        width = self.preview_canvas.winfo_width()

        if width <= 1:
            width = 400

        self.preview_text_id = self.preview_canvas.create_text(
            width // 2,
            110,
            text="请在左侧 PDF 中\n拖动鼠标框选公式",
            font=("Microsoft YaHei", 11),
            fill="gray",
            justify=tk.CENTER
        )


    # =============================================================
    # 取消框选
    # =============================================================

    def cancel_selection(self):

        self.selected_box = None
        self.selected_image = None

        if self.select_rect:

            self.canvas.delete(
                self.select_rect
            )

            self.select_rect = None

        self.clear_preview()

        self.confirm_button.config(
            state=tk.DISABLED
        )

        self.cancel_button.config(
            state=tk.DISABLED
        )

        self.status_label.config(
            text="已取消框选"


        )


    # =============================================================
    # 确认框选并 OCR
    # =============================================================

    def confirm_selection(self):

        if self.selected_image is None:

            messagebox.showwarning(
                "提示",
                "请先框选一个公式区域"
            )

            return

        self.status_label.config(
            text="正在进行公式 OCR，请稍候..."
        )

        self.root.update()

        try:

            latex = self.ocr_model(
                self.selected_image
            )

            self.last_latex = latex

            self.latex_text.delete(
                "1.0",
                tk.END
            )

            self.latex_text.insert(
                tk.END,
                latex
            )

            self.status_label.config(
                text="OCR 完成"
            )

        except Exception as e:

            messagebox.showerror(
                "OCR 失败",
                str(e)
            )

            self.status_label.config(
                text="OCR 失败"
            )


    # =============================================================
    # 复制 LaTeX
    # =============================================================

    def copy_latex(self):

        latex = self.latex_text.get(
            "1.0",
            tk.END
        ).strip()

        if not latex:
            return

        self.root.clipboard_clear()

        self.root.clipboard_append(
            latex
        )

        self.status_label.config(
            text="LaTeX 已复制到剪贴板"


        )


# =================================================================
# 主程序
# =================================================================

if __name__ == "__main__":

    root = tk.Tk()

    app = PDFFormulaOCR(
        root
    )

    root.mainloop()