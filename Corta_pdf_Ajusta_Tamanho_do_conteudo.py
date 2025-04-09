import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from PIL import Image, ImageTk
import io
import json


class PDFCropper:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Cropper")

        self.doc = None
        self.current_page = 0
        self.crop_rect = None
        self.drag_start = None
        self.rect_id = None
        self.overlay_id = None
        self.original_cropboxes = {}
        self.last_crop_rect = None
        self.page_states = {}
        self.current_item_id = None
        self.pan_start = None

        self.page_size = tk.StringVar(value="A4")
        self.display_scale = 0.75
        self.zoom_level = tk.DoubleVar(value=1.0)
        self.zoom_label_text = tk.StringVar(value="Zoom: 100%")
        self.crop_size_text = tk.StringVar(value="Tamanho do Recorte: N/A")
        self.page_label_text = tk.StringVar(value="Página 1 de 1")

        self.create_widgets()

    def create_widgets(self):
        control_frame = ttk.Frame(self.root)
        control_frame.pack(side=tk.TOP, fill=tk.X)

        ttk.Button(control_frame, text="Abrir PDF",
                   command=self.open_pdf).pack(side=tk.LEFT)
        ttk.Button(control_frame, text="Salvar PDF",
                   command=self.save_pdf).pack(side=tk.LEFT)
        ttk.Button(control_frame, text="Aplicar a todas",
                   command=self.apply_to_all_pages).pack(side=tk.LEFT)
        ttk.Button(control_frame, text="Aplicar às Páginas Selecionadas",
                   command=self.apply_to_selected_pages).pack(side=tk.LEFT)
        ttk.Button(control_frame, text="Exportar Páginas",
                   command=self.export_selected_pages).pack(side=tk.LEFT)
        ttk.Button(control_frame, text="Desfazer Corte",
                   command=self.reset_cropbox).pack(side=tk.LEFT)
        ttk.Button(control_frame, text="Limpar Recorte",
                   command=self.clear_crop).pack(side=tk.LEFT)

        settings_frame = ttk.Frame(self.root)
        settings_frame.pack(side=tk.TOP, fill=tk.X)

        ttk.Label(settings_frame, text="Tamanho alvo:").pack(side=tk.LEFT)
        self.scale_entry = ttk.Entry(settings_frame, width=10)
        self.scale_entry.pack(side=tk.LEFT)

        ttk.Label(settings_frame, text="Zoom:").pack(side=tk.LEFT)
        zoom_slider = ttk.Scale(settings_frame, from_=0.1, to=5.0, orient=tk.HORIZONTAL,
                                variable=self.zoom_level, command=self.update_zoom_from_slider)
        zoom_slider.pack(side=tk.LEFT, padx=5)

        ttk.Label(settings_frame, textvariable=self.zoom_label_text).pack(
            side=tk.LEFT, padx=5)

        ttk.Button(settings_frame, text="Redefinir Zoom",
                   command=self.reset_zoom).pack(side=tk.LEFT, padx=5)

        ttk.Label(settings_frame,
                  text="Páginas para Recorte (ex: 1,3-5):").pack(side=tk.LEFT)
        self.pages_entry = ttk.Entry(settings_frame, width=15)
        self.pages_entry.pack(side=tk.LEFT, padx=5)

        ttk.Button(settings_frame, text="Salvar Configurações",
                   command=self.save_settings).pack(side=tk.LEFT)
        ttk.Button(settings_frame, text="Carregar Configurações",
                   command=self.load_settings).pack(side=tk.LEFT)

        ttk.Button(settings_frame, text="Rotacionar 90°",
                   command=self.rotate_page).pack(side=tk.LEFT, padx=5)

        self.canvas_frame = ttk.Frame(self.root)
        self.canvas_frame.pack(fill=tk.BOTH, expand=False)

        self.canvas = tk.Canvas(
            self.canvas_frame, cursor="cross", bg="white", highlightthickness=0)
        self.canvas.pack(anchor="center")

        navigation_frame = ttk.Frame(self.root)
        navigation_frame.pack(side=tk.TOP, fill=tk.X)

        ttk.Button(navigation_frame, text="◄ Página Anterior",
                   command=self.prev_page).pack(side=tk.LEFT, padx=5)
        ttk.Label(navigation_frame, textvariable=self.page_label_text).pack(
            side=tk.LEFT, padx=5)
        ttk.Button(navigation_frame, text="Página Seguinte ►",
                   command=self.next_page).pack(side=tk.LEFT, padx=5)

        ttk.Label(navigation_frame, textvariable=self.crop_size_text).pack(
            side=tk.LEFT, padx=5)

        self.canvas.bind("<ButtonPress-1>", self.start_drag_or_pan)
        self.canvas.bind("<B1-Motion>", self.do_drag_or_pan)
        self.canvas.bind("<ButtonRelease-1>", self.stop_drag)

        self.canvas.bind("<Control-MouseWheel>", self.on_mousewheel)
        self.canvas.bind("<Control-Button-4>",
                         lambda event: self.on_mousewheel(event, delta=120))
        self.canvas.bind("<Control-Button-5>",
                         lambda event: self.on_mousewheel(event, delta=-120))
        self.root.bind("<Control-plus>", lambda event: self.adjust_zoom(1.1))
        self.root.bind("<Control-minus>", lambda event: self.adjust_zoom(0.9))

        self.root.bind("<Left>", self.prev_page)
        self.root.bind("<Right>", self.next_page)

    def get_target_size(self):
        target_sizes = {
            "A4": (595, 842),
            "A5": (420, 595),
            "Letter": (612, 792)
        }
        return target_sizes[self.page_size.get()]

    def parse_page_selection(self, page_input):
        if not page_input:
            return []
        pages = set()
        try:
            for part in page_input.split(","):
                part = part.strip()
                if "-" in part:
                    start, end = map(int, part.split("-"))
                    pages.update(range(start - 1, end))
                else:
                    pages.add(int(part) - 1)
        except ValueError:
            messagebox.showerror(
                "Erro", "Entrada de páginas inválida. Use o formato '1,3-5'.")
            return []
        return sorted(pages)

    def apply_to_all_pages(self):
        if not self.doc or not self.crop_rect:
            messagebox.showwarning(
                "Aviso", "Nenhum recorte definido para aplicar.")
            return

        page = self.doc.load_page(self.current_page)
        pix = page.get_pixmap(matrix=fitz.Matrix(1, 1))
        pix_width, pix_height = pix.width, pix.height

        media_width, media_height = page.rect.width, page.rect.height

        scale_x = media_width / (pix_width * self.display_scale)
        scale_y = media_height / (pix_height * self.display_scale)

        x1, y1, x2, y2 = self.crop_rect
        crop_rect_real = fitz.Rect(
            x1 * scale_x,
            y1 * scale_y,
            x2 * scale_x,
            y2 * scale_y
        )

        for page_num in range(len(self.doc)):
            page = self.doc.load_page(page_num)
            if page_num not in self.original_cropboxes:
                self.original_cropboxes[page_num] = page.cropbox
            page.set_cropbox(crop_rect_real)

        self.last_crop_rect = self.crop_rect
        self.show_page()

    def apply_to_selected_pages(self):
        if not self.doc or not self.crop_rect:
            messagebox.showwarning(
                "Aviso", "Nenhum recorte definido para aplicar.")
            return

        pages_input = self.pages_entry.get()
        selected_pages = self.parse_page_selection(pages_input)
        if not selected_pages:
            messagebox.showwarning("Aviso", "Nenhuma página selecionada.")
            return

        page = self.doc.load_page(self.current_page)
        pix = page.get_pixmap(matrix=fitz.Matrix(1, 1))
        pix_width, pix_height = pix.width, pix.height

        media_width, media_height = page.rect.width, page.rect.height

        scale_x = media_width / (pix_width * self.display_scale)
        scale_y = media_height / (pix_height * self.display_scale)

        x1, y1, x2, y2 = self.crop_rect
        crop_rect_real = fitz.Rect(
            x1 * scale_x,
            y1 * scale_y,
            x2 * scale_x,
            y2 * scale_y
        )

        for page_num in selected_pages:
            if 0 <= page_num < len(self.doc):
                page = self.doc.load_page(page_num)
                if page_num not in self.original_cropboxes:
                    self.original_cropboxes[page_num] = page.cropbox
                page.set_cropbox(crop_rect_real)

        self.last_crop_rect = self.crop_rect
        self.show_page()

    def export_selected_pages(self):
        if not self.doc:
            messagebox.showwarning("Aviso", "Nenhum PDF carregado.")
            return

        pages_input = tk.simpledialog.askstring(
            "Exportar Páginas", "Digite as páginas para exportar (ex: 1,3-5):")
        if not pages_input:
            return

        selected_pages = self.parse_page_selection(pages_input)
        if not selected_pages:
            messagebox.showwarning(
                "Aviso", "Nenhuma página selecionada para exportar.")
            return

        output_path = filedialog.asksaveasfilename(defaultextension=".pdf")
        if not output_path:
            return

        new_doc = fitz.open()
        base_crop = self.crop_rect if self.crop_rect else (
            0, 0, self.doc[0].rect.width, self.doc[0].rect.height)

        # Definir o DPI desejado para exportação
        desired_dpi = 400
        scale_factor = desired_dpi / 72  # 72 é o DPI base do PyMuPDF

        # Ajustar o target_size para o DPI desejado
        target_size = self.get_target_size()  # Ex.: (595, 842) para A4
        target_width_inch = target_size[0] / 72  # Largura em polegadas
        target_height_inch = target_size[1] / 72  # Altura em polegadas
        target_width_pixels = int(
            target_width_inch * desired_dpi)  # Largura em pixels
        target_height_pixels = int(
            target_height_inch * desired_dpi)  # Altura em pixels
        adjusted_target_size = (target_width_pixels, target_height_pixels)

        page = self.doc.load_page(self.current_page)
        pix = page.get_pixmap(matrix=fitz.Matrix(1, 1))
        pix_width, pix_height = pix.width, pix.height
        media_width, media_height = page.rect.width, page.rect.height

        scale_x = media_width / (pix_width * self.display_scale)
        scale_y = media_height / (pix_height * self.display_scale)

        x1, y1, x2, y2 = base_crop
        crop_rect_real = fitz.Rect(
            x1 * scale_x, y1 * scale_y, x2 * scale_x, y2 * scale_y)

        for page_num in selected_pages:
            if 0 <= page_num < len(self.doc):
                page = self.doc.load_page(page_num)
                cropped_page = page.get_pixmap(
                    clip=crop_rect_real, matrix=fitz.Matrix(scale_factor, scale_factor))
                img = Image.frombytes(
                    "RGB", (cropped_page.width, cropped_page.height), cropped_page.samples)
                resized_img = img.resize(
                    adjusted_target_size, resample=Image.LANCZOS)

                new_page = new_doc.new_page(
                    width=target_size[0], height=target_size[1])
                img_byte_arr = io.BytesIO()
                resized_img.save(img_byte_arr, format="PNG", quality=100)
                img_byte_arr = img_byte_arr.getvalue()
                new_page.insert_image(new_page.rect, stream=img_byte_arr)

        new_doc.save(output_path, deflate=True)
        new_doc.close()
        messagebox.showinfo(
            "Sucesso", f"Páginas exportadas para: {output_path}")

    def save_pdf(self):
        if not self.doc:
            messagebox.showwarning("Aviso", "Nenhum PDF carregado.")
            return

        output_path = filedialog.asksaveasfilename(defaultextension=".pdf")
        if not output_path:
            return

        new_doc = fitz.open()
        base_crop = self.crop_rect if self.crop_rect else (
            0, 0, self.doc[0].rect.width, self.doc[0].rect.height)

        # Definir o DPI desejado
        desired_dpi = 400
        scale_factor = desired_dpi / 72  # 72 é o DPI base do PyMuPDF

        # Ajustar o target_size para o DPI desejado
        target_size = self.get_target_size()  # Ex.: (595, 842) para A4
        target_width_inch = target_size[0] / 72  # Largura em polegadas
        target_height_inch = target_size[1] / 72  # Altura em polegadas
        target_width_pixels = int(
            target_width_inch * desired_dpi)  # Largura em pixels
        target_height_pixels = int(
            target_height_inch * desired_dpi)  # Altura em pixels
        adjusted_target_size = (target_width_pixels, target_height_pixels)

        page = self.doc.load_page(self.current_page)
        pix = page.get_pixmap(matrix=fitz.Matrix(1, 1))
        pix_width, pix_height = pix.width, pix.height
        media_width, media_height = page.rect.width, page.rect.height

        scale_x = media_width / (pix_width * self.display_scale)
        scale_y = media_height / (pix_height * self.display_scale)

        x1, y1, x2, y2 = base_crop
        crop_rect_real = fitz.Rect(
            x1 * scale_x, y1 * scale_y, x2 * scale_x, y2 * scale_y)

        for page_num in range(len(self.doc)):
            page = self.doc.load_page(page_num)
            cropped_page = page.get_pixmap(
                clip=crop_rect_real, matrix=fitz.Matrix(scale_factor, scale_factor))
            img = Image.frombytes(
                "RGB", (cropped_page.width, cropped_page.height), cropped_page.samples)
            resized_img = img.resize(
                adjusted_target_size, resample=Image.LANCZOS)

            new_page = new_doc.new_page(
                width=target_size[0], height=target_size[1])
            img_byte_arr = io.BytesIO()
            resized_img.save(img_byte_arr, format="PNG", quality=100)
            img_byte_arr = img_byte_arr.getvalue()
            new_page.insert_image(new_page.rect, stream=img_byte_arr)

        new_doc.save(output_path, deflate=True)
        new_doc.close()
        messagebox.showinfo("Sucesso", f"PDF salvo em: {output_path}")

    def save_settings(self):
        if not self.doc:
            messagebox.showwarning("Aviso", "Nenhum PDF carregado.")
            return

        output_path = filedialog.asksaveasfilename(
            defaultextension=".json", filetypes=[("JSON Files", "*.json")])
        if not output_path:
            return

        settings = {
            "crop_rect": self.crop_rect,
            "page_states": self.page_states
        }
        with open(output_path, "w") as f:
            json.dump(settings, f, indent=4)
        messagebox.showinfo(
            "Sucesso", f"Configurações salvas em: {output_path}")

    def load_settings(self):
        if not self.doc:
            messagebox.showwarning("Aviso", "Nenhum PDF carregado.")
            return

        input_path = filedialog.askopenfilename(
            filetypes=[("JSON Files", "*.json")])
        if not input_path:
            return

        try:
            with open(input_path, "r") as f:
                settings = json.load(f)
            self.crop_rect = settings.get("crop_rect")
            self.page_states = settings.get("page_states", {})
            self.zoom_level.set(self.page_states.get(
                self.current_page, {"scale_factor": 1.0})["scale_factor"])
            self.show_page()
            messagebox.showinfo(
                "Sucesso", f"Configurações carregadas de: {input_path}")
        except Exception as e:
            messagebox.showerror(
                "Erro", f"Erro ao carregar configurações: {str(e)}")

    def reset_cropbox(self):
        if not self.doc:
            return
        for page_num in range(len(self.doc)):
            page = self.doc.load_page(page_num)
            if page_num in self.original_cropboxes:
                page.set_cropbox(self.original_cropboxes[page_num])
            else:
                page.set_cropbox(page.rect)
            # Corrigido: Inclui a chave "rotation" ao redefinir o estado da página
            self.page_states[page_num] = {
                "scale_factor": 1.0, "x_offset": 0, "y_offset": 0, "rotation": 0}
        self.crop_rect = None
        self.last_crop_rect = None
        if self.rect_id:
            self.canvas.delete(self.rect_id)
            self.rect_id = None
        self.zoom_level.set(1.0)
        self.show_page()

    def clear_crop(self):
        self.crop_rect = None
        if self.rect_id:
            self.canvas.delete(self.rect_id)
            self.rect_id = None
        self.crop_size_text.set("Tamanho do Recorte: N/A")
        self.show_page()

    def open_pdf(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("PDF Files", "*.pdf")])
        if not file_path:
            return
        self.doc = fitz.open(file_path)
        self.original_cropboxes = {
            i: self.doc[i].cropbox for i in range(len(self.doc))}

        self.page_states = {
            i: {"scale_factor": 1.0, "x_offset": 0, "y_offset": 0, "rotation": 0}
            for i in range(len(self.doc))
        }

        self.current_page = 0
        self.zoom_level.set(1.0)
        self.update_page_label()
        self.show_page()

    def show_page(self):
        if not self.doc:
            return

        page = self.doc.load_page(self.current_page)
        page_state = self.page_states.get(
            self.current_page, {"scale_factor": 1.0, "x_offset": 0, "y_offset": 0, "rotation": 0})
        rotation = page_state["rotation"]
        page.set_rotation(rotation)

        pix = page.get_pixmap(matrix=fitz.Matrix(1, 1))
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)

        scale_factor = page_state["scale_factor"]
        x_offset = page_state["x_offset"]
        y_offset = page_state["y_offset"]

        base_width = int(pix.width * self.display_scale)
        base_height = int(pix.height * self.display_scale)

        resized_width = int(base_width * scale_factor)
        resized_height = int(base_height * scale_factor)
        img = img.resize((resized_width, resized_height),
                         Image.Resampling.LANCZOS)
        self.tk_img = ImageTk.PhotoImage(img)

        self.canvas.config(width=base_width, height=base_height)
        self.canvas.config(scrollregion=(0, 0, base_width, base_height))

        self.canvas.delete("all")
        self.current_item_id = self.canvas.create_image(
            x_offset, y_offset, anchor="nw", image=self.tk_img
        )

        if self.crop_rect and not self.rect_id:
            x1, y1, x2, y2 = self.crop_rect
            x1_adj = x1 * scale_factor
            y1_adj = y1 * scale_factor
            x2_adj = x2 * scale_factor
            y2_adj = y2 * scale_factor
            self.rect_id = self.canvas.create_rectangle(
                x1_adj + x_offset, y1_adj + y_offset,
                x2_adj + x_offset, y2_adj + y_offset,
                outline="red"
            )

        control_frame_height = 60
        window_width = base_width + 20
        window_height = base_height + control_frame_height + 20
        self.root.geometry(f"{window_width}x{window_height}")

        self.update_page_label()

    def update_page_label(self):
        if self.doc:
            self.page_label_text.set(
                f"Página {self.current_page + 1} de {len(self.doc)}")
        else:
            self.page_label_text.set("Página 1 de 1")

    def on_mousewheel(self, event, delta=None):
        if not self.doc:
            return

        delta = delta or event.delta
        zoom_factor = 1.1 if delta > 0 else 0.9
        self.adjust_zoom(zoom_factor)
        self.zoom_level.set(
            self.page_states[self.current_page]["scale_factor"])
        self.update_zoom_label()

    def update_zoom_from_slider(self, value):
        if not self.doc:
            return
        page_state = self.page_states.get(
            self.current_page, {"scale_factor": 1.0, "x_offset": 0, "y_offset": 0, "rotation": 0})
        page_state["scale_factor"] = float(value)
        self.page_states[self.current_page] = page_state
        self.show_page()
        self.update_zoom_label()

    def adjust_zoom(self, zoom_factor):
        page_state = self.page_states.get(
            self.current_page, {"scale_factor": 1.0, "x_offset": 0, "y_offset": 0, "rotation": 0})
        current_scale = page_state["scale_factor"]
        new_scale = current_scale * zoom_factor
        new_scale = max(0.1, min(new_scale, 5.0))
        page_state["scale_factor"] = new_scale
        self.page_states[self.current_page] = page_state
        self.show_page()

    def reset_zoom(self):
        if not self.doc:
            return
        self.zoom_level.set(1.0)
        page_state = self.page_states.get(
            self.current_page, {"scale_factor": 1.0, "x_offset": 0, "y_offset": 0, "rotation": 0})
        page_state["scale_factor"] = 1.0
        self.page_states[self.current_page] = page_state
        self.show_page()
        self.update_zoom_label()

    def update_zoom_label(self):
        zoom_percentage = int(self.zoom_level.get() * 100)
        self.zoom_label_text.set(f"Zoom: {zoom_percentage}%")

    def rotate_page(self):
        if not self.doc:
            return
        page_state = self.page_states.get(
            self.current_page, {"scale_factor": 1.0, "x_offset": 0, "y_offset": 0, "rotation": 0})
        current_rotation = page_state.get("rotation", 0)
        new_rotation = (current_rotation + 90) % 360
        page_state["rotation"] = new_rotation
        self.page_states[self.current_page] = page_state
        self.show_page()

    def start_drag_or_pan(self, event):
        if event.state & 0x0001:  # Shift para panning
            self.pan_start = (event.x, event.y)
        else:
            self.drag_start = (event.x, event.y)
            if self.rect_id:
                self.canvas.delete(self.rect_id)
            if self.overlay_id:
                self.canvas.delete(self.overlay_id)
            self.rect_id = self.canvas.create_rectangle(
                event.x, event.y, event.x, event.y, outline="red"
            )

    def do_drag_or_pan(self, event):
        if self.pan_start:
            delta_x = event.x - self.pan_start[0]
            delta_y = event.y - self.pan_start[1]

            page_state = self.page_states.get(
                self.current_page, {"scale_factor": 1.0, "x_offset": 0, "y_offset": 0, "rotation": 0})
            page_state["x_offset"] += delta_x
            page_state["y_offset"] += delta_y
            self.page_states[self.current_page] = page_state

            self.canvas.move(self.current_item_id, delta_x, delta_y)
            if self.rect_id:
                self.canvas.move(self.rect_id, delta_x, delta_y)

            self.pan_start = (event.x, event.y)
        elif self.drag_start:
            x0, y0 = self.drag_start
            x1, y1 = event.x, event.y
            self.canvas.coords(self.rect_id, x0, y0, x1, y1)

            width = abs(x1 - x0)
            height = abs(y1 - y0)
            self.crop_size_text.set(
                f"Tamanho do Recorte: {int(width)}x{int(height)} px")

            if self.overlay_id:
                self.canvas.delete(self.overlay_id)
            page_state = self.page_states.get(
                self.current_page, {"scale_factor": 1.0, "x_offset": 0, "y_offset": 0, "rotation": 0})
            scale_factor = page_state["scale_factor"]
            base_width = int(self.doc.load_page(self.current_page).get_pixmap(
                matrix=fitz.Matrix(1, 1)).width * self.display_scale)
            base_height = int(self.doc.load_page(self.current_page).get_pixmap(
                matrix=fitz.Matrix(1, 1)).height * self.display_scale)
            resized_width = int(base_width * scale_factor)
            resized_height = int(base_height * scale_factor)

            self.overlay_id = self.canvas.create_rectangle(
                0, 0, resized_width, resized_height,
                fill="gray", stipple="gray50", outline=""
            )
            self.canvas.tag_lower(self.overlay_id, self.rect_id)
            self.canvas.create_rectangle(
                min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1),
                fill="", outline=""
            )

    def stop_drag(self, event):
        if self.drag_start:
            x0, y0 = self.drag_start
            x1, y1 = event.x, event.y

            page_state = self.page_states.get(
                self.current_page, {"scale_factor": 1.0, "x_offset": 0, "y_offset": 0, "rotation": 0})
            scale_factor = page_state["scale_factor"]
            x_offset = page_state["x_offset"]
            y_offset = page_state["y_offset"]

            x0_adj = (x0 - x_offset) / scale_factor
            y0_adj = (y0 - y_offset) / scale_factor
            x1_adj = (x1 - x_offset) / scale_factor
            y1_adj = (y1 - y_offset) / scale_factor

            self.crop_rect = (min(x0_adj, x1_adj), min(y0_adj, y1_adj),
                              max(x0_adj, x1_adj), max(y0_adj, y1_adj))
            self.drag_start = None
            if self.overlay_id:
                self.canvas.delete(self.overlay_id)
                self.overlay_id = None
            self.show_page()
        elif self.pan_start:
            self.pan_start = None

    def apply_crop(self, img, crop_rect):
        x1, y1, x2, y2 = [int(v) for v in crop_rect]
        return img.crop((x1, y1, x2, y2))

    def next_page(self, event=None):
        if self.doc and self.current_page < len(self.doc) - 1:
            self.current_page += 1
            self.zoom_level.set(
                self.page_states[self.current_page]["scale_factor"])
            self.show_page()

    def prev_page(self, event=None):
        if self.doc and self.current_page > 0:
            self.current_page -= 1
            self.zoom_level.set(
                self.page_states[self.current_page]["scale_factor"])
            self.show_page()


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFCropper(root)
    root.mainloop()
