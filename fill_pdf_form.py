import PyPDF2
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
from PIL import Image, ImageTk
import re
import json  # <-- add this


def get_dominant_color(image_path):
    img = Image.open(image_path).convert("RGB")
    img = img.resize((50, 50))  # Speed up
    pixels = list(img.getdata())
    color_counts = {}
    for color in pixels:
        color_counts[color] = color_counts.get(color, 0) + 1
    dominant = max(color_counts, key=color_counts.get)
    # Convert to hex
    return "#%02x%02x%02x" % dominant


def get_form_fields(pdf_path):
    with open(pdf_path, "rb") as f:
        reader = PyPDF2.PdfReader(f)
        fields = {}
        if "/AcroForm" in reader.trailer["/Root"]:
            form = reader.trailer["/Root"]["/AcroForm"]
            if "/Fields" in form:
                for field in form["/Fields"]:
                    field_obj = field.get_object()
                    name = field_obj.get("/T")
                    if name:
                        fields[name] = ""
        return fields


def fill_pdf(input_pdf, output_pdf, field_values):
    with open(input_pdf, "rb") as f:
        reader = PyPDF2.PdfReader(f)
        writer = PyPDF2.PdfWriter()
        writer.clone_document_from_reader(reader)
        writer.update_page_form_field_values(writer.pages[0], field_values)
        with open(output_pdf, "wb") as out_f:
            writer.write(out_f)


def gui_main():
    logo_bg = "#f5f6fa"
    try:
        logo_bg = get_dominant_color("sondhitravels_ganeshalogo.jpg")
    except Exception:
        pass

    root = tk.Tk()
    root.title("Sondhi Travel's Form Filler")
    root.geometry("540x600")
    root.resizable(True, True)
    root.configure(bg=logo_bg)

    # Welcome window with logo and application buttons
    welcome_frame = tk.Frame(root, bg=logo_bg)
    welcome_frame.pack(expand=True, fill="both")

    # Use the provided logo image at its original size and ratio
    try:
        logo_img = Image.open("sondhitravels_ganeshalogo.jpg")
        logo_photo = ImageTk.PhotoImage(logo_img)
        logo_label = tk.Label(welcome_frame, image=logo_photo, bg=logo_bg)
        logo_label.image = logo_photo  # Keep a reference
        logo_label.pack(pady=(40, 10))
    except Exception:
        logo_label = tk.Label(welcome_frame, text="📝", font=("Arial", 64), bg=logo_bg)
        logo_label.pack(pady=(40, 10))

    title_label = tk.Label(
        welcome_frame,
        text="Welcome to Sondhi Travel's PDF Form Filler",
        font=("Segoe UI", 18, "bold"),
        fg="#273c75",
        bg=logo_bg,
    )
    title_label.pack(pady=(0, 8))

    subtitle_label = tk.Label(
        welcome_frame,
        text="Easily fill out PDF forms with a few clicks.",
        font=("Segoe UI", 12),
        fg="#353b48",
        bg=logo_bg,
    )
    subtitle_label.pack(pady=(0, 20))

    # Remove upload button, add three application buttons
    def app_btn_handler(app_num):
        if app_num == 1:
            input_pdf = "SchengenVisa/Application for Schengen Visa.pdf"
            # Load Schengen field descriptions
            try:
                with open("SchengenVisa/schengen_checkbox_fields_with_descriptions.json", "r", encoding="utf-8") as f:
                    schengen_checkbox_fields = json.load(f)
                with open("SchengenVisa/schengen_text_fields_with_descriptions.json", "r", encoding="utf-8") as f:
                    schengen_text_fields = json.load(f)
            except Exception as e:
                schengen_checkbox_fields = {}
                schengen_text_fields = {}
        else:
            input_pdf = filedialog.askopenfilename(
                title=f"Select PDF file for Application {app_num}",
                filetypes=[("PDF files", "*.pdf")]
            )
            if not input_pdf:
                messagebox.showinfo("Cancelled", "No PDF selected.")
                return
            schengen_checkbox_fields = {}
            schengen_text_fields = {}
        fields = get_form_fields(input_pdf)
        if not fields:
            messagebox.showinfo("No Fields", "No editable fields found in the PDF.")
            return
        welcome_frame.pack_forget()
        show_form_fields(root, input_pdf, fields, schengen_checkbox_fields, schengen_text_fields)

    btn_frame = tk.Frame(welcome_frame, bg=logo_bg)
    btn_frame.pack(pady=30)

    # Application 1: Always open the specific file
    btn1 = tk.Button(
        btn_frame,
        text="Application for Schengen Visa",
        font=("Segoe UI", 14, "bold"),
        bg="#44bd32",
        fg="white",
        activebackground="#4cd137",
        activeforeground="white",
        relief="flat",
        padx=30,
        pady=10,
        command=lambda: app_btn_handler(1),
    )
    btn1.pack(side="top", pady=10, fill="x", expand=True)

    # Application 2 and 3: Use file dialog
    for i in range(2, 4):
        btn = tk.Button(
            btn_frame,
            text=f"Application {i}",
            font=("Segoe UI", 14, "bold"),
            bg="#44bd32",
            fg="white",
            activebackground="#4cd137",
            activeforeground="white",
            relief="flat",
            padx=30,
            pady=10,
            command=lambda n=i: app_btn_handler(n),
        )
        btn.pack(side="top", pady=10, fill="x", expand=True)

    root.mainloop()


def show_form_fields(root, input_pdf, fields, schengen_checkbox_fields=None, schengen_text_fields=None):
    import os
    import threading

    if schengen_checkbox_fields is None:
        schengen_checkbox_fields = {}
    if schengen_text_fields is None:
        schengen_text_fields = {}

    try:
        logo_bg = get_dominant_color("sondhitravels_ganeshalogo.jpg")
    except Exception:
        logo_bg = "#f5f6fa"

    # Main frame with two columns: left for form, right for PDF preview
    main_frame = tk.Frame(root, bg=logo_bg)
    main_frame.pack(expand=True, fill="both")
    main_frame.rowconfigure(1, weight=1)
    main_frame.columnconfigure(0, weight=3)
    main_frame.columnconfigure(1, weight=2)

    # --- Top: PDF Name Label ---
    pdf_name = os.path.basename(input_pdf)
    pdf_name_label = tk.Label(
        main_frame,
        text=f"PDF: {pdf_name}",
        font=("Segoe UI", 14, "bold"),
        bg=logo_bg,
        fg="#273c75",
        anchor="w",
        padx=10,
        pady=8,
    )
    pdf_name_label.grid(row=0, column=0, columnspan=2, sticky="ew")

    # --- Left: Form Frame ---
    form_frame = tk.Frame(main_frame, bg=logo_bg)
    form_frame.grid(row=1, column=0, sticky="nsew")
    form_frame.rowconfigure(0, weight=1)
    form_frame.columnconfigure(0, weight=1)

    # --- Right: PDF Preview Frame ---
    preview_frame = tk.Frame(main_frame, bg=logo_bg)
    preview_frame.grid(row=1, column=1, sticky="nsew")
    preview_frame.rowconfigure(0, weight=1)
    preview_frame.columnconfigure(0, weight=1)

    try:
        import fitz  # PyMuPDF

        zoom_levels = [0.5, 0.75, 1.0, 1.25, 1.5, 2.0]
        zoom_idx = 2

        doc = fitz.open(input_pdf)
        total_pages = len(doc)
        current_page = [0]  # Use list for mutability in closures

        preview_outer = tk.Frame(preview_frame, bg=logo_bg)
        preview_outer.pack(expand=True, fill="both", padx=10, pady=(10, 0))

        controls = tk.Frame(preview_frame, bg=logo_bg)
        controls.pack(fill="x", padx=10, pady=(0, 10))

        pdf_canvas = tk.Canvas(preview_outer, bg=logo_bg, highlightthickness=0)
        pdf_canvas.grid(row=0, column=1, sticky="nsew")
        preview_outer.grid_rowconfigure(0, weight=1)
        preview_outer.grid_columnconfigure(1, weight=1)

        left_btn = tk.Button(
            preview_outer,
            text="←",
            font=("Segoe UI", 10, "bold"),
            width=4,
            height=1,
            bg="#2980ff",
            fg="white",
            activebackground="#407be7",
            activeforeground="white",
            relief="flat",
            command=lambda: goto_prev_page(),
        )
        left_btn.grid(row=0, column=0, sticky="ns", padx=(0, 5), pady=0)

        right_btn = tk.Button(
            preview_outer,
            text="→",
            font=("Segoe UI", 10, "bold"),
            width=4,
            height=1,
            bg="#2980ff",
            fg="white",
            activebackground="#407be7",
            activeforeground="white",
            relief="flat",
            command=lambda: goto_next_page(),
        )
        right_btn.grid(row=0, column=2, sticky="ns", padx=(5, 0), pady=0)

        # --- Centered page navigation controls ---
        nav_frame = tk.Frame(controls, bg=logo_bg)
        nav_frame.pack(side="top", expand=True)

        page_entry = tk.Entry(nav_frame, width=5, font=("Segoe UI", 11))
        page_entry.pack(side="left", padx=(0, 5))

        def goto_page(event=None):
            try:
                page = int(page_entry.get()) - 1
                if 0 <= page < total_pages:
                    current_page[0] = page
                    update_preview_fitz()
            except Exception:
                pass  # ignore invalid input

        page_entry.bind("<Return>", goto_page)

        page_label = tk.Label(nav_frame, text="", font=("Segoe UI", 11), bg=logo_bg)
        page_label.pack(side="left", padx=(0, 10))

        zoom_out_btn = tk.Button(
            controls,
            text="−",
            font=("Segoe UI", 12, "bold"),
            width=2,
            command=lambda: zoom_out(),
        )
        zoom_out_btn.pack(side="right", padx=(0, 5))
        zoom_label = tk.Label(controls, text="", font=("Segoe UI", 11), bg=logo_bg)
        zoom_label.pack(side="right")
        zoom_in_btn = tk.Button(
            controls,
            text="+",
            font=("Segoe UI", 12, "bold"),
            width=2,
            command=lambda: zoom_in(),
        )
        zoom_in_btn.pack(side="right", padx=(5, 0))

        # --- Fast PDF rendering with threading ---
        pdf_render_lock = threading.Lock()
        pdf_render_cache = {}

        def render_pdf_image_fitz(zoom, page_num):
            cache_key = (zoom, page_num)
            if cache_key in pdf_render_cache:
                return pdf_render_cache[cache_key]
            page = doc.load_page(page_num)
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat)
            mode = "RGBA" if pix.alpha else "RGB"
            img = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
            pdf_render_cache[cache_key] = img
            return img

        def update_preview_fitz():
            zoom = zoom_levels[zoom_idx]
            page_num = current_page[0]
            def render_and_update():
                with pdf_render_lock:
                    img = render_pdf_image_fitz(zoom, page_num)
                    photo = ImageTk.PhotoImage(img)
                    def update_canvas():
                        pdf_canvas.delete("all")
                        canvas_width = pdf_canvas.winfo_width()
                        canvas_height = pdf_canvas.winfo_height()
                        x = max((canvas_width - img.width) // 2, 0)
                        y = max((canvas_height - img.height) // 2, 0)
                        pdf_canvas.create_image(x, y, anchor="nw", image=photo)
                        pdf_canvas.image = photo
                        pdf_canvas.config(
                            scrollregion=(
                                0,
                                0,
                                max(canvas_width, img.width),
                                max(canvas_height, img.height),
                            )
                        )
                        zoom_label.config(text=f"Zoom: {int(zoom*100)}%")
                        page_label.config(text=f"Page {page_num+1} of {total_pages}")
                        page_entry.delete(0, tk.END)
                        page_entry.insert(0, str(page_num + 1))
                    root.after(0, update_canvas)
            threading.Thread(target=render_and_update, daemon=True).start()

        def goto_prev_page():
            if current_page[0] > 0:
                current_page[0] -= 1
                update_preview_fitz()

        def goto_next_page():
            if current_page[0] < total_pages - 1:
                current_page[0] += 1
                update_preview_fitz()

        def goto_page(event=None):
            try:
                page = int(page_entry.get()) - 1
                if 0 <= page < total_pages:
                    current_page[0] = page
                    update_preview_fitz()
            except Exception:
                pass  # ignore invalid input

        def zoom_in():
            nonlocal zoom_idx
            if zoom_idx < len(zoom_levels) - 1:
                zoom_idx += 1
                update_preview_fitz()

        def zoom_out():
            nonlocal zoom_idx
            if zoom_idx > 0:
                zoom_idx -= 1
                update_preview_fitz()

        # Mouse wheel zoom and vertical scroll for preview window
        def on_pdf_canvas_mousewheel(event):
            if event.state & 0x0004:  # Ctrl is pressed
                if event.delta > 0 or getattr(event, "num", None) == 4:
                    zoom_in()
                elif event.delta < 0 or getattr(event, "num", None) == 5:
                    zoom_out()
            else:
                pdf_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        def on_pdf_canvas_mousewheel_linux(event):
            if event.state & 0x0004:  # Ctrl is pressed
                if event.num == 4:
                    zoom_in()
                elif event.num == 5:
                    zoom_out()
            else:
                if event.state & 0x0001:  # Shift for horizontal
                    pdf_canvas.xview_scroll(-1 if event.num == 4 else 1, "units")
                else:
                    pdf_canvas.yview_scroll(-1 if event.num == 4 else 1, "units")

        pdf_canvas.bind("<Configure>", lambda e: update_preview_fitz())
        pdf_canvas.bind("<Enter>", lambda e: bind_pdf_canvas_mousewheel())
        pdf_canvas.bind("<Leave>", lambda e: unbind_pdf_canvas_mousewheel())

        def bind_pdf_canvas_mousewheel():
            pdf_canvas.bind_all("<MouseWheel>", on_pdf_canvas_mousewheel)
            pdf_canvas.bind_all("<Button-4>", on_pdf_canvas_mousewheel_linux)
            pdf_canvas.bind_all("<Button-5>", on_pdf_canvas_mousewheel_linux)

        def unbind_pdf_canvas_mousewheel():
            pdf_canvas.unbind_all("<MouseWheel>")
            pdf_canvas.unbind_all("<Button-4>")
            pdf_canvas.unbind_all("<Button-5>")

        pdf_canvas.config(xscrollincrement=20, yscrollincrement=20)

        update_preview_fitz()
    except Exception as e:
        pdf_label = tk.Label(
            preview_frame,
            text=f"PDF preview not available\n{e}",
            bg=logo_bg,
            fg="#888",
            font=("Segoe UI", 12, "italic"),
            justify="center",
        )
        pdf_label.pack(expand=True, fill="both", padx=10, pady=10)

    # Exit the program when the user presses the X icon
    def on_close():
        root.destroy()
        exit(0)

    root.protocol("WM_DELETE_WINDOW", on_close)

    # Scrollable canvas and frame
    canvas = tk.Canvas(form_frame, bg=logo_bg, highlightthickness=0)
    scrollbar = tk.Scrollbar(form_frame, orient="vertical", command=canvas.yview)
    scroll_frame = tk.Frame(canvas, bg=logo_bg)

    scroll_frame.bind(
        "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.grid(row=0, column=0, sticky="nsew")
    scrollbar.grid(row=0, column=1, sticky="ns")

    # Enable mousewheel scrolling when hovering over the canvas
    def _on_mousewheel(event):
        if event.delta:
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        else:  # For Linux systems
            if event.num == 4:
                canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                canvas.yview_scroll(1, "units")

    canvas.bind_all("<MouseWheel>", _on_mousewheel)  # Windows & Mac
    canvas.bind_all("<Button-4>", _on_mousewheel)  # Linux scroll up
    canvas.bind_all("<Button-5>", _on_mousewheel)  # Linux scroll down

    entries = {}

    label_font = ("Segoe UI", 11)
    entry_font = ("Segoe UI", 11)

    def auto_resize(event, text_widget):
        # Get number of lines and max line length
        content = text_widget.get("1.0", "end-1c")
        lines = content.split("\n")
        num_lines = len(lines)
        max_line_len = max((len(line) for line in lines), default=1)
        # Set min/max for usability
        new_height = max(3, min(num_lines, 20))
        new_width = max(40, min(max_line_len + 2, 120))
        text_widget.config(height=new_height, width=new_width)

    # --- Get field page mapping and type/options for reference ---
    field_page_map = {}
    field_types_map = {}
    field_options_map = {}
    try:
        with open(input_pdf, "rb") as f:
            reader = PyPDF2.PdfReader(f)
            if "/AcroForm" in reader.trailer["/Root"]:
                form = reader.trailer["/Root"]["/AcroForm"]
                if "/Fields" in form:
                    for field_obj in form["/Fields"]:
                        obj = field_obj.get_object()
                        name = obj.get("/T")
                        page_num = None
                        try:
                            if "/P" in obj:
                                page_obj = obj["/P"]
                                for idx, page in enumerate(reader.pages):
                                    if page_obj == page:
                                        page_num = idx + 1
                                        break
                        except Exception:
                            page_num = None
                        if name:
                            field_page_map[name] = page_num
                            ft = obj.get("/FT")
                            if ft == "/Ch":
                                field_types_map[name] = "select"
                                opts = obj.get("/Opt")
                                if opts:
                                    options = []
                                    for o in opts:
                                        if isinstance(o, str):
                                            options.append(o)
                                        elif hasattr(o, "get_object"):
                                            options.append(str(o.get_object()))
                                    field_options_map[name] = options
                            else:
                                field_types_map[name] = "text"
    except Exception:
        pass

    def humanize_field_name(field):
        # Replace underscores and dashes with space
        s = field.replace("_", " ").replace("-", " ")
        # Insert space before capital letters (except first)
        s = re.sub(r'(?<!^)(?=[A-Z])', ' ', s)
        # Insert space between letters and digits
        s = re.sub(r'([a-zA-Z])(\d)', r'\1 \2', s)
        s = re.sub(r'(\d)([a-zA-Z])', r'\1 \2', s)
        # Remove extra spaces and capitalize each word
        s = ' '.join(s.split())
        return s.title()

    # --- Sort fields by page number and order in PDF ---
    sorted_fields = []
    try:
        with open(input_pdf, "rb") as f:
            reader = PyPDF2.PdfReader(f)
            if "/AcroForm" in reader.trailer["/Root"]:
                form = reader.trailer["/Root"]["/AcroForm"]
                if "/Fields" in form:
                    for field_obj in form["/Fields"]:
                        obj = field_obj.get_object()
                        name = obj.get("/T")
                        if name in fields:
                            sorted_fields.append(name)
    except Exception:
        sorted_fields = list(fields.keys())

    # Group fields by page number, preserving order, and ignore None page numbers
    page_fields = {}
    for field in sorted_fields:
        page = field_page_map.get(field, 1)
        if page is None:
            continue  # skip fields with unknown page
        page_fields.setdefault(page, []).append(field)

    # --- Render all fields grouped by page, page 1 first, then 2, etc. ---
    row_idx = 0

    # --- For Schengen Visa, build reverse lookup for checkbox fields ---
    schengen_checkbox_lookup = {}
    schengen_checkbox_field_to_group = {}
    for group, mapping in schengen_checkbox_fields.items():
        for field_code, label in mapping.items():
            schengen_checkbox_lookup[field_code] = label
            schengen_checkbox_field_to_group[field_code] = group

    # For mandatory fields: mark all as mandatory for Schengen, or use a rule
    def is_mandatory(field_name, description):
        # You can customize this rule as needed
        # For now, mark all Schengen fields as mandatory
        return bool(schengen_checkbox_fields or schengen_text_fields)

    # --- Group checkboxes by group for Schengen ---
    checkbox_groups = {}
    for group, mapping in schengen_checkbox_fields.items():
        for field_code in mapping:
            checkbox_groups.setdefault(group, []).append(field_code)

    # --- Track which checkboxes have been rendered ---
    rendered_checkbox_groups = set()

    for page_num in sorted(page_fields.keys()):
        for field in page_fields[page_num]:
            # --- Schengen logic ---
            if schengen_checkbox_fields and field in schengen_checkbox_lookup:
                group = schengen_checkbox_field_to_group[field]
                if group in rendered_checkbox_groups:
                    continue  # Already rendered this group
                rendered_checkbox_groups.add(group)
                options = schengen_checkbox_fields[group]
                display_field = group.replace("_", " ").title()
                # Use * for mandatory
                label_text = f"{display_field} *" if is_mandatory(field, display_field) else display_field
                page_ref = field_page_map.get(field)
                if page_ref:
                    label_text = f"{label_text} (Page {page_ref})"
                tk.Label(
                    scroll_frame,
                    text=label_text,
                    font=("Segoe UI", 11, "bold"),
                    anchor="e",
                    bg=logo_bg,
                    fg="#353b48",
                ).grid(row=row_idx, column=0, sticky="e", padx=(20, 8), pady=6)
                # Render checkboxes for all options in this group
                var_dict = {}
                cb_frame = tk.Frame(scroll_frame, bg=logo_bg)
                for opt_field, opt_label in options.items():
                    var = tk.BooleanVar()
                    cb = tk.Checkbutton(
                        cb_frame,
                        text=opt_label,
                        variable=var,
                        font=("Segoe UI", 10),
                        bg="#f0f0f0",
                        selectcolor="#44bd32",
                        anchor="w"
                    )
                    cb.pack(side="left", padx=2, pady=2)
                    var_dict[opt_field] = var
                cb_frame.grid(row=row_idx, column=1, padx=(0, 20), pady=6, sticky="w")
                entries[group] = var_dict
                row_idx += 1
                continue

            # Schengen text fields
            if schengen_text_fields and field in schengen_text_fields:
                display_field = schengen_text_fields[field]
                label_text = f"{display_field} *" if is_mandatory(field, display_field) else display_field
                page_ref = field_page_map.get(field)
                if page_ref:
                    label_text = f"{label_text} (Page {page_ref})"
                tk.Label(
                    scroll_frame,
                    text=label_text,
                    font=("Segoe UI", 11, "bold"),
                    anchor="e",
                    bg=logo_bg,
                    fg="#353b48",
                ).grid(row=row_idx, column=0, sticky="e", padx=(20, 8), pady=6)
                text_widget = tk.Text(
                    scroll_frame,
                    width=80,
                    height=3,
                    font=entry_font,
                    bg="#f0f0f0",
                    relief="flat",
                    highlightthickness=1,
                    highlightbackground="#dcdde1",
                    wrap="word",
                )
                text_widget.grid(row=row_idx, column=1, padx=(0, 20), pady=6, sticky="ew")
                scroll_frame.columnconfigure(1, weight=1)
                text_widget.bind("<KeyRelease>", lambda e, tw=text_widget: auto_resize(e, tw))
                text_widget.bind("<<Paste>>", lambda e, tw=text_widget: auto_resize(e, tw))
                entries[field] = text_widget
                row_idx += 1
                continue

            # --- fallback: original logic ---
            display_field = humanize_field_name(
                field.replace("fillable", "").replace("Fillable", "").strip()
            )
            page_ref = field_page_map.get(field)
            if page_ref:
                display_field = f"{display_field} (Page {page_ref})"
            tk.Label(
                scroll_frame,
                text=display_field,
                font=("Segoe UI", 11, "bold"),
                anchor="e",
                bg=logo_bg,
                fg="#353b48",
            ).grid(row=row_idx, column=0, sticky="e", padx=(20, 8), pady=6)

            field_type = field_types_map.get(field, "text")
            if field_type == "select" and field_options_map.get(field):
                var = tk.StringVar()
                options = field_options_map[field]
                select_frame = tk.Frame(scroll_frame, bg=logo_bg)
                for opt in options:
                    cb = tk.Radiobutton(
                        select_frame,
                        text=opt,
                        variable=var,
                        value=opt,
                        indicatoron=False,
                        width=8,
                        font=("Segoe UI", 10),
                        bg="#f0f0f0",
                        selectcolor="#44bd32",
                        relief="ridge"
                    )
                    cb.pack(side="left", padx=2, pady=2)
                select_frame.grid(row=row_idx, column=1, padx=(0, 20), pady=6, sticky="w")
                entries[field] = var
            else:
                text_widget = tk.Text(
                    scroll_frame,
                    width=80,
                    height=3,
                    font=entry_font,
                    bg="#f0f0f0",
                    relief="flat",
                    highlightthickness=1,
                    highlightbackground="#dcdde1",
                    wrap="word",
                )
                text_widget.grid(row=row_idx, column=1, padx=(0, 20), pady=6, sticky="ew")
                scroll_frame.columnconfigure(1, weight=1)
                text_widget.bind("<KeyRelease>", lambda e, tw=text_widget: auto_resize(e, tw))
                text_widget.bind("<<Paste>>", lambda e, tw=text_widget: auto_resize(e, tw))
                entries[field] = text_widget
            row_idx += 1

    def on_submit():
        # Get all text from each Text widget
        field_values = {}
        # For Schengen, handle checkboxes as group
        if schengen_checkbox_fields:
            # Add checkbox values as checked/unchecked
            for group, var_dict in entries.items():
                if group in schengen_checkbox_fields:
                    for field_code, var in var_dict.items():
                        field_values[field_code] = "Yes" if var.get() else ""
            # Add text fields
            for field, widget in entries.items():
                if field in schengen_text_fields:
                    field_values[field] = widget.get("1.0", "end-1c")
        else:
            # fallback: original logic
            field_values = {
                field: entry.get("1.0", "end-1c") for field, entry in entries.items()
                if hasattr(entry, "get")
            }
            # Add select/radio fields
            for field, entry in entries.items():
                if hasattr(entry, "get") and not isinstance(entry, tk.Text):
                    continue
                if isinstance(entry, tk.StringVar):
                    field_values[field] = entry.get()
        output_pdf = filedialog.asksaveasfilename(
            title="Save filled PDF as",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
        )
        if not output_pdf:
            messagebox.showinfo("Cancelled", "No output file selected.")
            return
        fill_pdf(input_pdf, output_pdf, field_values)
        messagebox.showinfo("Success", f"Filled PDF saved as:\n{output_pdf}")
        root.destroy()
        exit(0)

    submit_btn = tk.Button(
        form_frame,
        text="Submit",
        font=("Segoe UI", 13, "bold"),
        bg="#273c75",
        fg="white",
        activebackground="#40739e",
        activeforeground="white",
        relief="flat",
        padx=20,
        pady=8,
        command=on_submit,
    )
    submit_btn.grid(row=1, column=0, columnspan=2, pady=18, sticky="ew")

    # Allow resizing
    form_frame.grid_rowconfigure(0, weight=1)
    form_frame.grid_columnconfigure(0, weight=1)
    scroll_frame.grid_columnconfigure(1, weight=1)

    # --- Mousewheel scroll logic for both sides ---
    mouse_side = {"left": False, "right": False}

    def on_entry_enter(event):
        mouse_side["left"] = True

    def on_entry_leave(event):
        mouse_side["left"] = False

    def on_pdf_enter(event):
        mouse_side["right"] = True

    def on_pdf_leave(event):
        mouse_side["right"] = False

    # Use form_frame (not just canvas) to detect mouse on left side
    form_frame.bind("<Enter>", on_entry_enter)
    form_frame.bind("<Leave>", on_entry_leave)
    pdf_canvas.bind("<Enter>", on_pdf_enter)
    pdf_canvas.bind("<Leave>", on_pdf_leave)

    def on_global_mousewheel(event):
        # Windows/Mac
        if mouse_side["right"]:
            img = getattr(pdf_canvas, "image", None)
            if img and (img.width() > pdf_canvas.winfo_width() or img.height() > pdf_canvas.winfo_height()):
                pdf_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        elif mouse_side["left"]:
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def on_global_mousewheel_linux(event):
        # Linux
        if mouse_side["right"]:
            img = getattr(pdf_canvas, "image", None)
            if img and (img.width() > pdf_canvas.winfo_width() or img.height() > pdf_canvas.winfo_height()):
                if event.num == 4:
                    pdf_canvas.yview_scroll(-1, "units")
                elif event.num == 5:
                    pdf_canvas.yview_scroll(1, "units")
        elif mouse_side["left"]:
            if event.num == 4:
                canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                canvas.yview_scroll(1, "units")

    # Bind toplevel for global mousewheel ONCE, persistently
    root.bind_all("<MouseWheel>", on_global_mousewheel)
    root.bind_all("<Button-4>", on_global_mousewheel_linux)
    root.bind_all("<Button-5>", on_global_mousewheel_linux)

    # --- Footer ---
    footer = tk.Label(
        root,
        text="Created with ♥ by Sukant Sondhi",
        font=("Segoe UI", 10, "italic"),
        fg="#273c75",
        bg=logo_bg,
        anchor="center",
        pady=6,
    )
    footer.pack(side="bottom", fill="x")


if __name__ == "__main__":
    gui_main()
