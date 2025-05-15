import PyPDF2
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
from PIL import Image, ImageTk

def get_dominant_color(image_path):
    img = Image.open(image_path).convert("RGB")
    img = img.resize((50, 50))  # Speed up
    pixels = list(img.getdata())
    color_counts = {}
    for color in pixels:
        color_counts[color] = color_counts.get(color, 0) + 1
    dominant = max(color_counts, key=color_counts.get)
    # Convert to hex
    return '#%02x%02x%02x' % dominant

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

    # Welcome window with logo and upload button
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
        logo_label = tk.Label(
            welcome_frame, text="📝", font=("Arial", 64), bg=logo_bg
        )
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

    def upload_pdf():
        input_pdf = filedialog.askopenfilename(
            title="Select PDF file", filetypes=[("PDF files", "*.pdf")]
        )
        if not input_pdf:
            messagebox.showinfo("Cancelled", "No PDF selected.")
            return

        fields = get_form_fields(input_pdf)
        if not fields:
            messagebox.showinfo("No Fields", "No editable fields found in the PDF.")
            return

        # Remove welcome frame and show form fields in the same window
        welcome_frame.pack_forget()
        show_form_fields(root, input_pdf, fields)

    upload_btn = tk.Button(
        welcome_frame,
        text="Upload PDF",
        font=("Segoe UI", 14, "bold"),
        bg="#44bd32",
        fg="white",
        activebackground="#4cd137",
        activeforeground="white",
        relief="flat",
        padx=30,
        pady=10,
        command=upload_pdf,
    )
    upload_btn.pack(pady=30)

    root.mainloop()

def show_form_fields(root, input_pdf, fields):
    import os

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
        pady=8
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

        # Scrollable canvas for all pages
        pdf_canvas = tk.Canvas(preview_frame, bg=logo_bg, highlightthickness=0)
        pdf_canvas.pack(expand=True, fill="both", padx=10, pady=(10,0))
        pdf_scrollbar = tk.Scrollbar(preview_frame, orient="vertical", command=pdf_canvas.yview)
        pdf_scrollbar.pack(side="right", fill="y")
        pdf_canvas.configure(yscrollcommand=pdf_scrollbar.set)
        pdf_pages_frame = tk.Frame(pdf_canvas, bg=logo_bg)
        pdf_canvas.create_window((0, 0), window=pdf_pages_frame, anchor="nw")

        controls = tk.Frame(preview_frame, bg=logo_bg)
        controls.pack(fill="x", padx=10, pady=(0,10))

        def render_pdf_images_fitz(zoom):
            doc = fitz.open(input_pdf)
            images = []
            for page_num in range(len(doc)):
                page = doc.load_page(page_num)
                mat = fitz.Matrix(zoom, zoom)
                pix = page.get_pixmap(matrix=mat)
                mode = "RGBA" if pix.alpha else "RGB"
                img = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
                images.append(img)
            return images

        def update_preview_fitz():
            zoom = zoom_levels[zoom_idx]
            # Clear previous images
            for widget in pdf_pages_frame.winfo_children():
                widget.destroy()
            images = render_pdf_images_fitz(zoom)
            for idx, img in enumerate(images):
                photo = ImageTk.PhotoImage(img)
                lbl = tk.Label(pdf_pages_frame, image=photo, bg=logo_bg)
                lbl.image = photo
                lbl.pack(pady=5)
            zoom_label.config(text=f"Zoom: {int(zoom*100)}%")
            pdf_canvas.update_idletasks()
            pdf_canvas.config(scrollregion=pdf_canvas.bbox("all"))

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

        zoom_out_btn = tk.Button(controls, text="−", font=("Segoe UI", 12, "bold"), width=2, command=zoom_out)
        zoom_out_btn.pack(side="left", padx=(0,5))
        zoom_label = tk.Label(controls, text="", font=("Segoe UI", 11), bg=logo_bg)
        zoom_label.pack(side="left")
        zoom_in_btn = tk.Button(controls, text="+", font=("Segoe UI", 12, "bold"), width=2, command=zoom_in)
        zoom_in_btn.pack(side="left", padx=(5,0))

        # Mouse wheel zoom and vertical scroll for preview window
        def on_pdf_canvas_mousewheel(event):
            if event.state & 0x0004:  # Ctrl is pressed
                if event.delta > 0 or getattr(event, 'num', None) == 4:
                    zoom_in()
                elif event.delta < 0 or getattr(event, 'num', None) == 5:
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

        # Bind mousewheel only when cursor is over the preview window
        def bind_pdf_canvas_mousewheel():
            pdf_canvas.bind_all("<MouseWheel>", on_pdf_canvas_mousewheel)
            pdf_canvas.bind_all("<Button-4>", on_pdf_canvas_mousewheel_linux)
            pdf_canvas.bind_all("<Button-5>", on_pdf_canvas_mousewheel_linux)

        def unbind_pdf_canvas_mousewheel():
            pdf_canvas.unbind_all("<MouseWheel>")
            pdf_canvas.unbind_all("<Button-4>")
            pdf_canvas.unbind_all("<Button-5>")

        pdf_canvas.bind("<Enter>", lambda e: bind_pdf_canvas_mousewheel())
        pdf_canvas.bind("<Leave>", lambda e: unbind_pdf_canvas_mousewheel())

        pdf_canvas.config(xscrollincrement=20, yscrollincrement=20)

        update_preview_fitz()
    except Exception as e:
        pdf_label = tk.Label(
            preview_frame,
            text=f"PDF preview not available\n{e}",
            bg=logo_bg,
            fg="#888",
            font=("Segoe UI", 12, "italic"),
            justify="center"
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

    for idx, field in enumerate(fields):
        # Remove the word "fillable" if present in the field name
        display_field = field.replace("fillable", "").replace("Fillable", "").strip()
        tk.Label(
            scroll_frame,
            text=display_field,
            font=label_font,
            anchor="e",
            bg=logo_bg,
            fg="#353b48",
        ).grid(row=idx, column=0, sticky="e", padx=(20, 8), pady=6)
        text_widget = tk.Text(
            scroll_frame,
            width=80,  # doubled from 40 to 80
            height=3,
            font=entry_font,
            bg="#f0f0f0",
            relief="flat",
            highlightthickness=1,
            highlightbackground="#dcdde1",
            wrap="word"
        )
        text_widget.grid(row=idx, column=1, padx=(0, 20), pady=6, sticky="ew")
        scroll_frame.columnconfigure(1, weight=1)
        # Bind resizing on key release and paste
        text_widget.bind("<KeyRelease>", lambda e, tw=text_widget: auto_resize(e, tw))
        text_widget.bind("<<Paste>>", lambda e, tw=text_widget: auto_resize(e, tw))
        entries[field] = text_widget

    def on_submit():
        # Get all text from each Text widget
        field_values = {field: entry.get("1.0", "end-1c") for field, entry in entries.items()}
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

if __name__ == "__main__":
    gui_main()
