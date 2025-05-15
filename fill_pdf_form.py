import PyPDF2
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
from PIL import Image, ImageTk


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
    root = tk.Tk()
    root.title("Sondhi Travel's Form Filler")
    root.geometry("540x600")
    root.resizable(True, True)
    root.configure(bg="#f5f6fa")

    # Welcome window with logo and upload button
    welcome_frame = tk.Frame(root, bg="#f5f6fa")
    welcome_frame.pack(expand=True, fill="both")

    # Use the provided logo image at its original size and ratio
    try:
        logo_img = Image.open("sondhitravels_ganeshalogo.jpg")
        logo_photo = ImageTk.PhotoImage(logo_img)
        logo_label = tk.Label(welcome_frame, image=logo_photo, bg="#f5f6fa")
        logo_label.image = logo_photo  # Keep a reference
        logo_label.pack(pady=(40, 10))
    except Exception:
        logo_label = tk.Label(
            welcome_frame, text="📝", font=("Arial", 64), bg="#f5f6fa"
        )
        logo_label.pack(pady=(40, 10))

    title_label = tk.Label(
        welcome_frame,
        text="Welcome to Sondhi Travel's PDF Form Filler",
        font=("Segoe UI", 18, "bold"),
        fg="#273c75",
        bg="#f5f6fa",
    )
    title_label.pack(pady=(0, 8))

    subtitle_label = tk.Label(
        welcome_frame,
        text="Easily fill out PDF forms with a few clicks.",
        font=("Segoe UI", 12),
        fg="#353b48",
        bg="#f5f6fa",
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
    form_frame = tk.Frame(root, bg="#f5f6fa")
    form_frame.pack(expand=True, fill="both")
    form_frame.rowconfigure(0, weight=1)
    form_frame.columnconfigure(0, weight=1)

    # Exit the program when the user presses the X icon
    def on_close():
        root.destroy()
        exit(0)

    root.protocol("WM_DELETE_WINDOW", on_close)

    # Scrollable canvas and frame
    canvas = tk.Canvas(form_frame, bg="#f5f6fa", highlightthickness=0)
    scrollbar = tk.Scrollbar(form_frame, orient="vertical", command=canvas.yview)
    scroll_frame = tk.Frame(canvas, bg="#f5f6fa")

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
        tk.Label(
            scroll_frame,
            text=field,
            font=label_font,
            anchor="e",
            bg="#f5f6fa",
            fg="#353b48",
        ).grid(row=idx, column=0, sticky="e", padx=(20, 8), pady=6)
        text_widget = tk.Text(
            scroll_frame,
            width=40,
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
