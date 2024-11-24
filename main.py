import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import PyPDF2
import os
from PIL import Image
from pdf2image import convert_from_path

class PDFToolsApp:
    def __init__(self, master):
        self.master = master
        master.title("PDF Tools")

        # Create "pdf" folder if it doesn't exist
        self.pdf_folder = os.path.join(os.getcwd(), "pdf")
        if not os.path.exists(self.pdf_folder):
            os.makedirs(self.pdf_folder)

        # Notebook for different functionalities
        self.notebook = ttk.Notebook(master)
        self.notebook.pack(fill="both", expand=True)

        # --- Merge Tab ---
        self.merge_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.merge_frame, text="Merge PDFs")

        self.merge_files = []
        self.merge_list_box = tk.Listbox(self.merge_frame)
        self.merge_list_box.pack(fill="both", expand=True)


        add_button = ttk.Button(self.merge_frame, text="Add Files", command=self.add_files_merge)
        add_button.pack()

        remove_button = ttk.Button(self.merge_frame, text="Remove Selected", command=self.remove_selected_merge)
        remove_button.pack()

        merge_button = ttk.Button(self.merge_frame, text="Merge", command=self.merge_pdfs)
        merge_button.pack()

        self.merge_list_box.bind("<ButtonRelease-1>", self.select_merge_file)
        self.merge_list_box.bind("<B1-Motion>", self.move_merge_file)

        self.selected_merge_file_index = None

        # --- Split Tab --- (More advanced - allows splitting by page ranges)
        self.split_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.split_frame, text="Split PDF")

        self.split_file_path = tk.StringVar()  # Store the selected file path
        split_file_label = ttk.Label(self.split_frame, text="Select PDF:")
        split_file_label.grid(row=0, column=0, sticky="w")

        split_file_entry = ttk.Entry(self.split_frame, textvariable=self.split_file_path, width=40) #, state="readonly"
        split_file_entry.grid(row=0, column=1, padx=5)


        split_file_button = ttk.Button(self.split_frame, text="Browse", command=self.browse_split_file)
        split_file_button.grid(row=0, column=2)


        split_ranges_label = ttk.Label(self.split_frame, text="Page Ranges (e.g., 1-3,5,7-9):")
        split_ranges_label.grid(row=1, column=0, sticky="w")
        self.split_ranges_entry = ttk.Entry(self.split_frame)
        self.split_ranges_entry.grid(row=1, column=1, padx=5, columnspan=2)

        split_button = ttk.Button(self.split_frame, text="Split", command=self.split_pdf)
        split_button.grid(row=2, column=0, columnspan=3, pady=10)

        # --- Rotate Tab ---
        self.rotate_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.rotate_frame, text="Rotate PDF")

        self.rotate_file_path = tk.StringVar()
        rotate_file_label = ttk.Label(self.rotate_frame, text="Select PDF:")
        rotate_file_label.grid(row=0, column=0, sticky="w")

        rotate_file_entry = ttk.Entry(self.rotate_frame, textvariable=self.rotate_file_path, width=40)
        rotate_file_entry.grid(row=0, column=1, columnspan=2, padx=5)

        rotate_file_button = ttk.Button(self.rotate_frame, text="Browse", command=self.browse_rotate_file)
        rotate_file_button.grid(row=0, column=3)

        rotate_angle_label = ttk.Label(self.rotate_frame, text="Rotation Angle:")
        rotate_angle_label.grid(row=1, column=0, sticky="w")

        self.rotate_90_button = ttk.Button(self.rotate_frame, text="90", command=lambda: self.set_rotation_angle(90))
        self.rotate_90_button.grid(row=2, column=0, padx=5, pady=5)

        self.rotate_180_button = ttk.Button(self.rotate_frame, text="180", command=lambda: self.set_rotation_angle(180))
        self.rotate_180_button.grid(row=2, column=1, padx=5, pady=5)

        self.rotate_270_button = ttk.Button(self.rotate_frame, text="270", command=lambda: self.set_rotation_angle(270))
        self.rotate_270_button.grid(row=2, column=2, padx=5, pady=5)

        self.rotation_angle = None  # Store the selected rotation angle

        rotate_button = ttk.Button(self.rotate_frame, text="Rotate", command=self.rotate_pdf)
        rotate_button.grid(row=2, column=3, padx=5, pady=5)

        # --- Extract Text Tab ---
        self.extract_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.extract_frame, text="Extract Text")

        self.extract_file_path = tk.StringVar()
        extract_file_label = ttk.Label(self.extract_frame, text="Select PDF:")
        extract_file_label.grid(row=0, column=0, sticky="w")

        extract_file_entry = ttk.Entry(self.extract_frame, textvariable=self.extract_file_path, width=40)
        extract_file_entry.grid(row=0, column=1, padx=5)

        extract_file_button = ttk.Button(self.extract_frame, text="Browse", command=self.browse_extract_file)
        extract_file_button.grid(row=0, column=2)

        extract_format_label = ttk.Label(self.extract_frame, text="Select Format:")
        extract_format_label.grid(row=1, column=0, sticky="w")

        self.extract_format = tk.StringVar(value="txt")
        extract_txt_radio = ttk.Radiobutton(self.extract_frame, text="TXT", variable=self.extract_format, value="txt")
        extract_txt_radio.grid(row=1, column=1, padx=5, pady=5)

        extract_csv_radio = ttk.Radiobutton(self.extract_frame, text="CSV", variable=self.extract_format, value="csv")
        extract_csv_radio.grid(row=1, column=2, padx=5, pady=5)

        extract_button = ttk.Button(self.extract_frame, text="Extract", command=self.extract_text)
        extract_button.grid(row=2, column=0, columnspan=3, pady=10)

        # --- Convert PDF to Images Tab ---
        self.convert_to_images_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.convert_to_images_frame, text="Convert PDF to Images")

        self.convert_to_images_file_path = tk.StringVar()
        convert_to_images_file_label = ttk.Label(self.convert_to_images_frame, text="Select PDF:")
        convert_to_images_file_label.grid(row=0, column=0, sticky="w")

        convert_to_images_file_entry = ttk.Entry(self.convert_to_images_frame, textvariable=self.convert_to_images_file_path, width=40)
        convert_to_images_file_entry.grid(row=0, column=1, padx=5)

        convert_to_images_file_button = ttk.Button(self.convert_to_images_frame, text="Browse", command=self.browse_convert_to_images_file)
        convert_to_images_file_button.grid(row=0, column=2)

        convert_to_images_button = ttk.Button(self.convert_to_images_frame, text="Convert", command=self.convert_pdf_to_images)
        convert_to_images_button.grid(row=1, column=0, columnspan=3, pady=10)

        # --- Convert Images to PDF Tab ---
        self.convert_from_images_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.convert_from_images_frame, text="Convert Images to PDF")

        self.convert_from_images_files = []
        self.convert_from_images_list_box = tk.Listbox(self.convert_from_images_frame)
        self.convert_from_images_list_box.pack(fill="both", expand=True)

        add_images_button = ttk.Button(self.convert_from_images_frame, text="Add Images", command=self.add_images)
        add_images_button.pack()

        remove_images_button = ttk.Button(self.convert_from_images_frame, text="Remove Selected", command=self.remove_selected_images)
        remove_images_button.pack()

        convert_from_images_button = ttk.Button(self.convert_from_images_frame, text="Convert", command=self.convert_images_to_pdf)
        convert_from_images_button.pack()

        self.convert_from_images_list_box.bind("<ButtonRelease-1>", self.select_image)
        self.convert_from_images_list_box.bind("<B1-Motion>", self.move_image)

        self.selected_image_index = None

        # --- Decrypt PDF Tab ---
        self.decrypt_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.decrypt_frame, text="Decrypt PDF")

        self.decrypt_file_path = tk.StringVar()
        decrypt_file_label = ttk.Label(self.decrypt_frame, text="Select PDF:")
        decrypt_file_label.grid(row=0, column=0, sticky="w")

        decrypt_file_entry = ttk.Entry(self.decrypt_frame, textvariable=self.decrypt_file_path, width=40)
        decrypt_file_entry.grid(row=0, column=1, padx=5)

        decrypt_file_button = ttk.Button(self.decrypt_frame, text="Browse", command=self.browse_decrypt_file)
        decrypt_file_button.grid(row=0, column=2)

        decrypt_password_label = ttk.Label(self.decrypt_frame, text="Enter Password:")
        decrypt_password_label.grid(row=1, column=0, sticky="w")

        self.decrypt_password_entry = ttk.Entry(self.decrypt_frame, show="*")
        self.decrypt_password_entry.grid(row=1, column=1, padx=5)

        decrypt_button = ttk.Button(self.decrypt_frame, text="Decrypt", command=self.decrypt_pdf)
        decrypt_button.grid(row=2, column=0, columnspan=3, pady=10)

        # --- Encrypt PDF Tab ---
        self.encrypt_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.encrypt_frame, text="Encrypt PDF")

        self.encrypt_file_path = tk.StringVar()
        encrypt_file_label = ttk.Label(self.encrypt_frame, text="Select PDF:")
        encrypt_file_label.grid(row=0, column=0, sticky="w")

        encrypt_file_entry = ttk.Entry(self.encrypt_frame, textvariable=self.encrypt_file_path, width=40)
        encrypt_file_entry.grid(row=0, column=1, padx=5)

        encrypt_file_button = ttk.Button(self.encrypt_frame, text="Browse", command=self.browse_encrypt_file)
        encrypt_file_button.grid(row=0, column=2)

        encrypt_password_label = ttk.Label(self.encrypt_frame, text="Enter Password:")
        encrypt_password_label.grid(row=1, column=0, sticky="w")

        self.encrypt_password_entry = ttk.Entry(self.encrypt_frame, show="*")
        self.encrypt_password_entry.grid(row=1, column=1, padx=5)

        encrypt_button = ttk.Button(self.encrypt_frame, text="Encrypt", command=self.encrypt_pdf)
        encrypt_button.grid(row=2, column=0, columnspan=3, pady=10)

    def add_files_merge(self):
        files = filedialog.askopenfilenames(initialdir=self.pdf_folder, filetypes=[("PDF Files", "*.pdf")])
        for file in files:
            self.merge_files.append(file)
            self.merge_list_box.insert(tk.END, os.path.basename(file))  # Display filenames

    def remove_selected_merge(self):
        try:
            selection = self.merge_list_box.curselection()[0]
            del self.merge_files[selection]
            self.merge_list_box.delete(selection)
        except IndexError:
            pass

    def merge_pdfs(self):
        if not self.merge_files:
            return

        output_filename = filedialog.asksaveasfilename(initialdir=self.pdf_folder, defaultextension=".pdf")
        if not output_filename:
            return

        merger = PyPDF2.PdfMerger()
        for pdf in self.merge_files:
            merger.append(pdf)

        merger.write(output_filename)
        merger.close()

        tk.messagebox.showinfo("Success", "PDFs merged successfully!")


    def browse_split_file(self):
        filepath = filedialog.askopenfilename(initialdir=self.pdf_folder, filetypes=[("PDF Files", "*.pdf")])
        if filepath:
            self.split_file_path.set(filepath)


    def split_pdf(self):
        filepath = self.split_file_path.get()
        ranges_str = self.split_ranges_entry.get()

        if not filepath or not ranges_str:
            tk.messagebox.showerror("Error", "Please select a PDF and enter page ranges.")
            return
        try:
           self._split_pdf_by_ranges(filepath, ranges_str)
           tk.messagebox.showinfo("Success", "PDF split successfully!")
        except Exception as e:
             tk.messagebox.showerror("Error splitting PDF", str(e))


    def _split_pdf_by_ranges(self, filepath, ranges_str):  # Helper function
        try:
            with open(filepath, "rb") as f:
                pdf = PyPDF2.PdfReader(f)
                num_pages = len(pdf.pages)

                ranges = self._parse_page_ranges(ranges_str, num_pages)  # Error handling inside

                output_dir = os.path.dirname(filepath) # Save in same directory
                filename_base = os.path.splitext(os.path.basename(filepath))[0]
                for i, (start, end) in enumerate(ranges):
                    output_filename = os.path.join(output_dir, f"{filename_base}_part{i+1}.pdf")
                    merger = PyPDF2.PdfMerger()
                    for page_num in range(start - 1, end):
                        merger.append(filepath, pages=(page_num, page_num + 1))  # Append individual pages
                    merger.write(output_filename)
                    merger.close()



        except FileNotFoundError:
            raise ValueError("Error: Input file not found.")
        except Exception as e: # Handle PyPDF2 errors or other issues
            raise ValueError(f"An error occurred during splitting: {str(e)}")


    def _parse_page_ranges(self, ranges_str, num_pages):  # Robust range parsing
        ranges = []
        range_parts = ranges_str.split(",")
        for part in range_parts:
            try:
                if "-" in part:
                    start, end = map(int, part.split("-"))
                    if not (1 <= start <= end <= num_pages):
                         raise ValueError(f"Invalid page range: {part}")
                    ranges.append((start, end))

                else:
                     page = int(part)
                     if not (1 <= page <= num_pages):
                         raise ValueError(f"Invalid page number: {page}")
                     ranges.append((page, page))  # Single page as range
            except ValueError as e:
               raise ValueError(f"Invalid page range specification: {e}") # Re-raise to calling function

        return ranges

    def browse_rotate_file(self):
        filepath = filedialog.askopenfilename(initialdir=self.pdf_folder, filetypes=[("PDF Files", "*.pdf")])
        if filepath:
            self.rotate_file_path.set(filepath)

    def set_rotation_angle(self, angle):
        self.rotation_angle = angle
        self.update_rotation_buttons()

    def update_rotation_buttons(self):
        buttons = [self.rotate_90_button, self.rotate_180_button, self.rotate_270_button]
        for button in buttons:
            button.state(["!pressed"])
        if self.rotation_angle == 90:
            self.rotate_90_button.state(["pressed"])
        elif self.rotation_angle == 180:
            self.rotate_180_button.state(["pressed"])
        elif self.rotation_angle == 270:
            self.rotate_270_button.state(["pressed"])

    def rotate_pdf(self):
        filepath = self.rotate_file_path.get()
        angle = self.rotation_angle

        if not filepath or angle is None:
            tk.messagebox.showerror("Error", "Please select a PDF and a rotation angle.")
            return

        try:
            angle = int(angle)
            if angle not in [90, 180, 270]:
                raise ValueError("Invalid rotation angle. Must be 90, 180, or 270.")

            with open(filepath, "rb") as f:
                pdf = PyPDF2.PdfReader(f)
                writer = PyPDF2.PdfWriter()

                for page in pdf.pages:
                    page.rotate(angle)
                    if angle in [90, 270]:
                        new_width = page.mediabox.height
                        new_height = page.mediabox.width
                        page.mediabox = PyPDF2.generic.RectangleObject([0, 0, new_height, new_width])
                    writer.add_page(page)

                output_filename = filedialog.asksaveasfilename(initialdir=self.pdf_folder, defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
                if not output_filename:
                    return

                with open(output_filename, "wb") as out_f:
                    writer.write(out_f)

            tk.messagebox.showinfo("Success", "PDF rotated successfully!")
        except Exception as e:
            tk.messagebox.showerror("Error rotating PDF", str(e))

    def browse_extract_file(self):
        filepath = filedialog.askopenfilename(initialdir=self.pdf_folder, filetypes=[("PDF Files", "*.pdf")])
        if filepath:
            self.extract_file_path.set(filepath)

    def extract_text(self):
        filepath = self.extract_file_path.get()
        file_format = self.extract_format.get()

        if not filepath:
            tk.messagebox.showerror("Error", "Please select a PDF.")
            return

        try:
            with open(filepath, "rb") as f:
                pdf = PyPDF2.PdfReader(f)
                text = ""
                for page in pdf.pages:
                    text += page.extract_text()

                output_filename = filedialog.asksaveasfilename(initialdir=self.pdf_folder, defaultextension=f".{file_format}", filetypes=[(f"{file_format.upper()} Files", f"*.{file_format}")])
                if not output_filename:
                    return

                with open(output_filename, "w", encoding="utf-8") as out_f:
                    out_f.write(text)

            tk.messagebox.showinfo("Success", f"Text extracted successfully as {file_format.upper()}!")
        except Exception as e:
            tk.messagebox.showerror("Error extracting text", str(e))

    def browse_convert_to_images_file(self):
        filepath = filedialog.askopenfilename(initialdir=self.pdf_folder, filetypes=[("PDF Files", "*.pdf")])
        if filepath:
            self.convert_to_images_file_path.set(filepath)

    def convert_pdf_to_images(self):
        filepath = self.convert_to_images_file_path.get()
        if not filepath:
            tk.messagebox.showerror("Error", "Please select a PDF.")
            return

        try:
            images = convert_from_path(filepath)
            output_dir = filedialog.askdirectory(initialdir=self.pdf_folder)
            if not output_dir:
                return

            for i, image in enumerate(images):
                image.save(os.path.join(output_dir, f"page_{i+1}.png"), "PNG")

            tk.messagebox.showinfo("Success", "PDF converted to images successfully!")
        except Exception as e:
            tk.messagebox.showerror("Error converting PDF to images", str(e))

    def add_images(self):
        files = filedialog.askopenfilenames(initialdir=self.pdf_folder, filetypes=[("Image Files", "*.png;*.jpg;*.jpeg")])
        for file in files:
            self.convert_from_images_files.append(file)
            self.convert_from_images_list_box.insert(tk.END, os.path.basename(file))

    def remove_selected_images(self):
        try:
            selection = self.convert_from_images_list_box.curselection()[0]
            del self.convert_from_images_files[selection]
            self.convert_from_images_list_box.delete(selection)
        except IndexError:
            pass

    def convert_images_to_pdf(self):
        if not self.convert_from_images_files:
            return

        output_filename = filedialog.asksaveasfilename(initialdir=self.pdf_folder, defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if not output_filename:
            return

        try:
            images = [Image.open(file).convert("RGB") for file in self.convert_from_images_files]
            images[0].save(output_filename, save_all=True, append_images=images[1:])
            tk.messagebox.showinfo("Success", "Images converted to PDF successfully!")
        except Exception as e:
            tk.messagebox.showerror("Error converting images to PDF", str(e))

    def select_image(self, event):
        self.selected_image_index = self.convert_from_images_list_box.nearest(event.y)

    def move_image(self, event):
        new_index = self.convert_from_images_list_box.nearest(event.y)
        if self.selected_image_index is not None and new_index != self.selected_image_index:
            self.convert_from_images_files.insert(new_index, self.convert_from_images_files.pop(self.selected_image_index))
            self.convert_from_images_list_box.delete(0, tk.END)
            for file in self.convert_from_images_files:
                self.convert_from_images_list_box.insert(tk.END, os.path.basename(file))
            self.selected_image_index = new_index

    def browse_decrypt_file(self):
        filepath = filedialog.askopenfilename(initialdir=self.pdf_folder, filetypes=[("PDF Files", "*.pdf")])
        if filepath:
            self.decrypt_file_path.set(filepath)

    def decrypt_pdf(self):
        filepath = self.decrypt_file_path.get()
        password = self.decrypt_password_entry.get()

        if not filepath or not password:
            tk.messagebox.showerror("Error", "Please select a PDF and enter a password.")
            return

        try:
            reader = PyPDF2.PdfReader(filepath)
            if not reader.decrypt(password):
                tk.messagebox.showerror("Error", "Incorrect password.")
                return

            writer = PyPDF2.PdfWriter()
            for page in reader.pages:
                writer.add_page(page)

            output_filename = filedialog.asksaveasfilename(initialdir=self.pdf_folder, defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
            if not output_filename:
                return

            with open(output_filename, "wb") as out_f:
                writer.write(out_f)

            tk.messagebox.showinfo("Success", "PDF decrypted successfully!")
        except Exception as e:
            tk.messagebox.showerror("Error decrypting PDF", str(e))

    def browse_encrypt_file(self):
        filepath = filedialog.askopenfilename(initialdir=self.pdf_folder, filetypes=[("PDF Files", "*.pdf")])
        if filepath:
            self.encrypt_file_path.set(filepath)

    def encrypt_pdf(self):
        filepath = self.encrypt_file_path.get()
        password = self.encrypt_password_entry.get()

        if not filepath or not password:
            tk.messagebox.showerror("Error", "Please select a PDF and enter a password.")
            return

        try:
            reader = PyPDF2.PdfReader(filepath)
            writer = PyPDF2.PdfWriter()

            for page in reader.pages:
                writer.add_page(page)

            writer.encrypt(password)

            output_filename = filedialog.asksaveasfilename(initialdir=self.pdf_folder, defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
            if not output_filename:
                return

            with open(output_filename, "wb") as out_f:
                writer.write(out_f)

            tk.messagebox.showinfo("Success", "PDF encrypted successfully!")
        except Exception as e:
            tk.messagebox.showerror("Error encrypting PDF", str(e))

    def select_merge_file(self, event):
        self.selected_merge_file_index = self.merge_list_box.nearest(event.y)

    def move_merge_file(self, event):
        new_index = self.merge_list_box.nearest(event.y)
        if self.selected_merge_file_index is not None and new_index != self.selected_merge_file_index:
            self.merge_files.insert(new_index, self.merge_files.pop(self.selected_merge_file_index))
            self.merge_list_box.delete(0, tk.END)
            for file in self.merge_files:
                self.merge_list_box.insert(tk.END, os.path.basename(file))
            self.selected_merge_file_index = new_index
        self.selected_merge_file_index = new_index

root = tk.Tk()
app = PDFToolsApp(root)
root.mainloop()