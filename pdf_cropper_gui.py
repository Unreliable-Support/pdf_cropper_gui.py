# Filename: pdf_cropper_gui.py

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from pypdf import PdfReader, PdfWriter

class PDFCropperApp:
    def __init__(self, master):
        self.master = master
        master.title("PDF Page Cropper Tool")
        master.geometry("700x550")

        self.pdf_files_paths = {} # Stores filename as key, full path as value
        self.output_folder = ""

        # --- Top Frame: File and Folder Selection ---
        top_frame = ttk.Frame(master, padding="10")
        top_frame.pack(fill=tk.X)

        self.btn_load_pdfs = ttk.Button(top_frame, text="Select PDF Files", command=self.load_pdfs)
        self.btn_load_pdfs.pack(side=tk.LEFT, padx=5)

        self.btn_select_output = ttk.Button(top_frame, text="Select Output Folder", command=self.select_output_folder)
        self.btn_select_output.pack(side=tk.LEFT, padx=5)

        self.lbl_output_folder = ttk.Label(top_frame, text="Output Folder: Not selected")
        self.lbl_output_folder.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        # --- Middle Frame: PDF List (Treeview) ---
        mid_frame = ttk.Frame(master, padding="10")
        mid_frame.pack(fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(mid_frame, columns=("filename", "page_range"), show="headings")
        self.tree.heading("filename", text="File Name")
        self.tree.heading("page_range", text="Page Range (e.g., 1-5, 7)")
        self.tree.column("filename", width=300)
        self.tree.column("page_range", width=150, anchor=tk.CENTER)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(mid_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)

        # --- Bottom Frame: Page Range Input ---
        bottom_frame = ttk.Frame(master, padding="10")
        bottom_frame.pack(fill=tk.X)

        ttk.Label(bottom_frame, text="Page Range for Selected PDF:").pack(side=tk.LEFT, padx=5)
        self.entry_page_range = ttk.Entry(bottom_frame, width=20)
        self.entry_page_range.pack(side=tk.LEFT, padx=5)
        self.btn_set_range = ttk.Button(bottom_frame, text="Set Range", command=self.set_page_range_for_selected)
        self.btn_set_range.pack(side=tk.LEFT, padx=5)
        self.btn_set_range["state"] = "disabled" # Initially disabled

        # --- Action Frame: Run and Status ---
        action_frame = ttk.Frame(master, padding="10")
        action_frame.pack(fill=tk.X)

        self.btn_run_crop = ttk.Button(action_frame, text="CROP and SAVE", command=self.crop_and_save_pdfs, style="Accent.TButton")
        self.btn_run_crop.pack(pady=10)
        
        style = ttk.Style()
        try:
            style.configure("Accent.TButton", foreground="white", background="blue")
        except tk.TclError:
            print("Note: The 'Accent.TButton' style is not available in your current theme. A standard button will be used.")

        self.lbl_status = ttk.Label(action_frame, text="Status: Ready")
        self.lbl_status.pack(fill=tk.X)

    def load_pdfs(self):
        filepaths = filedialog.askopenfilenames(
            title="Select PDF Files",
            filetypes=(("PDF Files", "*.pdf"), ("All Files", "*.*"))
        )
        if not filepaths:
            return

        for filepath in filepaths:
            filename = os.path.basename(filepath)
            if filename not in self.pdf_files_paths: # Prevent adding the same file twice
                self.pdf_files_paths[filename] = filepath
                try:
                    reader = PdfReader(filepath)
                    num_pages = len(reader.pages)
                    default_range = f"1-{num_pages}" if num_pages > 0 else "None"
                    self.tree.insert("", tk.END, values=(filename, default_range))
                except Exception as e:
                    self.tree.insert("", tk.END, values=(filename, "ERROR: Unreadable"))
                    messagebox.showerror("PDF Read Error", f"Error reading file {filename}: {e}")
            else:
                messagebox.showinfo("Information", f"{filename} is already in the list.")
        self.update_status(f"{len(filepaths)} PDF file(s) added to the list.")

    def select_output_folder(self):
        folder = filedialog.askdirectory(title="Select Folder to Save Cropped PDFs")
        if folder:
            self.output_folder = folder
            self.lbl_output_folder.config(text=f"Output Folder: {self.output_folder}")
            self.update_status(f"Output folder set to: {self.output_folder}")

    def on_tree_select(self, event):
        selected_items = self.tree.selection()
        if selected_items:
            item = self.tree.item(selected_items[0])
            filename, page_range = item["values"]
            self.entry_page_range.delete(0, tk.END)
            self.entry_page_range.insert(0, str(page_range))
            self.btn_set_range["state"] = "normal"
        else:
            self.entry_page_range.delete(0, tk.END)
            self.btn_set_range["state"] = "disabled"

    def set_page_range_for_selected(self):
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("Warning", "Please select a PDF file from the list first.")
            return

        item_id = selected_items[0]
        new_range = self.entry_page_range.get().strip()

        if not new_range:
            messagebox.showwarning("Warning", "Page range cannot be empty. Enter '1-X' for all pages or check the page count.")
            return
        
        if new_range.lower() == "all":
            filename = self.tree.item(item_id, "values")[0]
            filepath = self.pdf_files_paths.get(filename)
            if filepath:
                try:
                    reader = PdfReader(filepath)
                    num_pages = len(reader.pages)
                    new_range = f"1-{num_pages}" if num_pages > 0 else "1"
                except Exception:
                    new_range = "1-?" # In case of error
            else:
                new_range = "1-?"

        self.tree.item(item_id, values=(self.tree.item(item_id, "values")[0], new_range))
        self.update_status(f"Range for '{self.tree.item(item_id, 'values')[0]}' set to '{new_range}'.")

    def parse_page_range(self, range_str, max_pages):
        """
        Converts a string like "1-3,5,7-8" to a 0-indexed list of pages [0,1,2,4,6,7].
        Returns all pages if the string is empty or "all".
        """
        if not range_str or range_str.lower() == "all" or range_str == f"1-{max_pages}":
            return list(range(max_pages))

        pages_to_keep = set()
        parts = range_str.replace(" ", "").split(',')
        for part in parts:
            if '-' in part:
                try:
                    start, end = map(int, part.split('-'))
                    if start <= 0 or end <= 0 or start > end or end > max_pages:
                        raise ValueError("Invalid range values")
                    pages_to_keep.update(range(start - 1, end)) # 0-indexed
                except ValueError:
                    raise ValueError(f"Invalid range format: '{part}'. Start and end must be numbers and logical.")
            else:
                try:
                    page = int(part)
                    if page <= 0 or page > max_pages:
                        raise ValueError(f"Invalid page number: {page}. Must be between 1 and {max_pages}.")
                    pages_to_keep.add(page - 1) # 0-indexed
                except ValueError:
                    raise ValueError(f"Invalid page format: '{part}'. Must be a number.")
        
        return sorted(list(pages_to_keep))

    def crop_and_save_pdfs(self):
        if not self.pdf_files_paths:
            messagebox.showwarning("Warning", "There are no PDF files in the list to crop.")
            return
        if not self.output_folder:
            messagebox.showwarning("Warning", "Please select an output folder first.")
            return

        processed_count = 0
        error_count = 0

        for item_id in self.tree.get_children():
            filename, page_range_str = self.tree.item(item_id, "values")
            original_filepath = self.pdf_files_paths.get(filename)

            if not original_filepath:
                self.update_status(f"ERROR: Original file path for '{filename}' not found. Skipping.")
                error_count += 1
                continue
            
            if page_range_str == "ERROR: Unreadable" or not page_range_str:
                self.update_status(f"WARNING: No valid page range for '{filename}' or unreadable. Skipping.")
                error_count += 1
                continue

            self.update_status(f"Processing: {filename} (Range: {page_range_str})")
            self.master.update_idletasks() # Update the UI

            try:
                reader = PdfReader(original_filepath)
                writer = PdfWriter()
                max_pages = len(reader.pages)

                if max_pages == 0:
                    self.update_status(f"WARNING: '{filename}' is empty or contains no pages. Skipping.")
                    error_count +=1
                    continue

                try:
                    pages_to_include_indices = self.parse_page_range(page_range_str, max_pages)
                except ValueError as e_parse:
                    messagebox.showerror("Page Range Error", f"Could not parse page range for '{filename}': {e_parse}\nPlease check the range (e.g., 1-5, 7, 9-10).")
                    error_count += 1
                    continue
                
                if not pages_to_include_indices:
                    self.update_status(f"WARNING: No valid pages found in the specified range for '{filename}'. Skipping.")
                    error_count += 1
                    continue

                for page_index in pages_to_include_indices:
                    if 0 <= page_index < max_pages:
                        writer.add_page(reader.pages[page_index])
                    else:
                        self.update_status(f"WARNING: Invalid page index {page_index+1} for '{filename}' was skipped.")

                if len(writer.pages) > 0:
                    base, ext = os.path.splitext(filename)
                    output_filename = f"{base}_cropped{ext}"
                    output_filepath = os.path.join(self.output_folder, output_filename)
                    
                    with open(output_filepath, "wb") as f_out:
                        writer.write(f_out)
                    processed_count += 1
                else:
                    self.update_status(f"WARNING: No pages were left for '{filename}' after cropping. File not created.")
                    error_count += 1

            except Exception as e:
                self.update_status(f"ERROR: An issue occurred while processing '{filename}': {e}")
                messagebox.showerror("Processing Error", f"Error processing '{filename}': {e}")
                error_count += 1
        
        final_message = f"Process complete. {processed_count} PDF(s) successfully cropped."
        if error_count > 0:
            final_message += f" {error_count} PDF(s) failed or were skipped."
        
        self.update_status(final_message)
        messagebox.showinfo("Complete", final_message)

    def update_status(self, message):
        self.lbl_status.config(text=f"Status: {message}")
        print(message) # Also print to console for debugging

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFCropperApp(root)
    root.mainloop()