import tkinter as tk
from tkinter import ttk, filedialog, messagebox


class MaskKeywordApp:
    def __init__(self, root):
        self.root = root
        self.root.title("TJ-MaskKeyword")
        self.root.geometry("600x350")
        self.root.resizable(False, False)

        self.file_folder = tk.StringVar()
        self.mapping_folder = tk.StringVar()
        self.remove_header = tk.BooleanVar(value=False)

        self._setup_ui()

    def _setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        title = ttk.Label(main_frame, text="TJ-MaskKeyword",
                         font=("Segoe UI", 20, "bold"))
        title.pack(pady=(0, 20))

        folder_frame = ttk.Frame(main_frame)
        folder_frame.pack(fill=tk.X, pady=10)

        ttk.Label(folder_frame, text="Choose Files:").grid(row=0, column=0, sticky=tk.W, pady=10)
        ttk.Entry(folder_frame, textvariable=self.file_folder, width=45).grid(row=0, column=1, padx=10)
        ttk.Button(folder_frame, text="Browse", command=self.browse_file_folder).grid(row=0, column=2)

        ttk.Label(folder_frame, text="Choose Mapping Ref:").grid(row=1, column=0, sticky=tk.W, pady=10)
        ttk.Entry(folder_frame, textvariable=self.mapping_folder, width=45).grid(row=1, column=1, padx=10)
        ttk.Button(folder_frame, text="Browse", command=self.browse_mapping_folder).grid(row=1, column=2)

        options_frame = ttk.Frame(main_frame)
        options_frame.pack(fill=tk.X, pady=10)
        ttk.Checkbutton(options_frame, text="Remove Header", variable=self.remove_header).pack(anchor=tk.W)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=30)

        ttk.Button(btn_frame, text="Apply", command=self.apply).pack(side=tk.LEFT, padx=20)
        ttk.Button(btn_frame, text="Exit", command=self.root.quit).pack(side=tk.LEFT, padx=20)

    def browse_file_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.file_folder.set(folder)

    def browse_mapping_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.mapping_folder.set(folder)

    def apply(self):
        from logic import process_files

        if not self.file_folder.get():
            messagebox.showerror("Error", "Please select a files folder")
            return
        if not self.mapping_folder.get():
            messagebox.showerror("Error", "Please select a mapping ref folder")
            return

        status, msg = process_files(self.file_folder.get(), self.mapping_folder.get(), self.remove_header.get())

        if status == "success":
            messagebox.showinfo("Success", msg)
        elif status == "no_file":
            messagebox.showinfo("No File", msg)
        elif status == "missing_mapping":
            messagebox.showinfo("Missing Mapping", msg)
        else:
            messagebox.showerror("Error", msg)