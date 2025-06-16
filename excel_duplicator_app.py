import pandas as pd
import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext  # For a multi-line log

class ExcelDuplicatorApp:
    def __init__(self, master):
        # Main window setup
        self.master = master
        master.title("Excel Product File Duplicator")
        master.geometry("700x550")  # Increased size for better layout and log

        # Configure columns for better layout
        master.grid_columnconfigure(0, weight=0) # Labels column
        master.grid_columnconfigure(1, weight=1) # Entries column
        master.grid_columnconfigure(2, weight=0) # Buttons column

        # --- Variables to store paths ---
        self.lookup_file_path_var = tk.StringVar()
        self.source_folder_path_var = tk.StringVar()
        self.destination_folder_path_var = tk.StringVar()

        # --- Widgets for Lookup File ---
        tk.Label(master, text="Product Lookup File (.xlsx):").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        tk.Entry(master, textvariable=self.lookup_file_path_var, width=60).grid(row=0, column=1, padx=10, pady=5, sticky="ew")
        tk.Button(master, text="Browse", command=self.browse_lookup_file).grid(row=0, column=2, padx=10, pady=5)

        # --- Widgets for Source Folder ---
        tk.Label(master, text="Source Folder (Master Files):").grid(row=1, column=0, sticky="w", padx=10, pady=5)
        tk.Entry(master, textvariable=self.source_folder_path_var, width=60).grid(row=1, column=1, padx=10, pady=5, sticky="ew")
        tk.Button(master, text="Browse", command=self.browse_source_folder).grid(row=1, column=2, padx=10, pady=5)

        # --- Widgets for Destination Folder ---
        tk.Label(master, text="Destination Folder (New Files):").grid(row=2, column=0, sticky="w", padx=10, pady=5)
        tk.Entry(master, textvariable=self.destination_folder_path_var, width=60).grid(row=2, column=1, padx=10, pady=5, sticky="ew")
        tk.Button(master, text="Browse", command=self.browse_destination_folder).grid(row=2, column=2, padx=10, pady=5)

        # --- Execute Button ---
        tk.Button(
            master, text="Run Duplication", command=self.run_duplication,
            font=("Arial", 12, "bold"), bg="#4CAF50", fg="white"
        ).grid(row=3, column=0, columnspan=3, pady=20)

        # --- Status / Log Area ---
        tk.Label(master, text="Process Log:").grid(row=4, column=0, sticky="w", padx=10, pady=5)
        self.log_text = scrolledtext.ScrolledText(master, wrap=tk.WORD, width=80, height=15, font=("Courier New", 10))
        self.log_text.grid(row=5, column=0, columnspan=3, padx=10, pady=5, sticky="nsew")
        # Make the log area resizable
        master.grid_rowconfigure(5, weight=1)

    def log_message(self, message, color="black"):
        """Appends a message to the log area in the specified color."""
        self.log_text.insert(tk.END, message + "\n", color)
        self.log_text.see(tk.END)  # Scroll to the end
        # Color configuration for different log types
        self.log_text.tag_config("red", foreground="red")
        self.log_text.tag_config("green", foreground="green")
        self.log_text.tag_config("blue", foreground="blue")
        self.log_text.tag_config("orange", foreground="orange")
        self.master.update_idletasks()  # Update GUI immediately

    def browse_lookup_file(self):
        """Open file dialog to select the lookup Excel file."""
        file_path = filedialog.askopenfilename(
            title="Select Product Lookup Excel File",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            self.lookup_file_path_var.set(file_path)
            self.log_message(f"Selected lookup file: {file_path}", "blue")

    def browse_source_folder(self):
        """Open directory dialog to select the source folder."""
        folder_path = filedialog.askdirectory(title="Select Folder with Master Excel Files")
        if folder_path:
            self.source_folder_path_var.set(folder_path)
            self.log_message(f"Selected source folder: {folder_path}", "blue")

    def browse_destination_folder(self):
        """Open directory dialog to select the destination folder."""
        folder_path = filedialog.askdirectory(title="Select Destination Folder for New Files")
        if folder_path:
            self.destination_folder_path_var.set(folder_path)
            self.log_message(f"Selected destination folder: {folder_path}", "blue")

    def run_duplication(self):
        """Main logic to read lookup file and copy/rename files accordingly."""
        self.log_text.delete(1.0, tk.END)  # Clear previous log
        self.log_message("Starting Excel file duplication process...", "blue")

        lookup_file = self.lookup_file_path_var.get()
        source_folder = self.source_folder_path_var.get()
        destination_folder = self.destination_folder_path_var.get()

        # --- Basic validation ---
        if not lookup_file or not os.path.exists(lookup_file):
            self.log_message("Error: Please select a valid product lookup Excel file.", "red")
            messagebox.showerror("Input Error", "Please select a valid product lookup Excel file.")
            return
        if not source_folder or not os.path.isdir(source_folder):
            self.log_message("Error: Please select a valid source folder.", "red")
            messagebox.showerror("Input Error", "Please select a valid source folder.")
            return
        if not destination_folder:
            self.log_message("Error: Please select a destination folder.", "red")
            messagebox.showerror("Input Error", "Please select a destination folder.")
            return

        # Create destination folder if it doesn't exist
        if not os.path.exists(destination_folder):
            try:
                os.makedirs(destination_folder)
                self.log_message(f"Created destination folder: {destination_folder}")
            except Exception as e:
                self.log_message(f"Error creating destination folder '{destination_folder}': {e}", "red")
                messagebox.showerror("Folder Creation Error", f"Could not create destination folder:\n{e}")
                return

        # --- Read the lookup list ---
        try:
            lookup_df = pd.read_excel(lookup_file)
            self.log_message(f"Successfully loaded lookup list with {len(lookup_df)} entries.", "blue")
            if "Product Code" not in lookup_df.columns or "Product Name" not in lookup_df.columns:
                self.log_message("Error: Lookup file must contain 'Product Code' and 'Product Name' columns.", "red")
                self.log_message(f"Available columns: {lookup_df.columns.tolist()}", "red")
                messagebox.showerror("Lookup File Error", "The lookup file must have 'Product Code' and 'Product Name' columns.")
                return
        except FileNotFoundError:
            self.log_message(f"Error: Lookup file not found at '{lookup_file}'.", "red")
            messagebox.showerror("File Not Found", f"Lookup file not found:\n{lookup_file}")
            return
        except Exception as e:
            self.log_message(f"Error reading lookup file '{lookup_file}': {e}", "red")
            messagebox.showerror("File Read Error", f"Error reading lookup file:\n{e}")
            return

        processed_count = 0
        skipped_count = 0

        # --- Iterate through the lookup list ---
        for index, row in lookup_df.iterrows():
            product_code = str(row["Product Code"]).strip()
            product_name = str(row["Product Name"]).strip()

            source_file_name = f"{product_name}.xlsx"
            source_file_path = os.path.join(source_folder, source_file_name)
            destination_file_name = f"{product_code}.xlsx"
            destination_file_path = os.path.join(destination_folder, destination_file_name)

            self.log_message(f"Processing '{product_code}' - '{product_name}':")
            self.log_message(f"  Source: {source_file_name}")

            if os.path.exists(source_file_path):
                try:
                    shutil.copy2(source_file_path, destination_file_path)
                    self.log_message(f"  -> Created '{destination_file_name}'", "green")
                    processed_count += 1
                except Exception as e:
                    self.log_message(f"  Error copying file: {e}", "red")
                    skipped_count += 1
            else:
                self.log_message(f"  Source file NOT FOUND: '{source_file_name}'. Skipping.", "orange")
                skipped_count += 1

        self.log_message("\n--- Process Complete ---", "blue")
        self.log_message(f"Total entries processed: {processed_count}", "blue")
        self.log_message(f"Total entries skipped (source file not found/error): {skipped_count}", "blue")
        messagebox.showinfo(
            "Process Complete",
            f"Duplication process finished!\n"
            f"Processed: {processed_count}\n"
            f"Skipped: {skipped_count}\n"
            f"Files saved to: {destination_folder}"
        )

# Main part to run the GUI application
if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelDuplicatorApp(root)
    root.mainloop()
