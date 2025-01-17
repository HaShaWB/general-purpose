import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import re
import os

class CSVProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV Merger")

        # File list and order management
        self.file_paths = []
        self.file_gaps = []
        self.file_listbox = tk.Listbox(root, selectmode=tk.SINGLE, width=50, height=10)
        self.file_listbox.grid(row=0, column=0, columnspan=3, padx=10, pady=10)

        # Entry field for gap input
        self.gap_label = tk.Label(root, text="Gap (default: 400):")
        self.gap_label.grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.gap_entry = tk.Entry(root)
        self.gap_entry.grid(row=1, column=1, padx=5, pady=5)
        self.gap_entry.insert(0, "400")  # Default gap value

        # Buttons for file operations
        tk.Button(root, text="Add Files", command=self.add_files).grid(row=2, column=0, padx=5, pady=5)
        tk.Button(root, text="Move Up", command=lambda: self.move_file(-1)).grid(row=2, column=1, padx=5, pady=5)
        tk.Button(root, text="Move Down", command=lambda: self.move_file(1)).grid(row=2, column=2, padx=5, pady=5)

        # Confirm and Reset buttons
        tk.Button(root, text="Merge and Save", command=self.merge_and_save).grid(row=3, column=0, columnspan=2, padx=5, pady=10)
        tk.Button(root, text="Reset", command=self.reset).grid(row=3, column=2, padx=5, pady=10)

    def add_files(self):
        files = filedialog.askopenfilenames(filetypes=[("CSV Files", "*.csv")])
        if files:
            for file in files:
                self.file_paths.append(file)
                self.file_gaps.append(int(self.gap_entry.get()))  # Add default or entered gap
            self.update_file_list()

    def move_file(self, direction):
        selection = self.file_listbox.curselection()
        if selection:
            index = selection[0]
            new_index = index + direction
            if 0 <= new_index < len(self.file_paths):
                # Swap file paths and gaps
                self.file_paths[index], self.file_paths[new_index] = self.file_paths[new_index], self.file_paths[index]
                self.file_gaps[index], self.file_gaps[new_index] = self.file_gaps[new_index], self.file_gaps[index]
                self.update_file_list()
                self.file_listbox.select_set(new_index)

    def reset(self):
        self.file_paths = []
        self.file_gaps = []
        self.update_file_list()

    def update_file_list(self):
        self.file_listbox.delete(0, tk.END)
        for file_path, gap in zip(self.file_paths, self.file_gaps):
            display_text = f"{os.path.basename(file_path)} (Gap: {gap})"
            self.file_listbox.insert(tk.END, display_text)

    def merge_and_save(self):
        if not self.file_paths:
            messagebox.showerror("Error", "No files selected.")
            return

        # Process and merge files
        try:
            expanded_results = self.process_and_expand_files(self.file_paths, self.file_gaps)
            final_df = pd.DataFrame(expanded_results, columns=['english', 'korean'])

            # Save merged file
            output_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
            if output_path:
                final_df.to_csv(output_path, index=False, encoding='utf-8-sig')
                messagebox.showinfo("Success", f"File saved at {output_path}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def process_file(self, file_path):
        with open(file_path, 'r', encoding='euc-kr', errors='replace') as file:
            raw_lines = file.readlines()
        cleaned_lines = [line.strip() for line in raw_lines if line.strip() and not line.strip().startswith(';')]
        return self.clean_and_extract(cleaned_lines)

    def clean_and_extract(self, lines):
        valid_data = {}
        pattern = re.compile(r"^(\d+)\.\s+(.+)")

        for line in lines:
            cells = line.split(',')
            for cell in cells:
                cell = cell.strip().strip('"')
                match = pattern.match(cell)
                if match:
                    index = int(match.group(1))
                    word = match.group(2).strip()
                    word = re.sub(r"\s+", " ", word)
                    valid_data[index] = word

        return valid_data

    def process_and_expand_files(self, file_paths, gaps):
        expanded_data = []

        for idx, (file_path, gap) in enumerate(zip(file_paths, gaps)):
            file_data = self.process_file(file_path)

            base_offset = idx * gap
            for original_index, word in file_data.items():
                adjusted_index = original_index + base_offset
                if original_index <= gap:
                    paired_index = original_index + gap + base_offset
                    korean_word = file_data.get(original_index + gap, "")
                    expanded_data.append((word, korean_word))

        return expanded_data

# Run the GUI
if __name__ == "__main__":
    root = tk.Tk()
    app = CSVProcessorApp(root)
    root.mainloop()
