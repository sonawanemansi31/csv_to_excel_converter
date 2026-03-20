import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import os


class CSVToExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV to Excel Converter")
        self.root.geometry("700x500")
        self.root.resizable(False, False)

        # Variables
        self.input_file = tk.StringVar()
        self.output_file = tk.StringVar()
        self.rename_columns = tk.StringVar()

        # UI
        self.create_widgets()

    def create_widgets(self):
        title = tk.Label(
            self.root,
            text="CSV to Excel Converter",
            font=("Arial", 18, "bold"),
            fg="darkblue"
        )
        title.pack(pady=10)

        # Input file
        input_frame = tk.Frame(self.root)
        input_frame.pack(pady=10, padx=20, fill="x")

        tk.Label(input_frame, text="Select CSV File:", font=("Arial", 11)).pack(anchor="w")
        tk.Entry(input_frame, textvariable=self.input_file, width=70).pack(side="left", padx=(0, 10), pady=5)
        tk.Button(input_frame, text="Browse", command=self.browse_input, bg="#4CAF50", fg="white").pack(side="left")

        # Output file
        output_frame = tk.Frame(self.root)
        output_frame.pack(pady=10, padx=20, fill="x")

        tk.Label(output_frame, text="Save Excel File As:", font=("Arial", 11)).pack(anchor="w")
        tk.Entry(output_frame, textvariable=self.output_file, width=70).pack(side="left", padx=(0, 10), pady=5)
        tk.Button(output_frame, text="Browse", command=self.browse_output, bg="#2196F3", fg="white").pack(side="left")

        # Rename columns
        rename_frame = tk.Frame(self.root)
        rename_frame.pack(pady=10, padx=20, fill="x")

        tk.Label(rename_frame, text='Rename Columns (optional): old1:new1,old2:new2', font=("Arial", 11)).pack(anchor="w")
        tk.Entry(rename_frame, textvariable=self.rename_columns, width=90).pack(pady=5)

        # Convert button
        tk.Button(
            self.root,
            text="Convert CSV to Excel",
            command=self.convert_file,
            font=("Arial", 12, "bold"),
            bg="#FF9800",
            fg="white",
            width=25,
            height=2
        ).pack(pady=15)

        # Log area
        tk.Label(self.root, text="Status / Logs:", font=("Arial", 11, "bold")).pack(anchor="w", padx=20)
        self.log_box = scrolledtext.ScrolledText(self.root, width=80, height=12, state="disabled")
        self.log_box.pack(padx=20, pady=10)

    def log(self, message):
        self.log_box.config(state="normal")
        self.log_box.insert(tk.END, message + "\n")
        self.log_box.see(tk.END)
        self.log_box.config(state="disabled")

    def browse_input(self):
        file_path = filedialog.askopenfilename(
            title="Select CSV File",
            filetypes=[("CSV Files", "*.csv")]
        )
        if file_path:
            self.input_file.set(file_path)
            self.log(f"Selected input file: {file_path}")

    def browse_output(self):
        file_path = filedialog.asksaveasfilename(
            title="Save Excel File As",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if file_path:
            self.output_file.set(file_path)
            self.log(f"Selected output file: {file_path}")

    def parse_rename_columns(self, rename_str):
        rename_dict = {}
        if rename_str.strip():
            pairs = rename_str.split(",")
            for pair in pairs:
                if ":" in pair:
                    old, new = pair.split(":", 1)
                    rename_dict[old.strip()] = new.strip()
        return rename_dict

    def clean_dataframe(self, df):
        # Clean column names
        df.columns = [col.strip() for col in df.columns]

        # Strip whitespace from string values
        for col in df.select_dtypes(include=["object"]).columns:
            df[col] = df[col].astype(str).str.strip()

        # Replace 'nan' strings with actual missing values
        df.replace("nan", pd.NA, inplace=True)

        # Fill missing values
        for col in df.columns:
            if pd.api.types.is_numeric_dtype(df[col]):
                df[col] = df[col].fillna(0)
            else:
                df[col] = df[col].fillna("Unknown")

        return df

    def parse_dates(self, df):
        for col in df.columns:
            if "date" in col.lower() or "time" in col.lower():
                try:
                    df[col] = pd.to_datetime(df[col], errors="coerce")
                    self.log(f"Parsed date column: {col}")
                except Exception as e:
                    self.log(f"Could not parse date column '{col}': {e}")
        return df

    def convert_file(self):
        input_path = self.input_file.get().strip()
        output_path = self.output_file.get().strip()
        rename_str = self.rename_columns.get().strip()

        # Validation
        if not input_path:
            messagebox.showerror("Error", "Please select an input CSV file.")
            return

        if not output_path:
            messagebox.showerror("Error", "Please select an output Excel file path.")
            return

        if not os.path.exists(input_path):
            messagebox.showerror("Error", "Input file does not exist.")
            return

        if not input_path.lower().endswith(".csv"):
            messagebox.showerror("Error", "Selected input file is not a CSV file.")
            return

        try:
            self.log("Reading CSV file...")
            df = pd.read_csv(input_path)

            self.log("Cleaning data...")
            df = self.clean_dataframe(df)

            self.log("Parsing date columns...")
            df = self.parse_dates(df)

            rename_dict = self.parse_rename_columns(rename_str)
            if rename_dict:
                self.log(f"Renaming columns: {rename_dict}")
                df.rename(columns=rename_dict, inplace=True)

            if not output_path.lower().endswith(".xlsx"):
                output_path += ".xlsx"

            self.log("Saving Excel file...")
            df.to_excel(output_path, index=False, engine="openpyxl")

            self.log("Conversion completed successfully!")
            messagebox.showinfo("Success", f"Excel file saved successfully:\n{output_path}")

        except pd.errors.EmptyDataError:
            messagebox.showerror("Error", "The CSV file is empty.")
            self.log("Error: The CSV file is empty.")

        except pd.errors.ParserError:
            messagebox.showerror("Error", "Error parsing the CSV file. Please check the file format.")
            self.log("Error: Invalid CSV format.")

        except PermissionError:
            messagebox.showerror("Error", "Permission denied. Close the Excel file if it is open.")
            self.log("Error: Permission denied.")

        except Exception as e:
            messagebox.showerror("Error", f"Unexpected error:\n{e}")
            self.log(f"Unexpected error: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = CSVToExcelApp(root)
    root.mainloop()