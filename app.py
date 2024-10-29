import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# Function to load Excel file and find duplicates within the specified column
def find_duplicates_in_excel(file_path, column_name):
    df = pd.read_excel(file_path, header=0)  # Load the Excel file with headers
    duplicates = df[df.duplicated(subset=[column_name], keep=False)]  # Find all duplicates
    unique_values = set(df[column_name].dropna().astype(str))  # Extract unique passport numbers
    return duplicates, unique_values

# Function to select file
def select_file(entry_widget):
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    entry_widget.delete(0, tk.END)
    entry_widget.insert(0, file_path)

# Function to compare values between two Excel files
def compare_files(excel_data1, excel_data2):
    excel1_only = excel_data1 - excel_data2  # Values only in the first Excel file
    excel2_only = excel_data2 - excel_data1  # Values only in the second Excel file
    return excel1_only, excel2_only

# Function to run the duplicate check
def check_duplicates():
    excel_file1 = file1_entry.get()
    excel_file2 = file2_entry.get()
    column_to_check = 'passport'  # Assuming the column name is "passport"
    
    if not excel_file1 or not excel_file2:
        messagebox.showerror("Error", "Please select both Excel files!")
        return
    
    # Find duplicates and unique values in both Excel files
    duplicates_excel1, unique_excel1 = find_duplicates_in_excel(excel_file1, column_to_check)
    duplicates_excel2, unique_excel2 = find_duplicates_in_excel(excel_file2, column_to_check)

    # Compare values between Excel files
    excel1_only, excel2_only = compare_files(unique_excel1, unique_excel2)
    
    # Display the results in the text box
    result_text.delete(1.0, tk.END)  # Clear the text box first
    result_text.insert(tk.END, f"Duplicates in Excel 1:\n{duplicates_excel1}\n\n")
    result_text.insert(tk.END, f"Duplicates in Excel 2:\n{duplicates_excel2}\n\n")
    result_text.insert(tk.END, f"Values only in Excel 1:\n{excel1_only}\n\n")
    result_text.insert(tk.END, f"Values only in Excel 2:\n{excel2_only}\n\n")

# Create main application window
root = tk.Tk()
root.title("Duplicate Finder")

# Labels and buttons to select Excel files
file1_label = tk.Label(root, text="Select first Excel file:")
file1_label.pack(pady=5)
file1_entry = tk.Entry(root, width=50)
file1_entry.pack(pady=5)
file1_button = tk.Button(root, text="Browse", command=lambda: select_file(file1_entry))
file1_button.pack(pady=5)

file2_label = tk.Label(root, text="Select second Excel file:")
file2_label.pack(pady=5)
file2_entry = tk.Entry(root, width=50)
file2_entry.pack(pady=5)
file2_button = tk.Button(root, text="Browse", command=lambda: select_file(file2_entry))
file2_button.pack(pady=5)

# Button to start the duplicate check
check_button = tk.Button(root, text="Check Duplicates", command=check_duplicates)
check_button.pack(pady=20)

# Text box to display results
result_text = tk.Text(root, height=15, width=80)
result_text.pack(pady=10)

# Run the application
root.mainloop()