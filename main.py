import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import random
import time

def rolling_effect():
    if data_frame is not None:
        for _ in range(20):  # Number of rolls
            random_name = data_frame.sample(n=1)
            result_text.delete('1.0', tk.END)
            result_text.insert(tk.END, random_name.to_string(index=False))
            root.update()
            time.sleep(0.1)  # Delay between rolls

def select_winners_with_effect():
    try:
        event_name = event_name_entry.get()
        prize_count = int(prize_count_entry.get())
        if not event_name or prize_count <= 0:
            messagebox.showerror("Error", "Please enter a valid event name and number of prizes.")
            return
        if data_frame is None:
            messagebox.showerror("Error", "Please upload an Excel file first.")
            return
        rolling_effect()  # Add rolling names effect here
 
        # Select winners
        winners = data_frame.sample(n=prize_count)
        winners['Event Name'] = event_name
        result_text.delete('1.0', tk.END)
        result_text.insert(tk.END, winners.to_string(index=False))
        # Save result to Excel
        output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                                   filetypes=[("Excel files", "*.xlsx")])
        if output_file:
            winners.to_excel(output_file, index=False, engine='openpyxl')
            messagebox.showinfo("Success", f"Results saved to {output_file}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# Function to select winners
def select_winners():
    try:
        # Validate inputs
        event_name = event_name_entry.get()
        prize_count = int(prize_count_entry.get())
        if not event_name or prize_count <= 0:
            messagebox.showerror("Error", "Please enter a valid event name and number of prizes.")
            return
        if data_frame is None:
            messagebox.showerror("Error", "Please upload an Excel file first.")
            return
        # Select winners
        winners = data_frame.sample(n=prize_count)
        winners['Event Name'] = event_name
        result_text.delete('1.0', tk.END)
        result_text.insert(tk.END, winners.to_string(index=False))
        # Save result to Excel
        output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                                   filetypes=[("Excel files", "*.xlsx")])
        if output_file:
            winners.to_excel(output_file, index=False, engine='openpyxl')
            messagebox.showinfo("Success", f"Results saved to {output_file}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
 
# Function to load Excel file
def upload_file():
    global data_frame
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        try:
            data_frame = pd.read_excel(file_path, sheet_name=0)
            required_columns = ["Id", "中文全名", "英文全名", "員工編號", "營運單位"]
            if not all(col in data_frame.columns for col in required_columns):
                raise ValueError("Uploaded Excel file does not have the required columns.")
            messagebox.showinfo("Success", "File uploaded successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {e}")
 
# Initialize GUI
root = tk.Tk()
root.title("Lucky Draw Application")
 
data_frame = None
 
# Event name input
tk.Label(root, text="Event Name:").grid(row=0, column=0, padx=10, pady=5, sticky="e")
event_name_entry = tk.Entry(root, width=30)
event_name_entry.grid(row=0, column=1, padx=10, pady=5)
 
# Prize count input
tk.Label(root, text="Number of Prizes:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
prize_count_entry = tk.Entry(root, width=30)
prize_count_entry.grid(row=1, column=1, padx=10, pady=5)
 
# Upload file button
upload_button = tk.Button(root, text="Upload Excel File", command=upload_file)
upload_button.grid(row=2, column=0, columnspan=2, pady=10)
 
# Select winners button
draw_button = tk.Button(root, text="Select Winners", command=select_winners_with_effect)
draw_button.grid(row=3, column=0, columnspan=2, pady=10)
 
# Result display
result_text = tk.Text(root, height=30, width=100)
result_text.grid(row=4, column=0, columnspan=2, padx=10, pady=10)
 
root.mainloop()