import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import random
import time
 
assigned_prizes = []  # Track assigned prize numbers
results = []  # Store assigned prizes and winners
selected_winner = None  # Global variable to store the currently selected winner
 
# Function to perform rolling effect
def rolling_effect(display_text):
    for _ in range(20):  # Number of rolling iterations
        random_name = data_frame.sample(n=1)  # Temporary display name
        result_text.delete('1.0', tk.END)
        result_text.insert(tk.END, f"Rolling...\n{display_text}: {random_name.to_string(index=False)}")
        root.update()
        time.sleep(0.1)  # Pause for visual effect
 
# Function to draw a name
def draw_name():
    global selected_winner  # Ensure global declaration
    try:
        if data_frame is None:
            messagebox.showerror("Error", "Please upload an Excel file first.")
            return
 
        # Randomly select a winner
        selected_winner = data_frame.sample(n=1)
        rolling_effect("Winner")
 
        # Display the winner
        result_text.delete('1.0', tk.END)
        result_text.insert(tk.END, "Final Result:\n")
        result_text.insert(tk.END, f"Winner:\n{selected_winner.to_string(index=False)}")
 
        prize_button.config(state=tk.NORMAL)  # Enable prize draw button
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
 
# Function to draw a prize number
def draw_prize():
    global selected_winner  # Ensure global declaration
    try:
        if selected_winner is None:
            messagebox.showerror("Error", "Please draw a name first.")
            return
 
        # Get prize number range
        try:
            start_prize = int(start_prize_entry.get())
            end_prize = int(end_prize_entry.get())
            if start_prize > end_prize or start_prize < 1:
                raise ValueError("Invalid prize range.")
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid prize number range.")
            return
 
        # Generate the prize number range
        prize_numbers = list(range(start_prize, end_prize + 1))
        available_prizes = [p for p in prize_numbers if p not in assigned_prizes]
        if not available_prizes:
            messagebox.showinfo("Info", "All prizes have been assigned.")
            return
 
        # Randomly select a prize number from available ones
        selected_prize = random.choice(available_prizes)
        assigned_prizes.append(selected_prize)
 
        # Record the result
        winner_data = selected_winner.iloc[0].to_dict()
        winner_data['Prize Number'] = selected_prize
        results.append(winner_data)
 
        # Display the prize number with the winner
        result_text.insert(tk.END, f"\nPrize Number: {selected_prize}\n")
        result_text.insert(tk.END, f"Assigned to Winner:\n{selected_winner.to_string(index=False)}\n\n")
 
        # Reset for the next draw
        selected_winner = None
        prize_button.config(state=tk.DISABLED)  # Disable prize draw button
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
 
# Function to export results to Excel
def export_results():
    if not results:
        messagebox.showerror("Error", "No results to export.")
        return
 
    # Create a DataFrame from results
    results_df = pd.DataFrame(results)
 
    # Save to Excel
    output_file = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                               filetypes=[("Excel files", "*.xlsx")])
    if output_file:
        try:
            results_df.to_excel(output_file, index=False, engine='openpyxl')
            messagebox.showinfo("Success", f"Results saved to {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save results: {e}")
 
# Function to upload an Excel file
def upload_file():
    global data_frame  # Ensure global declaration
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
 
# Fullscreen mode
root.state('zoomed')  # Windows fullscreen
# For cross-platform fullscreen, use: root.attributes('-fullscreen', True)
 
data_frame = None  # Initialize as None to load data later
 
# Set larger font
default_font = ("Arial", 16)
 
# Prize number range input
tk.Label(root, text="Prize Number Range Start:", font=default_font).grid(row=0, column=0, padx=10, pady=5, sticky="e")
start_prize_entry = tk.Entry(root, width=10, font=default_font)
start_prize_entry.grid(row=0, column=1, padx=10, pady=5)
 
tk.Label(root, text="Prize Number Range End:", font=default_font).grid(row=0, column=2, padx=10, pady=5, sticky="e")
end_prize_entry = tk.Entry(root, width=10, font=default_font)
end_prize_entry.grid(row=0, column=3, padx=10, pady=5)
 
# Upload file button
upload_button = tk.Button(root, text="Upload Excel File", command=upload_file, font=default_font)
upload_button.grid(row=1, column=0, columnspan=4, pady=10)
 
# Draw name button
name_button = tk.Button(root, text="Draw Name", command=draw_name, font=default_font)
name_button.grid(row=2, column=0, columnspan=2, pady=10)
 
# Draw prize button
prize_button = tk.Button(root, text="Draw Prize Number", command=draw_prize, state=tk.DISABLED, font=default_font)
prize_button.grid(row=2, column=2, columnspan=2, pady=10)
 
# Export results button
export_button = tk.Button(root, text="Export Results", command=export_results, font=default_font)
export_button.grid(row=3, column=0, columnspan=4, pady=10)
 
# Result display
result_text = tk.Text(root, height=20, width=80, font=("Arial", 18))
result_text.grid(row=4, column=0, columnspan=4, padx=10, pady=10)
 
root.mainloop()