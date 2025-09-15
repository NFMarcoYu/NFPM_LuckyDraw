import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import random
import time

assigned_prizes = []  # Track assigned prize numbers
results = []  # Store assigned prizes and winners
selected_winner = None  # Global variable to store the currently selected winner

data_frame = None  # Initialize as None to load data later

draw_count = 0  # Counter for the number of draws

# Function to perform rolling effect
def rolling_effect(display_text, items):
    for _ in range(20):  # Number of rolling iterations
        random_item = random.choice(items)  # Temporary display item
        result_text.delete('1.0', tk.END)
        result_text.insert(tk.END, f"Rolling...\n{display_text}: {random_item}")
        root.update()
        time.sleep(0.1)  # Pause for visual effect

# Function to draw a name
def draw_name():
    global selected_winner  # Ensure global declaration
    try:
        if data_frame is None:
            messagebox.showerror("Error", "Please upload an Excel file first.")
            return

        # Filter out already selected winners
        available_names = data_frame[~data_frame['Id'].isin([result['Id'] for result in results])]
        if available_names.empty:
            messagebox.showinfo("Info", "All names have been drawn.")
            return

        # Randomly select a winner
        selected_winner = available_names.sample(n=1)
        rolling_effect(":", available_names['中文全名'].tolist())

        # Display the winner
        result_text.delete('1.0', tk.END)
        winner_data = selected_winner.iloc[0]
        result_text.insert(tk.END, f"{winner_data['員工編號']} {winner_data['營運單位']} {winner_data['英文全名']}")

        # Update name list and dim the selected winner
        update_name_list(dim_selected=True)

        prize_button.config(state=tk.NORMAL)  # Enable prize draw button
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# Function to draw a prize number
def draw_prize():
    global selected_winner, draw_count  # Ensure global declaration
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

        # Rolling effect for prize number
        rolling_effect("Prize Number", available_prizes)

        # Randomly select a prize number from available ones
        selected_prize = random.choice(available_prizes)
        assigned_prizes.append(selected_prize)

        # Record the result
        winner_data = selected_winner.iloc[0].to_dict()
        winner_data['Prize Number'] = selected_prize
        results.append(winner_data)

        # Display the winner and prize
        result_text.delete('1.0', tk.END)
        result_text.insert(tk.END, f"{winner_data['員工編號']} {winner_data['營運單位']} {winner_data['英文全名']}\nPrize: {selected_prize}")

        # Update prize list and history list
        update_prize_list()
        update_history_list()

        # Reset for the next draw
        selected_winner = None
        prize_button.config(state=tk.DISABLED)  # Disable prize draw button
        draw_count += 1  # Increment the draw count
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
            # Replace NaN values with 'N/A'
            data_frame.fillna("N/A", inplace=True)
            messagebox.showinfo("Success", "File uploaded successfully.")
            update_name_list()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {e}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {e}")

# Function to update the name list display
def update_name_list(dim_selected=False):
    name_list.delete(0, tk.END)
    for index, row in data_frame.iterrows():
        display_text = f"{row['員工編號']:<15}{row['營運單位']:<20}{row['英文全名']:<20}"
        if row['Id'] not in [result['Id'] for result in results]:
            name_list.insert(tk.END, display_text)
        else:
            name_list.insert(tk.END, display_text)
            name_list.itemconfig(tk.END, {'bg': 'lightgray'})
    if dim_selected and selected_winner is not None:
        # Dim the selected winner immediately after being chosen
        for i in range(name_list.size()):
            selected_display_text = f"{selected_winner.iloc[0]['員工編號']:<15}{selected_winner.iloc[0]['營運單位']:<20}{selected_winner.iloc[0]['英文全名']:<20}"
            if name_list.get(i) == selected_display_text:
                name_list.itemconfig(i, {'bg': 'lightgray'})
                break

# Function to update the prize list display
def update_prize_list():
    prize_list.delete(0, tk.END)
    try:
        start_prize = int(start_prize_entry.get())
        end_prize = int(end_prize_entry.get())
        prize_numbers = list(range(start_prize, end_prize + 1))
        for prize in prize_numbers:
            if prize not in assigned_prizes:
                prize_list.insert(tk.END, prize)
            else:
                prize_list.insert(tk.END, prize)
                prize_list.itemconfig(tk.END, {'bg': 'lightgray'})
    except ValueError:
        return

# Function to update the history list display
def update_history_list():
    history_list.delete(0, tk.END)
    for count, result in enumerate(results, start=1):
        display_text = f"{count:<10}{result['英文全名']:<20}{result['Prize Number']:<10}"
        history_list.insert(tk.END, display_text)

# Function to reset the application
def reset_app():
    global assigned_prizes, results, selected_winner, draw_count
    assigned_prizes = []
    results = []
    selected_winner = None
    draw_count = 0
    result_text.delete('1.0', tk.END)
    name_list.delete(0, tk.END)
    prize_list.delete(0, tk.END)
    history_list.delete(0, tk.END)
    prize_button.config(state=tk.DISABLED)
    if data_frame is not None:
        update_name_list()
        update_prize_list()

# Function to set prize number range
def set_prize_range():
    try:
        start_prize = int(start_prize_entry.get())
        end_prize = int(end_prize_entry.get())
        if start_prize > end_prize or start_prize < 1:
            raise ValueError("Invalid prize range.")
        update_prize_list()
    except ValueError:
        messagebox.showerror("Error", "Please enter a valid prize number range.")

        
        
# Initialize GUI
root = tk.Tk()
root.title("Lucky Draw Application")

# Fullscreen mode
root.state('zoomed')  # Windows fullscreen
# For cross-platform fullscreen, use: root.attributes('-fullscreen', True)

# Set larger font
default_font = ("Arial", 16)
button_width = 25  # Standard button width, increased for larger buttons
entry_width = 15   # Standard width for entry widgets
label_width = 20   # Standard width for labels

# Main frame to align all widgets to the left
main_frame = tk.Frame(root)
main_frame.pack(expand=True, fill='both')

# Upload file, Export results, and Reset buttons in the same row
upload_button = tk.Button(main_frame, text="Upload Excel File", command=upload_file, font=default_font, width=button_width)
upload_button.grid(row=0, column=0, padx=5, pady=10, sticky="w", columnspan=2)

export_button = tk.Button(main_frame, text="Export Results", command=export_results, font=default_font, width=button_width)
export_button.grid(row=0, column=2, padx=5, pady=10, sticky="w", columnspan=2)

reset_button = tk.Button(main_frame, text="Reset", command=reset_app, font=default_font, width=button_width)
reset_button.grid(row=0, column=4, padx=5, pady=10, sticky="w")

# Prize number range input
tk.Label(main_frame, text="Prize Number Range", font=default_font, width=label_width, anchor='w').grid(row=1, column=0, padx=5, pady=5, sticky="w")

# From, To, and Go button in the next row
tk.Label(main_frame, text="From", font=default_font, width=5, anchor='w').grid(row=2, column=0, padx=5, pady=5, sticky="w")
start_prize_entry = tk.Entry(main_frame, width=10, font=default_font)  # Decreased width for narrower entry
start_prize_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")

tk.Label(main_frame, text="To", font=default_font, width=5, anchor='w').grid(row=2, column=2, padx=5, pady=5, sticky="w")
end_prize_entry = tk.Entry(main_frame, width=10, font=default_font)  # Decreased width for narrower entry
end_prize_entry.grid(row=2, column=3, padx=5, pady=5, sticky="w")

# Set prize range button
go_button = tk.Button(main_frame, text="Go", command=set_prize_range, font=default_font, width=button_width)
go_button.grid(row=2, column=4, padx=5, pady=5, sticky="w")

# Draw name button
name_button = tk.Button(main_frame, text="Draw Name", command=draw_name, font=("Arial", 18), bg="#ff6b00", fg="white", width=button_width + 5)
name_button.grid(row=3, column=0, columnspan=3, padx=5, pady=10, sticky="w")

# Draw prize button
prize_button = tk.Button(main_frame, text="Draw Prize Number", command=draw_prize, state=tk.DISABLED, font=("Arial", 18), bg="#ff6b00", fg="white", width=button_width + 5)
prize_button.grid(row=3, column=3, columnspan=2, padx=5, pady=10, sticky="w")

# Result display
result_text = tk.Text(main_frame, height=10, width=95, font=("Arial", 18))
result_text.grid(row=4, column=0, columnspan=6, padx=10, pady=10, sticky="w")

# Name list display
tk.Label(main_frame, text="Name List", font=default_font, width=label_width, anchor='w').grid(row=5, column=0, columnspan=2, sticky="w")
name_list = tk.Listbox(main_frame, height=15, width=50, font=default_font)
name_list.grid(row=6, column=0, columnspan=2, padx=10, pady=10, sticky="w")

# Prize list display
tk.Label(main_frame, text="Prize List", font=default_font, width=label_width, anchor='w').grid(row=5, column=2, columnspan=1, sticky="w")
prize_list = tk.Listbox(main_frame, height=15, width=15, font=default_font)
prize_list.grid(row=6, column=2, columnspan=1, padx=10, pady=10, sticky="w")

# History list display
tk.Label(main_frame, text="History List", font=default_font, width=label_width, anchor='w').grid(row=5, column=3, columnspan=2, sticky="w")
history_list = tk.Listbox(main_frame, height=15, width=30, font=default_font)
history_list.grid(row=6, column=3, columnspan=2, padx=10, pady=10, sticky="w")

# Run the main loop
root.mainloop()
