import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import random
import time

def analyze_attendance():
    if data_frame is None:
        messagebox.showerror("Error", "Please upload an Excel file first.")
        return
    
    # Analyze attendance distribution
    attendance_dist = data_frame['出席人數'].value_counts().sort_index()
    total_attendance = data_frame['出席人數'].sum()
    
    analysis_text = f"Total attendance capacity: {total_attendance}\n\n"
    analysis_text += "Distribution of attendance numbers:\n"
    for attendance, count in attendance_dist.items():
        analysis_text += f"Attendance {attendance}: {count} entries (total capacity: {attendance * count})\n"
    
    result_text.delete('1.0', tk.END)
    result_text.insert(tk.END, analysis_text)

def rolling_effect():
    if data_frame is not None:
        for _ in range(20):  # Number of rolls
            random_name = data_frame.sample(n=1)
            result_text.delete('1.0', tk.END)
            result_text.insert(tk.END, random_name.to_string(index=False))
            root.update()
            time.sleep(0.1)  # Delay between rolls

def select_winners_with_effect():
    """
    Selects winners for a lucky draw using a fair, unbiased algorithm.
    """
    try:
        # --- 1. Input Validation ---
        event_name = event_name_entry.get()
        target_prize_count = int(prize_count_entry.get())

        if not event_name or target_prize_count <= 0:
            messagebox.showerror("Error", "Please enter a valid event name and a positive number of prizes.")
            return
        if data_frame is None:
            messagebox.showerror("Error", "Please upload an Excel file first.")
            return

        # --- 2. Visual Effect ---
        # This is purely for show and does not impact the final result's fairness.
        rolling_effect()

        # --- 3. Fair Selection Logic ---
        remaining_prizes = target_prize_count
        winners = pd.DataFrame()
        available_df = data_frame.copy()  # Use a copy to keep the original data intact.

        # Loop until all prizes are awarded or no eligible participants are left.
        while remaining_prizes > 0 and not available_df.empty:
            
            # STEP A: Identify ALL eligible candidates for the current round.
            # An eligible candidate's attendance count must be less than or equal to the remaining prizes.
            valid_candidates = available_df[available_df['出席人數'] <= remaining_prizes]

            # STEP B: If no one is eligible, stop the draw.
            # This happens if all remaining groups are too large for the prize pool.
            if valid_candidates.empty:
                messagebox.showinfo("Draw Concluding",
                                  f"No more eligible participants found.\n\n"
                                  f"There are {remaining_prizes} prizes left, but all remaining groups "
                                  f"are larger than this number. The draw will now end.")
                break

            # STEP C: Select ONE winner randomly from the ENTIRE pool of eligible candidates.
            # This is the core of a fair draw: every entry in 'valid_candidates' has an equal chance.
            candidate = valid_candidates.sample(n=1)

            # STEP D: Process the selected winner.
            attendance = candidate['出席人數'].values[0]

            # Add the winner to the winners list
            winners = pd.concat([winners, candidate])

            # Subtract their attendance count from the remaining prizes
            remaining_prizes -= attendance

            # Remove the winner from the available pool so they cannot be drawn again
            available_df = available_df[available_df['Id'] != candidate['Id'].values[0]]

        # --- 4. Display Results ---
        if remaining_prizes > 0:
            messagebox.showwarning("Warning", 
                                 f"Could not allocate all prizes. {remaining_prizes} prizes remaining.")

        total_allocated = winners['出席人數'].sum()
        winners['Event Name'] = event_name
        
        # Clear previous results and show the new ones
        result_text.delete('1.0', tk.END)
        result_text.insert(tk.END, f"🏆 Draw Results for {event_name} 🏆\n")
        result_text.insert(tk.END, "---------------------------------\n")
        result_text.insert(tk.END, f"Total Prizes to Award: {target_prize_count}\n")
        result_text.insert(tk.END, f"Total Prizes Allocated: {total_allocated}\n")
        result_text.insert(tk.END, f"Prizes Remaining: {remaining_prizes}\n\n")
        result_text.insert(tk.END, "WINNERS:\n")
        result_text.insert(tk.END, winners.to_string(index=False))
        
        # Show distribution of winners by attendance number
        winner_dist = winners['出席人數'].value_counts().sort_index()
        result_text.insert(tk.END, "\n\n--- Winner Distribution ---\n")
        for attendance, count in winner_dist.items():
            result_text.insert(tk.END, f"Groups of {attendance}: {count} winner(s)\n")

        # --- 5. Save Results to File ---
        output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                                     filetypes=[("Excel files", "*.xlsx")])
        if output_file:
            winners.to_excel(output_file, index=False, engine='openpyxl')
            messagebox.showinfo("Success", f"Results successfully saved to {output_file}")

    except ValueError:
        messagebox.showerror("Error", "Number of prizes must be a valid integer.")
    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred: {e}")


def upload_file():
    global data_frame
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        try:
            data_frame = pd.read_excel(file_path, sheet_name=0)
            required_columns = ["Id", "中文全名", "英文全名", "員工編號", "營運單位", "區域", "出席人數", "額外購買門票"]
            if not all(col in data_frame.columns for col in required_columns):
                raise ValueError("Uploaded Excel file does not have the required columns.")
            messagebox.showinfo("Success", "File uploaded successfully.")
            
            # Show initial analysis
            analyze_attendance()
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

# Analyze button
analyze_button = tk.Button(root, text="Analyze Attendance", command=analyze_attendance)
analyze_button.grid(row=3, column=0, columnspan=2, pady=10)

# Select winners button
draw_button = tk.Button(root, text="Select Winners", command=select_winners_with_effect)
draw_button.grid(row=4, column=0, columnspan=2, pady=10)

# Result display
result_text = tk.Text(root, height=30, width=100)
result_text.grid(row=5, column=0, columnspan=2, padx=10, pady=10)

root.mainloop()