import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import random
import time

# Global set to track selected IDs across all events
selected_ids_global = set()

def rolling_effect():
    if data_frame is not None:
        for _ in range(20):  # Number of rolls
            random_name = data_frame.sample(n=1)
            result_text.delete('1.0', tk.END)
            result_text.insert(tk.END, random_name.to_string(index=False))
            root.update()
            time.sleep(0.1)  # Delay between rolls

def select_winners_for_event(event_name, prize_count, rules):
    try:
        if not event_name or prize_count <= 0:
            messagebox.showerror("Error", f"Please enter a valid event name and number of prizes for {event_name}.")
            return
        if data_frame is None:
            messagebox.showerror("Error", "Please upload an Excel file first.")
            return

        # Select winners based on rules
        total_people = 0
        winners = pd.DataFrame()
        attempts = 0
        max_attempts = 1000  # Limit the number of attempts to prevent infinite loop

        for people_count, count in rules.items():
            for _ in range(count):
                if total_people >= prize_count:
                    break
                remaining_count = prize_count - total_people
                candidates = data_frame[(data_frame["出席人數"] == people_count) & (~data_frame["Id"].isin(selected_ids_global))]
                if not candidates.empty:
                    candidate = candidates.sample(n=1)
                    candidate_id = candidate["Id"].values[0]
                    if candidate_id not in selected_ids_global:
                        winners = pd.concat([winners, candidate])
                        total_people += people_count
                        selected_ids_global.add(candidate_id)
                attempts += 1
                if attempts >= max_attempts:
                    break

        if total_people < prize_count:
            messagebox.showwarning("Warning", f"Could not find enough candidates for event {event_name}. Only {total_people} people selected.")

        winners['Event Name'] = event_name
        result_text.insert(tk.END, f"\nEvent: {event_name}\n")
        result_text.insert(tk.END, winners.to_string(index=False))
        return winners
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
        return None

def select_winners_with_effect():
    try:
        all_winners = pd.DataFrame()
        event_rules = {
            "Sea Turtle (GAST)": {4: 1, 2: 7},
            "Coral (GA)": {4: 1, 2: 7},
            "Meerkat & Aldabra Tortoise (MTA)": {4: 2, 2: 4},
            "Sea Lion (PP)": {3: 1, 2: 15},
            "Dolphin (DEx)": {3: 1, 2: 12},
            "Sea Jelly (SJS)": {3: 1, 2: 5},
            "Shark & Ray (SM)": {3: 1, 2: 5},
            "Rainforest (RF)": {4: 1, 2: 7},
            "Seal (NP) 1445": {3: 7, 1: 1, 2:50},
            "Seal (NP) 1525": {4: 1, 1: 4, 2:57},
        }

        for i in range(10):
            event_name = event_name_entries[i].get()
            prize_count = int(prize_count_entries[i].get())
            rolling_effect()  # Add rolling names effect here
            winners = select_winners_for_event(event_name, prize_count, event_rules[event_name])
            if winners is not None:
                all_winners = pd.concat([all_winners, winners])

        # Save result to Excel
        output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                                   filetypes=[("Excel files", "*.xlsx")])
        if output_file:
            all_winners.to_excel(output_file, index=False, engine='openpyxl')
            messagebox.showinfo("Success", f"Results saved to {output_file}")

        # List out candidates not picked for any events
        all_ids = set(data_frame["Id"].values)
        not_picked_ids = all_ids - selected_ids_global
        not_picked_candidates = data_frame[data_frame["Id"].isin(not_picked_ids)]
        result_text.insert(tk.END, "\nCandidates not picked for any events:\n")
        result_text.insert(tk.END, not_picked_candidates.to_string(index=False))
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# Function to load Excel file
def upload_file():
    global data_frame
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        try:
            data_frame = pd.read_excel(file_path, sheet_name=0)
            required_columns = ["Id", "中文全名", "英文全名", "員工編號", "營運單位", "出席人數"]
            if not all(col in data_frame.columns for col in required_columns):
                raise ValueError("Uploaded Excel file does not have the required columns.")
            messagebox.showinfo("Success", "File uploaded successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {e}")

# Initialize GUI
root = tk.Tk()
root.title("Lucky Draw Application")

data_frame = None

# Event name and prize count inputs
event_name_entries = []
prize_count_entries = []

events = [
    ("Sea Turtle (GAST)", 18),
    ("Coral (GA)", 18),
    ("Meerkat & Aldabra Tortoise (MTA)", 16),
    ("Sea Lion (PP)", 33),
    ("Dolphin (DEx)", 27),
    ("Sea Jelly (SJS)", 13),
    ("Shark & Ray (SM)", 13),
    ("Rainforest (RF)", 18),
    ("Seal (NP) 1445", 122),
    ("Seal (NP) 1525", 122)
]

for i, (event_name, prize_count) in enumerate(events):
    tk.Label(root, text=f"Event {i+1} Name:").grid(row=i, column=0, padx=10, pady=5, sticky="e")
    event_name_entry = tk.Entry(root, width=30)
    event_name_entry.insert(0, event_name)
    event_name_entry.grid(row=i, column=1, padx=10, pady=5)
    event_name_entries.append(event_name_entry)

    tk.Label(root, text=f"Number of Prizes for Event {i+1}:").grid(row=i, column=2, padx=10, pady=5, sticky="e")
    prize_count_entry = tk.Entry(root, width=10)
    prize_count_entry.insert(0, prize_count)
    prize_count_entry.grid(row=i, column=3, padx=10, pady=5)
    prize_count_entries.append(prize_count_entry)

# Upload file button
upload_button = tk.Button(root, text="Upload Excel File", command=upload_file)
upload_button.grid(row=10, column=0, columnspan=4, pady=10)

# Select winners button
draw_button = tk.Button(root, text="Select Winners", command=select_winners_with_effect)
draw_button.grid(row=11, column=0, columnspan=4, pady=10)

# Result display
result_text = tk.Text(root, height=30, width=100)
result_text.grid(row=12, column=0, columnspan=4, padx=10, pady=10)

root.mainloop()