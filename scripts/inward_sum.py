import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import pandas as pd
import os
import pyautogui
import time
import sys
import subprocess

def get_resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        # base_path = root folder (one level up from current script folder)
        base_path = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    return os.path.join(base_path, relative_path)

# Load and process data
df = pd.read_json(get_resource_path("assets/selected_products.json"), dtype={'Product': str})
df_grouped = df.groupby(["Product", "Long Description", "Brand", "Pack Size", "Cost Price"], as_index=False)[['Qty Free']].sum()

# Calculate Grand Total
grand_total = df_grouped["Qty Free"].sum()

# Function to automate purchase order entry using PyAutoGUI
def update_purchase_order():
    # Disable button and show processing state
    update_button.config(text="Processing...", state=tk.DISABLED)

    # Popup instruction for the user
    messagebox.showinfo("Switch to Sage Screen", "Please switch to the Sage Purchase Order screen before continuing.")

    # Add a short delay to give the user time to switch screens
    time.sleep(5)  # Adjust the number of seconds as needed

    # Automate keystrokes for each row in the dataframe
    for index, row in df_grouped.iterrows():
        product_code = row['Product']
        qty_free = row['Qty Free']
        cost_price = row['Cost Price']

        # Type warehouse code "01"
        pyautogui.write("01")  # Enter "01" (warehouse code)
        time.sleep(0.2)  # Ensure "01" is registered

        # Move to the product code field
        pyautogui.press("tab")  # Tab to product code field
        time.sleep(0.5)  # Slightly longer delay for Sage to process

        # Enter the product code one character at a time
        for char in product_code:
            pyautogui.write(char)
            time.sleep(0.05)  # Add small delay for stability

        pyautogui.press("enter")  # Confirm product code
        time.sleep(0.2)

        # Enter the quantity free
        pyautogui.write(str(qty_free))  # Enter quantity
        pyautogui.press("enter")  # Confirm quantity
        time.sleep(0.2)
        pyautogui.press("enter")  # Extra confirm
        time.sleep(0.2)

        # Enter the cost price
        pyautogui.write(str(cost_price))  # Enter cost price
        pyautogui.press("enter")  # Confirm cost price
        time.sleep(0.5)  # Delay before processing next row

    # Reset button to original state
    update_button.config(text="Update Purchase Order", state=tk.NORMAL)

# Function to run "paste_sage.py" and close current GUI
def run_paste_sage():
    script_path = get_resource_path("scripts/paste_sage.py")
    subprocess.run([sys.executable, script_path])
    root.destroy()

# Create Tkinter GUI
def create_gui():
    global root, update_button

    root = tk.Tk()
    root.title("Product Summary")

    # Create Treeview for the DataFrame
    tree = ttk.Treeview(root, columns=list(df_grouped.columns), show="headings", height=15)
    for col in df_grouped.columns:
        tree.heading(col, text=col)
        tree.column(col, width=120, anchor=tk.CENTER)

    for _, row in df_grouped.iterrows():
        tree.insert("", tk.END, values=list(row))

    tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    # Add Text field for Grand Total
    total_frame = tk.Frame(root)
    total_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    total_label = tk.Label(total_frame, text="Grand Total Qty:", font=("Arial", 12))
    total_label.pack(side=tk.LEFT, padx=5)

    total_value = tk.Text(total_frame, height=1, width=20, font=("Arial", 12))
    total_value.insert(tk.END, f'{grand_total}')
    total_value.config(state=tk.DISABLED)
    total_value.pack(side=tk.LEFT)

    # Add buttons
    button_frame = tk.Frame(root)
    button_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    # Update Purchase Order Button
    update_button = tk.Button(button_frame, text="Update Purchase Order", font=("Arial", 12), command=update_purchase_order)
    update_button.pack(side=tk.LEFT, padx=10, pady=10)

    # Bin Summary Button
    bin_button = tk.Button(button_frame, text="Bin Summary", font=("Arial", 12), command=run_paste_sage)
    bin_button.pack(side=tk.LEFT, padx=10, pady=10)

    root.mainloop()

# Run the GUI
if __name__ == "__main__":
    create_gui()
