import json
import tkinter as tk
from tkinter import ttk
import pyautogui
import time
from collections import defaultdict
import os
import sys

def get_resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        # base_path = root folder (one level up from current script folder)
        base_path = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    return os.path.join(base_path, relative_path)

# Load the JSON data
with open(get_resource_path("assets/selected_products.json"), "r") as file:
    data = json.load(file)

# Group data by Product
product_bins = defaultdict(list)
for record in data:
    product_bins[record["Product"]].append({
        "Bin Number": record["Bin Number"],
        "Quantity": int(record["Qty Free"]),  # Ensure quantity is stored as an integer
        "Best Before": record["Best Before"]
    })

# Function to update the table based on selected product
def update_table(product):
    # Clear the table
    for row in table.get_children():
        table.delete(row)
    # Add rows for the selected product
    if product in product_bins:
        for entry in product_bins[product]:
            table.insert(
                "",
                "end",
                values=(
                    entry["Bin Number"].upper(),  # Convert to uppercase
                    entry["Quantity"],           # Quantity is already an integer
                    entry["Best Before"]
                )
            )

# Function to automate data entry using PyAutoGUI
def automate_entry():
    selected_product = product_dropdown.get()
    if not selected_product or selected_product not in product_bins:
        print("Please select a valid product.")
        return

    print("Switch to Sage within 5 seconds...")
    time.sleep(5)  # Allow time to switch to Sage application

    # Automate entering bin data into Sage
    for entry in product_bins[selected_product]:
        pyautogui.typewrite(entry["Bin Number"].upper())  # Convert to uppercase
        pyautogui.press("tab")
        pyautogui.typewrite(str(entry["Quantity"]))  # Ensure integer
        pyautogui.press("tab")
        pyautogui.typewrite(entry["Best Before"])
        pyautogui.press("enter")  # Single Enter to move to the next row
        time.sleep(1)  # Add a delay between entries to ensure smooth operation

# Create the main window
root = tk.Tk()
root.title("Bin Management GUI")
root.geometry("600x400")

# Dropdown for products
product_label = tk.Label(root, text="Select Product")
product_label.pack(pady=5)

product_dropdown = ttk.Combobox(root, values=list(product_bins.keys()))
product_dropdown.pack(pady=5)

# Table for displaying bins, quantities, and best-before dates
columns = ("Bin Number", "Quantity", "Best Before")
table = ttk.Treeview(root, columns=columns, show="headings")
for col in columns:
    table.heading(col, text=col)
    table.column(col, anchor="center", width=150)
table.pack(pady=10, fill="x")

# Button to refresh the table based on selected product
refresh_button = tk.Button(root, text="Show Bins", command=lambda: update_table(product_dropdown.get()))
refresh_button.pack(pady=5)

# Button to automate data entry into Sage
automate_button = tk.Button(root, text="Paste to Sage", command=automate_entry)
automate_button.pack(pady=5)

# Run the GUI
root.mainloop()
