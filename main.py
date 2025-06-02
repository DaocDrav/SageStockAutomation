import tkinter as tk
import subprocess
import os
import sys

def get_resource_path(relative_path):
    """Get absolute path to resource, works for dev and PyInstaller."""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except AttributeError:
        # when running in dev environment, just use current file's dir
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def run_pending_iqos():
    subprocess.run(["python", get_resource_path("scripts/pending_iqos.py")])

def run_un_obsolete():
    subprocess.run(["python", get_resource_path("scripts/unobsolete_product_codes.py")])

def run_promo_cost():
    subprocess.run(["python", get_resource_path("scripts/promo_cost.py")])


# Function to trigger count.py
def run_count():
    subprocess.Popen(["python", get_resource_path("scripts/count.py")])


# Create the Tkinter window
root = tk.Tk()
root.iconbitmap(get_resource_path(os.path.join("assets", "ibco_logo.ico")))
root.title("Ibco Stock Control Automation")
# Set the geometry (width x height)
root.geometry("600x600")  # Adjust these values for desired width and height

# Create a frame for "Stock Entry"
stock_entry_frame = tk.LabelFrame(root, text="Stock Entry", padx=10, pady=10)
stock_entry_frame.pack(padx=10, pady=10, fill="both")

# Add the "Run Count" button to the "Stock Entry" section
count_button = tk.Button(stock_entry_frame, text="Goods Inward Count", command=run_count, width=20)
count_button.pack(pady=5)

# Create a frame for "IQOS Transfers"
iqos_transfers_frame = tk.LabelFrame(root, text="IQOS Transfers", padx=10, pady=10)
iqos_transfers_frame.pack(padx=10, pady=10, fill="both")

# Add the "Run Pending IQOS" button to the "IQOS Transfers" section
pending_iqos_button = tk.Button(iqos_transfers_frame, text="Pending IQOS", command=run_pending_iqos, width=20)
pending_iqos_button.pack(pady=5)

# Create a frame for "Stock Code Maintenance"
stock_entry_frame = tk.LabelFrame(root, text="Stock Code Maintenance", padx=10, pady=10)
stock_entry_frame.pack(padx=10, pady=10, fill="both")

# Add the "Un-obsolete" button to the "Stock Code Maintenance" section
unobsolete_button = tk.Button(stock_entry_frame, text="Un-obsolete product codes in stock", command=run_un_obsolete, width=30)
unobsolete_button.pack(pady=5)

# Add a label with the text "Master Slave After Running this"
master_slave_label = tk.Label(stock_entry_frame, text="Master Slave after running this")
master_slave_label.pack(pady=5)

# Create a frame for "Promotions"
stock_entry_frame = tk.LabelFrame(root, text="Promotion Cost Maintenance", padx=10, pady=10)
stock_entry_frame.pack(padx=10, pady=10, fill="both")

# Add the "Promotion Cost Fix" button to the "Promotions" section
prom_cost_button = tk.Button(stock_entry_frame, text="Fix Promo WH Cost", command=run_promo_cost, width=30)
prom_cost_button.pack(pady=5)

# Add a label with the text "Fixes IQOS Promotions Cost Issue"
prom_cost_label = tk.Label(stock_entry_frame, text="Fixes IQOS Promotions Cost Issue")
prom_cost_label.pack(pady=5)

# Start the GUI loop
root.mainloop()
