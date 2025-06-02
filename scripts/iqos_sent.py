from sqlalchemy import create_engine
from read_sql import Read_sql_from_file
import pandas as pd
import tkinter as tk
from tkinter import ttk
import sys
import json
import os  # To run the external script
import subprocess

def get_resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        # base_path = root folder (one level up from current script folder)
        base_path = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    return os.path.join(base_path, relative_path)

# Load the configuration file
with open(get_resource_path('assets/db_config.json'), 'r') as config_file:
    config = json.load(config_file)

# Construct the SQLAlchemy connection string
conn_str = (
    f"mssql+pyodbc://{config['database']['username']}:{config['database']['password']}@"
    f"{config['database']['server']}/{config['database']['database_name']}?"
    f"driver=ODBC+Driver+17+for+SQL+Server"
)

# Create an engine using SQLAlchemy
engine = create_engine(conn_str)

# Create a connection using the engine
with engine.connect() as conn:
    # Define the SQL query
    sql_query = Read_sql_from_file(get_resource_path("SQL/iqos_sent.sql"))

    # Load data into the dataframe
    df_iqos_sent = pd.read_sql(sql_query, conn)

# Filter and rename the columns
columns_to_display = {
    "order_no": "Order No",
    "customer_order_no": "Customer Order No",
    "date_entered": "Date Entered",
    "product": "Product",
    "long_description": "Long Description",
    "analysis_b": "Pack Size",
    "bin": "Bin",
    "warehouse": "Warehouse",
    "quantity": "Quantity",
    "bin_number": "Iqos Site",
    "physical_qty": "Iqos Qty"
}

# Select the relevant columns and rename them
df_filtered = df_iqos_sent[list(columns_to_display.keys())].rename(columns=columns_to_display)

# Remove rows where any column in 'columns_to_display' values is blank or null
df_filtered.dropna(subset=columns_to_display.values(), inplace=True)

# Format the 'Date Entered' column
df_filtered["Date Entered"] = pd.to_datetime(df_filtered["Date Entered"]).dt.date

# Reset the index to be numeric
df_filtered.reset_index(drop=True, inplace=True)

# Create the main window
root = tk.Tk()
root.title("IQOS Sent Transfers")

# Create a treeview widget
tree = ttk.Treeview(root, height=20)


# Define the columns
tree["columns"] = list(df_filtered.columns)

# Format the Treeview columns
tree.column("#0", width=50, anchor=tk.W)  # Numeric index as the default column
tree.heading("#0", text="Index", anchor=tk.W)  # Add a heading for the default column

for column in tree["columns"]:
    tree.column(column, anchor=tk.W, width=100)
    tree.heading(column, text=column, anchor=tk.W)

# Add data to the Treeview
for idx, row in df_filtered.iterrows():
    tree.insert("", tk.END, text=idx, values=row.tolist())  # Use the index as the text for the #0 column

# Pack the Treeview
tree.pack(fill=tk.BOTH, expand=1)

# Define the function to run the external script
def run_send_email():
    script_path = get_resource_path("scripts/send_pend_email.py")
    subprocess.run(["python", script_path])

# Add the button to run the script
button = ttk.Button(root, text="Send IQOS Email", command=run_send_email)
button.pack(pady=10)  # Adds some padding around the button

# Run the application
root.mainloop()
