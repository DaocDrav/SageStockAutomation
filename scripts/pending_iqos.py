from sqlalchemy import create_engine
from read_sql import Read_sql_from_file
import pandas as pd
from scipy.stats import mode
import tkinter as tk
from tkinter import ttk
import time
import subprocess
from tkinter import messagebox
import os
import sys
import json

def get_resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        # base_path = root folder (one level up from current script folder)
        base_path = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    return os.path.join(base_path, relative_path)


# Use the helper function to locate db_config.json
config_path = get_resource_path('assets/db_config.json')

print(config_path)

with open(config_path, 'r') as config_file:
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
    sql_query = Read_sql_from_file(get_resource_path("SQL/ToIqos.sql"))

    # Load data into the dataframe
    df_iqos = pd.read_sql(sql_query, conn)

with engine.connect() as conn:
    # Define the SQL query
    sql_query = Read_sql_from_file(get_resource_path("SQL/pallet_size.sql"))

    # Load data into the dataframe
    df_pallet = pd.read_sql(sql_query, conn)

with engine.connect() as conn:
    # Define the SQL query
    sql_query = Read_sql_from_file(get_resource_path("SQL/next_iqos_bin.sql"))

    # Load data into the dataframe
    df_next = pd.read_sql(sql_query, conn)

# Group by product and calculate the mode (most common quantity)
common_quantities = df_pallet.groupby('product')['movement_quantity'].apply(lambda x: mode(x, keepdims=True).mode[0]).reset_index()
common_quantities.columns = ['product', 'common_quantity']

# Ensure only one row per product
unique_products = common_quantities.drop_duplicates(subset=['product']).reset_index(drop=True)

# Merge the unique_products dataframe with df_iqos on the 'product' column
df_final = pd.merge(df_iqos, unique_products, on='product', how='left')

# Check the result
#print(df_final)

#print(df_next)





# Filter df_next by products in df_final
filtered_df_next = df_next[df_next['product'].isin(df_final['product'])]

# Group by 'warehouse' and 'product', summing 'quantity_free', and counting unique bins for 'total_bins_for_product'
grouped_df_next = (
    filtered_df_next.groupby(['warehouse', 'product'], as_index=False).agg({
        'quantity_free': 'sum',
        'bin_number': 'nunique'  # Assuming 'bin_number' represents individual bins
    })
)


# Rename 'bin_number' to 'total_bins_for_product' for clarity
grouped_df_next.rename(columns={'bin_number': 'total_bins_for_product'}, inplace=True)

# Drop unnecessary columns
columns_to_drop = ['best_before_date', 'common_quantity', 'cumulative_quantity']
cleaned_df_next = grouped_df_next.drop(columns=[col for col in columns_to_drop if col in filtered_df_next.columns], errors='ignore')

# Check the result
print(cleaned_df_next)

# Save cleaned_df_next to a JSON file
json_path = get_resource_path("assets/cleaned_df_next.json")
cleaned_df_next.to_json(json_path, index=False)

# Warn if the same product exists in multiple warehouses
conflicting_products = (
    filtered_df_next.groupby('product')['warehouse'].nunique()
    .reset_index()
    .query('warehouse > 1')
)

if not conflicting_products.empty:
    products_list = "\n".join(conflicting_products['product'].tolist())
    messagebox.showwarning(
        "Warehouse Conflict Detected",
        f"The following products exist in multiple warehouses:\n\n{products_list}\n\n"
        "Please transfer these products into the same warehouse before continuing."
    )


# Create the main window
root = tk.Tk()
root.iconbitmap(get_resource_path("assets/ibco_logo.ico"))

# Maximize the window
root.state('zoomed')
root.title("Pending IQOS")

# Exclude unnecessary columns
columns_to_display = [col for col in df_final.columns if col not in ['TotalAllocated', 'OppickBin', 'OrderNo']]

# Update Treeview to only include the necessary columns
tree = ttk.Treeview(root)
tree["columns"] = columns_to_display

# Format the columns
tree.column("#0", width=0, stretch=tk.NO)
for column in tree["columns"]:
    tree.column(column, anchor=tk.W, width=100)
    tree.heading(column, text=column, anchor=tk.W)

# Insert the filtered data into the Treeview
for index, row in df_final[columns_to_display].iterrows():
    tree.insert("", tk.END, values=list(row))

# Add the Treeview to the window
tree.pack(fill=tk.BOTH, expand=False)



# Create the second Treeview widget for df_next
tree_next = ttk.Treeview(root)

# Define the columns for df_next
tree_next["columns"] = filtered_df_next.columns.tolist()

print(filtered_df_next)
# Format the columns for df_next
tree_next.column("#0", width=0, stretch=tk.NO)
for column in tree_next["columns"]:
    tree_next.column(column, anchor=tk.W, width=100)
    tree_next.heading(column, text=column, anchor=tk.W)

# Insert the data from cleaned_df_next into the second Treeview
for index, row in filtered_df_next.iterrows():
    tree_next.insert("", tk.END, values=list(row))

# Add some space between the two Treeviews
spacer = tk.Label(root, text="")
spacer.pack()

# Add the second Treeview to the window
tree_next.pack(fill=tk.BOTH, expand=False)

# Function to handle the "Execute IQOS" button click
def execute_iqos():
    # Show a dialog box to prompt the user
    messagebox.showinfo("Switch to Sage", "Please switch to Sage. Execution will start in 2 seconds.")

    # Wait for 2 seconds to allow the user to switch
    time.sleep(2)

    # Call the execute_iqos.py script
    script_path = get_resource_path("scripts/execute_iqos.py")
    subprocess.run([sys.executable, script_path])

# Create a button below the Treeviews
execute_button = tk.Button(root, text="Execute IQOS", command=execute_iqos)
execute_button.pack(pady=10)

def iqos_sent():

    # Call the iqos_sent.py script
    script_path = get_resource_path("scripts/iqos_sent.py")
    subprocess.run([sys.executable, script_path])




# Function to refresh connections and reload data
def refresh_connections():
    # Reconnect to the database and reload the data
    with engine.connect() as conn:
        # Reload data from ToIqos.sql
        sql_query = Read_sql_from_file(get_resource_path("SQL/ToIqos.sql"))
        refreshed_df_iqos = pd.read_sql(sql_query, conn)

        # Reload data from pallet_size.sql
        sql_query = Read_sql_from_file(get_resource_path("SQL/pallet_size.sql"))
        refreshed_df_pallet = pd.read_sql(sql_query, conn)

        # Reload data from next_iqos_bin.sql
        sql_query = Read_sql_from_file(get_resource_path("SQL/next_iqos_bin.sql"))
        refreshed_df_next = pd.read_sql(sql_query, conn)

    # Process the refreshed data
    common_quantities = refreshed_df_pallet.groupby('product')['movement_quantity'].apply(
        lambda x: mode(x, keepdims=True).mode[0]
    ).reset_index()
    common_quantities.columns = ['product', 'common_quantity']

    unique_products = common_quantities.drop_duplicates(subset=['product']).reset_index(drop=True)
    refreshed_df_final = pd.merge(refreshed_df_iqos, unique_products, on='product', how='left')

    filtered_df_next = refreshed_df_next[refreshed_df_next['product'].isin(refreshed_df_final['product'])]

    # Update the Treeviews with refreshed data
    for item in tree.get_children():
        tree.delete(item)  # Clear existing data from the Treeview
    for index, row in refreshed_df_final[columns_to_display].iterrows():
        tree.insert("", tk.END, values=list(row))  # Insert refreshed data into Treeview

    for item in tree_next.get_children():
        tree_next.delete(item)  # Clear existing data from the second Treeview
    for index, row in filtered_df_next.iterrows():
        tree_next.insert("", tk.END, values=list(row))  # Insert refreshed data into Treeview

    # Save cleaned_df_next to a JSON file
    cleaned_df_next.to_json("cleaned_df_next.json", index=False)

    # Warn if the same product exists in multiple warehouses
    conflicting_products = (
        filtered_df_next.groupby('product')['warehouse'].nunique()
        .reset_index()
        .query('warehouse > 1')
    )

    if not conflicting_products.empty:
        products_list = "\n".join(conflicting_products['product'].tolist())
        messagebox.showwarning(
            "Warehouse Conflict Detected",
            f"The following products exist in multiple warehouses:\n\n{products_list}\n\n"
            "Please transfer these products into the same warehouse before continuing."
        )

    messagebox.showinfo("Refresh Complete", "Data connections and display have been refreshed!")





# Add a Refresh button to the GUI
refresh_button = tk.Button(root, text="Refresh", command=refresh_connections)
refresh_button.pack(side=tk.RIGHT, padx=10, pady=10)

# Add the button for sent IQOS transfers to the GUI
refresh_button = tk.Button(root, text="Pending Iqos", command=iqos_sent)
refresh_button.pack(side=tk.RIGHT, padx=10, pady=5)

# Start the GUI event loop
root.mainloop()

# Close the connection
conn.close()