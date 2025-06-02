import tkinter as tk
from tkinter import ttk, font
import pandas as pd
import os
import json
import pyodbc
import sys
import subprocess


def get_resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        # base_path = root folder (one level up from current script folder)
        base_path = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    return os.path.join(base_path, relative_path)

# Function to execute "Inward_Sum.py"
def run_inward_sum():
    script_path = get_resource_path("scripts/inward_sum.py")
    subprocess.run([sys.executable, script_path])

# Database connection class
class ProductsManager:
    def __init__(self, config_file=None):
        if config_file is None:
            config_file = get_resource_path("assets/db_config.json")
        else:
            config_file = get_resource_path(config_file)
        # Load database configuration
        with open(config_file, "r") as file:
            config = json.load(file)
        self.conn_str = (
            f"DRIVER={{ODBC Driver 17 for SQL Server}};"
            f"SERVER={config['database']['server']};"
            f"DATABASE={config['database']['database_name']};"
            f"UID={config['database']['username']};"
            f"PWD={config['database']['password']}"
        )

    def get_products(self):
        """Fetch product codes, descriptions, and supplier names for filtering."""
        with pyodbc.connect(self.conn_str) as conn:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT product, long_description, analysis_c AS supplier FROM scheme.stockm WHERE stockm.warehouse='zz'"
            )
            return [(row[0], row[1].strip(), row[2].strip() if row[2] else "") for row in cursor.fetchall()]

    def get_product_details(self, product_code):
        """Fetch product details based on the selected product code."""
        query = """
        SELECT 
            stockm.product, 
            stockm.alpha AS group_type, 
            stockm.long_description AS product_description, 
            stockm.analysis_a AS brand, 
            stockm.analysis_b AS pack_size, 
            stockm.analysis_c AS supplier, 
            SUM(stquem.quantity_free) AS qty_free
        FROM 
            scheme.stockm
        LEFT JOIN 
            scheme.stquem 
        ON 
            stockm.product = stquem.product AND stockm.warehouse = stquem.warehouse
        WHERE 
            stockm.product = ?
        GROUP BY 
            stockm.product, stockm.alpha, stockm.long_description, 
            stockm.analysis_a, stockm.analysis_b, stockm.analysis_c;
        """
        with pyodbc.connect(self.conn_str) as conn:
            cursor = conn.cursor()
            cursor.execute(query, (product_code,))
            return cursor.fetchone()


class ProductApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Goods Inward Count")
        self.manager = ProductsManager()
        self.all_products = self.manager.get_products()

        #Connection String
        config_path = get_resource_path('assets/db_config.json')
        with open(config_path, 'r') as config_file:
            config = json.load(config_file)
        self.conn_str = (
            f"DRIVER={{ODBC Driver 17 for SQL Server}};"
            f"SERVER={config['database']['server']};"
            f"DATABASE={config['database']['database_name']};"
            f"UID={config['database']['username']};"
            f"PWD={config['database']['password']}"
        )

        # Search field
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(root, textvariable=self.search_var, width=50)
        self.search_entry.grid(row=0, column=0, padx=10, pady=5)
        self.search_entry.bind("<KeyRelease>", self.update_dropdown)
        # Search field label
        self.search_label = ttk.Label(root, text="Product Search")
        self.search_label.grid(row=0, column=0, padx=10, pady=(45, 0))


        # Supplier filter field
        self.supplier_var = tk.StringVar()
        self.supplier_entry = ttk.Entry(root, textvariable=self.supplier_var, width=30)
        self.supplier_entry.grid(row=0, column=1, padx=10, pady=5)
        self.supplier_entry.bind("<KeyRelease>", self.update_dropdown)
        #Supplier field label
        self.supplier_label = ttk.Label(root, text="Supplier Filter")
        self.supplier_label.grid(row=0, column=1, padx=10, pady=(45, 0))

        self.listbox = tk.Listbox(root, width=100)
        self.listbox.grid(row=1, column=0, columnspan=2, padx=10, pady=5)
        self.listbox.bind("<<ListboxSelect>>", self.select_product)

        # Initialize Treeview
        self.tree = ttk.Treeview(root, columns=(
            "Product", "Description", "Brand", "Pack Size",
            "Cost Price", "Qty Free", "Bin Number", "Best Before"), show="headings", height=26)

        columns = ["Product", "Description", "Brand", "Pack Size",
                   "Cost Price", "Qty Free", "Bin Number", "Best Before"]

        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="center")  # Center the values

        self.tree.grid(row=0, column=2, rowspan=10, padx=80, pady=5, sticky='e')
        self.tree.bind("<Double-1>", self.edit_cost_price)

        self.auto_resize_columns()

        # Running Summary (Right Side)
        self.df = pd.DataFrame(columns=["Product", "Long Description", "Brand", "Pack Size", "Cost Price", "Qty Free", "Bin Number", "Best Before"])
        self.json_file = get_resource_path(os.path.join("assets/selected_products.json"))
        self.load_from_json()  # Ensure Treeview is initialized before this is called

        # Remaining UI setup (Labels, buttons, etc.)
        self.labels = ["Product", "Group Type", "Description", "Brand", "Pack Size", "Supplier", "Qty Free"]
        self.entries = {}
        for i, label in enumerate(self.labels):
            tk.Label(root, text=label).grid(row=i+2, column=0, padx=10, pady=5, sticky="e")
            self.entries[label] = tk.Entry(root, width=50)
            self.entries[label].grid(row=i+2, column=1, padx=10, pady=5)

        # Bin Number input
        tk.Label(root, text="Bin Number").grid(row=len(self.labels)+2, column=0, padx=10, pady=5, sticky="e")
        self.bin_entry = tk.Entry(root, width=50)
        self.bin_entry.grid(row=len(self.labels)+2, column=1, padx=10, pady=5)

        # Best Before input
        tk.Label(root, text="Best Before").grid(row=len(self.labels)+3, column=0, padx=10, pady=5, sticky="e")
        self.best_before_entry = tk.Entry(root, width=50)
        self.best_before_entry.grid(row=len(self.labels)+3, column=1, padx=10, pady=5)

        # Quantity input
        tk.Label(root, text="Quantity").grid(row=len(self.labels)+4, column=0, padx=10, pady=5, sticky="e")
        self.qty_entry = tk.Entry(root, width=10)
        self.qty_entry.grid(row=len(self.labels)+4, column=1, padx=10, pady=5)

        # Save button
        self.save_button = tk.Button(root, text="Add to Summary", command=self.add_to_summary)
        self.save_button.grid(row=len(self.labels)+5, column=0, columnspan=2, pady=10)

        # Delete button
        self.delete_button = tk.Button(root, text="Delete Selected Entry", command=self.delete_selected)
        self.delete_button.grid(row=len(self.labels)+6, column=0, columnspan=2, pady=10)

        # Clear Form button
        self.clear_button = tk.Button(root, text="Clear Form", command=self.clear_form)
        self.clear_button.grid(row=len(self.labels) + 7, column=0, columnspan=2, pady=10)

        # Button to execute Inward_Sum.py
        inward_sum_button = tk.Button(root, text="Purchase Order Summary", font=("Arial", 12), command=run_inward_sum)
        inward_sum_button.grid(
            row=len(self.labels) + 8, column=0, columnspan=2, padx=10, pady=10)


    def auto_resize_columns(self):
        """ Adjust column width to fit largest value with a minimum width """
        f = font.Font()  # Get the default font
        min_widths = {
            "Product": 100,
            "Description": 200,
            "Brand": 100,
            "Pack Size": 100,
            "Cost Price": 100,
            "Qty Free": 100,
            "Bin Number": 100,
            "Best Before": 120,
        }  # Set minimum widths for better readability

        for col in self.tree["columns"]:
            max_width = f.measure(col)  # Start with column title width
            for item in self.tree.get_children():
                text = self.tree.item(item, "values")[self.tree["columns"].index(col)]
                max_width = max(max_width, f.measure(str(text)))  # Measure text width

            # Set column width, ensuring it doesn't go below the minimum width
            self.tree.column(col, width=max(max_width + 20, min_widths.get(col, 100)))


        # Running Summary (Right Side)
        self.df = pd.DataFrame(columns=["Product", "Long Description", "Brand", "Pack Size", "Cost Price", "Qty Free", "Bin Number", "Best Before"])
        self.json_file = get_resource_path(os.path.join("assets/selected_products.json"))
        self.load_from_json()  # Ensure Treeview is initialized before this is called

        # Remaining UI setup (Labels, buttons, etc.)
        self.labels = ["Product", "Group Type", "Description", "Brand", "Pack Size", "Supplier", "Qty Free"]
        self.entries = {}
        for i, label in enumerate(self.labels):
            tk.Label(root, text=label).grid(row=i+2, column=0, padx=10, pady=5, sticky="e")
            self.entries[label] = tk.Entry(root, width=50)
            self.entries[label].grid(row=i+2, column=1, padx=10, pady=5)

        # Bin Number input
        tk.Label(root, text="Bin Number").grid(row=len(self.labels)+2, column=0, padx=10, pady=5, sticky="e")
        self.bin_entry = tk.Entry(root, width=50)
        self.bin_entry.grid(row=len(self.labels)+2, column=1, padx=10, pady=5)

        # Best Before input
        tk.Label(root, text="Best Before").grid(row=len(self.labels)+3, column=0, padx=10, pady=5, sticky="e")
        self.best_before_entry = tk.Entry(root, width=50)
        self.best_before_entry.grid(row=len(self.labels)+3, column=1, padx=10, pady=5)

        # Quantity input
        tk.Label(root, text="Quantity").grid(row=len(self.labels)+4, column=0, padx=10, pady=5, sticky="e")
        self.qty_entry = tk.Entry(root, width=10)
        self.qty_entry.grid(row=len(self.labels)+4, column=1, padx=10, pady=5)

        # Save button
        self.save_button = tk.Button(root, text="Add to Summary", command=self.add_to_summary)
        self.save_button.grid(row=len(self.labels)+5, column=0, columnspan=2, pady=10)

        # Delete button
        self.delete_button = tk.Button(root, text="Delete Selected Entry", command=self.delete_selected)
        self.delete_button.grid(row=len(self.labels)+6, column=0, columnspan=2, pady=10)

        # Clear Form button
        self.clear_button = tk.Button(root, text="Clear Form", command=self.clear_form)
        self.clear_button.grid(row=len(self.labels) + 7, column=0, columnspan=2, pady=10)

        # Button to execute Inward_Sum.py
        inward_sum_button = tk.Button(root, text="Purchase Order Summary", font=("Arial", 12), command=run_inward_sum)
        inward_sum_button.grid(
            row=len(self.labels) + 8, column = 0, columnspan = 2, padx = 10, pady = 10)


    def update_dropdown(self, event):
        """Update dropdown based on product search input and supplier filter."""
        search_term = self.search_var.get().lower()
        supplier_filter = self.supplier_var.get().lower()

        self.listbox.delete(0, tk.END)

        for product_code, description, supplier in self.all_products:
            if not supplier_filter or supplier_filter in supplier.lower():
                if search_term in description.lower() or search_term in product_code.lower():
                    self.listbox.insert(tk.END, f"{product_code} - {description}")

    def select_product(self, event):
        """Populate fields when a product is selected."""
        selected_index = self.listbox.curselection()
        if not selected_index:
            return

        # Get selected value (e.g., "010010150 - KOI CLEAN BLOCK")
        selected_product = self.listbox.get(selected_index[0])

        # Extract product code (everything before the ' - ')
        product_code = selected_product.split(" - ")[0]

        # Fetch product details based on product code
        result = self.manager.get_product_details(product_code)

        # Populate the GUI fields with the retrieved data
        if result:
            for i, label in enumerate(self.labels):
                self.entries[label].delete(0, tk.END)
                self.entries[label].insert(0, str(result[i]))

    def add_to_summary(self):
        """Add selected product details to the summary or create new entries for each bin."""
        product_code = self.entries["Product"].get().strip()
        bin_number = self.bin_entry.get().strip()  # Get Bin Number
        best_before = self.best_before_entry.get().strip()  # Get Best Before
        entered_qty = float(self.qty_entry.get() or 0)  # Default to 0 if empty

        if self.df.empty:
            self.df = pd.DataFrame(
                columns=["Product", "Long Description", "Brand", "Pack Size", "Cost Price", "Qty Free", "Bin Number",
                         "Best Before"])

        if product_code and bin_number:  # Ensure required fields are filled
            # Check if this Product + Bin combination already exists
            existing_entry = self.df[
                (self.df["Product"] == product_code) & (self.df["Bin Number"] == bin_number)
                ]

            if not existing_entry.empty:
                # If it exists, update the quantity and best before date
                existing_qty = existing_entry["Qty Free"].values[0]
                total_qty = existing_qty + entered_qty
                self.df.loc[
                    (self.df["Product"] == product_code) & (self.df["Bin Number"] == bin_number),
                    ["Qty Free", "Best Before"]
                ] = [total_qty, best_before]
            else:
                # If it doesn't exist, create a new entry
                product_data = {
                    "Product": product_code,
                    "Long Description": self.entries["Description"].get(),
                    "Brand": self.entries["Brand"].get(),
                    "Pack Size": self.entries["Pack Size"].get(),
                    "Cost Price": "",
                    "Qty Free": entered_qty,
                    "Bin Number": bin_number,
                    "Best Before": best_before
                }
                new_row = pd.DataFrame([product_data])

                # Check for empty or all-NA data in new_row
                if not new_row.empty:  # Ensure new_row isn't empty before concatenation
                    self.df = pd.concat([self.df, new_row], ignore_index=True)

            self.refresh_treeview()
            self.save_to_json()

    def delete_selected(self):
        """Delete the selected entry from Treeview, DataFrame, and JSON."""
        selected_item = self.tree.selection()  # Get the selected item in Treeview
        if not selected_item:
            print("No item selected to delete.")  # Debugging message
            return

        # Fetch the values of the selected row
        item_values = self.tree.item(selected_item, "values")
        if not item_values:
            print("Failed to fetch item values.")
            return

        product_code = str(item_values[0]).strip()  # Ensure it's a string and strip spaces
        long_description = str(item_values[1]).strip()  # Ensure it's a string and strip spaces
        brand = str(item_values[2]).strip()  # Ensure it's a string and strip spaces
        pack_size = str(item_values[3]).strip()  # Ensure it's a string and strip spaces
        cost_price = str(item_values[4]).strip()  # Ensure it's a string and strip spaces
        qty_free = str(item_values[5]).strip()  # Ensure it's a string and strip spaces
        bin_number = str(item_values[6]).strip()  # Ensure it's a string and strip spaces
        best_before = str(item_values[7]).strip()  # Ensure it's a string and strip spaces

        print(f"Attempting to delete Product: {product_code}, Bin Number: {bin_number}")

        # Ensure columns in DataFrame are strings and trimmed
        self.df["Product"] = self.df["Product"].astype(str).str.strip()
        self.df["Bin Number"] = self.df["Bin Number"].astype(str).str.strip()

        # Remove the entry from the DataFrame
        self.df = self.df[~((self.df["Product"] == product_code) & (self.df["Bin Number"] == bin_number))]

        # Refresh the Treeview
        self.refresh_treeview()

        # Save the updated DataFrame to JSON
        self.save_to_json()

        print(f"Deleted Product {product_code} with Bin Number {bin_number}.")

    def clear_form(self):
        """Clear all data from the Treeview, DataFrame, and JSON file."""
        # Clear the DataFrame
        self.df = self.df.iloc[0:0]  # This will clear all rows but keep the column structure intact

        # Clear the Treeview
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Save the cleared DataFrame to JSON
        self.save_to_json()

        print("All data has been cleared.")

    def refresh_treeview(self):
        """Update the Treeview with DataFrame contents."""
        self.tree.delete(*self.tree.get_children())
        for _, row in self.df.iterrows():
            self.tree.insert("", "end", values=(
            row["Product"], row["Long Description"], row["Brand"], row["Pack Size"], row["Cost Price"], row["Qty Free"],
            row["Bin Number"], row["Best Before"]))

    def edit_cost_price(self, event):
        """Allow user to edit cost price in Treeview."""
        selected_item = self.tree.selection()
        if not selected_item:
            return
        item_values = self.tree.item(selected_item, "values")
        product_code = item_values[0]

        def save_price():
            new_price = price_entry.get()
            is_per_kg = per_kg_var.get()  # Get the state of the checkbox
            if is_per_kg:
                # Get the weight of the product from the database
                query = "SELECT weight FROM scheme.stockm WHERE product = ?"
                with pyodbc.connect(self.conn_str) as conn:
                    cursor = conn.cursor()
                    cursor.execute(query, (product_code,))
                    weight = cursor.fetchone()[0]
                new_price = float(new_price) * weight  # Multiply by weight
            new_price = round(float(new_price), 4)  # Round to four decimal places
            self.df.loc[self.df["Product"] == product_code, "Cost Price"] = new_price
            self.refresh_treeview()
            self.save_to_json()
            popup.destroy()

        popup = tk.Toplevel(self.root)
        popup.title("Edit Cost Price")
        tk.Label(popup, text="Enter Cost Price:").pack(pady=5)
        price_entry = tk.Entry(popup)
        price_entry.pack(pady=5)
        price_entry.insert(0, item_values[4])

        per_kg_var = tk.BooleanVar()  # Create a variable to hold the checkbox state
        per_kg_checkbox = tk.Checkbutton(popup, text="Per KG", variable=per_kg_var)
        per_kg_checkbox.pack(pady=5)

        tk.Button(popup, text="Save", command=save_price).pack(pady=5)

    def save_to_json(self):
        """Save summary to JSON."""
        self.df.to_json(self.json_file, orient="records", indent=4)

    def load_from_json(self):
        """Load summary from JSON."""
        if os.path.exists(self.json_file):
            self.df = pd.read_json(self.json_file, dtype={'Product': str})
        else:
            # Create an empty DataFrame and save it to JSON
            self.df = pd.DataFrame(
                columns=["Product", "Long Description", "Brand", "Pack Size", "Cost Price", "Qty Free", "Bin Number"])
            self.save_to_json()  # Save the empty DataFrame to initialize the file
        self.refresh_treeview()


if __name__ == "__main__":
    root = tk.Tk()
    ico_path = get_resource_path("assets/ibco_logo.ico")
    root.iconbitmap(ico_path)
    # Maximize the window
    root.state('zoomed')
    app = ProductApp(root)
    root.mainloop()

