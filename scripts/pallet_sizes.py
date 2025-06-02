from sqlalchemy import create_engine
from read_sql import Read_sql_from_file
import pandas as pd
from scipy.stats import mode
import json
import sys
import os

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

# Print or use the connection string
print(conn_str)


# Create an engine using SQLAlchemy
engine = create_engine(conn_str)

# Create a connection using the engine
with engine.connect() as conn:
    # Define the SQL query
    sql_query = Read_sql_from_file(get_resource_path("SQL/pallet_size.sql"))

    # Load data into the dataframe
    df = pd.read_sql(sql_query, conn)

# Group by product and calculate the mode (most common quantity)
common_quantities = df.groupby('product')['movement_quantity'].apply(lambda x: mode(x, keepdims=True).mode[0]).reset_index()
common_quantities.columns = ['product', 'common_quantity']

# Ensure only one row per product
unique_products = common_quantities.drop_duplicates(subset=['product']).reset_index(drop=True)

# Save or print the result
print(unique_products.head())  # View the result

# Close the connection
conn.close()
