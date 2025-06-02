from sqlalchemy import create_engine, text
import json
import os
import sys

def get_resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        # base_path = root folder (one level up from current script folder)
        base_path = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    return os.path.join(base_path, relative_path)

# Function to read the SQL file
def read_sql_from_file(file_path):
    with open(file_path, 'r') as file:
        return file.read()


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
    # Read the SQL query from the file
    sql_query = read_sql_from_file(get_resource_path("SQL/to_un_obsolete.sql"))


    # Execute the SQL query using the text() function
    conn.execute(text(sql_query))
    conn.commit()

print("SQL query executed successfully.")

# Close the connection
conn.close()