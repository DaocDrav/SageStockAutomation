import os
import sys

def get_resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        # base_path = root folder (one level up from current script folder)
        base_path = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    return os.path.join(base_path, relative_path)

def Read_sql_from_file(raw_file_path):
    # Ensure the file path works both during development and as an .exe
    absolute_path = get_resource_path(raw_file_path)
    with open(absolute_path, 'r') as file:
        sql_query = file.read()
    return sql_query
