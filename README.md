# SageStockAutomation

A Python-based automation tool designed to streamline stock control, goods inward processing & goods transfers for Sage ERP 1000. This project features a Tkinter GUI interface, SQL Server database integration, and utilities to simplify stock management workflows.

## Features

- Interactive Tkinter GUI for product and supplier filtering
- SQL Server connectivity via pyodbc for dynamic data querying
- Modular scripts including:
  - `count.py` — Main stock counting interface
  - `inward_sum.py` — Summary and reporting of goods inward data
  - Other utility scripts for data processing and automation
- JSON-based persistence for session data
- Easy extendibility for future enhancements
- Count stock, update purchase orders, and paste sites/bins, quantities & best befores into Sage Line 1000 ERP
- Automate Goods Transfers into Sage Line 1000 ERP using specified minimum reorder levels into picking bays.

## Video Demonstration
[![Watch the demo](https://img.youtube.com/vi/jaglMFYmt6U/0.jpg)](https://youtu.be/jaglMFYmt6U)

Click the image above to watch a short video demo of the application in action.

## Requirements


- Python 3.8 or higher
- Windows OS (tested on Windows 10/11)
- Microsoft SQL Server database with proper access credentials
- Required Python packages:
  - `pyodbc`
  - `pandas`
  - `tkinter` (usually included with Python)
  - `openpyxl` (for Excel integration)
  - `json` (standard library)


## Setup

1. **Clone the repository:**


   git clone https://github.com/DaocDrav/SageStockAutomation.git
   cd SageStockAutomation

2. **Create and activate a virtual environment (recommended):**

python -m venv venv
# On Windows:
venv\Scripts\activate
# On macOS/Linux:
source venv/bin/activate

3. **Install required packages**

   pip install pyodbc pandas numpy openpyxl pyautogui scipy pillow pyscreeze opencv-python

4. **Configure db_config.JSON**

   Enter SQL database details into assets/db_config.JSON to connect to your database

5. **Usage**
   Launch the main application:-
   python main.py
   
   

## License

MIT License 


