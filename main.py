import sys
import pandas as pd
from sqlalchemy import create_engine, types
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import json
import openpyxl
import psycopg2

def infer_sqlalchemy_types(df):
    """Infer SQLAlchemy column types based on DataFrame column data types."""
    dtype_map = {}
    for column in df.columns:
        if pd.api.types.is_integer_dtype(df[column]):
            dtype_map[column] = types.Integer
        elif pd.api.types.is_float_dtype(df[column]):
            dtype_map[column] = types.Float
        elif pd.api.types.is_bool_dtype(df[column]):
            dtype_map[column] = types.Boolean
        elif pd.api.types.is_datetime64_any_dtype(df[column]):
            dtype_map[column] = types.DateTime
        elif pd.api.types.is_string_dtype(df[column]):
            max_length = df[column].str.len().max()
            if pd.isna(max_length) or max_length > 255:
                dtype_map[column] = types.Text()
            else:
                dtype_map[column] = types.String(length=max(255, int(max_length)))
        else:
            dtype_map[column] = types.String()
    return dtype_map

def read_file(file_path):
    """Read CSV or Excel file into a DataFrame."""
    file_extension = os.path.splitext(file_path)[1].lower()
    
    if file_extension == '.csv':
        return pd.read_csv(file_path)
    elif file_extension in ['.xls', '.xlsx']:
        return pd.read_excel(file_path)
    else:
        raise ValueError("Unsupported file format. Please provide a .csv or .xls/.xlsx file.")

def import_data_to_db(username, password, db_name, db_url, file_path, table_name,server_type):
    if server_type.get() == "PostgreSQL":
        db_url_full = f'postgresql+psycopg2://{username}:{password}@{db_url}/{db_name}'
    elif server_type.get() == "MySQL":
        db_url_full = f'mysql+pymysql://{username}:{password}@{db_url}/{db_name}'
    engine = create_engine(db_url_full)
    df = read_file(file_path)
    dtype_map = infer_sqlalchemy_types(df)
    
    try:
        df.to_sql(
            table_name,
            con=engine,
            if_exists='replace',
            index=False,
            dtype=dtype_map
        )
        messagebox.showinfo("Success", f"Data from {file_path} imported successfully into table '{table_name}'!")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to import data: {str(e)}")

def browse_file(entry_file, entry_table):
    file_path = filedialog.askopenfilename(filetypes=[("CSV and Excel files", "*.csv;*.xls;*.xlsx")])
    if file_path:
        entry_file.delete(0, tk.END)
        entry_file.insert(0, file_path)
        table_name = os.path.splitext(os.path.basename(file_path))[0]
        entry_table.delete(0, tk.END)
        entry_table.insert(0, table_name)

def save_credentials(username, password, db_name, db_url):
    credentials = {
        "username": username,
        "password": password,
        "db_name": db_name,
        "db_url": db_url
    }
    with open("db_credentials.json", "w") as f:
        json.dump(credentials, f)

def load_credentials():
    if os.path.exists("db_credentials.json"):
        with open("db_credentials.json", "r") as f:
            return json.load(f)
    return None

def create_gui():
    root = tk.Tk()
    root.title("Excel to DB Data Import Tool")
    root.geometry("630x500")
    root.configure(bg="#f0f0f0")
    root.eval('tk::PlaceWindow . center')

    style = ttk.Style()
    style.configure("TLabel", font=("Helvetica", 12), background="#f0f0f0")
    style.configure("TEntry", font=("Helvetica", 12))
    style.configure("TButton", font=("Helvetica", 12), padding=5)

    saved_credentials = load_credentials()

    ttk.Label(root, text="Username:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
    entry_username = ttk.Entry(root)
    entry_username.grid(row=0, column=1, padx=10, pady=10)
    entry_username.insert(0, saved_credentials["username"] if saved_credentials else "postgres")

    ttk.Label(root, text="Password:").grid(row=1, column=0, padx=10, pady=10, sticky="w")
    entry_password = ttk.Entry(root, show="*")
    entry_password.grid(row=1, column=1, padx=10, pady=10)
    entry_password.insert(0, saved_credentials["password"] if saved_credentials else "")

    ttk.Label(root, text="Database Name:").grid(row=2, column=0, padx=10, pady=10, sticky="w")
    entry_dbname = ttk.Entry(root)
    entry_dbname.grid(row=2, column=1, padx=10, pady=10)
    entry_dbname.insert(0, saved_credentials["db_name"] if saved_credentials else "")

    ttk.Label(root, text="Database URL (host:port):").grid(row=3, column=0, padx=10, pady=10, sticky="w")
    entry_dburl = ttk.Entry(root)
    entry_dburl.grid(row=3, column=1, padx=10, pady=10)
    entry_dburl.insert(0, saved_credentials["db_url"] if saved_credentials else "localhost:5432")
    
    server_type = tk.StringVar(value="MySQL")
    ttk.Label(root, text="Database Server:").grid(row=4, column=0, padx=10, pady=10, sticky="w")
    ttk.Radiobutton(root, text='MySQL', variable=server_type, value='MySQL').grid(row=4, column=1) 
    ttk.Radiobutton(root, text='PostgreSQL', variable=server_type, value='PostgreSQL').grid(row=4, column=2 ) 
    
    ttk.Label(root, text="File Path:").grid(row=5, column=0, padx=10, pady=10, sticky="w")
    entry_file = ttk.Entry(root, width=40)
    entry_file.grid(row=5, column=1, padx=10, pady=10)
    ttk.Button(root, text="Browse", command=lambda: browse_file(entry_file, entry_table)).grid(row=5, column=2, padx=10, pady=10)

    ttk.Label(root, text="Table Name:").grid(row=6, column=0, padx=10, pady=10, sticky="w")
    entry_table = ttk.Entry(root)
    entry_table.grid(row=6, column=1, padx=10, pady=10)

    ttk.Button(root, text="Import Data", command=lambda: [
        import_data_to_db(
            entry_username.get(),
            entry_password.get(),
            entry_dbname.get(),
            entry_dburl.get(),
            entry_file.get(),
            entry_table.get(),
            server_type
        ),
        save_credentials(
            entry_username.get(),
            entry_password.get(),
            entry_dbname.get(),
            entry_dburl.get()
        )
    ]).grid(row=7, column=0, columnspan=3, pady=20)
    ttk.Label(root, text="").grid(row=9, column=1, sticky="w")
    ttk.Label(root, text="Developed By - Md Abdullah Al Baki").grid(row=10, column=1, sticky="w")
    
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    icon_path = os.path.join(base_path, 'icon.ico')

    root.iconbitmap(icon_path)
    root.mainloop()


if __name__ == "__main__":
    create_gui()
