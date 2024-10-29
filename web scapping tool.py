import os
import sys
import subprocess
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import messagebox

# Auto-install function
def install_and_import(package):
    try:
        __import__(package)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])

# Install required libraries
install_and_import("requests")
install_and_import("bs4")
install_and_import("pandas")
install_and_import("openpyxl")
install_and_import("python-docx")

import requests
from bs4 import BeautifulSoup
import pandas as pd
from docx import Document

def create_directory(custom_path):
    try:
        main_folder = Path(custom_path) / "web scrapped files"
        main_folder.mkdir(parents=True, exist_ok=True)

        folder_name = datetime.now().strftime("%Y%m%d_%H%M%S")
        folder_path = main_folder / folder_name
        folder_path.mkdir()
        return folder_path
    except FileNotFoundError:
        messagebox.showerror("Error", "Could not find or create the specified path. Please check the folder path and try again.")
        return None

def fetch_content(url):
    response = requests.get(url)
    response.raise_for_status()
    return response.content

def parse_and_save_data(content, folder_path):
    soup = BeautifulSoup(content, 'html.parser')
    
    tables = soup.find_all('table')
    if tables:
        data_frames = pd.read_html(str(tables))
        for i, df in enumerate(data_frames):
            excel_path = folder_path / f"table_data_{i+1}.xlsx"
            df.to_excel(excel_path, index=False)
        messagebox.showinfo("Success", "Data saved in Excel files.")
    else:
        paragraphs = soup.find_all('p')
        if paragraphs:
            doc = Document()
            for para in paragraphs:
                doc.add_paragraph(para.get_text())
            word_path = folder_path / "text_data.docx"
            doc.save(word_path)
            messagebox.showinfo("Success", "Data saved in a Word document.")
        else:
            messagebox.showinfo("Info", "No suitable data found on the page.")

def run_scraping(url, custom_path):
    # Strip quotes and whitespace from the custom path
    custom_path = custom_path.strip().strip('"').strip("'")
    folder_path = create_directory(custom_path)
    if folder_path:
        try:
            content = fetch_content(url)
            parse_and_save_data(content, folder_path)
        except requests.RequestException as e:
            messagebox.showerror("Error", f"Failed to retrieve data from the URL: {e}")

def start_scraping():
    url = url_entry.get().strip()  # Strip whitespace from URL
    custom_path = path_entry.get()  # Get the path input
    run_scraping(url, custom_path)

# Create the GUI window
root = tk.Tk()
root.title("Web Scraping Tool")

# Create and place the URL input label and entry
tk.Label(root, text="Enter the URL to scrape data from:").pack(pady=10)
url_entry = tk.Entry(root, width=50)
url_entry.pack(pady=5)

# Create and place the storage path label and entry
tk.Label(root, text="Enter the path for storing the file after scraping:").pack(pady=10)
path_entry = tk.Entry(root, width=50)
path_entry.pack(pady=5)

# Create and place the Start button
start_button = tk.Button(root, text="Start Scraping", command=start_scraping)
start_button.pack(pady=20)

# Start the GUI loop
root.mainloop()


