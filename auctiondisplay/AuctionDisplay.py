import tkinter as tk
from tkinter import font
import logging

import pandas as pd

__DATA_FILE = 'Buyer-List.xlsx'

data = {}

# Get our logging instance
log = logging.getLogger(__name__)

# Set up the Window interface
root = tk.Tk()
root.attributes('-fullscreen', True)

buyer_id_font_style = font.Font(size=150, weight='bold')
buyer_font_style = font.Font(size=100)

buyer_id_label = tk.Label(root, text="", font=buyer_id_font_style, wraplength=root.winfo_screenwidth(),
                          justify='center')
buyer_id_label.pack(side=tk.TOP, pady=50)

buyer_label = tk.Label(root, text="", font=buyer_font_style, wraplength=root.winfo_screenwidth(), justify='center')
buyer_label.pack(expand=True)


def read_excel_data(filename):
    global data
    data = {}
    df = pd.read_excel(filename)
    for index, row in df.iterrows():
        data[int(row['Identifier'])] = row['Name']  # Convert the Buyer ID to integer
    print(f"Number of records loaded: {len(data)}")
    return data  # Return the data dictionary and the count of records loaded


def update_display(value):
    global data  # Declare data as a global variable inside the function
    if value == '0':
        buyer_id_label.config(text="")
        buyer_label.config(text="")
    elif value.lower() == 'quit':
        root.destroy()
    elif value.lower() == 'read':
        data = read_excel_data(__DATA_FILE)  # Re-read the Excel file and update data
    else:
        try:
            buyer_id = int(value)  # Convert user input to integer
            buyer = data.get(buyer_id, "N/A")
            buyer_id_label.config(text=f"{buyer_id}")
            buyer_label.config(text=f"{buyer}")
            print(f"{buyer_id}: {buyer}")
        except ValueError:
            buyer_id_label.config(text="")
            buyer_label.config(text="")


def run():
    global data
    data = read_excel_data(__DATA_FILE)  # Initialize data with the initial read
    print("Enter a Buyer ID ('0' to clear, 'read' to re-read Excel, 'quit' to exit):")
    while True:
        root.update()
        user_input = input("> ")
        update_display(user_input.lower())

        if user_input.lower() == 'quit':
            break



