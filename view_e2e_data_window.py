import tkinter as tk
from tkinter import ttk
import pandas as pd
def View_E2E_Data(root):
  new_window = tk.Toplevel(root)
  new_window.title("Garments Status Data")
  new_window.geometry("1920x1080")

  # --- Your Data ---
  headers = [
    ('Buyer', 12),
    ('Order', 9),
    ('Grey Book. Qty', 15),
    ('Grey Recv. Qty', 15),
    ('Grey Bal. Qty', 15),
    ('Finish Book. Qty', 15),
    ('Finish Recv. Qty', 15),
    ('Finish Bal. Qty', 15),
    ('Order Qty', 15),
    ('Cutting Qty', 15),
    ('Input Qty', 15),
    ('Output Qty', 15),
    ('Poly Qty', 15),
    ('Shipped Qty', 15),

  ]

  for col_idx, (header, width) in enumerate(headers):
      headerL = tk.Label(
          new_window, text=header,
          font=("Calibri", 11, 'normal', 'underline'),
          fg="darkgreen",
          width=width,   # <-- custom width per column
          anchor="w"
      )
      headerL.grid(row=0, column=col_idx, sticky="w", padx=2, pady=2)
  data = pd.read_excel("E2E_output.xlsx")
  # Create a frame for the table and add a vertical scrollbar
  table_frame = tk.Frame(new_window)
  table_frame.grid(row=1, column=0, columnspan=len(headers) + 1, sticky="nsew")

  canvas = tk.Canvas(table_frame, height=690)
  scrollbar = tk.Scrollbar(table_frame, orient="vertical", command=canvas.yview)
  scrollable_frame = tk.Frame(canvas)

  scrollable_frame.bind(
    "<Configure>",
    lambda _: canvas.configure(
      scrollregion=canvas.bbox("all")
    )
  )

  canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
  canvas.configure(yscrollcommand=scrollbar.set)

  canvas.pack(side="left", fill="both", expand=True)
  scrollbar.pack(side="right", fill="y")

  # Populate the table with data
  for row_idx, row in data.iterrows():
    grey_book = row['Grey Book. Qty']
    finish_book = row['Finish Book. Qty']
    order_qty = row['Order Qty']

    for col_idx, (header, width) in enumerate(headers):
      if col_idx < 2:  # First three columns are left-aligned
        value = row.get(header, "")
        cell = tk.Label(scrollable_frame, text=str(value), font=("Calibri", 11), padx=2, pady=2, width=width, anchor="w")
        cell.grid(row=row_idx, column=col_idx, sticky="w")
      elif col_idx >= 2 and col_idx <= 4:
        value = row.get(header, "")
        progress = ttk.Progressbar(scrollable_frame, length=150, mode='determinate')
        progress['maximum'] = grey_book if grey_book > 0 else 1
        progress['value'] = value if pd.notna(value) else 0
        progress.grid(row=row_idx, column=col_idx, sticky="w", padx=2, pady=2)
      elif col_idx >= 5 and col_idx <= 7:
        value = row.get(header, "")
        progress = ttk.Progressbar(scrollable_frame, length=150, mode='determinate')
        progress['maximum'] = finish_book if finish_book > 0 else 1
        progress['value'] = value if pd.notna(value) else 0
        progress.grid(row=row_idx, column=col_idx, sticky="w", padx=2, pady=2)
      else:
        value = row.get(header, "")
        progress = ttk.Progressbar(scrollable_frame, length=150, mode='determinate')
        progress['maximum'] = order_qty if order_qty > 0 else 1
        progress['value'] = value if pd.notna(value) else 0
        progress.grid(row=row_idx, column=col_idx, sticky="w", padx=2, pady=2)
 
