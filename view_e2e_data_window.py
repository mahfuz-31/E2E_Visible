import tkinter as tk
from tkinter import ttk
import pandas as pd
def View_E2E_Data(root):
  new_window = tk.Toplevel(root)
  new_window.title("Garments Status Data")
  new_window.geometry("1920x1080")

  # --- Your Data ---
  headers = [
    ('Buyer', 15),
    ('Order', 9),
    ('Style', 9),
    ('Grey Book. Qty', 18),
    ('Grey Recv. Qty', 18),
    ('Grey Bal. Qty', 18),
    ('Finish Book. Qty', 18),
    ('Finish Recv. Qty', 18),
    ('Finish Bal. Qty', 18),
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
    order_qty = row['Order Qty']
    for col_idx, (header, width) in enumerate(headers):
      if col_idx < 3:  # First three columns are left-aligned
        value = row.get(header, "")
        cell = tk.Label(scrollable_frame, text=str(value), font=("Calibri", 11), padx=2, pady=2, width=width, anchor="w")
        cell.grid(row=row_idx, column=col_idx, sticky="w")
      else: 
        value = row.get(header, "") # Other columns are right-aligned
        progress = ttk.Progressbar(scrollable_frame, length=150, mode='determinate')
        progress['maximum'] = order_qty if order_qty > 0 else 1
        progress['value'] = value if pd.notna(value) else 0
        progress.grid(row=row_idx, column=col_idx, sticky="w", padx=2, pady=2)
  