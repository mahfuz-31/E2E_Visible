import tkinter as tk
from tkinter import ttk
import pandas as pd

def View_E2E_Data(root):
  new_window = tk.Toplevel(root)
  new_window.title("Garments Status Data")
  new_window.geometry("1920x1080")

  # --- Horizontal Scrollbar Setup ---
  main_frame = tk.Frame(new_window)
  main_frame.pack(fill="both", expand=True)

  main_canvas = tk.Canvas(main_frame)
  h_scrollbar = tk.Scrollbar(main_frame, orient="horizontal", command=main_canvas.xview)
  
  main_canvas.configure(xscrollcommand=h_scrollbar.set)
  
  h_scrollbar.pack(side="bottom", fill="x")
  main_canvas.pack(side="left", fill="both", expand=True)

  wrapper_frame = tk.Frame(main_canvas)
  main_canvas.create_window((0, 0), window=wrapper_frame, anchor="nw")

  def on_frame_configure(event):
    main_canvas.configure(scrollregion=main_canvas.bbox("all"))
  wrapper_frame.bind("<Configure>", on_frame_configure)

  # --- Your Data ---
  headers = [
    ('Buyer', 12),
    ('Order', 9),
    ('Grey Recv. Qty', 18),
    ('Finish Recv. Qty', 18),
    ('Cutting Qty', 18),
    ('Input Qty', 18),
    ('Output Qty', 18),
    ('Poly Qty', 18),
    ('Shipped Qty', 20),
  ]

  for col_idx, (header, width) in enumerate(headers):
    headerL = tk.Label(
      wrapper_frame, text=header,
      font=("Calibri", 11, 'normal', 'underline'),
      fg="darkgreen",
      width=width,   # <-- custom width per column
      anchor="w"
    )
    headerL.grid(row=0, column=col_idx, sticky="w", padx=2, pady=2)
  
  data = pd.read_excel("E2E_output.xlsx")

  # Create a frame for the table and add a vertical scrollbar
  # Note: table_frame is now inside wrapper_frame
  table_frame = tk.Frame(wrapper_frame)
  table_frame.grid(row=1, column=0, columnspan=len(headers) + 1, sticky="nsew")

  # Existing vertical scroll logic
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
        if pd.notna(value) and grey_book > 0 and value < 0.3 * grey_book:
            progress.configure(style='red.Horizontal.TProgressbar')
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
