from tkinter import *
from tkinter.ttk import Progressbar
import spc_search_color_wise
from tkinter import messagebox
import threading
import garments_sp_search
import view_e2e_data_window

root = Tk()
root.title('Mahfuz\'s End to End Followup')
w = Label(root, text='Mahfuz\'s End to End Followup', fg='#054270', font=("Cooper Black", 18, ))
w.pack(pady=5)

frs_label = Label(root, text="Enter your list of FRS No:", fg='#2B6936', font=("Calibri", 11, 'italic', 'bold'))
frs_label.place(x=35, y=120)

# Create a frame to hold text + scrollbar together
text_frame = Frame(root)
text_frame.place(x=35, y=145)

# Text widget
text_area = Text(text_frame, wrap=WORD, width=20, height=23, font=("Arial", 11))
text_area.pack(side=LEFT, fill=BOTH, expand=True)

# Add a Scrollbar
scrollbar = Scrollbar(text_frame, command=text_area.yview)
scrollbar.pack(side=RIGHT, fill=Y)
text_area.config(yscrollcommand=scrollbar.set)



# Progress bar
progress = Progressbar(root, orient=HORIZONTAL, length=350, mode='determinate')
progress.pack(pady=10)
progress.pack_configure(anchor='center')


def download_thread():
    """Function to run in separate thread"""
    try:
        # Update UI to show processing state
        root.after(0, lambda: download_btn.config(state='disabled', text='Downloading...'))
        
        # Get the orders from text area
        orders = text_area.get("1.0", END).strip().split('\n')
        
        # Run the main function
        spc_search_color_wise.spc_search_color_wise(root, progress, orders)
        progress['value'] = 0
        root.after(0, lambda: messagebox.showinfo("Completed", "Color wise fabric data download completed!"))
        garments_sp_search.garments_sp_search(root, progress, orders)
        root.after(0, lambda: messagebox.showinfo("Completed", "Garments data download completed!"))
        
        root.after(0, lambda: text_area.delete("1.0", END))
        
    except Exception as e:
        # Error handling - update UI on main thread
        root.after(0, lambda: messagebox.showerror("Error", f"An error occurred: {str(e)}"))
    
    finally:
        # Re-enable button and reset text
        root.after(0, lambda: download_btn.config(state='normal', text='Download Data'))

def on_download():
    """Start download in separate thread"""
    if text_area.get("1.0", END).strip():
        thread = threading.Thread(target=download_thread, daemon=True)
        thread.start()
    else:
        messagebox.showwarning("Warning", "Please enter some FRS numbers!")
        # only hide if warning

download_btn = Button(
    root, 
    font=('Calibri', 9, 'bold'),
    text='Download Data', 
    bg='#2E763A',     
    fg='white',
    activebackground='lightgreen', 
    activeforeground='white',
    command=on_download
)
download_btn.pack(side=LEFT, anchor='sw', pady=20, padx=20)

view_data_btn = Button(root, text='View Fabric Data', font=('Calibri', 9, 'bold'), bg='#3B974A', fg='white',
                       activebackground='lightgreen', activeforeground='white', command=lambda: view_e2e_data_window.View_E2E_Data(root))
view_data_btn.pack(side=LEFT, anchor='se', pady=20, padx=0)



root.geometry('410x600')
root.mainloop()