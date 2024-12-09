import pandas as pd
from tkinter import Tk, Label, Button, messagebox
from tkinter.filedialog import askopenfilename

def load_excel_file():
    # Open file dialog to select the Excel file
    file_path = askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])

    if not file_path:
        messagebox.showinfo("File Selection", "No file selected!")
        return

    try:
        # Load the Excel file using pandas
        data = pd.read_excel(file_path)
        messagebox.showinfo("Success", "File loaded successfully!")
        print(data.head())  # Display the first few rows of the file
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while loading the file: {e}")

def create_gui():
    # Initialize the root window
    root = Tk()
    root.title("Excel File Loader")
    root.geometry("300x150")

    # Add a label and button to the window
    label = Label(root, text="Upload an Excel file")
    label.pack(pady=10)

    button = Button(root, text="Upload File", command=load_excel_file)
    button.pack(pady=10)

    # Start the Tkinter main loop
    root.mainloop()

if __name__ == "__main__":
    create_gui()

