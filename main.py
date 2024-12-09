import pandas as pd
from tkinter import Tk, Label, Button, filedialog, messagebox
from datetime import datetime, timedelta
import calendar
import os

# Get the past month name in Spanish
month_names = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
current_month = datetime.now().month
past_month = (current_month - 1) if current_month > 1 else 12
past_month_name = month_names[past_month - 1]

# Calculate the cutoff date (24th of the past month)
cutoff_date = datetime(datetime.now().year, past_month, 24)

# Initialize the Excel writer
excel_file_path = f"Complementos {past_month_name}.xlsx"
excel_writer = pd.ExcelWriter(excel_file_path, engine='xlsxwriter')

def load_excel_file(sheet_name, filter_date=False, skip_rows=False):
    try:
        file_path = filedialog.askopenfilename()
        if file_path.endswith(('.xlsx', '.xls')):
            data = pd.read_excel(file_path, skiprows=5 if skip_rows else 0)  # Conditionally skip the first 5 rows
        elif file_path.endswith('.csv'):
            data = pd.read_csv(file_path, skiprows=5 if skip_rows else 0)  # Conditionally skip the first 5 rows
        else:
            messagebox.showerror("Error", "¡Formato de archivo no soportado!")
            return

        if filter_date:
            if 'FECHA' in data.columns:
                # Convert 'FECHA' column to datetime, invalid parsing will be set as NaT
                data['FECHA'] = pd.to_datetime(data['FECHA'], format='%Y-%m-%d', errors='coerce')
                # Filter out rows with NaT in 'FECHA' column
                data = data.dropna(subset=['FECHA'])
                # Filter transactions based on the cutoff date
                data = data[data['FECHA'] > cutoff_date]
                # Format 'FECHA' column to only include the date
                data['FECHA'] = data['FECHA'].dt.strftime('%Y-%m-%d')
            else:
                messagebox.showerror("Error", "¡La columna 'FECHA' no se encuentra en el archivo!")
                return

        data.to_excel(excel_writer, sheet_name=sheet_name, index=False)
        messagebox.showinfo("Éxito", f"¡Archivo subido correctamente a la hoja {sheet_name}!")
        print(data.head())  # Display the first few rows of the file
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al cargar el archivo: {e}")

def create_gui():
    # Initialize the root window
    root = Tk()
    root.title("Cargador de Archivos Excel")
    root.geometry("400x200")

    # Define the sequence of uploads
    uploads = [
        ("Subir Movimientos de la Tarjeta Platino a partir del 14 de " + past_month_name, "Subir Archivo", f"Complementos Platino {past_month}"),
        ("Subir Movimientos de la Tarjeta Oro a partir del 14 de " + past_month_name, "Subir Archivo", f"Complementos Oro {past_month}"),
        ("Subir Movimientos de la cuenta de débito a partir del 24 de " + past_month_name, "Subir Archivo", f"Complementos Débito {past_month}")
    ]

    def update_gui(index):
        if index < len(uploads):
            label.config(text=uploads[index][0])
            button.config(text=uploads[index][1], command=lambda: handle_upload(index))
        else:
            excel_writer.close()
            messagebox.showinfo("Éxito", "¡Todos los archivos han sido subidos y guardados correctamente!")
            root.quit()

    def handle_upload(index):
        filter_date = (index == 2)  # Only filter the last upload
        skip_rows = (index == 2)  # Only skip rows for the last upload
        load_excel_file(uploads[index][2], filter_date, skip_rows)
        update_gui(index + 1)

    # Add a label and button to the window
    label = Label(root, text="")
    label.pack(pady=10)

    button = Button(root, text="")
    button.pack(pady=10)

    # Start the sequence
    update_gui(0)

    # Start the Tkinter main loop
    root.mainloop()

if __name__ == "__main__":
    create_gui()

