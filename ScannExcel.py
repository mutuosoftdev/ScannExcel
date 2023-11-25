import tkinter as tk
from tkinter import filedialog, ttk
import pandas as pd
import os






def open_file():
    file_path = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx *.xls')])
    if file_path:
        file_path_var.set(file_path)  # Actualizar la variable StringVar
        load_sheets(file_path)
        save_button.config(state='normal')  # Habilitar el botón "Guardar como TXT"

def load_sheets(file_path):
    global xls
    xls = pd.ExcelFile(file_path)
    sheet_names = xls.sheet_names
    sheets_combobox['values'] = sheet_names

def on_sheet_selected(event):
    # Leer la hoja seleccionada y contar el número de registros
    sheet_name = sheets_combobox.get()
    df = pd.read_excel(xls, sheet_name=sheet_name)
    num_records = len(df)
    records_var.set("REGISTROS ENCONTRADOS EN LA HOJA: " + str(num_records))  # Actualizar la variable StringVar
    

def save_txt():
    sheet_name = sheets_combobox.get()
    try:
        df = pd.read_excel(xls, sheet_name=sheet_name, usecols=[2, 8, 10, 12])
    except ValueError:
        tk.messagebox.showerror("Error", "Verifique si la Hoja de Calculo tiene las Columnas Necesarias (Er013Col)")
        return
    if df.isnull().values.any():  # Verificar si hay celdas vacías
        tk.messagebox.showerror("Error", "ErrCCel03: Por favor revise el documento e intente nuevamente // Datos faltantes || Columnas || Celdas")
        save_button.config(state='disabled')  # Desactivar el botón "Guardar como TXT"
        return
    df = df.apply(pd.to_numeric, errors='coerce').dropna()
    df = df.astype('int64')  # Convertir los datos a enteros de 64 bits
    df.insert(0, 'FC', 'FC')  # Insertar la columna 'FC' al principio del DataFrame
    df.insert(3, 'Nueva columna', '1')  # Insertar la nueva columna después de la segunda columna
    fifth_col = df.pop(df.columns[4])  # Remover la quinta columna
    df.insert(df.shape[1], fifth_col.name, fifth_col)  # Insertar la quinta columna al final
    df[df.columns[4]] = df[df.columns[4]].astype(str) + '00'  # Agregar '00' al final de la quinta columna
    df.insert(df.shape[1] - 1, 'DB', 'DB')  # Insertar la columna 'DB' en la penúltima posición
    df.columns = ['FC', 'COLUMNA "C"', 'COLUMNA "I"', 'Nueva columna', 'COLUMNA "M"', 'DB', 'COLUMNA "K"']
    file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
    if file_path:
        df.to_csv(file_path, sep='|', index=False, header=False)  # Guardar el archivo sin la fila de encabezado
        os.startfile(file_path)  # Abrir el archivo de texto

root = tk.Tk()
root.title("ScannExcel V3.1")
root.geometry('555x160')
root.resizable(False, False)
root.iconbitmap('./logo.ico')


# Centrar la ventana en la pantalla
window_width = root.winfo_reqwidth()
window_height = root.winfo_reqheight()
position_right = int(root.winfo_screenwidth()/2 - window_width/2)
position_down = int(root.winfo_screenheight()/2 - window_height/2)
root.geometry("+{}+{}".format(position_right, position_down))

# Crear un label para el título
title_label = tk.Label(root, text="ScannExcel", font=("Helvetica", 14, 'bold'))
title_label.pack()

open_button = tk.Button(root, text="Buscar archivos", command=open_file)
open_button.place(x=25, y=40, height=65)  # Posicionar el botón a 15px del margen izquierdo y a 10px del margen superior

# Crear una variable StringVar para la ruta del archivo
file_path_var = tk.StringVar()

# Crear un campo de entrada para la ruta del archivo
file_path_entry = tk.Entry(root, textvariable=file_path_var)
file_path_entry.place(x=140, y=45, width=250, height=25)  # Posicionar la ruta del archivo a 25px del margen izquierdo y a 40px del margen superior
file_path_entry.config(state='readonly')  # Hacer que el campo de entrada sea de solo lectura

sheets_combobox = ttk.Combobox(root)
sheets_combobox.place(x=420, y=47, width=110, height=25)  # Posicionar el combobox a 250px del margen izquierdo y a 40px del margen superior
sheets_combobox.bind("<<ComboboxSelected>>", on_sheet_selected)  # Agregar un controlador de eventos al combobox

save_button = tk.Button(root, text="Guardar como TXT", command=save_txt, state='disabled')  # Desactivar inicialmente el botón "Guardar como TXT"
save_button.pack()
save_button.place(x=420, y=80)


# Crear una variable StringVar para los registros encontrados
records_var = tk.StringVar()

# Crear un campo de entrada para los registros encontrados
records_entry = tk.Entry(root, textvariable=records_var)
records_entry.place(x=140, y=80, width=250, height=25)  # Posicionar el campo de entrada a 130px del margen izquierdo, a 65px del margen superior, con un ancho de 250px y una altura de 25px
records_entry.config(state='readonly')  # Hacer que el campo de entrada sea de solo lectura

copyright_label = tk.Label(root, text="© 2023 Catalitico || Dev. HCastro", anchor='e')
copyright_label.place(relx=1.0, rely=1.0, anchor='se')


root.mainloop()
