import os
import tkinter as tk
from tkinter import filedialog, messagebox
from docx2pdf import convert
import pythoncom
from tkinter import ttk
import os

# Si el archivo Python está en la misma carpeta
config_file_path = os.path.join(os.path.dirname(__file__), 'config', 'app_config.json')
print(config_file_path)

# Función para convertir el archivo .docx a PDF
def convert_docx_to_pdf(docx_file, pdf_path, progress_bar, progress_var, total_files, current_file):
    try:
        # Inicializar la librería COM
        pythoncom.CoInitialize()

        # Verificar si el archivo .docx existe y no está vacío
        if not os.path.exists(docx_file) or os.path.getsize(docx_file) == 0:
            messagebox.showerror("Error", f"El archivo {docx_file} está vacío o no se encuentra.")
            return

        # Verificar si la carpeta de salida existe
        if not os.path.exists(os.path.dirname(pdf_path)):
            os.makedirs(os.path.dirname(pdf_path))  # Crear la carpeta si no existe

        try:
            convert(docx_file, pdf_path)
            # Actualizar barra de progreso
            progress_var.set((current_file / total_files) * 100)
            progress_bar.update_idletasks()  # Actualiza la barra de progreso en la interfaz
        finally:
            # Finalizar la librería COM
            pythoncom.CoUninitialize()
    except Exception as e:
        messagebox.showerror("Error", f"Error al convertir {docx_file} a PDF: {e}")

# Función para seleccionar archivos .docx
def select_files():
    files = filedialog.askopenfilenames(filetypes=[("Archivos Word", "*.docx")])
    if files:
        file_listbox.delete(0, tk.END)  # Limpiar lista previa
        for file in files:
            file_listbox.insert(tk.END, file)

# Función para seleccionar la carpeta de salida
def select_output_folder():
    folder = filedialog.askdirectory()
    if folder:
        output_folder_entry.delete(0, tk.END)  # Limpiar campo
        output_folder_entry.insert(0, folder)

# Función para convertir los archivos seleccionados
def convert_files():
    files = file_listbox.get(0, tk.END)
    output_folder = output_folder_entry.get()

    if not files:
        messagebox.showwarning("Advertencia", "Por favor, selecciona al menos un archivo Word.")
        return

    if not output_folder:
        output_folder = os.getcwd()

    os.makedirs(output_folder, exist_ok=True)

    total_files = len(files)
    progress_var.set(0)  # Reiniciar barra de progreso

    # Mostrar la barra de progreso
    progress_bar.pack(pady=20)

    # Procesar cada archivo y convertir
    for idx, file in enumerate(files, start=1):
        # Definir la ruta de salida del PDF
        pdf_path = os.path.join(output_folder, f"{os.path.splitext(os.path.basename(file))[0]}.pdf")

        # Convertir .docx a PDF
        convert_docx_to_pdf(file, pdf_path, progress_bar, progress_var, total_files, idx)

    messagebox.showinfo("Éxito", f"¡Conversión completada! Los archivos PDF se han guardado en: {output_folder}")

# Crear la ventana principal
root = tk.Tk()
root.title("Conversor de Word a PDF")
root.geometry("700x750")  # Aumentar tamaño de la ventana

# Título
title_label = tk.Label(root, text="Conversor de Word a PDF", font=("Arial", 16, "bold"), fg="#003366", pady=20)
title_label.pack()

# Subtítulo
subtitle_label = tk.Label(root, text="Selecciona los archivos Word para convertirlos a PDF.", font=("Arial", 12))
subtitle_label.pack()

# Botón para seleccionar archivos
select_files_button = tk.Button(root, text="Seleccionar Archivos Word", command=select_files, font=("Arial", 12), bg="#007BFF", fg="white", relief="flat")
select_files_button.pack(pady=15, ipadx=10, ipady=5)

# Lista de archivos seleccionados
file_listbox = tk.Listbox(root, width=50, height=6, selectmode=tk.MULTIPLE, font=("Arial", 10), bd=1, relief="solid")
file_listbox.pack(pady=10)

# Campo de texto para la ruta de la carpeta de salida
output_folder_label = tk.Label(root, text="Carpeta de salida (opcional):", font=("Arial", 10))
output_folder_label.pack()

output_folder_entry = tk.Entry(root, width=40, font=("Arial", 12), relief="solid")
output_folder_entry.pack(pady=5)

# Botón para seleccionar la carpeta de salida
select_folder_button = tk.Button(root, text="Seleccionar Carpeta", command=select_output_folder, font=("Arial", 12), bg="#007BFF", fg="white", relief="flat")
select_folder_button.pack(pady=10, ipadx=10, ipady=5)

# Barra de progreso
progress_var = tk.DoubleVar()
progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100, length=400)

# Botón para convertir archivos
convert_button = tk.Button(root, text="Convertir a PDF", command=convert_files, font=("Arial", 12), bg="#28A745", fg="white", relief="flat")
convert_button.pack(pady=20, ipadx=10, ipady=5)

# Botón para cerrar la aplicación
close_button = tk.Button(root, text="Cerrar", command=root.quit, font=("Arial", 12), bg="#FF5733", fg="white", relief="flat")
close_button.pack(pady=15, ipadx=10, ipady=5)

# Ejecutar la ventana
root.mainloop()
