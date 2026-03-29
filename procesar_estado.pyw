import tkinter as tk
from tkinter import filedialog, messagebox
import zipfile
import PyPDF2
import io
import re
import pandas as pd
import os

def extraer_movimientos_banbajio(zip_path):
    # Generar el nombre del archivo de salida basado en el ZIP
    directorio_base = os.path.dirname(zip_path)
    nombre_base = os.path.splitext(os.path.basename(zip_path))[0]
    output_excel = os.path.join(directorio_base, f"{nombre_base}.xlsx")

    pattern = re.compile(r'^(\d{1,2}\s+[A-Z]{3})\s+(.*?)\s+\$\s*([\d,]+\.\d{2})\s+\$\s*([\d,]+\.\d{2})$')
    depositos, retiros = [], []

    with zipfile.ZipFile(zip_path, 'r') as z:
        # Buscar el PDF dentro del ZIP
        pdf_filename = next((name for name in z.namelist() if name.endswith('.pdf')), None)
        
        if not pdf_filename:
            raise ValueError("No se encontró ningún archivo PDF dentro del ZIP.")

        with z.open(pdf_filename) as f:
            pdf_file = io.BytesIO(f.read())
            reader = PyPDF2.PdfReader(pdf_file)
            
            prev_saldo = 0.0
            
            for i in range(len(reader.pages)):
                text = reader.pages[i].extract_text()
                for line in text.split('\n'):
                    if "SALDO INICIAL" in line:
                        m = re.search(r'\$\s*([\d,]+\.\d{2})', line)
                        if m: prev_saldo = float(m.group(1).replace(',', ''))
                            
                    m = pattern.search(line)
                    if m:
                        fecha, desc = m.group(1), m.group(2).strip()
                        monto = float(m.group(3).replace(',', ''))
                        saldo = float(m.group(4).replace(',', ''))
                        diff = saldo - prev_saldo
                        
                        row = {'Fecha': fecha, 'Descripción': desc, 'Monto ($)': monto, 'Saldo ($)': saldo}
                        
                        if abs(diff - monto) < 0.05: depositos.append(row)
                        elif abs(diff + monto) < 0.05: retiros.append(row)
                        else:
                            if "DEPOSITO" in desc or "ABONO" in desc: depositos.append(row)
                            else: retiros.append(row)
                            
                        prev_saldo = saldo

    df_depositos = pd.DataFrame(depositos)
    df_retiros = pd.DataFrame(retiros)

    with pd.ExcelWriter(output_excel) as writer:
        df_depositos.to_excel(writer, sheet_name="Depósitos", index=False)
        df_retiros.to_excel(writer, sheet_name="Retiros", index=False)

    return output_excel, len(df_depositos), len(df_retiros)

# --- CONFIGURACIÓN DE LA INTERFAZ GRÁFICA (GUI) ---

def buscar_archivo():
    ruta = filedialog.askopenfilename(
        title="Selecciona el estado de cuenta",
        filetypes=[("Archivos ZIP", "*.zip")]
    )
    if ruta:
        entrada_ruta.delete(0, tk.END)
        entrada_ruta.insert(0, ruta)

def iniciar_proceso(event=None):
    # Limpiar la ruta por si el usuario arrastró el archivo y se pegó con comillas
    ruta_zip = entrada_ruta.get().strip().strip('"').strip("'")
    
    if not ruta_zip or not ruta_zip.lower().endswith('.zip'):
        messagebox.showwarning("Atención", "Por favor, ingresa un archivo .zip válido.")
        return
    
    if not os.path.exists(ruta_zip):
        messagebox.showerror("Error", "El archivo especificado no existe.")
        return

    try:
        boton_procesar.config(text="Procesando...", state=tk.DISABLED)
        root.update()
        
        archivo_salida, cant_dep, cant_ret = extraer_movimientos_banbajio(ruta_zip)
        
        mensaje = f"¡Proceso completado con éxito!\n\nSe guardó como:\n{os.path.basename(archivo_salida)}\n\nDepósitos: {cant_dep}\nRetiros: {cant_ret}"
        messagebox.showinfo("¡Éxito!", mensaje)
        
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un problema al procesar:\n{str(e)}")
    finally:
        boton_procesar.config(text="Procesar Excel", state=tk.NORMAL)

# Crear la ventana principal
root = tk.Tk()
root.title("Procesador de BanBajío")
root.geometry("550x180")
root.resizable(False, False)

# Componentes visuales
tk.Label(root, text="Arrastra la ruta del archivo .zip aquí o búscalo:", font=("Arial", 10)).pack(pady=(15, 5))

frame_input = tk.Frame(root)
frame_input.pack(pady=5, padx=20, fill="x")

entrada_ruta = tk.Entry(frame_input, font=("Arial", 10))
entrada_ruta.pack(side=tk.LEFT, fill="x", expand=True, padx=(0, 10), ipady=4)

boton_buscar = tk.Button(frame_input, text="Buscar...", command=buscar_archivo)
boton_buscar.pack(side=tk.RIGHT)

boton_procesar = tk.Button(root, text="Procesar Excel", font=("Arial", 10, "bold"), bg="#4CAF50", fg="white", command=iniciar_proceso)
boton_procesar.pack(pady=15, ipady=5, ipadx=10)

# Permitir que la tecla "Enter" ejecute el proceso
root.bind('<Return>', iniciar_proceso)

# Iniciar la aplicación
root.mainloop()