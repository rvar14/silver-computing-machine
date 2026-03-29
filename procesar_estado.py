import zipfile
import PyPDF2
import io
import re
import pandas as pd

def extraer_movimientos_banbajio(zip_path, output_excel):
    # Esta expresión regular busca las filas de transacciones
    # Ejemplo: "1 ENE 8029374 DEPOSITO NEGOCIOS... $ 5,049.00 $ 197,098.80"
    pattern = re.compile(r'^(\d{1,2}\s+[A-Z]{3})\s+(.*?)\s+\$\s*([\d,]+\.\d{2})\s+\$\s*([\d,]+\.\d{2})$')

    depositos = []
    retiros = []

    # Abrimos el archivo ZIP
    with zipfile.ZipFile(zip_path, 'r') as z:
        for name in z.namelist():
            # Buscamos el archivo PDF dentro del ZIP
            if name.endswith('.pdf'):
                with z.open(name) as f:
                    pdf_file = io.BytesIO(f.read())
                    reader = PyPDF2.PdfReader(pdf_file)
                    
                    prev_saldo = 0.0
                    
                    # Leemos cada página del PDF
                    for i in range(len(reader.pages)):
                        text = reader.pages[i].extract_text()
                        lines = text.split('\n')
                        
                        for line in lines:
                            # Intentamos capturar el saldo inicial por si acaso
                            if "SALDO INICIAL" in line:
                                m = re.search(r'\$\s*([\d,]+\.\d{2})', line)
                                if m:
                                    prev_saldo = float(m.group(1).replace(',', ''))
                                    
                            m = pattern.search(line)
                            if m:
                                fecha = m.group(1)
                                desc = m.group(2).strip()
                                monto = float(m.group(3).replace(',', ''))
                                saldo = float(m.group(4).replace(',', ''))
                                
                                # Calculamos la diferencia de saldo para saber si entró o salió dinero
                                diff = saldo - prev_saldo
                                
                                row = {
                                    'Fecha': fecha,
                                    'Descripción': desc,
                                    'Monto ($)': monto,
                                    'Saldo tras operación ($)': saldo
                                }
                                
                                # Si la diferencia positiva coincide con el monto, es Depósito
                                if abs(diff - monto) < 0.05:
                                    depositos.append(row)
                                # Si la diferencia negativa coincide con el monto, es Retiro
                                elif abs(diff + monto) < 0.05:
                                    retiros.append(row)
                                else:
                                    # Alternativa por si se pierde un renglón en el salto de página
                                    if "DEPOSITO" in desc or "ABONO" in desc:
                                        depositos.append(row)
                                    else:
                                        retiros.append(row)
                                    
                                prev_saldo = saldo

    # Creamos DataFrames de Pandas
    df_depositos = pd.DataFrame(depositos)
    df_retiros = pd.DataFrame(retiros)

    # Exportamos a Excel con dos hojas
    with pd.ExcelWriter(output_excel) as writer:
        df_depositos.to_excel(writer, sheet_name="Depósitos", index=False)
        df_retiros.to_excel(writer, sheet_name="Retiros", index=False)

    print(f"¡Listo! Se ha creado el archivo: {output_excel}")
    print(f"Total encontrados -> Depósitos: {len(df_depositos)} | Retiros: {len(df_retiros)}")

# Ejecución del script
archivo_zip_banco = "Estados_20260321195534.zip" # Cambia esto por el nombre de tu ZIP
archivo_excel_salida = "Reporte_Movimientos.xlsx"

extraer_movimientos_banbajio(archivo_zip_banco, archivo_excel_salida)