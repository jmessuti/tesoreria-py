import streamlit as st
import pandas as pd
import pdfplumber
import re
from datetime import datetime
import io

# --- CONFIGURACIÓN DE PARÁMETROS 2025 ---
JORNAL_DIARIO = 111502
SALARIO_MIN_IRE = 2798309
UMBRAL_IVA_10_JORNALES = JORNAL_DIARIO * 10

st.set_page_config(page_title="Tesorería PY - Retenciones", layout="wide")

st.title("🏦 Gestión de Retenciones - Tesorería")

# Función para convertir texto de guaraníes a número (básico para validación)
def letras_a_numero(texto):
    texto = texto.upper()
    # Buscamos el patrón numérico que suele acompañar al texto entre paréntesis o cerca
    # Ej: "TOTAL A PAGAR: GUARANIES DOCE MILLONES ... (12.205.479)"
    match = re.search(r'([\d\.]+)', texto)
    if match:
        num_str = match.group(1).replace(".", "")
        if len(num_str) >= 4:
            return int(num_str)
    return 0

def extraer_datos_factura(file):
    with pdfplumber.open(file) as pdf:
        texto = ""
        for page in pdf.pages:
            texto += page.extract_text() + "\n"
    
    # 1. Buscar RUC (Emisor)
    ruc_match = re.findall(r'(\d+-\d+)', texto)
    ruc_emisor = ruc_match[-1] if ruc_match else "No detectado"
    
    # 2. Buscar Número de Factura
    nro_match = re.search(r'(\d{3}-\d{3}-\d{7})', texto)
    nro_factura = nro_match.group(1) if nro_match else "S/N"
    
    # 3. EXTRAER MONTO TOTAL (Estrategia por etiquetas de texto)
    total_factura = 0
    
    # Buscamos específicamente la sección de "TOTAL A PAGAR"
    lineas = texto.split('\n')
    for i, linea in enumerate(lineas):
        if "TOTAL A PAGAR" in linea.upper() or "TOTAL GS" in linea.upper():
            # Intentamos sacar el número de esa línea o la siguiente
            # Buscamos el formato 00.000.000
            monto_detectado = re.findall(r'(\d{1,3}(?:\.\d{3})+)', linea)
            if not monto_detectado and i+1 < len(lineas):
                monto_detectado = re.findall(r'(\d{1,3}(?:\.\d{3})+)', lineas[i+1])
            
            if monto_detectado:
                # Tomamos el último número encontrado en esa sección de totales
                total_factura = int(monto_detectado[-1].replace(".", ""))
                break

    # Si aún no detecta, buscamos el CDC o Timbrado para descartarlos y quedarnos con el resto
    if total_factura == 0:
        numeros_largos = re.findall(r'(\d{1,3}(?:\.\d{3})+)', texto)
        if numeros_largos:
            # El total suele ser el último o penúltimo número con puntos del documento
            total_factura = int(numeros_largos[-1].replace(".", ""))

    # 4. Cálculos
    iva_10 = round(total_factura / 11)
    gravada_10 = total_factura - iva_10

    nombre = "Proveedor Desconocido"
    if "ITAU" in texto.upper(): nombre = "BANCO ITAÚ PARAGUAY S.A."
    elif "CRYSALIS" in texto.upper(): nombre = "CRYSALIS S.R.L."
    else:
        lineas_n = [l.strip() for l in lineas if len(l.strip()) > 5]
        nombre = lineas_n[0][:40] if lineas_n else "Detectar manualmente"

    return {
        "Proveedor": nombre,
        "RUC": ruc_emisor,
        "Factura": nro_factura,
        "Total": total_factura,
        "Base_Imponible": gravada_10,
        "IVA_10": iva_10
    }

# --- INTERFAZ Y PROCESAMIENTO ---
with st.sidebar:
    st.header("Configuración")
    fecha_pago = st.date_input("Fecha de Pago", datetime.now())

archivos = st.file_uploader("Subir Facturas (PDF)", type="pdf", accept_multiple_files=True)

if archivos:
    datos_crudos = []
    for f in archivos:
        try:
            res = extraer_datos_factura(f)
            datos_crudos.append(res)
        except:
            st.error(f"Error procesando {f.name}")

    if datos_crudos:
        df_temp = pd.DataFrame(datos_crudos).sort_values(by="Factura")
        resultados = []
        acumulado_ruc = {}

        for _, fila in df_temp.iterrows():
            ruc = fila['RUC']
            bi = fila['Base_Imponible']
            if ruc not in acumulado_ruc: acumulado_ruc[ruc] = 0
            
            ret_iva = (fila['IVA_10'] * 0.30) if (acumulado_ruc[ruc] + bi) >= UMBRAL_IVA_10_JORNALES else 0
            ret_ire = (bi * 0.04) if bi >= SALARIO_MIN_IRE else 0
            acumulado_ruc[ruc] += bi
            
            resultados.append({
                "Proveedor": fila['Proveedor'],
                "RUC": ruc,
                "Factura": fila['Factura'],
                "Monto Factura": fila['Total'],
                "Base Imponible": bi,
                "Retención IVA (30%)": round(ret_iva),
                "Retención IRE (4%)": round(ret_ire),
                "Total Retenido": round(ret_iva + ret_ire),
                "Neto a Pagar": fila['Total'] - round(ret_iva + ret_ire)
            })

        df_final = pd.DataFrame(resultados)
        st.subheader("📋 Resultados")
        st.dataframe(df_final.style.format(precision=0, thousands="."))

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Retenciones')
        
        st.download_button("📥 Descargar Excel", buffer.getvalue(), f"Retenciones_{fecha_pago}.xlsx", "application/vnd.ms-excel")
