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
st.markdown(f"**Configuración 2025:** Jornal Gs. {JORNAL_DIARIO:,} | Salario Mínimo Gs. {SALARIO_MIN_IRE:,}".replace(",", "."))

# --- INTERFAZ LATERAL ---
with st.sidebar:
    st.header("Configuración de Pago")
    fecha_pago = st.date_input("Fecha de Pago", datetime.now())
    st.info("La app ordenará las facturas por número y calculará la retención acumulada del mes.")

archivos = st.file_uploader("Subir Facturas (PDF)", type="pdf", accept_multiple_files=True)

# --- FUNCIÓN DE EXTRACCIÓN MEJORADA ---
def extraer_datos_factura(file):
    with pdfplumber.open(file) as pdf:
        texto = ""
        for page in pdf.pages:
            texto += page.extract_text() + "\n"
    
    # Limpieza de texto para facilitar búsqueda
    texto_limpio = texto.replace(".", "") # Quitamos puntos de miles para capturar números limpios
    
    # 1. Buscar RUC (Emisor)
    # Buscamos el RUC que suele estar cerca de 'TIMBRADO' o al inicio/final
    ruc_match = re.findall(r'RUC[:\s]*(\d+-\d+)', texto)
    # En facturas paraguayas el segundo RUC suele ser el del Emisor si el primero es del Cliente
    ruc_emisor = ruc_match[-1] if ruc_match else "No detectado"
    
    # 2. Buscar Número de Factura (XXX-XXX-XXXXXXX)
    nro_match = re.search(r'(\d{3}-\d{3}-\d{7})', texto)
    nro_factura = nro_match.group(1) if nro_match else "S/N"
    
    # 3. Extraer Nombre del Proveedor (Simplificado)
    nombre = "Proveedor Desconocido"
    if "ITAU" in texto.upper(): nombre = "BANCO ITAÚ PARAGUAY S.A."
    elif "CRYSALIS" in texto.upper(): nombre = "CRYSALIS S.R.L."
    else:
        # Intenta sacar la primera línea que suele ser el nombre
        lineas = [l.strip() for l in texto.split('\n') if len(l.strip()) > 5]
        nombre = lineas[0][:40] if lineas else "Detectar manualmente"

    # 4. Capturar Totales (Exentas, 5%, 10%)
    # Esta lógica busca la fila de totales que suele tener 3 columnas al final
    totales_raw = re.findall(r'(\d{4,12})', texto_limpio)
    
    # Lógica de seguridad para capturar montos
    # Asumimos valores por defecto si no detecta la tabla perfectamente
    try:
        total_factura = int(totales_raw[-1]) if totales_raw else 0
        # Estimación de base imponible para el cálculo (Total / 1.1)
        # En una versión avanzada, mapearíamos exactamente la columna del PDF
        gravada_10 = round(total_factura / 1.1) 
        iva_10 = total_factura - gravada_10
    except:
        total_factura = 0; gravada_10 = 0; iva_10 = 0

    return {
        "Proveedor": nombre,
        "RUC": ruc_emisor,
        "Factura": nro_factura,
        "Total": total_factura,
        "Base_Imponible": gravada_10,
        "IVA_10": iva_10
    }

# --- PROCESAMIENTO Y LÓGICA DE RETENCIÓN ---
if archivos:
    datos_crudos = []
    for f in archivos:
        try:
            datos_crudos.append(extraer_datos_factura(f))
        except Exception as e:
            st.error(f"Error leyendo {f.name}: {e}")

    # Convertir a Tabla y Ordenar por Número de Factura (Criterio de Tesorería)
    df_temp = pd.DataFrame(datos_crudos).sort_values(by="Factura")
    
    resultados = []
    acumulado_por_ruc = {} # Para rastrear los 10 jornales por mes

    for _, fila in df_temp.iterrows():
        ruc = fila['RUC']
        bi = fila['Base_Imponible']
        iva = fila['IVA_10']
        
        if ruc not in acumulado_por_ruc:
            acumulado_por_ruc[ruc] = 0
        
        # --- REGLA 1: RETENCIÓN IVA (30%) ---
        # Se retiene si la suma del mes >= 10 Jornales
        ret_iva = 0
        if (acumulado_por_ruc[ruc] + bi) >= UMBRAL_IVA_10_JORNALES:
            ret_iva = iva * 0.30
        
        # --- REGLA 2: RETENCIÓN IRE (4%) ---
        # Se retiene si la factura individual >= 1 Salario Mínimo
        ret_ire = bi * 0.04 if bi >= SALARIO_MIN_IRE else 0
        
        acumulado_por_ruc[ruc] += bi
        
        resultados.append({
            "Proveedor": fila['Proveedor'],
            "RUC": ruc,
            "Nro Factura": fila['Factura'],
            "Monto Total": fila['Total'],
            "Base Imponible": bi,
            "Retención IVA (30%)": round(ret_iva),
            "Retención IRE (4%)": round(ret_ire),
            "Total Retenido": round(ret_iva + ret_ire),
            "Neto a Pagar": fila['Total'] - round(ret_iva + ret_ire)
        })

    # Mostrar Tabla en Web
    df_final = pd.DataFrame(resultados)
    st.subheader("📋 Resumen de Retenciones a Aplicar")
    st.dataframe(df_final.style.format({
        "Monto Total": "{:,.0f}", "Base Imponible": "{:,.0f}", 
        "Retención IVA (30%)": "{:,.0f}", "Retención IRE (4%)": "{:,.0f}",
        "Total Retenido": "{:,.0f}", "Neto a Pagar": "{:,.0f}"
    }).set_properties(**{'text-align': 'right'}))

    # --- DESCARGA EXCEL ---
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df_final.to_excel(writer, index=False, sheet_name='Retenciones_Dia')
        # Formateo básico de celdas puede ir aquí
    
    st.download_button(
        label="📥 Descargar Excel para OneDrive",
        data=buffer.getvalue(),
        file_name=f"Retenciones_{fecha_pago.strftime('%Y%m%d')}.xlsx",
        mime="application/vnd
