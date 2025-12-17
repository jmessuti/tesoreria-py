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

with st.sidebar:
    st.header("Configuración de Pago")
    fecha_pago = st.date_input("Fecha de Pago", datetime.now())

archivos = st.file_uploader("Subir Facturas (PDF)", type="pdf", accept_multiple_files=True)

def extraer_datos_factura(file):
    with pdfplumber.open(file) as pdf:
        texto = ""
        for page in pdf.pages:
            texto += page.extract_text() + "\n"
    
    # 1. Buscar RUC del Emisor
    ruc_match = re.findall(r'(\d+-\d+)', texto)
    ruc_emisor = ruc_match[-1] if ruc_match else "No detectado"
    
    # 2. Buscar Número de Factura
    nro_match = re.search(r'(\d{3}-\d{3}-\d{7})', texto)
    nro_factura = nro_match.group(1) if nro_match else "S/N"
    
    # 3. EXTRAER MONTO TOTAL (Mejorado)
    # Buscamos palabras clave y capturamos el número que sigue
    total_factura = 0
    # Patrón para buscar números con puntos o comas después de palabras de total
    patron_total = r'(?:TOTAL|SUBTOTAL|TOTAL A PAGAR|VALOR VENTA|TOTAL Gs\.?)[:\s]+([\d\.]+)'
    match_total = re.search(patron_total, texto, re.IGNORECASE)
    
    if match_total:
        total_str = match_total.group(1).replace(".", "")
        total_factura = int(total_str)
    else:
        # Si falla el patrón, buscamos el último número grande del documento (que suele ser el total final)
        numeros = re.findall(r'\b\d{1,3}(?:\.\d{3})+\b', texto)
        if numeros:
            total_factura = int(numeros[-1].replace(".", ""))

    # 4. Cálculos derivados
    # En Paraguay, el IVA 10% se saca dividiendo el total por 11
    iva_10 = round(total_factura / 11)
    gravada_10 = total_factura - iva_10

    # Identificar Proveedor por texto
    nombre = "Proveedor Desconocido"
    if "ITAU" in texto.upper(): nombre = "BANCO ITAÚ PARAGUAY S.A."
    elif "CRYSALIS" in texto.upper(): nombre = "CRYSALIS S.R.L."
    else:
        lineas = [l.strip() for l in texto.split('\n') if len(l.strip()) > 5]
        nombre = lineas[0][:40] if lineas else "Detectar manualmente"

    return {
        "Proveedor": nombre,
        "RUC": ruc_emisor,
        "Factura": nro_factura,
        "Total": total_factura,
        "Base_Imponible": gravada_10,
        "IVA_10": iva_10
    }

if archivos:
    datos_crudos = []
    for f in archivos:
        try:
            datos_crudos.append(extraer_datos_factura(f))
        except Exception as e:
            st.error(f"Error en {f.name}: {e}")

    if datos_crudos:
        df_temp = pd.DataFrame(datos_crudos).sort_values(by="Factura")
        resultados = []
        acumulado_por_ruc = {}

        for _, fila in df_temp.iterrows():
            ruc = fila['RUC']
            bi = fila['Base_Imponible']
            iva = fila['IVA_10']
            
            if ruc not in acumulado_por_ruc:
                acumulado_por_ruc[ruc] = 0
            
            ret_iva = 0
            if (acumulado_por_ruc[ruc] + bi) >= UMBRAL_IVA_10_JORNALES:
                ret_iva = iva * 0.30
            
            ret_ire = bi * 0.04 if bi >= SALARIO_MIN_IRE else 0
            acumulado_por_ruc[ruc] += bi
            
            resultados.append({
                "Proveedor": fila['Proveedor'],
                "RUC": ruc,
                "Nro Factura": fila['Factura'],
                "Monto Factura": fila['Total'],
                "Base Imponible": bi,
                "Retención IVA (30%)": round(ret_iva),
                "Retención IRE (4%)": round(ret_ire),
                "Total Retenido": round(ret_iva + ret_ire),
                "Neto a Pagar": fila['Total'] - round(ret_iva + ret_ire)
            })

        df_final = pd.DataFrame(resultados)
        st.subheader("📋 Resumen de Retenciones")
        st.dataframe(df_final.style.format(precision=0, thousands="."))

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Retenciones')
        
        st.download_button(
            label="📥 Descargar Excel para OneDrive",
            data=buffer.getvalue(),
            file_name=f"Retenciones_{fecha_pago.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.ms-excel"
        )
