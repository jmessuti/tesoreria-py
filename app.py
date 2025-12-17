def extraer_datos_profesional(file):
    with pdfplumber.open(file) as pdf:
        texto = "".join([p.extract_text() for p in pdf.pages if p.extract_text()])
    
    # 1. Buscador de RUC más flexible (acepta R.U.C, RUC:, etc)
    ruc_pattern = r'(?:RUC|R\.U\.C)[:\s]*(\d+-\d+|\d+)'
    ruc_match = re.search(ruc_pattern, texto, re.IGNORECASE)
    
    # 2. Buscador de Número de Factura (XXX-XXX-XXXXXXX)
    nro_pattern = r'(\d{3}-\d{3}-\d{7})'
    nro_match = re.search(nro_pattern, texto)
    
    # 3. Buscador de Montos (Lógica mejorada)
    # Buscamos números largos que terminen en la fila de totales
    montos = re.findall(r'(\d{1,3}(?:\.\d{3})*)', texto)
    valores = [int(m.replace('.', '')) for m in montos if len(m) >= 4] # Filtra números pequeños
    
    # Asumimos que el valor más alto cerca del final es el Total
    total = max(valores) if valores else 0
    
    # Si no detecta RUC, intentamos buscar el nombre del proveedor (ej. CRYSALIS)
    prov_nombre = "Proveedor Genérico"
    if "CRYSALIS" in texto.upper(): prov_nombre = "CRYSALIS S.R.L."
    if "ITAU" in texto.upper(): prov_nombre = "BANCO ITAU PY"

    return {
        "Proveedor": prov_nombre,
        "RUC": ruc_match.group(1) if ruc_match else "No detectado",
        "Factura": nro_match.group(1) if nro_match else "No detectado",
        "Total": total,
        "Base_Imponible": round(total / 1.1) if total > 0 else 0,
        "IVA_10": round((total / 1.1) * 0.1) if total > 0 else 0
    }
