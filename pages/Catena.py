import io
import os
import re
from datetime import datetime
import pandas as pd
from pypdf import PdfReader
import streamlit as st

# Configuración básica de pantalla plana
st.set_page_config(page_title="Procesador Contable", page_icon="📊", layout="wide")

st.title("📊 Panel de Control de Cuentas Corrientes")
st.markdown("Cargue los archivos para cruzar clientes y descargar el informe general por vendedor.")

def extraer_datos_pdf_final(pdf_file):
    """Parsea el PDF capturando comprobantes, importes, localidad, límite de crédito y días en calle."""
    datos_comprobantes = []
    patron_cliente = re.compile(r"Cliente\s+([A-Z0-9]+)")
    patron_cuit = re.compile(r"Nro\.\s+CUIT\s+([\d-]+)")
    patron_localidad = re.compile(r"Localidad\s+(.*?)(?:\s{2,}|Provincia|$)", re.IGNORECASE)
    patron_fecha = re.compile(r"(\d{2}/\d{2}/\d{2,4})")
    
    patron_credito = re.compile(r"(?:Cr[eé]dito|L[ií]mite)\s*[:]?\s*\$?\s*([\d\.,]+)", re.IGNORECASE)

    reader = PdfReader(pdf_file)
    fecha_actual_sistema = datetime.now()

    cliente_actual = None
    cuit_actual = None
    localidad_actual = "SIN LOCALIDAD"
    credito_actual = 0.0
    dias_calle_actual = 0.0
    capturar_siguiente_dias_calle = False

    for num_pagina, pagina in enumerate(reader.pages, start=1):
        texto = pagina.extract_text()
        if not texto:
            continue

        lineas = texto.split("\n")
        for idx, linea in enumerate(lineas):
            linea_strip = linea.strip()

            if "Proyección" in linea_strip:
                continue

            if "Cliente" in linea_strip and not "Desde" in linea_strip and not "Hasta" in linea_strip:
                match_cli = patron_cliente.search(linea_strip)
                if match_cli:
                    cliente_actual = match_cli.group(1).strip().upper()
                continue

            if "Nro. CUIT" in linea_strip:
                match_cuit = patron_cuit.search(linea_strip)
                if match_cuit:
                    cuit_actual = match_cuit.group(1).strip()
                continue

            # Captura de Localidad del PDF
            if "Localidad" in linea_strip:
                match_loc = patron_localidad.search(linea_strip)
                if match_loc:
                    localidad_actual = match_loc.group(1).strip().upper()
                else:
                    partes_loc = linea_strip.split("Localidad")
                    if len(partes_loc) > 1 and partes_loc[1].strip():
                        localidad_actual = partes_loc[1].split("Provincia")[0].strip().upper()
                continue

            # Captura de Límite de Crédito (se mantiene intacta tal como estaba)
            match_cred = patron_credito.search(linea_strip)
            if match_cred:
                aux_cred = match_cred.group(1).replace(",", "")
                try:
                    credito_actual = float(aux_cred)
                except:
                    pass
                continue

            # Captura flexible de Días en Calle
            if "dias en calle" in linea_strip.lower():
                capturar_siguiente_dias_calle = True
                continue

            if capturar_siguiente_dias_calle:
                partes_linea = linea_strip.split()
                if partes_linea:
                    candidatos_floats = []
                    for p in partes_linea:
                        p_clean = p.replace(",", "")
                        try:
                            val = float(p_clean)
                            candidatos_floats.append(val)
                        except:
                            pass
                    if candidatos_floats:
                        dias_calle_actual = candidatos_floats[-1]
                capturar_siguiente_dias_calle = False
                continue

            # Patrón flexible para comprobantes
            match_comp = re.search(r"([A-Z]-\d{4}[-_\s]\w+)", linea_strip)
            fechas_encontradas = patron_fecha.findall(linea_strip)

            if match_comp and fechas_encontradas and cliente_actual:
                comprobante = match_comp.group(0)
                fecha_comprobante_str = str(fechas_encontradas[0]).strip()

                tipo_comprobante = "Factura"
                es_credito = False
                if "NC" in linea_strip or "Nota de Credito" in linea_strip:
                    tipo_comprobante = "Nota de Crédito"
                    es_credito = True
                elif "RC" in linea_strip or "Recibo" in linea_strip:
                    tipo_comprobante = "Recibo"
                    es_credito = True
                elif "ND" in linea_strip or "Nota de Debito" in linea_strip:
                    tipo_comprobante = "Nota de Débito"

                partes = linea_strip.split()
                montos_candidatos = []
                
                for p in partes:
                    p_clean = "".join([c for c in p if c.isdigit() or c in [".", ",", "-", "(", ")"]]).strip()
                    if p_clean and any(char.isdigit() for char in p_clean):
                        if "/" not in p_clean and p_clean != comprobante and ("." in p_clean or "," in p_clean):
                            montos_candidatos.append(p_clean)

                def limpiar_monto_contable_americano(cadena_monto):
                    if not cadena_monto: return 0.0
                    aux = str(cadena_monto).replace(",", "")
                    if "(" in aux or ")" in aux:
                        aux = "-" + aux.replace("(", "").replace(")", "")
                    try: return float(aux)
                    except: return 0.0

                importe_original = 0.0
                saldo_remanente = 0.0

                if len(montos_candidatos) >= 2:
                    importe_original = limpiar_monto_contable_americano(montos_candidatos[-1])
                    saldo_remanente = limpiar_monto_contable_americano(montos_candidatos[0])
                elif len(montos_candidatos) == 1:
                    importe_original = limpiar_monto_contable_americano(montos_candidatos[0])
                    saldo_remanente = importe_original

                if es_credito:
                    if importe_original > 0: importe_original = -importe_original
                    if saldo_remanente > 0: saldo_remanente = -saldo_remanente

                dias_atraso = 0
                if saldo_remanente > 0:
                    try:
                        componentes = fecha_comprobante_str.split("/")
                        formato_anio = "%y" if len(componentes[-1]) == 2 else "%Y"
                        fecha_comp_dt = datetime.strptime(fecha_comprobante_str, f"%d/%m/{formato_anio}")
                        diferencia = fecha_actual_sistema - fecha_comp_dt
                        dias_atraso = max(0, diferencia.days)
                    except:
                        dias_atraso = 0

                datos_comprobantes.append(
                    {
                        "Cliente_PDF": cliente_actual,
                        "Tipo Comprobante": tipo_comprobante,
                        "Nro Comprobante": comprobante,
                        "Fecha Emisión": fecha_comprobante_str,
                        "Importe Original": importe_original,
                        "Saldo Deuda (Imp. Total)": saldo_remanente,
                        "Días de Atraso": dias_atraso,
                        "Localidad": localidad_actual,
                        "Limite Credito": credito_actual,
                        "Dias en Calle": dias_calle_actual,
                        "Pagina PDF": num_pagina,
                    }
                )
                
            if "Saldo Final" in linea_strip:
                cliente_actual = None
                cuit_actual = None
                localidad_actual = "SIN LOCALIDAD"
                credito_actual = 0.0
                dias_calle_actual = 0.0
                capturar_siguiente_dias_calle = False
                
    return pd.DataFrame(datos_comprobantes)

# --- BARRA LATERAL GLOBAL ---
st.sidebar.header("📁 Cargar Documentos")
excel_subido = st.sidebar.file_uploader("1. Base maestra de Clientes (.xlsx)", type=["xlsx"])
pdf_subido = st.sidebar.file_uploader("2. Listado de Cuenta Corriente (.pdf)", type=["pdf"])

if 'df_final' not in st.session_state:
    st.session_state['df_final'] = None

if excel_subido and pdf_subido and st.session_state['df_final'] is None:
    with st.spinner("Cruzando datos en limpio..."):
        try:
            df_excel = pd.read_excel(excel_subido, engine="openpyxl")
            df_excel.columns = df_excel.columns.astype(str).str.strip()
            
            df_excel["Cliente_Clean"] = df_excel["Cliente"].astype(str).str.strip().str.upper()
            df_maestro = df_excel[["Cliente_Clean", "Razon Social", "VENDEDOR"]].drop_duplicates()

            df_pdf = extraer_datos_pdf_final(pdf_subido)
            df_res = pd.merge(df_pdf, df_maestro, left_on="Cliente_PDF", right_on="Cliente_Clean", how="left")
            
            # Parche explícito de control para Valyva
            es_valyva = df_res["Cliente_PDF"].str.contains("8604405", na=False) | df_res["Cliente_PDF"].str.contains("860440S", na=False)
            df_res.loc[es_valyva, "VENDEDOR"] = "Claudio"
            df_res.loc[es_valyva, "Razon Social"] = "VALYVA MERCADOS INTEGRADOS SA"
            
            df_res["Razon Social"] = df_res["Razon Social"].fillna("Cliente Nuevo No Encontrado")
            df_res["VENDEDOR"] = df_res["VENDEDOR"].fillna("Sin Vendedor Asignado")
            df_res["Cliente"] = df_res["Cliente_PDF"]
            
            df_final_render = df_res[[
                "Cliente", "Razon Social", "Localidad", "Tipo Comprobante", "Nro Comprobante", 
                "Fecha Emisión", "Importe Original", "Saldo Deuda (Imp. Total)", 
                "Días de Atraso", "Limite Credito", "Dias en Calle", "VENDEDOR", "Pagina PDF"
            ]].rename(columns={"VENDEDOR": "Vendedor"})
            
            st.session_state['df_final'] = df_final_render
        except Exception as e:
            st.error(f"Error al cruzar los archivos: {e}")

# --- PANTALLA PRINCIPAL DIRECTA Y LIMPIA ---
if st.session_state['df_final'] is not None:
    df_final = st.session_state['df_final']
    st.success(f"✅ ¡Proceso completado con éxito! Se unificaron los registros comerciales.")
    
    output_completo = io.BytesIO()
    with pd.ExcelWriter(output_completo, engine='openpyxl') as writer:
        df_final.to_excel(writer, index=False, sheet_name="General")
        
    st.download_button(
        label="📥 Descargar Excel con Clientes Asignados (.XLSX)", 
        data=output_completo.getvalue(), 
        file_name="Cuentas_Corrientes_Vendedores.xlsx", 
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    st.markdown("---")
    st.markdown("**Vista previa rápida de los datos cruzados:**")
    st.dataframe(df_final, use_container_width=True, hide_index=True)
else:
    st.info("Por favor, cargue la base de clientes (Excel de 3 columnas) y el PDF de cuenta corriente en el panel izquierdo.")
  
