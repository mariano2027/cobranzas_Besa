import io
import os
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

# Configuración de la página web
st.set_page_config(page_title="Panel Séptima", page_icon="📈", layout="wide")

st.title("📈 Panel Séptima - Control de Cuentas Corrientes y Morosidad")
st.markdown("Cargue el archivo consolidado generado para analizar la performance de cobros por vendedor.")

# --- INICIALIZAR SESSION STATE PARA PERSISTENCIA ---
if "archivo_cargado" not in st.session_state:
    st.session_state.archivo_cargado = None
if "df_global" not in st.session_state:
    st.session_state.df_global = None

# --- COMPONENTE DE CARGA (Persistente) ---
archivo_subido = st.file_uploader("Adjunte la planilla consolidada (.xlsx)", type=["xlsx"])

# Si el usuario sube un archivo nuevo, lo guardamos en el session_state
if archivo_subido is not None:
    st.session_state.archivo_cargado = archivo_subido

# Usamos el archivo almacenado en la sesión (si existe)
if st.session_state.archivo_cargado is not None:
    try:
        # Si el DataFrame no está cargado en memoria, lo leemos
        if st.session_state.df_global is None:
            df = pd.read_excel(st.session_state.archivo_cargado, engine="openpyxl")
            df.columns = df.columns.astype(str).str.strip()
            
            # Validar columnas críticas
            columnas_necesarias = ["Vendedor", "Saldo Deuda (Imp. Total)", "Días de Atraso", "Cliente", "Razon Social"]
            for col in columnas_necesarias:
                if col not in df.columns:
                    st.error(f"❌ Falta la columna obligatoria '{col}' en el archivo subido.")
                    st.stop()
                    
            # Asegurar tipos de datos correctos para cálculos matemáticos
            df["Saldo Deuda (Imp. Total)"] = pd.to_numeric(df["Saldo Deuda (Imp. Total)"], errors="coerce").fillna(0.0)
            df["Días de Atraso"] = pd.to_numeric(df["Días de Atraso"], errors="coerce").fillna(0).astype(int)
            
            # --- CLASIFICACIÓN ESTRICTA DE TRAMOS DE MOROSIDAD ---
            def asignar_tramo(dias):
                if dias <= 60: return "Menos de 60 días"
                elif 61 <= dias <= 75: return "61 a 75 días"
                elif 76 <= dias <= 90: return "76 a 90 días"
                else: return "Más de 90 días"
                
            df["Tramo Morosidad"] = df["Días de Atraso"].apply(asignar_tramo)
            st.session_state.df_global = df

        # Recuperamos el DataFrame desde la memoria de sesión
        df = st.session_state.df_global

        # --- FILTRO GENERAL POR RAZÓN SOCIAL (BARRA LATERAL) ---
        st.sidebar.markdown("---")
        st.sidebar.subheader("🔎 Filtro Global - Panel Séptima")
        lista_razon_social = sorted(df['Razon Social'].dropna().unique())
        razon_social_global = st.sidebar.selectbox("Filtrar por Razón Social:", ["Todos"] + list(lista_razon_social))

        # Aplicar el filtro global al dataframe principal
        if razon_social_global != "Todos":
            df_trabajo = df[df['Razon Social'] == razon_social_global]
        else:
            df_trabajo = df

        # --- SECCIÓN 1: ALERTAS Y DESCARGAS > 75 DÍAS ---
        st.markdown("## 🚨 Alertas de Morosidad Crítica (> 75 Días)")
        df_critico_global = df_trabajo[df_trabajo["Días de Atraso"] > 75]
        
        if not df_critico_global.empty:
            st.warning(f"Se detectaron deudas vencidas críticas bajo los filtros seleccionados.")
            
            vendedores_criticos = sorted(df_critico_global["Vendedor"].unique())
            vendedor_sel = st.selectbox("Seleccione un Vendedor para descargar sus deudas mayores a 75 días:", vendedores_criticos)
            
            if vendedor_sel:
                df_vend_critico = df_critico_global[df_critico_global["Vendedor"] == vendedor_sel]
                
                buf_critico = io.BytesIO()
                with pd.ExcelWriter(buf_critico, engine="openpyxl") as writer:
                    df_vend_critico.to_excel(writer, index=False, sheet_name="Morosidad_Critica")
                    
                st.download_button(
                    label=f"📥 Descargar Excel Morosos >75 días - {vendedor_sel}",
                    data=buf_critico.getvalue(),
                    file_name=f"Morosidad_75dias_{str(vendedor_sel).replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                columnas_mostrar = [c for c in ["Cliente", "Razon Social", "Nro Comprobante", "Fecha Emisión", "Saldo Deuda (Imp. Total)", "Días de Atraso"] if c in df_vend_critico.columns]
                st.dataframe(df_vend_critico[columnas_mostrar], use_container_width=True, hide_index=True)
        else:
            st.success("✅ ¡Excelente! No se detectaron deudas mayores a 75 días bajo los filtros actuales.")

        st.markdown("---")

        # --- SECCIÓN 2: PERFORMANCE Y GRÁFICOS DINÁMICOS ---
        st.markdown("## 📊 Análisis Visual por Vendedor (Panel Séptima)")
        
        vendedores_todos = sorted(df_trabajo["Vendedor"].unique())
        if vendedores_todos:
            vendedor_grafico = st.selectbox("Seleccione el Vendedor que desea auditar:", vendedores_todos)
        else:
            vendedor_grafico = None
        
        if vendedor_grafico:
            df_v = df_trabajo[df_trabajo["Vendedor"] == vendedor_grafico]
            
            # Agrupar los saldos totales por cada tramo de deudas
            resumen_tramos = df_v.groupby("Tramo Morosidad")["Saldo Deuda (Imp. Total)"].sum().reset_index()
            
            todos_tramos = ["Menos de 60 días", "61 a 75 días", "76 a 90 días", "Más de 90 días"]
            saldos_fijos = []
            
            for tramo in todos_tramos:
                match = resumen_tramos[resumen_tramos["Tramo Morosidad"] == tramo]
                if not match.empty:
                    saldos_fijos.append(float(match["Saldo Deuda (Imp. Total)"].iloc[0]))
                else:
                    saldos_fijos.append(0.0)
            
            total_vendedor = sum(saldos_fijos)
            
            if total_vendedor > 0:
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown(f"#### Resumen de Performance de Cobros: **{vendedor_grafico}**")
                    df_tabla_ver = pd.DataFrame({
                        "Tramo de Vencimiento": todos_tramos,
                        "Monto Total": saldos_fijos,
                        "Porcentaje": [(s / total_vendedor) * 100 for s in saldos_fijos]
                    })
                    st.dataframe(df_tabla_ver.style.format({"Monto Total": "${:,.2f}", "Porcentaje": "{:.1f}%"}), use_container_width=True, hide_index=True)
                    st.metric(label="Total Cuenta Corriente Asignada", value=f"${total_vendedor:,.2f}")
                
                with col2:
                    fig, ax = plt.subplots(figsize=(5, 5))
                    colores_tramos = ["#2F855A", "#ECC94B", "#DD6B20", "#C53030"] # Verde, Amarillo, Naranja, Rojo
                    
                    labels_filtrados = [todos_tramos[i] for i in range(4) if saldos_fijos[i] > 0]
                    saldos_filtrados = [saldos_fijos[i] for i in range(4) if saldos_fijos[i] > 0]
                    colores_filtrados = [colores_tramos[i] for i in range(4) if saldos_fijos[i] > 0]
                    
                    ax.pie(
                        saldos_filtrados, 
                        labels=labels_filtrados, 
                        autopct='%1.1f%%', 
                        startangle=90, 
                        colors=colores_filtrados,
                        textprops={'fontsize': 10, 'weight': 'bold'}
                    )
                    ax.axis('equal')
                    st.pyplot(fig)

                # --- MATRIZ DINÁMICA DE SALDOS Y CLIENTES ---
                st.markdown(f"### 📌 Matriz Dinámica de Saldos y Clientes - Vendedor: {vendedor_grafico}")
                
                df_dinamico_v = df_v.copy()
                
                df_dinamico_v["Menos de 60 Días"] = df_dinamico_v.apply(lambda r: r["Saldo Deuda (Imp. Total)"] if r["Días de Atraso"] <= 60 else 0.0, axis=1)
                df_dinamico_v["61-75 Días"] = df_dinamico_v.apply(lambda r: r["Saldo Deuda (Imp. Total)"] if 61 <= r["Días de Atraso"] <= 75 else 0.0, axis=1)
                df_dinamico_v["76-90 Días"] = df_dinamico_v.apply(lambda r: r["Saldo Deuda (Imp. Total)"] if 76 <= r["Días de Atraso"] <= 90 else 0.0, axis=1)
                df_dinamico_v["Mayor a 90 Días"] = df_dinamico_v.apply(lambda r: r["Saldo Deuda (Imp. Total)"] if r["Días de Atraso"] > 90 else 0.0, axis=1)
                
                df_resumen_v = df_dinamico_v.groupby(["Cliente", "Razon Social"]).agg({
                    "Menos de 60 Días": "sum",
                    "61-75 Días": "sum",
                    "76-90 Días": "sum",
                    "Mayor a 90 Días": "sum",
                    "Saldo Deuda (Imp. Total)": "sum"
                }).reset_index()
                
                df_resumen_v = df_resumen_v.sort_values(by=["Saldo Deuda (Imp. Total)"], ascending=False)
                df_resumen_v.columns = ["Código Cliente", "Razón Social", "Menos de 60 Días", "61-75 Días", "76-90 Días", "Mayor a 90 Días", "Total General"]
                
                output_dinamico_v = io.BytesIO()
                with pd.ExcelWriter(output_dinamico_v, engine='openpyxl') as writer:
                    df_resumen_v.to_excel(writer, index=False, sheet_name=f"Matriz_{str(vendedor_grafico)[:10]}")
                    
                st.download_button(
                    label=f"📥 Descargar Matriz Dinámica en Excel de {vendedor_grafico} (.XLSX)",
                    data=output_dinamico_v.getvalue(),
                    file_name=f"Matriz_Morosidad_{str(vendedor_grafico).replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                st.dataframe(
                    df_resumen_v.style.format({
                        "Menos de 60 Días": "${:,.2f}",
                        "61-75 Días": "${:,.2f}",
                        "76-90 Días": "${:,.2f}",
                        "Mayor a 90 Días": "${:,.2f}",
                        "Total General": "${:,.2f}"
                    }).background_gradient(subset=["Total General", "Mayor a 90 Días"], cmap="Reds"),
                    use_container_width=True,
                    hide_index=True
                )

                # --- NUEVA SECCIÓN: DETALLE DE FACTURAS / COMPROBANTES ---
                st.markdown(f"### 📄 Detalle de Comprobantes Individuales")
                cols_detalle = [c for c in ["Cliente", "Razon Social", "Tipo Comprobante", "Nro Comprobante", "Fecha Emisión", "Importe Original", "Saldo Deuda (Imp. Total)", "Días de Atraso", "Tramo Morosidad"] if c in df_v.columns]
                st.dataframe(
                    df_v[cols_detalle].style.format({
                        "Importe Original": "${:,.2f}",
                        "Saldo Deuda (Imp. Total)": "${:,.2f}"
                    }),
                    use_container_width=True,
                    hide_index=True
                )

            else:
                st.info(f"El vendedor {vendedor_grafico} no registra saldo de deudas con los filtros actuales.")
                
    except Exception as e:
        st.error(f"Ocurrió un error al procesar la planilla: {e}")
else:
    st.info("A la espera del archivo Excel unificado para desplegar las métricas del Panel Séptima.")
