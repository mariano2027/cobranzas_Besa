import os
import base64
import io
import re
import dash
from dash import dcc, html, Input, Output, State, ctx
import dash_bootstrap_components as dbc
import dash_ag_grid as dag
import pandas as pd
import plotly.express as px

app = dash.Dash(
    __name__,
    external_stylesheets=[
        dbc.themes.SLATE,
        "https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.5/font/bootstrap-icons.css",
        "https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap"
    ],
    title="Dashboard Ejecutivo de Cuentas Corrientes",
    suppress_callback_exceptions=True
)

app.index_string = '''
<!DOCTYPE html>
<html>
    <head>
        {%metas%}
        <title>{%title%}</title>
        {%favicon%}
        {%css%}
        <style>
            body { font-family: 'Inter', sans-serif; background-color: #0f172a; color: #f8fafc; }
            .exec-card { background: linear-gradient(145deg, #1e293b, #0f172a); border: 1px solid #334155; border-radius: 12px; }
            .kpi-title { font-size: 0.75rem; text-transform: uppercase; color: #94a3b8; font-weight: 600; }
            .kpi-value { font-size: 1.5rem; font-weight: 700; color: #f8fafc; }
            .upload-box { border: 2px dashed #475569; border-radius: 12px; background-color: #1e293b; }
            .upload-box:hover { border-color: #38bdf8; background-color: #334155; }
            
            .ag-theme-alpine-dark, 
            .ag-theme-alpine-dark .ag-root-wrapper,
            .ag-theme-alpine-dark .ag-root,
            .ag-theme-alpine-dark .ag-body,
            .ag-theme-alpine-dark .ag-body-viewport,
            .ag-theme-alpine-dark .ag-center-cols-container,
            .ag-theme-alpine-dark .ag-row,
            .ag-theme-alpine-dark .ag-cell {
                background-color: #162032 !important;
                background: #162032 !important;
                color: #f8fafc !important;
            }
            .ag-theme-alpine-dark .ag-row-odd {
                background-color: #1b263b !important;
                background: #1b263b !important;
            }
            .ag-theme-alpine-dark .ag-header,
            .ag-theme-alpine-dark .ag-header-row {
                background-color: #0f172a !important;
                background: #0f172a !important;
                color: #38bdf8 !important;
            }
            .ag-theme-alpine-dark .ag-header-cell {
                background-color: #0f172a !important;
                color: #38bdf8 !important;
            }
            .ag-theme-alpine-dark .ag-paging-panel {
                background-color: #1e293b !important;
                color: #f8fafc !important;
                border-top: 1px solid #334155 !important;
            }
            .ag-root-wrapper { border: 1px solid #334155 !important; border-radius: 8px !important; }
            .ag-cell { border-color: #283548 !important; display: flex; align-items: center; }
            
            .grid-cell-red { background-color: rgba(239, 68, 68, 0.25) !important; color: #fca5a5 !important; font-weight: 700; border-left: 4px solid #ef4444 !important; }
            .grid-cell-orange { background-color: rgba(249, 115, 22, 0.25) !important; color: #fdba74 !important; font-weight: 600; border-left: 4px solid #f97316 !important; }
            .grid-cell-yellow { background-color: rgba(245, 158, 11, 0.2) !important; color: #fde047 !important; font-weight: 600; }
            .grid-cell-green { background-color: rgba(16, 185, 129, 0.15) !important; color: #6ee7b7 !important; }
        </style>
    </head>
    <body>
        {%app_entry%}
        <footer>
            {%config%}
            {%scripts%}
            {%renderer%}
        </footer>
    </body>
</html>
'''

def clean_currency_series(series):
    def parse_val(v):
        if pd.isna(v): return 0.0
        if isinstance(v, (int, float)): return float(v)
        s = str(v).strip()
        s = re.sub(r'[^\d,\.-]', '', s)
        if not s: return 0.0
        if ',' in s and '.' in s:
            if s.rfind(',') > s.rfind('.'): s = s.replace('.', '').replace(',', '.')
            else: s = s.replace(',', '')
        elif ',' in s: s = s.replace(',', '.')
        try: return float(s)
        except ValueError: return 0.0
    return series.apply(parse_val)

def format_currency_human(value):
    if abs(value) >= 1e9: return f"${value/1e9:,.2f} B"
    elif abs(value) >= 1e6: return f"${value/1e6:,.2f} M"
    elif abs(value) >= 1e3: return f"${value/1e3:,.1f} K"
    return f"${value:,.2f}"

def format_currency_full(value):
    return f"${value:,.2f}"

app.layout = dbc.Container([
    dcc.Download(id="download-criticos-excel"),
    dcc.Download(id="download-matriz-excel"),
    dcc.Download(id="download-dinamica-excel"),
    dcc.Store(id='stored-data'),

    dbc.Row([
        dbc.Col([
            html.Div([
                html.H1([html.I(className="bi bi-graph-up-arrow me-3 text-info"), "Dashboard Ejecutivo de Cuentas Corrientes"], className="fw-bold text-white mb-1 h2"),
                html.P("Auditoría de Morosidad, Control de Cartera y Performance por Vendedor", className="text-secondary mb-0 small")
            ], className="py-3 px-2")
        ], width=12)
    ], className="border-bottom border-secondary mb-4"),

    dbc.Row([
        dbc.Col([
            html.Div([
                html.Div([
                    html.H5([html.I(className="bi bi-sliders me-2"), "Panel de Control y Filtros"], className="fw-bold text-info mb-0"),
                    dbc.Button([html.I(className="bi bi-arrow-counterclockwise me-1"), "Limpiar Filtros"], id="btn-reset-filters", color="outline-warning", size="sm", className="py-1 px-2")
                ], className="d-flex justify-content-between align-items-center mb-3"),
                
                dbc.Row([
                    dbc.Col([
                        dcc.Upload(
                            id='upload-data',
                            children=html.Div([
                                html.I(className="bi bi-cloud-arrow-up fs-4 text-info d-block mb-1"),
                                html.Span("Arrastrá o selecciona la planilla (.xlsx)", className="small text-light fw-medium")
                            ], className="p-2 text-center"),
                            className="upload-box cursor-pointer",
                            multiple=False
                        ),
                        html.Div(id='upload-status-message', className="mt-2")
                    ], xs=12, lg=3, className="mb-3 mb-lg-0"),

                    dbc.Col([
                        html.Label("Filtrar Vendedor:", className="fw-semibold text-light mb-1 small"),
                        dcc.Dropdown(
                            id='vendedor-select',
                            options=[{'label': '🌐 TODOS LOS VENDEDORES', 'value': 'TODOS'}],
                            value='TODOS',
                            clearable=False,
                            style={'color': '#0f172a'}
                        )
                    ], xs=12, lg=3, className="mb-3 mb-lg-0"),

                    dbc.Col([
                        html.Label("Filtrar Localidad:", className="fw-semibold text-light mb-1 small"),
                        dcc.Dropdown(
                            id='localidad-select',
                            options=[{'label': '📍 TODAS LAS LOCALIDADES', 'value': 'TODOS'}],
                            value='TODOS',
                            clearable=False,
                            style={'color': '#0f172a'}
                        )
                    ], xs=12, lg=3, className="mb-3 mb-lg-0"),

                    dbc.Col([
                        html.Label("Filtrar Cliente / Razón Social:", className="fw-semibold text-light mb-1 small"),
                        dcc.Dropdown(
                            id='cliente-select',
                            options=[{'label': '🔍 TODOS LOS CLIENTES', 'value': 'TODOS'}],
                            value='TODOS',
                            clearable=False,
                            placeholder="Escribe o selecciona cliente...",
                            style={'color': '#0f172a'}
                        )
                    ], xs=12, lg=3)
                ])
            ], className="p-3 exec-card")
        ], width=12, className="mb-4")
    ]),

    dbc.Row([
        dbc.Col([
            dbc.Row(id='kpi-cards-row', className="g-3 mb-4"),

            dbc.Tabs(id="tabs-main", active_tab="tab-graficos", children=[
                dbc.Tab(label="📊 Resumen General & Gráficos", tab_id="tab-graficos"),
                dbc.Tab(label="🏆 Ranking Top Clientes", tab_id="tab-top-clientes"),
                dbc.Tab(label="🚨 Morosidad Crítica (>75 Días)", tab_id="tab-criticos"),
                dbc.Tab(label="📌 Ctas Corrientes", tab_id="tab-matriz"),
                dbc.Tab(label="📈 Dinámica por Tramo", tab_id="tab-dinamica")
            ], className="border-bottom border-secondary mb-3"),

            html.Div(id="tab-content-area")
        ], width=12)
    ])
], fluid=True, className="p-3 p-md-4")


@app.callback(
    Output('vendedor-select', 'value'),
    Output('localidad-select', 'value'),
    Output('cliente-select', 'value'),
    Input('btn-reset-filters', 'n_clicks'),
    prevent_initial_call=True
)
def reset_filters(n_clicks):
    if n_clicks:
        return 'TODOS', 'TODOS', 'TODOS'
    return dash.no_update, dash.no_update, dash.no_update


@app.callback(
    Output('stored-data', 'data'),
    Output('upload-status-message', 'children'),
    Output('vendedor-select', 'options'),
    Output('localidad-select', 'options'),
    Output('cliente-select', 'options'),
    Input('upload-data', 'contents'),
    State('upload-data', 'filename')
)
def parse_excel_and_store(contents, filename):
    default_vends = [{'label': '🌐 TODOS LOS VENDEDORES', 'value': 'TODOS'}]
    default_locs = [{'label': '📍 TODAS LAS LOCALIDADES', 'value': 'TODOS'}]
    default_clis = [{'label': '🔍 TODOS LOS CLIENTES', 'value': 'TODOS'}]

    if contents is None:
        return None, dbc.Alert("Cargue la planilla Excel.", color="info", className="py-1 px-2 small mb-0"), default_vends, default_locs, default_clis

    try:
        content_type, content_string = contents.split(',')
        decoded = base64.b64decode(content_string)
        df = pd.read_excel(io.BytesIO(decoded), engine="openpyxl")
        
        df.columns = df.columns.astype(str).str.strip()
        df = df.loc[:, ~df.columns.duplicated()].copy()
        
        cols_lower = [c.lower() for c in df.columns]
        if 'Cliente' not in df.columns and len(cols_lower) > 0:
            df['Cliente'] = df.iloc[:, 0]
        if 'Razon Social' not in df.columns and len(cols_lower) > 1:
            df['Razon Social'] = df.iloc[:, 1]

        if 'Vendedor' not in df.columns:
            df['Vendedor'] = "Vendedor General"
        else:
            df['Vendedor'] = df['Vendedor'].fillna("Vendedor General").astype(str).str.strip()

        # Detección y limpieza de Localidad
        col_localidad = next((c for c in df.columns if any(k in c.lower() for k in ['localidad', 'ciudad', 'municipio'])), None)
        if col_localidad:
            df['Localidad'] = df[col_localidad].fillna("Sin Localidad").astype(str).str.strip()
            # Si el nombre de la columna original era distinto de 'Localidad', normalizamos
            if col_localidad != 'Localidad':
                df.drop(columns=[col_localidad], inplace=True)
                df.rename(columns={col_localidad: 'Localidad'}, inplace=True) # por si acaso
        else:
            df['Localidad'] = "General"

        df['Cliente'] = df['Cliente'].astype(str).str.strip()
        df['Razon Social'] = df['Razon Social'].astype(str).str.strip()

        if 'Nro Comprobante' not in df.columns: df['Nro Comprobante'] = '-'
        else: df['Nro Comprobante'] = df['Nro Comprobante'].astype(str).str.strip()

        if 'Fecha Emisión' not in df.columns: df['Fecha Emisión'] = '-'

        col_importe = next((c for c in df.columns if any(k in c.lower() for k in ['importe', 'imp.', 'total'])), None)
        s_importe = clean_currency_series(df[col_importe]) if col_importe else pd.Series(0.0, index=df.index)

        col_saldo = next((c for c in df.columns if 'saldo' in c.lower()), None)
        if col_saldo:
            df['Saldo Deuda'] = clean_currency_series(df[col_saldo])
            mask_zero = (df['Saldo Deuda'] == 0) & (s_importe > 0)
            df.loc[mask_zero, 'Saldo Deuda'] = s_importe[mask_zero]
        else:
            df['Saldo Deuda'] = s_importe if s_importe.sum() > 0 else 0.0      

        col_atraso = next((c for c in df.columns if any(k in c.lower() for k in ['dias', 'atraso', 'antiguedad', 'vencido', 'mora']) and 'calle' not in c.lower()), None)
        if col_atraso:
            df['Días de Atraso'] = pd.to_numeric(
                df[col_atraso].astype(str).str.replace(r'[^\d-]', '', regex=True),
                errors='coerce'
            ).fillna(0).astype(int)
        else:
            df['Días de Atraso'] = 0

        col_limite = next((c for c in df.columns if any(k in c.lower() for k in ['limite', 'crédito', 'credito'])) , None)
        if col_limite:
            df['Limite Credito'] = clean_currency_series(df[col_limite])
        else:
            df['Limite Credito'] = 0.0

        col_dias_calle = next((c for c in df.columns if any(k in c.lower() for k in ['dias en calle', 'días en calle', 'calle'])) , None)
        if col_dias_calle:
            df['Dias en Calle'] = pd.to_numeric(
                df[col_dias_calle].astype(str).str.replace(r'[^\d\.,-]', '', regex=True).str.replace(',', '.'),
                errors='coerce'
            ).fillna(0.0)
        else:
            df['Dias en Calle'] = 0.0

        def asignar_tramo(dias):
            if dias <= 60: return "Menos de 60 días"
            elif 61 <= dias <= 75: return "61-75 Días"
            elif 76 <= dias <= 90: return "76-90 Días"
            else: return "Mayor a 90 Días"

        df["Tramo Morosidad"] = df["Días de Atraso"].apply(asignar_tramo)

        vendedores = sorted([str(v) for v in df["Vendedor"].dropna().unique() if str(v).strip() != ''])
        opts_vendedores = [{'label': '🌐 TODOS LOS VENDEDORES', 'value': 'TODOS'}] + [{'label': f"👤 {v}", 'value': v} for v in vendedores]

        localidades = sorted([str(l) for l in df["Localidad"].dropna().unique() if str(l).strip() != ''])
        opts_localidades = [{'label': '📍 TODAS LAS LOCALIDADES', 'value': 'TODOS'}] + [{'label': f"📍 {l}", 'value': l} for l in localidades]

        df_cli_unique = df[['Cliente', 'Razon Social']].drop_duplicates().sort_values(by='Razon Social')
        opts_clientes = [{'label': '🔍 TODOS LOS CLIENTES', 'value': 'TODOS'}]
        for _, r in df_cli_unique.iterrows():
            cli_code = str(r['Cliente'])
            cli_name = str(r['Razon Social'])
            lbl = f"{cli_code} - {cli_name}" if cli_code != cli_name else cli_code
            opts_clientes.append({'label': lbl, 'value': cli_code})

        return df.to_dict('records'), dbc.Alert(f"Cargado correctamente: {filename}", color="success", className="py-1 px-2 small mb-0"), opts_vendedores, opts_localidades, opts_clientes

    except Exception as e:
        return None, dbc.Alert(f"Error procesando archivo: {str(e)}", color="danger", className="py-1 px-2 small"), default_vends, default_locs, default_clis


@app.callback(
    Output('vendedor-select', 'value', allow_duplicate=True),
    Input('bar-chart-vendedor', 'clickData'),
    State('vendedor-select', 'value'),
    prevent_initial_call=True
)
def drilldown_vendedor(clickData, current_vendedor):
    if clickData and 'points' in clickData:
        clicked_vend = clickData['points'][0].get('y') or clickData['points'][0].get('x')
        if clicked_vend and clicked_vend != current_vendedor:
            return clicked_vend
    return dash.no_update


def filtrar_dataframe(records, vendedor_sel, localidad_sel, cliente_sel):
    if not records:
        return pd.DataFrame()
    df = pd.DataFrame(records)
    if vendedor_sel and vendedor_sel != 'TODOS':
        df = df[df["Vendedor"].astype(str) == str(vendedor_sel)]
    if localidad_sel and localidad_sel != 'TODOS':
        df = df[df["Localidad"].astype(str) == str(localidad_sel)]
    if cliente_sel and cliente_sel != 'TODOS':
        df = df[df["Cliente"].astype(str) == str(cliente_sel)]
    return df


@app.callback(
    Output('kpi-cards-row', 'children'),
    Input('stored-data', 'data'),
    Input('vendedor-select', 'value'),
    Input('localidad-select', 'value'),
    Input('cliente-select', 'value')
)
def render_kpis(records, vendedor_sel, localidad_sel, cliente_sel):
    df = filtrar_dataframe(records, vendedor_sel, localidad_sel, cliente_sel)
    if df.empty:
        return []

    total_cartera = df["Saldo Deuda"].sum()
    df_critico = df[df["Días de Atraso"] > 75]
    total_critico = df_critico["Saldo Deuda"].sum()
    pct_critico = (total_critico / total_cartera * 100) if total_cartera > 0 else 0.0
    
    total_clientes = len(df["Cliente"].unique())
    clientes_mayores_60 = len(df[df["Días de Atraso"] > 60]["Cliente"].unique())
    pct_clientes_60 = (clientes_mayores_60 / total_clientes * 100) if total_clientes > 0 else 0.0

    if total_cartera > 0:
        atraso_ponderado = (df["Días de Atraso"] * df["Saldo Deuda"]).sum() / total_cartera
    else:
        atraso_ponderado = 0.0

    return [
        dbc.Col([
            html.Div([
                html.Div([html.Span("Cartera Total Gestionada", className="kpi-title"), html.I(className="bi bi-wallet2 text-info fs-4")], className="d-flex justify-content-between align-items-center mb-2"),
                html.Div(format_currency_human(total_cartera), className="kpi-value text-info", id="kpi-1-val"),
                dbc.Tooltip(format_currency_full(total_cartera), target="kpi-1-val", placement="bottom"),
                html.Span("Saldo total acumulado en cartera", className="text-secondary small")
            ], className="p-3 exec-card h-100")
        ], xs=12, md=3),

        dbc.Col([
            html.Div([
                html.Div([html.Span("Deuda Crítica (>75 Días)", className="kpi-title"), html.I(className="bi bi-exclamation-triangle text-danger fs-4")], className="d-flex justify-content-between align-items-center mb-2"),
                html.Div(format_currency_human(total_critico), className="kpi-value text-danger", id="kpi-2-val"),
                dbc.Tooltip(format_currency_full(total_critico), target="kpi-2-val", placement="bottom"),
                html.Div([dbc.Badge(f"{pct_critico:.1f}% del total", color="danger", className="me-2"), html.Span("Riesgo alto", className="text-secondary small")])
            ], className="p-3 exec-card h-100")
        ], xs=12, md=3),

        dbc.Col([
            html.Div([
                html.Div([html.Span("Antigüedad Promedio Mora", className="kpi-title"), html.I(className="bi bi-clock-history text-warning fs-4")], className="d-flex justify-content-between align-items-center mb-2"),
                html.Div(f"{atraso_ponderado:.1f} días", className="kpi-value text-warning"),
                html.Span("Promedio ponderado por deuda", className="text-secondary small")
            ], className="p-3 exec-card h-100")
        ], xs=12, md=3),

        dbc.Col([
            html.Div([
                html.Div([html.Span("Clientes > 60 Días", className="kpi-title"), html.I(className="bi bi-person-exclamation text-danger fs-4")], className="d-flex justify-content-between align-items-center mb-2"),
                html.Div(f"{clientes_mayores_60:,}", className="kpi-value text-danger"),
                html.Div([dbc.Badge(f"{pct_clientes_60:.1f}% del total", color="danger", className="me-2"), html.Span(f"De {total_clientes} clientes totales", className="text-secondary small")])
            ], className="p-3 exec-card h-100")
        ], xs=12, md=3)
    ]


@app.callback(
    Output('tab-content-area', 'children'),
    Input('tabs-main', 'active_tab'),
    Input('stored-data', 'data'),
    Input('vendedor-select', 'value'),
    Input('localidad-select', 'value'),
    Input('cliente-select', 'value')
)
def render_tab_content(active_tab, records, vendedor_sel, localidad_sel, cliente_sel):
    if not records:
        return html.Div([
            dbc.Alert("Por favor, suba una planilla de Excel para habilitar los reportes visuales.", color="dark", className="text-center py-4 border border-secondary")
        ])

    df = filtrar_dataframe(records, vendedor_sel, localidad_sel, cliente_sel)
    vendedor_actual = vendedor_sel if vendedor_sel else 'TODOS'

    cliente_header = None
    if cliente_sel and cliente_sel != 'TODOS' and not df.empty:
        r_cli = df.iloc[0]
        limite_val = r_cli.get('Limite Credito', 0.0)
        dias_val = r_cli.get('Dias en Calle', 0.0)
        cliente_header = dbc.Alert([
            html.Div([
                html.Div([
                    html.I(className="bi bi-person-check-fill me-2 fs-5"),
                    html.Strong(f"Cliente Seleccionado: [{r_cli['Cliente']}] {r_cli['Razon Social']}"),
                    html.Span(f" — Comprobantes activos: {len(df)}", className="ms-2 small text-muted")
                ], className="mb-2"),
                dbc.Row([
                    dbc.Col([
                        html.Span("💳 Límite de Crédito: ", className="text-secondary small"),
                        html.Strong(f"${limite_val:,.2f}", className="text-info")
                    ], xs=12, md=6),
                    dbc.Col([
                        html.Span("⏱️ Días en Calle: ", className="text-secondary small"),
                        html.Strong(f"{dias_val:,.1f} días", className="text-warning")
                    ], xs=12, md=6)
                ])
            ])
        ], color="primary", className="py-2 px-3 mb-3")

    if active_tab == "tab-graficos":
        todos_tramos = ["Menos de 60 días", "61-75 Días", "76-90 Días", "Mayor a 90 Días"]
        resumen_tramos = df.groupby("Tramo Morosidad")["Saldo Deuda"].sum().reset_index()

        df_tabla_ver = pd.DataFrame({"Tramo de Vencimiento": todos_tramos})
        df_tabla_ver = df_tabla_ver.merge(resumen_tramos, left_on="Tramo de Vencimiento", right_on="Tramo Morosidad", how="left").fillna(0)
        df_tabla_ver.rename(columns={"Saldo Deuda": "Monto Total"}, inplace=True)
        total_v = df_tabla_ver["Monto Total"].sum()
        df_tabla_ver["Porcentaje"] = (df_tabla_ver["Monto Total"] / total_v * 100) if total_v > 0 else 0

        colors_map = {
            "Menos de 60 días": "#10b981",
            "61-75 Días": "#f59e0b",
            "76-90 Días": "#f97316",
            "Mayor a 90 Días": "#ef4444"
        }

        fig_donut = px.pie(
            df_tabla_ver[df_tabla_ver["Monto Total"] > 0],
            names="Tramo de Vencimiento",
            values="Monto Total",
            hole=0.5,
            color="Tramo de Vencimiento",
            color_discrete_map=colors_map
        )
        fig_donut.update_traces(textposition='inside', textinfo='percent+label')
        fig_donut.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font=dict(color='#f8fafc'), margin=dict(t=20, b=20, l=20, r=20), showlegend=False)

        cant_vendedores_unicos = pd.DataFrame(records)["Vendedor"].nunique()
        if vendedor_actual == 'TODOS' and (not cliente_sel or cliente_sel == 'TODOS') and cant_vendedores_unicos > 1:
            df_vend_pivot = pd.DataFrame(records).pivot_table(index="Vendedor", columns="Tramo Morosidad", values="Saldo Deuda", aggfunc="sum", fill_value=0).reset_index()
            for t in todos_tramos:
                if t not in df_vend_pivot.columns: df_vend_pivot[t] = 0
            
            fig_bar = px.bar(
                df_vend_pivot, 
                y="Vendedor", 
                x=todos_tramos, 
                title="Performance por Vendedor (Haga clic en una barra para filtrar)", 
                barmode='stack',
                color_discrete_map=colors_map
            )
            fig_bar.update_layout(xaxis_title="Monto Total ($)", yaxis_title="Vendedor", legend_title="Tramo Morosidad", clickmode='event+select')
        else:
            fig_bar = px.bar(df_tabla_ver, x="Tramo de Vencimiento", y="Monto Total", title="Distribución de Deuda por Tramo de Vencimiento", color="Tramo de Vencimiento", color_discrete_map=colors_map)

        fig_bar.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font=dict(color='#f8fafc'), margin=dict(t=40, b=20, l=20, r=40))

        df_temp = df.copy()
        df_temp['Razon Social'] = df_temp['Razon Social'].astype(str).str.strip()
        
        df_top20 = df_temp.groupby('Razon Social')['Saldo Deuda'].sum().reset_index()
        df_top20.rename(columns={'Razon Social': 'Cliente_Label'}, inplace=True)
        df_top20 = df_top20.sort_values(by='Saldo Deuda', ascending=False).head(20)
        df_top20 = df_top20.sort_values(by='Saldo Deuda', ascending=True)
        
        fig_top20 = px.bar(
            df_top20,
            x="Saldo Deuda",
            y="Cliente_Label",
            orientation='h',
            title="Top 20 Clientes con Mayor Deuda Acumulada",
            color="Saldo Deuda",
            color_continuous_scale="Reds",
            text_auto='.2s'
        )
        fig_top20.update_layout(
            paper_bgcolor='rgba(0,0,0,0)', 
            plot_bgcolor='rgba(0,0,0,0)', 
            font=dict(color='#f8fafc'), 
            margin=dict(t=40, b=20, l=20, r=20),
            coloraxis_showscale=False,
            xaxis_title="Deuda Total ($)",
            yaxis_title=""
        )

        return html.Div([
            cliente_header if cliente_header else html.Div(),
            dbc.Row([
                dbc.Col([
                    html.Div([
                        html.H5([html.I(className="bi bi-pie-chart-fill text-info me-2"), "Distribución de Cartera por Tramo de Morosidad"], className="fw-bold text-light mb-3"),
                        dcc.Graph(figure=fig_donut, config={'displayModeBar': False})
                    ], className="p-3 exec-card mb-4")
                ], xs=12, lg=5),
                dbc.Col([
                    html.Div([
                        dcc.Graph(id='bar-chart-vendedor', figure=fig_bar, config={'displayModeBar': False})
                    ], className="p-3 exec-card mb-4")
                ], xs=12, lg=7)
            ]),
            dbc.Row([
                dbc.Col([
                    html.Div([
                        dcc.Graph(figure=fig_top20, config={'displayModeBar': False})
                    ], className="p-3 exec-card")
                ], width=12)
            ])
        ])

    elif active_tab == "tab-top-clientes":
        df_temp = df.copy()
        df_temp['Razon Social'] = df_temp['Razon Social'].astype(str).str.strip()
        
        df_top_all = df_temp.groupby('Razon Social')['Saldo Deuda'].sum().reset_index()
        df_top_all.rename(columns={'Razon Social': 'Cliente_Label'}, inplace=True)
        df_top_all = df_top_all.sort_values(by='Saldo Deuda', ascending=False)
        
        df_top_plot = df_top_all.head(20).sort_values(by='Saldo Deuda', ascending=True)
        
        fig_top_ranking = px.bar(
            df_top_plot,
            x="Saldo Deuda",
            y="Cliente_Label",
            orientation='h',
            title="Ranking de Clientes con Mayor Deuda Acumulada",
            color="Saldo Deuda",
            color_continuous_scale="Reds"
        )
        fig_top_ranking.update_layout(
            paper_bgcolor='rgba(0,0,0,0)', 
            plot_bgcolor='rgba(0,0,0,0)', 
            font=dict(color='#f8fafc'), 
            margin=dict(t=40, b=20, l=20, r=40),
            coloraxis_showscale=False,
            xaxis_title="Deuda Total ($)",
            yaxis_title="",
            yaxis=dict(type='category'),
            height=600
        )
        fig_top_ranking.update_traces(texttemplate='%{x:$,.2f}', textposition='outside')

        column_defs_ranking = [
            {"field": "Cliente_Label", "headerName": "Razón Social", "filter": True, "flex": 2, "minWidth": 200},
            {"field": "Saldo Deuda", "headerName": "Deuda Total ($)", "filter": "agNumberColumnFilter", "sortable": True, "valueFormatter": {"function": "d3.format('$,.2f')(params.value)"}, "flex": 1, "minWidth": 150}
        ]

        return html.Div([
            cliente_header if cliente_header else html.Div(),
            dbc.Row([
                dbc.Col([
                    html.Div([
                        html.H5([html.I(className="bi bi-trophy-fill text-warning me-2"), "Ranking Detallado de Cartera por Cliente"], className="fw-bold text-light mb-3"),
                        dcc.Graph(figure=fig_top_ranking, config={'displayModeBar': False})
                    ], className="p-3 exec-card mb-4")
                ], width=12)
            ]),
            dbc.Row([
                dbc.Col([
                    html.Div([
                        html.H6([html.I(className="bi bi-list-columns-reverse text-info me-2"), "Tabla Consolidada de Deuda por Cliente"], className="fw-bold text-light mb-3"),
                        dag.AgGrid(
                            id="grid-ranking-clientes",
                            rowData=df_top_all.to_dict("records"),
                            columnDefs=column_defs_ranking,
                            defaultColDef={"resizable": True, "sortable": True, "filter": True},
                            className="ag-theme-alpine-dark",
                            style={"height": "400px", "width": "100%"},
                            dashGridOptions={
                                "pagination": True,
                                "paginationPageSize": 10,
                                "enableBrowserTooltips": True
                            }
                        )
                    ], className="p-3 exec-card")
                ], width=12)
            ])
        ])

    elif active_tab == "tab-criticos":
        df_crit = df[df["Días de Atraso"] > 75].sort_values(by="Días de Atraso", ascending=False)
        
        column_defs_criticos = []
        for col in df.columns:
            c_def = {
                "field": col, 
                "headerName": col, 
                "filter": True, 
                "sortable": True, 
                "flex": 1, 
                "minWidth": 130
            }
            if col in ["Saldo Deuda", "Limite Credito"]:
                c_def["valueFormatter"] = {"function": "d3.format('$,.2f')(params.value)"}
                c_def["filter"] = "agNumberColumnFilter"
            elif col == "Días de Atraso":
                c_def["filter"] = "agNumberColumnFilter"
                c_def["cellClassRules"] = {
                    "grid-cell-red": "params.value > 90",
                    "grid-cell-orange": "params.value > 75 && params.value <= 90"
                }
            column_defs_criticos.append(c_def)

        return html.Div([
            cliente_header if cliente_header else html.Div(),
            html.Div([
                html.Div([
                    html.H5([html.I(className="bi bi-exclamation-octagon text-danger me-2"), f"Casos Críticos de Morosidad (>75 Días) - Total: {len(df_crit)} registros"], className="fw-bold text-light mb-0"),
                    dbc.Button([html.I(className="bi bi-file-earmark-excel me-2"), "Exportar Críticos Excel"], id="btn-export-criticos", color="danger", size="sm", n_clicks=0)
                ], className="d-flex justify-content-between align-items-center mb-3"),

                dag.AgGrid(
                    id="grid-criticos",
                    rowData=df_crit.to_dict("records"),
                    columnDefs=column_defs_criticos,
                    defaultColDef={"resizable": True, "sortable": True, "filter": True},
                    className="ag-theme-alpine-dark",
                    style={"height": "auto", "width": "100%"},
                    dashGridOptions={
                        "pagination": True,
                        "paginationPageSize": 15,
                        "domLayout": "autoHeight",
                        "enableBrowserTooltips": True
                    }
                )
            ], className="p-3 exec-card")
        ])

    elif active_tab == "tab-matriz":
        column_defs_matriz = []
        for col in df.columns:
            c_def = {
                "field": col, 
                "headerName": col, 
                "filter": True, 
                "sortable": True, 
                "flex": 1, 
                "minWidth": 130
            }
            if col in ["Saldo Deuda", "Limite Credito"]:
                c_def["valueFormatter"] = {"function": "d3.format('$,.2f')(params.value)"}
                c_def["filter"] = "agNumberColumnFilter"
            elif col == "Días de Atraso":
                c_def["filter"] = "agNumberColumnFilter"
                c_def["cellClassRules"] = {
                    "grid-cell-red": "params.value > 90",
                    "grid-cell-orange": "params.value > 75 && params.value <= 90",
                    "grid-cell-yellow": "params.value > 60 && params.value <= 75",
                    "grid-cell-green": "params.value <= 60"
                }
            column_defs_matriz.append(c_def)

        return html.Div([
            cliente_header if cliente_header else html.Div(),
            html.Div([
                html.Div([
                    html.H5([html.I(className="bi bi-table text-info me-2"), "Matriz General de Cuentas Corrientes"], className="fw-bold text-light mb-0"),
                    dbc.Button([html.I(className="bi bi-file-earmark-excel me-2"), "Exportar Matriz Excel"], id="btn-export-matriz", color="success", size="sm", n_clicks=0)
                ], className="d-flex justify-content-between align-items-center mb-3"),

                dag.AgGrid(
                    id="grid-matriz",
                    rowData=df.to_dict("records"),
                    columnDefs=column_defs_matriz,
                    defaultColDef={"resizable": True, "sortable": True, "filter": True},
                    className="ag-theme-alpine-dark",
                    style={"height": "auto", "width": "100%"},
                    dashGridOptions={
                        "pagination": True,
                        "paginationPageSize": 15,
                        "domLayout": "autoHeight",
                        "enableBrowserTooltips": True
                    }
                )
            ], className="p-3 exec-card")
        ])

    elif active_tab == "tab-dinamica":
        todos_tramos = ["Menos de 60 días", "61-75 Días", "76-90 Días", "Mayor a 90 Días"]
        
        df_pivot = df.pivot_table(
            index=["Cliente", "Razon Social"],
            columns="Tramo Morosidad",
            values="Saldo Deuda",
            aggfunc="sum",
            fill_value=0
        ).reset_index()

        for t in todos_tramos:
            if t not in df_pivot.columns:
                df_pivot[t] = 0.0

        cols_ordenadas = ["Cliente", "Razon Social"] + todos_tramos
        df_pivot = df_pivot[[c for c in cols_ordenadas if c in df_pivot.columns]]
        df_pivot["Total General"] = df_pivot[todos_tramos].sum(axis=1)

        column_defs_dinamica = [
            {"field": "Cliente", "headerName": "Código Cliente", "filter": True, "sortable": True, "flex": 1, "minWidth": 130},
            {"field": "Razon Social", "headerName": "Razón Social", "filter": True, "sortable": True, "flex": 2, "minWidth": 200},
            {"field": "Menos de 60 días", "headerName": "Menos de 60 días", "filter": "agNumberColumnFilter", "sortable": True, "valueFormatter": {"function": "d3.format('$,.2f')(params.value)"}, "flex": 1, "minWidth": 140},
            {"field": "61-75 Días", "headerName": "61-75 Días", "filter": "agNumberColumnFilter", "sortable": True, "valueFormatter": {"function": "d3.format('$,.2f')(params.value)"}, "flex": 1, "minWidth": 130},
            {"field": "76-90 Días", "headerName": "76-90 Días", "filter": "agNumberColumnFilter", "sortable": True, "valueFormatter": {"function": "d3.format('$,.2f')(params.value)"}, "flex": 1, "minWidth": 130},
            {"field": "Mayor a 90 Días", "headerName": "Mayor a 90 Días", "filter": "agNumberColumnFilter", "sortable": True, "valueFormatter": {"function": "d3.format('$,.2f')(params.value)"}, "flex": 1, "minWidth": 140},
            {"field": "Total General", "headerName": "Total General", "filter": "agNumberColumnFilter", "sortable": True, "valueFormatter": {"function": "d3.format('$,.2f')(params.value)"}, "flex": 1, "minWidth": 140}
        ]

        return html.Div([
            cliente_header if cliente_header else html.Div(),
            html.Div([
                html.Div([
                    html.H5([html.I(className="bi bi-bar-chart-steps text-info me-2"), "Dinámica por Tramo (Matriz por Cliente)"], className="fw-bold text-light mb-0"),
                    dbc.Button([html.I(className="bi bi-file-earmark-excel me-2"), "Exportar Dinámica Excel"], id="btn-export-dinamica", color="success", size="sm", n_clicks=0)
                ], className="d-flex justify-content-between align-items-center mb-3"),

                dag.AgGrid(
                    id="grid-dinamica",
                    rowData=df_pivot.to_dict("records"),
                    columnDefs=column_defs_dinamica,
                    defaultColDef={"resizable": True, "sortable": True, "filter": True},
                    className="ag-theme-alpine-dark",
                    style={"height": "auto", "width": "100%"},
                    dashGridOptions={
                        "pagination": True,
                        "paginationPageSize": 15,
                        "domLayout": "autoHeight",
                        "enableBrowserTooltips": True
                    }
                )
            ], className="p-3 exec-card")
        ])

    return html.Div()


@app.callback(
    Output("download-criticos-excel", "data"),
    Input("btn-export-criticos", "n_clicks"),
    State('stored-data', 'data'),
    State('vendedor-select', 'value'),
    State('localidad-select', 'value'),
    State('cliente-select', 'value'),
    prevent_initial_call=True
)
def export_criticos(n_clicks, records, vendedor_sel, localidad_sel, cliente_sel):
    if not n_clicks or not records:
        return dash.no_update
    df = filtrar_dataframe(records, vendedor_sel, localidad_sel, cliente_sel)
    df_crit = df[df["Días de Atraso"] > 75].sort_values(by="Días de Atraso", ascending=False)
    if df_crit.empty:
        df_crit = df
    return dcc.send_data_frame(df_crit.to_excel, "casos_criticos_morosidad.xlsx", sheet_name="Criticos", index=False)


@app.callback(
    Output("download-matriz-excel", "data"),
    Input("btn-export-matriz", "n_clicks"),
    State('stored-data', 'data'),
    State('vendedor-select', 'value'),
    State('localidad-select', 'value'),
    State('cliente-select', 'value'),
    prevent_initial_call=True
)
def export_matriz(n_clicks, records, vendedor_sel, localidad_sel, cliente_sel):
    if not n_clicks or not records:
        return dash.no_update
    df = filtrar_dataframe(records, vendedor_sel, localidad_sel, cliente_sel)
    return dcc.send_data_frame(df.to_excel, "matriz_cuentas_corrientes.xlsx", sheet_name="Matriz", index=False)


@app.callback(
    Output("download-dinamica-excel", "data"),
    Input("btn-export-dinamica", "n_clicks"),
    State('stored-data', 'data'),
    State('vendedor-select', 'value'),
    State('localidad-select', 'value'),
    State('cliente-select', 'value'),
    prevent_initial_call=True
)
def export_dinamica(n_clicks, records, vendedor_sel, localidad_sel, cliente_sel):
    if not n_clicks or not records:
        return dash.no_update
    df = filtrar_dataframe(records, vendedor_sel, localidad_sel, cliente_sel)
    
    todos_tramos = ["Menos de 60 días", "61-75 Días", "76-90 Días", "Mayor a 90 Días"]
    
    df_pivot = df.pivot_table(
        index=["Cliente", "Razon Social"],
        columns="Tramo Morosidad",
        values="Saldo Deuda",
        aggfunc="sum",
        fill_value=0
    ).reset_index()

    for t in todos_tramos:
        if t not in df_pivot.columns:
            df_pivot[t] = 0.0

    cols_ordenadas = ["Cliente", "Razon Social"] + todos_tramos
    df_pivot = df_pivot[[c for c in cols_ordenadas if c in df_pivot.columns]]
    df_pivot["Total General"] = df_pivot[todos_tramos].sum(axis=1)

    return dcc.send_data_frame(df_pivot.to_excel, "dinamica_por_tramo.xlsx", sheet_name="Dinamica", index=False)
    

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port, debug=False)