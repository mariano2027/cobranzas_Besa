import os
import glob
import re
import dash
from dash import dcc, html, Input, Output, State, ctx
import dash_bootstrap_components as dbc
import dash_ag_grid as dag
import pandas as pd
import plotly.express as px

# Inicialización de la aplicación Dash
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

server = app.server

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

CREDENCIALES_USUARIOS = {
    "Administrador": "admin123",
    "Claudio": "claudio2026",
    "Miguel": "miguel2026",
    "Sala": "sala2026",
    "Vellini": "vellini2026",
    "Fede": "fede2026",
    "Ana": "ana2026",
    "JP Sexto": "jp2026",
    "Paola": "paola2026"
}

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

def cargar_excel_automatico():
    archivos_xlsx = glob.glob("*.xlsx")
    if not archivos_xlsx:
        raise FileNotFoundError("No se encontró ningún archivo .xlsx en la carpeta del script.")
    
    ruta_archivo = archivos_xlsx[0]
    df = pd.read_excel(ruta_archivo, engine="openpyxl")
    
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

    col_localidad = next((c for c in df.columns if any(k in c.lower() for k in ['localidad', 'ciudad', 'municipio'])), None)
    if col_localidad:
        df['Localidad'] = df[col_localidad].fillna("Sin Localidad").astype(str).str.strip()
        if col_localidad != 'Localidad':
            df.drop(columns=[col_localidad], inplace=True)
            df.rename(columns={col_localidad: 'Localidad'}, inplace=True)
    else:
        df['Localidad'] = "General"

    df['Cliente'] = df['Cliente'].astype(str).str.strip()
    df['Razon Social'] = df['Razon Social'].astype(str).str.strip()

    if 'Nro Comprobante' not in df.columns: df['Nro Comprobante'] = '-'
    else: df['Nro Comprobante'] = df['Nro Comprobante'].astype(str).str.strip()

    if 'Fecha Emisión' not in df.columns: df['Fecha Emisión'] = '-'

    col_importe = next((c for c in df.columns if any(k in c.lower() for k in ['importe', 'imp.', 'total'])), None)
    s_importe = clean_currency_series(df[col_importe]) if col_importe else pd.Series(0.0, index=df.index)
    
    df['Importe Original'] = s_importe

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
    return df, ruta_archivo

try:
    df_global, nombre_archivo_encontrado = cargar_excel_automatico()
    datos_globales_dict = df_global.to_dict('records')
    error_carga = None
except Exception as e:
    datos_globales_dict = []
    error_carga = str(e)

def crear_layout_login():
    return dbc.Container([
        dbc.Row([
            dbc.Col([
                html.Div([
                    html.I(className="bi bi-shield-lock-fill text-info display-4 mb-3"),
                    html.H3("Control de Acceso - Dashboard", className="fw-bold text-white mb-2"),
                    html.P("Seleccione su usuario e ingrese su contraseña personal.", className="text-secondary mb-4"),
                    
                    html.Div([
                        html.Label("Usuario:", className="fw-semibold text-light mb-2 text-start d-block"),
                        dcc.Dropdown(
                            id='login-usuario-select',
                            options=[{'label': u, 'value': u} for u in CREDENCIALES_USUARIOS.keys()],
                            placeholder="Seleccione usuario...",
                            style={'color': '#0f172a'}
                        ),
                    ], className="mb-3 text-start"),

                    html.Div([
                        html.Label("Contraseña:", className="fw-semibold text-light mb-2 text-start d-block"),
                        dbc.Input(
                            id='login-password-input',
                            type="password",
                            placeholder="Ingrese su contraseña...",
                            className="bg-dark text-white border-secondary"
                        ),
                    ], className="mb-4 text-start"),

                    dbc.Button("Ingresar al Sistema", id="btn-login", color="info", className="w-100 fw-bold py-2"),
                    html.Div(id="login-error-msg", className="mt-3")
                ], className="exec-card p-5 text-center shadow-lg")
            ], xs=12, md=6, lg=4)
        ], className="vh-100 align-items-center justify-content-center")
    ], fluid=True)

def crear_layout_dashboard(usuario_actual):
    vendedores = sorted([str(v) for v in df_global["Vendedor"].dropna().unique() if str(v).strip() != '']) if not df_global.empty else []
    opts_vendedores = [{'label': '🌐 TODOS LOS VENDEDORES', 'value': 'TODOS'}] + [{'label': f"👤 {v}", 'value': v} for v in vendedores]

    localidades = sorted([str(l) for l in df_global["Localidad"].dropna().unique() if str(l).strip() != '']) if not df_global.empty else []
    opts_localidades = [{'label': '📍 TODAS LAS LOCALIDADES', 'value': 'TODOS'}] + [{'label': f"📍 {l}", 'value': l} for l in localidades]

    opts_clientes = [{'label': '🔍 TODOS LOS CLIENTES', 'value': 'TODOS'}]
    if not df_global.empty:
        df_cli_unique = df_global[['Cliente', 'Razon Social']].drop_duplicates().sort_values(by='Razon Social')
        for _, r in df_cli_unique.iterrows():
            cli_code = str(r['Cliente'])
            cli_name = str(r['Razon Social'])
            lbl = f"{cli_code} - {cli_name}" if cli_code != cli_name else cli_code
            opts_clientes.append({'label': lbl, 'value': cli_code})

    val_vendedor_inicial = usuario_actual if usuario_actual != 'Administrador' else 'TODOS'
    vendedor_deshabilitado = (usuario_actual != 'Administrador')

    modal_detalle_tramo = dbc.Modal([
        dbc.ModalHeader(dbc.ModalTitle("Clientes en el Tramo", id="modal-tramo-titulo"), className="bg-dark text-white border-secondary"),
        dbc.ModalBody([
            html.Div(id="modal-tramo-contenido")
        ], style={'backgroundColor': '#162032'}),
        dbc.ModalFooter(
            dbc.Button("Cerrar", id="btn-cerrar-modal", color="secondary", className="fw-bold")
        )
    ], id="modal-tramo", size="xl", is_open=False, centered=True, backdrop="static")

    return dbc.Container([
        dcc.Store(id='stored-data', data=datos_globales_dict),
        dcc.Store(id='session-usuario', data=usuario_actual),
        modal_detalle_tramo,

        dbc.Row([
            dbc.Col([
                html.Div([
                    html.Div([
                        html.H1([html.I(className="bi bi-graph-up-arrow me-3 text-info"), "Dashboard Ejecutivo de Cuentas Corrientes"], className="fw-bold text-white mb-1 h2"),
                        html.P(f"Archivo analizado: {nombre_archivo_encontrado if not error_carga else 'Error'} — Sesión Activa: {usuario_actual}", className="text-secondary mb-0 small")
                    ]),
                    dbc.Button([html.I(className="bi bi-box-arrow-left me-1"), "Cerrar Sesión"], id="btn-logout", color="outline-danger", size="sm")
                ], className="py-3 px-2 d-flex justify-content-between align-items-center")
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
                            html.Label("Filtrar Vendedor:", className="fw-semibold text-light mb-1 small"),
                            dcc.Dropdown(
                                id='vendedor-select',
                                options=opts_vendedores,
                                value=val_vendedor_inicial,
                                disabled=vendedor_deshabilitado,
                                clearable=False,
                                style={'color': '#0f172a'}
                            )
                        ], xs=12, lg=4, className="mb-3 mb-lg-0"),

                        dbc.Col([
                            html.Label("Filtrar Localidad:", className="fw-semibold text-light mb-1 small"),
                            dcc.Dropdown(
                                id='localidad-select',
                                options=opts_localidades,
                                value='TODOS',
                                clearable=False,
                                style={'color': '#0f172a'}
                            )
                        ], xs=12, lg=4, className="mb-3 mb-lg-0"),

                        dbc.Col([
                            html.Label("Filtrar Cliente / Razón Social:", className="fw-semibold text-light mb-1 small"),
                            dcc.Dropdown(
                                id='cliente-select',
                                options=opts_clientes,
                                value='TODOS',
                                clearable=False,
                                placeholder="Escribe o selecciona cliente...",
                                style={'color': '#0f172a'}
                            )
                        ], xs=12, lg=4)
                    ])
                ], className="p-3 exec-card")
            ], width=12, className="mb-4")
        ]),

        dbc.Row([
            dbc.Col([
                dbc.Row(id='kpi-cards-row', className="g-3 mb-4"),

                dbc.Tabs(id="tabs-main", active_tab="tab-graficos", children=[
                    dbc.Tab(label="📊 Resumen General & Gráficos", tab_id="tab-graficos"),
                    dbc.Tab(label="🗺️ Mapa por Zonas", tab_id="tab-zonas"),
                    dbc.Tab(label="🏆 Ranking Top Clientes", tab_id="tab-top-clientes"),
                    dbc.Tab(label="🚨 Morosidad Crítica (>75 Días)", tab_id="tab-criticos"),
                    dbc.Tab(label="📌 Ctas Corrientes", tab_id="tab-matriz"),
                    dbc.Tab(label="📈 Dinámica por Tramo", tab_id="tab-dinamica")
                ], className="border-bottom border-secondary mb-3"),

                html.Div(id="tab-content-area")
            ], width=12)
        ])
    ], fluid=True, className="p-3 p-md-4")

app.layout = html.Div([
    dcc.Store(id='auth-store', data=None),
    html.Div(id='page-content')
])

@app.callback(
    Output('auth-store', 'data'),
    Output('login-error-msg', 'children'),
    Input('btn-login', 'n_clicks'),
    State('login-usuario-select', 'value'),
    State('login-password-input', 'value'),
    prevent_initial_call=True
)
def procesar_login(n_clicks, usuario, password):
    if n_clicks:
        if not usuario:
            return None, dbc.Alert("Por favor seleccione un usuario.", color="danger", className="py-1 px-2 small")
        if not password:
            return None, dbc.Alert("Por favor ingrese su contraseña.", color="danger", className="py-1 px-2 small")
        
        if usuario in CREDENCIALES_USUARIOS and CREDENCIALES_USUARIOS[usuario] == password:
            if error_carga:
                return None, dbc.Alert(f"Error cargando Excel: {error_carga}", color="danger", className="py-1 px-2 small")
            return usuario, None
        else:
            return None, dbc.Alert("Contraseña incorrecta.", color="danger", className="py-1 px-2 small")
            
    return dash.no_update, None

@app.callback(
    Output('page-content', 'children'),
    Input('auth-store', 'data')
)
def render_pantalla(usuario_autenticado):
    if not usuario_autenticado:
        return crear_layout_login()
    return crear_layout_dashboard(usuario_autenticado)

def filtrar_dataframe(records, usuario_sesion, vendedor_sel, localidad_sel, cliente_sel):
    if not records:
        return pd.DataFrame()
    df = pd.DataFrame(records)
    
    if usuario_sesion != 'Administrador':
        df = df[df["Vendedor"].astype(str).str.lower() == str(usuario_sesion).lower()]
    elif vendedor_sel and vendedor_sel != 'TODOS':
        df = df[df["Vendedor"].astype(str).str.lower() == str(vendedor_sel).lower()]
        
    if localidad_sel and localidad_sel != 'TODOS':
        df = df[df["Localidad"].astype(str) == str(localidad_sel)]
    if cliente_sel and cliente_sel != 'TODOS':
        df = df[df["Cliente"].astype(str) == str(cliente_sel)]
    return df

@app.callback(
    Output('auth-store', 'data', allow_duplicate=True),
    Input('btn-logout', 'n_clicks'),
    prevent_initial_call=True
)
def cerrar_sesion(n_clicks):
    if n_clicks:
        return None
    return dash.no_update

@app.callback(
    Output('vendedor-select', 'value', allow_duplicate=True),
    Output('localidad-select', 'value', allow_duplicate=True),
    Output('cliente-select', 'value', allow_duplicate=True),
    Input('btn-reset-filters', 'n_clicks'),
    State('session-usuario', 'data'),
    prevent_initial_call=True
)
def reset_filters(n_clicks, usuario_sesion):
    if n_clicks:
        val_vend = usuario_sesion if usuario_sesion != 'Administrador' else 'TODOS'
        return val_vend, 'TODOS', 'TODOS'
    return dash.no_update, dash.no_update, dash.no_update

@app.callback(
    Output('kpi-cards-row', 'children'),
    Input('stored-data', 'data'),
    Input('session-usuario', 'data'),
    Input('vendedor-select', 'value'),
    Input('localidad-select', 'value'),
    Input('cliente-select', 'value')
)
def render_kpis(records, usuario_sesion, vendedor_sel, localidad_sel, cliente_sel):
    df = filtrar_dataframe(records, usuario_sesion, vendedor_sel, localidad_sel, cliente_sel)
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
    Input('session-usuario', 'data'),
    Input('vendedor-select', 'value'),
    Input('localidad-select', 'value'),
    Input('cliente-select', 'value')
)
def render_tab_content(active_tab, records, usuario_sesion, vendedor_sel, localidad_sel, cliente_sel):
    df = filtrar_dataframe(records, usuario_sesion, vendedor_sel, localidad_sel, cliente_sel)
    
    if df.empty:
        return dbc.Alert("No hay datos cargados o que coincidan con los filtros seleccionados.", color="warning", className="mt-3 text-center")

    if active_tab == "tab-graficos":
        df_tramo = df.groupby("Tramo Morosidad", as_index=False)["Saldo Deuda"].sum()
        fig_tramo = px.pie(df_tramo, names="Tramo Morosidad", values="Saldo Deuda", title="Distribución de Cartera por Tramo de Morosidad", hole=0.4, template="plotly_dark")
        fig_tramo.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')

        df_vend_tramo = df.groupby(["Vendedor", "Tramo Morosidad"], as_index=False)["Saldo Deuda"].sum()
        fig_vend = px.bar(df_vend_tramo, x="Saldo Deuda", y="Vendedor", color="Tramo Morosidad", orientation='h', title="Performance por Vendedor (Haga clic en una barra para filtrar)", template="plotly_dark")
        fig_vend.update_layout(barmode='stack', plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')

        return dbc.Row([
            dbc.Col([html.Div([dcc.Graph(figure=fig_tramo)], className="p-3 exec-card")], xs=12, lg=6, className="mb-3"),
            dbc.Col([html.Div([dcc.Graph(figure=fig_vend, id='bar-chart-vendedor')], className="p-3 exec-card")], xs=12, lg=6, className="mb-3")
        ])

    elif active_tab == "tab-zonas":
        # Agregamos conteo de clientes únicos por Vendedor y Localidad
        df_zona = df.groupby(["Vendedor", "Localidad"], as_index=False).agg({
            "Saldo Deuda": "sum",
            "Cliente": "nunique"
        }).rename(columns={"Cliente": "Cant Clientes"})
        
        # Filtramos saldos mayores a 0 para que el Treemap no colapse con negativos o ceros
        df_zona_plot = df_zona[df_zona["Saldo Deuda"] > 0].copy()
        
        fig_mapa = px.treemap(
            df_zona_plot, 
            path=["Vendedor", "Localidad"], 
            values="Saldo Deuda", 
            custom_data=["Cant Clientes", "Saldo Deuda"],
            title="🗺️ Mapa Jerárquico de Zonas, Vendedores y Cantidad de Clientes",
            template="plotly_dark",
            color="Saldo Deuda",
            color_continuous_scale="Viridis"
        )
        fig_mapa.update_traces(
            hovertemplate="<b>%{label}</b><br>Volumen Deuda: $%{customdata[1]:,.2f}<br>Clientes Únicos: %{customdata[0]}<extra></extra>"
        )
        fig_mapa.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', margin=dict(t=50, l=25, r=25, b=25))

        df_localidad_bar = df.groupby("Localidad", as_index=False).agg({"Saldo Deuda": "sum", "Cliente": "nunique"}).sort_values(by="Saldo Deuda", ascending=True)
        fig_loc_bar = px.bar(
            df_localidad_bar, 
            x="Saldo Deuda", 
            y="Localidad", 
            orientation='h', 
            hover_data={"Cliente": True, "Saldo Deuda": ":$,.2f"},
            title="📍 Volumen de Dinero Total por Localidad",
            template="plotly_dark"
        )
        fig_loc_bar.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')

        return dbc.Row([
            dbc.Col([html.Div([dcc.Graph(figure=fig_mapa, style={"height": "500px"})], className="p-3 exec-card")], xs=12, lg=12, className="mb-3"),
            dbc.Col([html.Div([dcc.Graph(figure=fig_loc_bar, style={"height": "450px"})], className="p-3 exec-card")], xs=12, lg=12, className="mb-3")
        ])
    elif active_tab == "tab-top-clientes":
        df_top = df.groupby(["Cliente", "Razon Social", "Vendedor"], as_index=False)["Saldo Deuda"].sum().sort_values(by="Saldo Deuda", ascending=False).head(15)
        
        columnas_grid = [
            {"field": "Cliente", "headerName": "Código", "sortable": True, "filter": True, "width": 120},
            {"field": "Razon Social", "headerName": "Razón Social", "sortable": True, "filter": True, "flex": 2},
            {"field": "Vendedor", "headerName": "Vendedor", "sortable": True, "filter": True, "flex": 1},
            {"field": "Saldo Deuda", "headerName": "Saldo Total", "sortable": True, "filter": True, "type": "rightAligned", "valueFormatter": {"function": "d3.format('$,.2f')(params.value)"}, "flex": 1}
        ]
        
        return html.Div([
            html.H5("🏆 Top 15 Clientes con Mayor Deuda", className="text-white mb-3"),
            dag.AgGrid(
                id="grid-top-clientes",
                rowData=df_top.to_dict("records"),
                columnDefs=columnas_grid,
                className="ag-theme-alpine-dark",
                style={"height": "450px", "width": "100%"},
                columnSize="sizeToFit",
                dashGridOptions={"pagination": True, "paginationPageSize": 10}
            )
        ], className="p-3 exec-card")

    elif active_tab == "tab-criticos":
        df_crit = df[df["Días de Atraso"] > 75].copy()
        
        columnas_crit = [
            {"field": "Cliente", "headerName": "Código", "sortable": True, "filter": True, "width": 110},
            {"field": "Razon Social", "headerName": "Razón Social", "sortable": True, "filter": True, "flex": 2},
            {"field": "Nro Comprobante", "headerName": "Comprobante", "sortable": True, "filter": True, "flex": 1},
            {"field": "Fecha Emisión", "headerName": "Emisión", "sortable": True, "filter": True, "width": 110},
            {"field": "Días de Atraso", "headerName": "Atraso", "sortable": True, "filter": True, "width": 90},
            {"field": "Importe Original", "headerName": "Importe Original", "sortable": True, "filter": True, "type": "rightAligned", "valueFormatter": {"function": "d3.format('$,.2f')(params.value)"}, "flex": 1},
            {"field": "Saldo Deuda", "headerName": "Saldo Vencido", "sortable": True, "filter": True, "type": "rightAligned", "valueFormatter": {"function": "d3.format('$,.2f')(params.value)"}, "flex": 1},
            {"field": "Vendedor", "headerName": "Vendedor", "sortable": True, "filter": True, "flex": 1}
        ]

        return html.Div([
            html.Div([
                html.H5("🚨 Comprobantes en Morosidad Crítica (> 75 Días)", className="text-danger mb-0 fw-bold"),
            ], className="d-flex justify-content-between align-items-center mb-3"),
            dag.AgGrid(
                id="grid-criticos",
                rowData=df_crit.to_dict("records"),
                columnDefs=columnas_crit,
                className="ag-theme-alpine-dark",
                style={"height": "480px", "width": "100%"},
                columnSize="sizeToFit",
                dashGridOptions={"pagination": True, "paginationPageSize": 12}
            )
        ], className="p-3 exec-card")

    elif active_tab == "tab-matriz":
        columnas_matriz = [
            {"field": "Cliente", "headerName": "Código", "sortable": True, "filter": True, "width": 110},
            {"field": "Razon Social", "headerName": "Razón Social", "sortable": True, "filter": True, "flex": 2},
            {"field": "Nro Comprobante", "headerName": "Comprobante", "sortable": True, "filter": True, "flex": 1},
            {"field": "Fecha Emisión", "headerName": "Emisión", "sortable": True, "filter": True, "width": 110},
            {"field": "Días de Atraso", "headerName": "Atraso", "sortable": True, "filter": True, "width": 90},
            {"field": "Importe Original", "headerName": "Importe Original", "sortable": True, "filter": True, "type": "rightAligned", "valueFormatter": {"function": "d3.format('$,.2f')(params.value)"}, "flex": 1},
            {"field": "Saldo Deuda", "headerName": "Saldo", "sortable": True, "filter": True, "type": "rightAligned", "valueFormatter": {"function": "d3.format('$,.2f')(params.value)"}, "flex": 1},
            {"field": "Vendedor", "headerName": "Vendedor", "sortable": True, "filter": True, "flex": 1}
        ]

        return html.Div([
            html.H5("📌 Detalle Completo de Cuentas Corrientes", className="text-white mb-3"),
            dag.AgGrid(
                id="grid-matriz",
                rowData=df.to_dict("records"),
                columnDefs=columnas_matriz,
                className="ag-theme-alpine-dark",
                style={"height": "480px", "width": "100%"},
                columnSize="sizeToFit",
                dashGridOptions={"pagination": True, "paginationPageSize": 12}
            )
        ], className="p-3 exec-card")

    elif active_tab == "tab-dinamica":
        df_din = df.groupby("Tramo Morosidad", as_index=False).agg({"Saldo Deuda": "sum", "Cliente": "nunique"}).rename(columns={"Cliente": "Cant Clientes"})
        
        columnas_din = [
            {"field": "Tramo Morosidad", "headerName": "Tramo de Antigüedad", "sortable": True, "filter": True, "flex": 2},
            {"field": "Cant Clientes", "headerName": "Clientes Únicos", "sortable": True, "filter": True, "flex": 1},
            {"field": "Saldo Deuda", "headerName": "Deuda Total", "sortable": True, "filter": True, "type": "rightAligned", "valueFormatter": {"function": "d3.format('$,.2f')(params.value)"}, "flex": 1}
        ]

        return html.Div([
            html.H5("📈 Dinámica Agregada por Tramo (Haga clic en una fila para ver clientes)", className="text-white mb-3"),
            dag.AgGrid(
                id="grid-dinamica",
                rowData=df_din.to_dict("records"),
                columnDefs=columnas_din,
                className="ag-theme-alpine-dark",
                style={"height": "400px", "width": "100%"},
                columnSize="sizeToFit",
                dashGridOptions={"rowSelection": "single"}
            )
        ], className="p-3 exec-card")

    return html.Div()

@app.callback(
    Output('vendedor-select', 'value', allow_duplicate=True),
    Input('bar-chart-vendedor', 'clickData'),
    State('session-usuario', 'data'),
    State('vendedor-select', 'value'),
    prevent_initial_call=True
)
def drilldown_vendedor(clickData, usuario_sesion, current_vendedor):
    if usuario_sesion == 'Administrador' and clickData and 'points' in clickData:
        clicked_vend = clickData['points'][0].get('y') or clickData['points'][0].get('x')
        if clicked_vend and clicked_vend != current_vendedor:
            return clicked_vend
    return dash.no_update

@app.callback(
    Output("modal-tramo", "is_open"),
    Output("modal-tramo-titulo", "children"),
    Output("modal-tramo-contenido", "children"),
    Input("grid-dinamica", "rowClicked"),
    Input("btn-cerrar-modal", "n_clicks"),
    State("stored-data", "data"),
    State("session-usuario", "data"),
    State("vendedor-select", "value"),
    State("localidad-select", "value"),
    State("cliente-select", "value"),
    prevent_initial_call=True
)
def mostrar_detalle_tramo(row_click, n_cerrar, records, usuario_sesion, vendedor_sel, localidad_sel, cliente_sel):
    trigger = ctx.triggered_id
    if trigger == "btn-cerrar-modal":
        return False, dash.no_update, dash.no_update
        
    if trigger == "grid-dinamica" and row_click:
        tramo_seleccionado = row_click.get("data", {}).get("Tramo Morosidad")
        if not tramo_seleccionado:
            return dash.no_update, dash.no_update, dash.no_update
            
        df = filtrar_dataframe(records, usuario_sesion, vendedor_sel, localidad_sel, cliente_sel)
        df_filtrado_tramo = df[df["Tramo Morosidad"] == tramo_seleccionado].groupby(
            ["Cliente", "Razon Social", "Vendedor"], as_index=False
        ).agg({"Saldo Deuda": "sum", "Días de Atraso": "max"}).sort_values(by="Saldo Deuda", ascending=False)
        
        columnas_subgrid = [
            {"field": "Cliente", "headerName": "Código", "sortable": True, "filter": True, "width": 110},
            {"field": "Razon Social", "headerName": "Razón Social", "sortable": True, "filter": True, "flex": 2},
            {"field": "Vendedor", "headerName": "Vendedor", "sortable": True, "filter": True, "flex": 1},
            {"field": "Días de Atraso", "headerName": "Máx Atraso", "sortable": True, "filter": True, "width": 110},
            {"field": "Saldo Deuda", "headerName": "Saldo Total", "sortable": True, "filter": True, "type": "rightAligned", "valueFormatter": {"function": "d3.format('$,.2f')(params.value)"}, "flex": 1}
        ]
        
        subgrid = dag.AgGrid(
            rowData=df_filtrado_tramo.to_dict("records"),
            columnDefs=columnas_subgrid,
            className="ag-theme-alpine-dark",
            style={"height": "400px", "width": "100%"},
            columnSize="sizeToFit",
            dashGridOptions={"pagination": True, "paginationPageSize": 8}
        )
        
        return True, f"📋 Clientes en el Tramo: {tramo_seleccionado}", subgrid
        
    return dash.no_update, dash.no_update, dash.no_update

if __name__ == '__main__':
    app.run(debug=True, port=8050)
