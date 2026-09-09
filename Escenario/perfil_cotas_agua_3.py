"""
Visualizador interactivo de cotas de agua MÁXIMAS sobre un perfil longitudinal.

FORMATO DEL EXCEL
-----------------
- Una hoja por escenario, nombrada "Escenario 0", "Escenario 1", ... "Escenario N"
  (se detectan automáticamente todas las que existan en el archivo).
  * SIN fila de encabezado: la primera fila ya es dato.
  * Primera columna: Progresiva (m).
  * Segunda columna: Cota de agua máxima (m s.n.m.).

- Hoja "Perfil" CON encabezado: columnas "Progresiva" y "Cota_terreno".

Instalación de dependencias:
    pip install dash pandas openpyxl plotly

Ejecución:
    python perfil_cotas_agua.py
    (se abre en http://127.0.0.1:8050)
"""

import re
import itertools

import pandas as pd
import dash
from dash import dcc, html, Input, Output
import plotly.graph_objects as go

# ------------------------------------------------------------------
# CONFIGURACIÓN: cambiar por la ruta real del Excel
# ------------------------------------------------------------------
EXCEL_PATH = "Escenarios.xlsx"

PATRON_ESCENARIO = re.compile(r"^Escenario\s+(\d+)$", re.IGNORECASE)

# Paleta legible sobre fondo oscuro, un color fijo por escenario
PALETA = [
    "#F2A65A", "#5FD3BC", "#F26D6D", "#E8C547", "#8E9AAF",
    "#C77DFF", "#5FA8D3", "#FF8FA3", "#94D2BD", "#EE9B00",
    "#B8F2E6", "#FFD6A5",
]
COLOR_TERRENO = "#E4572E"  # rojo óxido, como una corrección a mano sobre un plano

# Paleta de la interfaz (tema "plano técnico")
BG = "#0E2A42"
BG_PANEL = "#0B2338"
INK = "#EAF2F8"
INK_DIM = "#9FB8CB"
BORDE = "rgba(234,242,248,0.18)"
GRILLA = "rgba(234,242,248,0.10)"
ACENTO = "#5FA8D3"


def cargar_datos(path):
    xls = pd.ExcelFile(path)

    encontrados = []
    for nombre in xls.sheet_names:
        m = PATRON_ESCENARIO.match(nombre.strip())
        if m:
            encontrados.append((int(m.group(1)), nombre))
    encontrados.sort(key=lambda x: x[0])

    data = {}
    for _, nombre in encontrados:
        df = pd.read_excel(xls, sheet_name=nombre, header=None, usecols=[0, 1])
        df.columns = ["Progresiva", "Cota"]
        df = df.sort_values("Progresiva")
        data[nombre] = df

    perfil = None
    if "Perfil" in xls.sheet_names:
        perfil = pd.read_excel(xls, sheet_name="Perfil").sort_values("Progresiva")

    return data, perfil


data, perfil = cargar_datos(EXCEL_PATH)

if not data:
    raise ValueError(
        "No se encontró ninguna hoja con formato 'Escenario N' en el Excel. "
        "Revisá los nombres de las hojas."
    )

colores = dict(zip(data.keys(), itertools.cycle(PALETA)))

if perfil is not None:
    prog_min, prog_max = perfil["Progresiva"].min(), perfil["Progresiva"].max()
else:
    todas_prog = pd.concat([df["Progresiva"] for df in data.values()])
    prog_min, prog_max = todas_prog.min(), todas_prog.max()

app = dash.Dash(__name__)
app.title = "Perfil longitudinal - Cotas de agua"

app.index_string = """
<!DOCTYPE html>
<html>
    <head>
        {%metas%}
        <title>{%title%}</title>
        {%favicon%}
        <link rel="preconnect" href="https://fonts.googleapis.com">
        <link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;600&family=IBM+Plex+Sans:wght@400;500&display=swap" rel="stylesheet">
        {%css%}
        <style>
            :root {
                --bg: """ + BG + """;
                --bg-panel: """ + BG_PANEL + """;
                --ink: """ + INK + """;
                --ink-dim: """ + INK_DIM + """;
                --borde: """ + BORDE + """;
                --acento: """ + ACENTO + """;
            }
            * { box-sizing: border-box; }
            body {
                margin: 0;
                background-color: var(--bg);
                background-image:
                    linear-gradient(rgba(234,242,248,0.05) 1px, transparent 1px),
                    linear-gradient(90deg, rgba(234,242,248,0.05) 1px, transparent 1px);
                background-size: 28px 28px;
                color: var(--ink);
                font-family: 'IBM Plex Sans', -apple-system, Segoe UI, sans-serif;
            }
            .app-shell { max-width: 1280px; margin: 0 auto; padding: 32px 24px 48px; }
            .titleblock {
                display: flex; flex-wrap: wrap;
                border: 1px solid var(--borde); border-radius: 2px;
                background: var(--bg-panel); margin-bottom: 24px;
            }
            .titleblock > div {
                padding: 16px 22px; border-right: 1px solid var(--borde);
            }
            .titleblock > div:last-child { border-right: none; }
            .titleblock .campo-titulo { flex: 1 1 320px; display: flex; align-items: center; }
            .titleblock .campo-titulo h1 {
                font-family: 'IBM Plex Mono', monospace; font-weight: 600;
                font-size: 19px; margin: 0; letter-spacing: 0.01em;
            }
            .titleblock .campo-dato { flex: 0 0 auto; min-width: 150px; }
            .campo-label {
                font-family: 'IBM Plex Mono', monospace; font-size: 11px;
                color: var(--ink-dim); margin-bottom: 4px;
            }
            .campo-valor { font-size: 14px; }
            .layout { display: flex; gap: 22px; align-items: flex-start; }
            .panel {
                width: 250px; flex-shrink: 0; background: var(--bg-panel);
                border: 1px solid var(--borde); border-radius: 2px; padding: 18px;
            }
            .panel h2 {
                font-family: 'IBM Plex Mono', monospace; font-size: 12px;
                font-weight: 500; color: var(--ink-dim); margin: 0 0 10px 0;
            }
            .botones { margin-top: 14px; }
            .botones button {
                font-family: 'IBM Plex Sans', sans-serif; background: transparent;
                border: 1px solid var(--borde); color: var(--ink);
                padding: 7px 13px; border-radius: 2px; cursor: pointer;
                font-size: 13px; margin-right: 8px; margin-top: 6px;
                transition: border-color 0.15s ease, color 0.15s ease;
            }
            .botones button:hover { border-color: var(--acento); color: var(--acento); }
            .grafico-container {
                flex: 1; min-width: 0; background: var(--bg-panel);
                border: 1px solid var(--borde); border-radius: 2px; padding: 10px;
            }
            .tip {
                font-family: 'IBM Plex Mono', monospace; font-size: 12px;
                color: var(--ink-dim); margin-top: 16px;
            }
            #check-escenarios label {
                display: flex; align-items: center; padding: 5px 0;
                font-size: 14px; cursor: pointer;
            }
            #check-escenarios input[type="checkbox"] {
                accent-color: var(--acento); width: 14px; height: 14px; margin-right: 9px;
            }
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
"""

app.layout = html.Div(
    className="app-shell",
    children=[
        html.Div(
            className="titleblock",
            children=[
                html.Div(
                    className="campo-titulo",
                    children=[html.H1("Perfil longitudinal: cotas de agua máxima")],
                ),
                html.Div(
                    className="campo-dato",
                    children=[
                        html.Div("escenarios", className="campo-label"),
                        html.Div(str(len(data)), className="campo-valor"),
                    ],
                ),
                html.Div(
                    className="campo-dato",
                    children=[
                        html.Div("progresivas", className="campo-label"),
                        html.Div(f"{prog_min:,.0f} – {prog_max:,.0f} m", className="campo-valor"),
                    ],
                ),
                html.Div(
                    className="campo-dato",
                    children=[
                        html.Div("unidades", className="campo-label"),
                        html.Div("m s.n.m.", className="campo-valor"),
                    ],
                ),
            ],
        ),
        html.Div(
            className="layout",
            children=[
                html.Div(
                    className="panel",
                    children=[
                        html.H2("Escenarios"),
                        dcc.Checklist(
                            id="check-escenarios",
                            options=[{"label": esc, "value": esc} for esc in data.keys()],
                            value=list(data.keys()),
                        ),
                        html.Div(
                            className="botones",
                            children=[
                                html.Button("Todos", id="btn-todos", n_clicks=0),
                                html.Button("Ninguno", id="btn-ninguno", n_clicks=0),
                            ],
                        ),
                    ],
                ),
                html.Div(
                    className="grafico-container",
                    children=[dcc.Graph(id="grafico-perfil", style={"height": "660px"})],
                ),
            ],
        ),
        html.P(
            "Arrastrá sobre el gráfico para hacer zoom. Doble click para volver a la vista completa.",
            className="tip",
        ),
    ],
)


@app.callback(
    Output("check-escenarios", "value"),
    Input("btn-todos", "n_clicks"),
    Input("btn-ninguno", "n_clicks"),
    prevent_initial_call=True,
)
def toggle_todos(n_todos, n_ninguno):
    ctx = dash.callback_context
    boton = ctx.triggered[0]["prop_id"].split(".")[0]
    return list(data.keys()) if boton == "btn-todos" else []


@app.callback(
    Output("grafico-perfil", "figure"),
    Input("check-escenarios", "value"),
)
def actualizar_grafico(escenarios_sel):
    escenarios_sel = escenarios_sel or []
    fig = go.Figure()
    valores_y = []

    if perfil is not None:
        fig.add_trace(
            go.Scatter(
                x=perfil["Progresiva"], y=perfil["Cota_terreno"],
                mode="lines", name="Terreno",
                line=dict(color=COLOR_TERRENO, width=2),
            )
        )
        valores_y.append(perfil["Cota_terreno"])

    for esc in escenarios_sel:
        df = data[esc]
        fig.add_trace(
            go.Scatter(
                x=df["Progresiva"], y=df["Cota"],
                mode="lines", name=esc,
                line=dict(color=colores[esc], width=2),
            )
        )
        valores_y.append(df["Cota"])

    if valores_y:
        y_min = min(serie.min() for serie in valores_y)
        y_max = max(serie.max() for serie in valores_y)
        margen = max((y_max - y_min) * 0.08, 0.1)
        rango_y = [y_min - margen, y_max + margen]
    else:
        rango_y = None

    fig.update_layout(
        xaxis_title="Progresiva (m)",
        yaxis_title="Cota (m s.n.m.)",
        xaxis=dict(gridcolor=GRILLA, zerolinecolor=GRILLA),
        yaxis=dict(range=rango_y, gridcolor=GRILLA, zerolinecolor=GRILLA),
        legend=dict(bgcolor="rgba(0,0,0,0)"),
        paper_bgcolor=BG_PANEL,
        plot_bgcolor=BG_PANEL,
        font=dict(family="IBM Plex Sans, sans-serif", color=INK),
        hovermode="x unified",
        hoverlabel=dict(bgcolor=BG_PANEL, font=dict(family="IBM Plex Mono, monospace", color=INK), bordercolor=BORDE),
        margin=dict(l=60, r=30, t=20, b=50),
    )
    return fig


server = app.server

if __name__ == "__main__":
    app.run(debug=False)