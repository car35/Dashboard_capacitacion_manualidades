"""Dashboard Web de Productividad de Gestion de Llamadas."""

from __future__ import annotations

import io
from datetime import datetime

import dash
import dash_bootstrap_components as dbc
import pandas as pd
import plotly.graph_objects as go
from dash import Input, Output, State, dash_table, dcc, html
from dash.exceptions import PreventUpdate

import charts
import kpi_engine
from data_loader import cargar_desde_bytes, cargar_desde_ruta, decodificar_contenido_upload
from excel_export import generar_reporte_excel

RUTA_DATOS_EJEMPLO = "data/sample_llamadas.xlsx"

app = dash.Dash(__name__, external_stylesheets=[dbc.themes.FLATLY, dbc.icons.BOOTSTRAP], title="Dashboard de Productividad de Llamadas", suppress_callback_exceptions=True)
server = app.server


def tarjeta_kpi(kpi_id, etiqueta, icono):
  return dbc.Card(dbc.CardBody([html.Div([html.I(className=f"bi {icono} kpi-icono"), html.Span(etiqueta, className="kpi-etiqueta")], className="kpi-encabezado"), html.Div(id=kpi_id, className="kpi-valor")]), className="tarjeta-kpi shadow-sm")


FILA_KPIS = [("kpi-total-llamadas", "Total de llamadas", "bi-telephone"), ("kpi-total-si", "Llamadas efectivas (SI)", "bi-check-circle"), ("kpi-total-no", "Llamadas no efectivas (NO)", "bi-x-circle"), ("kpi-pct-efectividad", "% Efectividad general", "bi-speedometer2"), ("kpi-total-responsables", "Total responsables", "bi-people"), ("kpi-total-comunas", "Total comunas", "bi-geo-alt"), ("kpi-promedio-efectividad", "Promedio de efectividad", "bi-bar-chart-line"), ("kpi-mejor-comuna", "Mejor comuna", "bi-trophy"), ("kpi-peor-comuna", "Comuna a reforzar", "bi-exclamation-triangle"), ("kpi-mejor-responsable", "Responsable mas productivo", "bi-star"), ("kpi-peor-responsable", "Responsable a reforzar", "bi-arrow-down-circle")]


def barra_filtros():
  return dbc.Card(dbc.CardBody(dbc.Row([dbc.Col([html.Label("Rango de fechas", className="filtro-label"), dcc.DatePickerRange(id="filtro-fechas", display_format="YYYY-MM-DD", className="w-100")], md=4), dbc.Col([html.Label("Responsable(s)", className="filtro-label"), dcc.Dropdown(id="filtro-responsable", multi=True, placeholder="Todos")], md=4), dbc.Col([html.Label("Comuna(s)", className="filtro-label"), dcc.Dropdown(id="filtro-comuna", multi=True, placeholder="Todas")], md=4)], className="g-3")), className="shadow-sm mb-3")


def encabezado():
  return dbc.Navbar(dbc.Container([html.Div([html.I(className="bi bi-graph-up-arrow me-2"), html.Span("Productividad de Gestion de Llamadas", className="fw-bold")], className="navbar-brand d-flex align-items-center"), html.Div([dcc.Upload(id="upload-datos", children=dbc.Button([html.I(className="bi bi-upload me-2"), "Cargar archivo"], color="light", outline=True, size="sm"), multiple=False), dbc.Button([html.I(className="bi bi-file-earmark-excel me-2"), "Exportar reporte Excel"], id="btn-exportar-excel", color="success", size="sm", className="ms-2"), dcc.Download(id="descarga-excel")], className="d-flex align-items-center")], fluid=True), color="dark", dark=True, className="mb-3 shadow-sm")


def panel_comuna():
  return html.Div([dbc.Row([dbc.Col(dcc.Graph(id="grafico-comuna-si", config={"displayModeBar": False}), md=6), dbc.Col(dcc.Graph(id="grafico-comuna-no", config={"displayModeBar": False}), md=6)], className="g-3 mb-3"), dbc.Row([dbc.Col(dcc.Graph(id="grafico-comuna-participacion", config={"displayModeBar": False}), md=6), dbc.Col(dcc.Graph(id="grafico-comuna-heatmap", config={"displayModeBar": False}), md=6)], className="g-3 mb-3"), dbc.Row([dbc.Col(dcc.Graph(id="grafico-comuna-top-mejores", config={"displayModeBar": False}), md=6), dbc.Col(dcc.Graph(id="grafico-comuna-top-peores", config={"displayModeBar": False}), md=6)], className="g-3 mb-3"), dbc.Card(dbc.CardBody([html.H6("Detalle por comuna", className="mb-3"), html.Div(id="tabla-comunas")]), className="shadow-sm")])


def panel_responsable():
  return html.Div([dbc.Row([dbc.Col(dcc.Graph(id="grafico-resp-ranking", config={"displayModeBar": False}), md=6), dbc.Col(dcc.Graph(id="grafico-resp-comparativo", config={"displayModeBar": False}), md=6)], className="g-3 mb-3"), dbc.Row(dbc.Col(dcc.Graph(id="grafico-resp-tendencia", config={"displayModeBar": False}), md=12), className="g-3 mb-3"), dbc.Card(dbc.CardBody([html.H6("Detalle por responsable", className="mb-3"), html.Div(id="tabla-responsables")]), className="shadow-sm")])


def panel_relacion():
  return html.Div([dbc.Row(dbc.Col(dcc.Graph(id="grafico-pivote-heatmap", config={"displayModeBar": False}), md=12), className="g-3 mb-3"), dbc.Card(dbc.CardBody([html.H6("Tabla dinamica: % de efectividad por Responsable y Comuna", className="mb-3"), html.Div(id="tabla-pivote")]), className="shadow-sm")])


app.layout = html.Div([dcc.Store(id="store-datos"), dcc.Store(id="store-nombre-archivo"), encabezado(), dbc.Container([html.Div(id="zona-alertas"), barra_filtros(), dbc.Row([dbc.Col(tarjeta_kpi(kpi_id, etiqueta, icono), lg=3, md=4, sm=6, xs=12, className="mb-3") for kpi_id, etiqueta, icono in FILA_KPIS], className="g-3 mb-2"), dbc.Tabs([dbc.Tab(panel_comuna(), label="Analisis por Comuna", tab_id="tab-comuna"), dbc.Tab(panel_responsable(), label="Analisis por Responsable", tab_id="tab-responsable"), dbc.Tab(panel_relacion(), label="Relacion Responsable - Comuna", tab_id="tab-relacion")], id="tabs-principales", active_tab="tab-comuna", className="mt-2"), html.Footer("Dashboard de Productividad - generado con Dash, Plotly y Pandas", className="text-muted text-center small py-4")], fluid=True)])


@app.callback(Output("store-datos", "data"), Output("store-nombre-archivo", "data"), Output("zona-alertas", "children"), Input("upload-datos", "contents"), State("upload-datos", "filename"), prevent_initial_call=False)
def cargar_datos(contenido_upload, nombre_upload):
  if contenido_upload is None:
    resultado = cargar_desde_ruta(RUTA_DATOS_EJEMPLO)
    nombre = "Datos de ejemplo (sample_llamadas.xlsx)"
  else:
    datos_bytes = decodificar_contenido_upload(contenido_upload)
    resultado = cargar_desde_bytes(datos_bytes, nombre_upload)
    nombre = nombre_upload
  if not resultado.exito:
    alerta = dbc.Alert([html.B("No se pudo cargar el archivo. "), *[html.Div(e) for e in resultado.errores]], color="danger", dismissable=True)
    if contenido_upload is None:
      raise PreventUpdate
    return dash.no_update, dash.no_update, alerta
  alertas = []
  if resultado.advertencias:
    alertas.append(dbc.Alert([html.B("Archivo cargado con observaciones: ")] + [html.Div(a) for a in resultado.advertencias], color="warning", dismissable=True))
  else:
    alertas.append(dbc.Alert(f"Archivo '{nombre}' cargado correctamente ({len(resultado.df)} filas).", color="success", dismissable=True, duration=4000))
  return resultado.df.to_json(date_format="iso", orient="split"), nombre, alertas
          


@app.callback(Output("filtro-responsable", "options"), Output("filtro-comuna", "options"), Output("filtro-fechas", "min_date_allowed"), Output("filtro-fechas", "max_date_allowed"), Output("filtro-fechas", "start_date"), Output("filtro-fechas", "end_date"), Input("store-datos", "data"))
def poblar_filtros(datos_json):
  if not datos_json:
    raise PreventUpdate
  df = pd.read_json(io.StringIO(datos_json), orient="split")
  opciones_resp = sorted(df["Responsable"].dropna().unique())
  opciones_comuna = sorted(df["Comuna"].dropna().unique())
  opts_resp = [{"label": r, "value": r} for r in opciones_resp]
  opts_comuna = [{"label": c, "value": c} for c in opciones_comuna]
  if "Fecha" in df.columns and not df["Fecha"].isna().all():
    fechas = pd.to_datetime(df["Fecha"]).dropna()
    return opts_resp, opts_comuna, fechas.min().date(), fechas.max().date(), fechas.min().date(), fechas.max().date()
  return opts_resp, opts_comuna, None, None, None, None


def _filtrar(df, responsables, comunas, fecha_ini, fecha_fin):
  resultado = df.copy()
  if responsables:
    resultado = resultado[resultado["Responsable"].isin(responsables)]
  if comunas:
    resultado = resultado[resultado["Comuna"].isin(comunas)]
  if "Fecha" in resultado.columns and fecha_ini and fecha_fin:
    fechas = pd.to_datetime(resultado["Fecha"])
    mascara = (fechas >= pd.Timestamp(fecha_ini)) & (fechas <= pd.Timestamp(fecha_fin) + pd.Timedelta(days=1))
    resultado = resultado[mascara]
  return resultado


def _tabla_dash(df, columnas_pct=None):
  columnas_pct = columnas_pct or []
  columnas = []
  for col in df.columns:
    col_def = {"name": col.replace("_", " "), "id": col}
    if col in columnas_pct:
      col_def["type"] = "numeric"
      col_def["format"] = {"specifier": ".1f"}
      columnas.append(col_def)
      return dash_table.DataTable(data=df.to_dict("records"), columns=columnas, sort_action="native", filter_action="native", page_size=10, style_table={"overflowX": "auto"}, style_header={"backgroundColor": "#1B4965", "color": "white", "fontWeight": "bold"}, style_cell={"fontFamily": "Inter, Arial, sans-serif", "fontSize": 13, "padding": "8px"}, style_data_conditional=[{"if": {"row_index": "odd"}, "backgroundColor": "#F5F8FA"}])
      


OUTPUTS_PRINCIPALES = [Output("kpi-total-llamadas", "children"), Output("kpi-total-si", "children"), Output("kpi-total-no", "children"), Output("kpi-pct-efectividad", "children"), Output("kpi-total-responsables", "children"), Output("kpi-total-comunas", "children"), Output("kpi-promedio-efectividad", "children"), Output("kpi-mejor-comuna", "children"), Output("kpi-peor-comuna", "children"), Output("kpi-mejor-responsable", "children"), Output("kpi-peor-responsable", "children"), Output("grafico-comuna-si", "figure"), Output("grafico-comuna-no", "figure"), Output("grafico-comuna-participacion", "figure"), Output("grafico-comuna-heatmap", "figure"), Output("grafico-comuna-top-mejores", "figure"), Output("grafico-comuna-top-peores", "figure"), Output("tabla-comunas", "children"), Output("grafico-resp-ranking", "figure"), Output("grafico-resp-comparativo", "figure"), Output("grafico-resp-tendencia", "figure"), Output("tabla-responsables", "children"), Output("grafico-pivote-heatmap", "figure"), Output("tabla-pivote", "children")]


@app.callback(OUTPUTS_PRINCIPALES, Input("store-datos", "data"), Input("filtro-responsable", "value"), Input("filtro-comuna", "value"), Input("filtro-fechas", "start_date"), Input("filtro-fechas", "end_date"))
def actualizar_dashboard(datos_json, responsables, comunas, fecha_ini, fecha_fin):
  if not datos_json:
    raise PreventUpdate
  df = pd.read_json(io.StringIO(datos_json), orient="split")
  if "Fecha" in df.columns:
    df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce")
  df_filtrado = _filtrar(df, responsables, comunas, fecha_ini, fecha_fin)
  if df_filtrado.empty:
    fig_vacia = go.Figure()
    fig_vacia.add_annotation(text="Sin datos para los filtros seleccionados", showarrow=False, font=dict(size=14))
    fig_vacia.update_xaxes(visible=False)
    fig_vacia.update_yaxes(visible=False)
    mensaje = html.Div("No hay datos para los filtros seleccionados.", className="text-muted p-3")
    return "0", "0", "0", "0.0%", "0", "0", "0.0%", "-", "-", "-", "-", fig_vacia, fig_vacia, fig_vacia, fig_vacia, fig_vacia, fig_vacia, mensaje, fig_vacia, fig_vacia, fig_vacia, mensaje, fig_vacia, mensaje
    kpis = kpi_engine.calcular_kpis_generales(df_filtrado)
    tabla_comunas = kpi_engine.calcular_tabla_comunas(df_filtrado)
    tabla_resp = kpi_engine.calcular_tabla_responsables(df_filtrado)
    pivote_pct = kpi_engine.calcular_pivote_responsable_comuna(df_filtrado, metrica="pct")
    tendencia = kpi_engine.calcular_tendencia_diaria(df_filtrado)
    tabla_pivote_reset = pivote_pct.reset_index()
    return f"{kpis.total_llamadas:,}", f"{kpis.total_si:,}", f"{kpis.total_no:,}", f"{kpis.pct_efectividad:.1f}%", f"{kpis.total_responsables:,}", f"{kpis.total_comunas:,}", f"{kpis.promedio_efectividad_responsables:.1f}%", f"{kpis.mejor_comuna} ({kpis.mejor_comuna_pct:.1f}%)", f"{kpis.peor_comuna} ({kpis.peor_comuna_pct:.1f}%)", f"{kpis.mejor_responsable} ({kpis.mejor_responsable_pct:.1f}%)", f"{kpis.peor_responsable} ({kpis.peor_responsable_pct:.1f}%)", charts.grafico_barras_comuna(tabla_comunas, "Total_SI", charts.COLOR_SI, "Llamadas efectivas (SI) por comuna"), charts.grafico_barras_comuna(tabla_comunas, "Total_NO", charts.COLOR_NO, "Llamadas no efectivas (NO) por comuna"), charts.grafico_participacion_comuna(tabla_comunas), charts.grafico_heatmap_comuna(tabla_comunas), charts.grafico_top_comunas(tabla_comunas, mejores=True), charts.grafico_top_comunas(tabla_comunas, mejores=False), _tabla_dash(tabla_comunas, columnas_pct=["Pct_Efectividad", "Pct_Participacion"]), charts.grafico_ranking_responsables(tabla_resp), charts.grafico_comparativo_responsables(tabla_resp), charts.grafico_tendencia(tendencia), _tabla_dash(tabla_resp, columnas_pct=["Pct_Productividad"]), charts.grafico_pivote_heatmap(pivote_pct), _tabla_dash(tabla_pivote_reset)
        
