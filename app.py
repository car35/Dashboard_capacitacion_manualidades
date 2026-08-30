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
