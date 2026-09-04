"""Construccion de las figuras Plotly usadas en el dashboard."""

from __future__ import annotations

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

COLOR_SI = "#2BB673"
COLOR_NO = "#E4572E"
COLOR_PRIMARIO = "#1B4965"
COLOR_ACENTO = "#5FA8D3"
COLOR_NEUTRO = "#8895A7"

ESCALA_CALOR = ["#F4511E", "#FDD835", "#43A047"]

FONT_FAMILY = "Inter, Segoe UI, Helvetica, Arial, sans-serif"

LAYOUT_BASE = dict(
font=dict(family=FONT_FAMILY, size=13, color="#2B2F38"),
paper_bgcolor="rgba(0,0,0,0)",
plot_bgcolor="rgba(0,0,0,0)",
margin=dict(l=10, r=10, t=48, b=10),
legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
hoverlabel=dict(bgcolor="white", font_size=13, font_family=FONT_FAMILY),
)


def _aplicar_layout_base(fig, titulo):
  fig.update_layout(**LAYOUT_BASE, title=dict(text=titulo, x=0.02, xanchor="left", font=dict(size=16)))
  return fig


def grafico_barras_comuna(tabla_comunas, columna, color, titulo):
  datos = tabla_comunas.sort_values(columna, ascending=True)
  fig = go.Figure(
  go.Bar(
  x=datos[columna], y=datos["Comuna"], orientation="h",
  marker_color=color, text=datos[columna], textposition="outside",
  )
  )
  _aplicar_layout_base(fig, titulo)
  fig.update_yaxes(title=None)
  fig.update_xaxes(title=None)
  return fig


def grafico_participacion_comuna(tabla_comunas):
  fig = px.pie(
  tabla_comunas, names="Comuna", values="Total_Llamadas", hole=0.55,
  color_discrete_sequence=px.colors.sequential.Teal[::-1],
  )
  fig.update_traces(textposition="inside", textinfo="percent+label", showlegend=False)
  _aplicar_layout_base(fig, "Participacion por comuna (% del total de llamadas)")
  return fig




def grafico_heatmap_comuna(tabla_comunas):
  datos = tabla_comunas.sort_values("Pct_Efectividad", ascending=True)
  fig = go.Figure(
  go.Heatmap(
  z=[datos["Pct_Efectividad"].tolist()],
  x=datos["Comuna"].tolist(),
  y=["% Efectividad"],
  colorscale=[[0, ESCALA_CALOR[0]], [0.5, ESCALA_CALOR[1]], [1, ESCALA_CALOR[2]]],
  text=[[f"{v:.1f}%" for v in datos["Pct_Efectividad"]]],
  texttemplate="%{text}",
  showscale=True,
  zmin=0, zmax=100,
  )
  )
  _aplicar_layout_base(fig, "Mapa de calor de desempeno por comuna")
  fig.update_layout(height=220)
  fig.update_xaxes(tickangle=-35)
  return fig


def grafico_top_comunas(tabla_comunas, mejores=True, n=10):
  datos = tabla_comunas.sort_values("Pct_Efectividad", ascending=not mejores).head(n)
  datos = datos.sort_values("Pct_Efectividad", ascending=True)
  color = COLOR_SI if mejores else COLOR_NO
  titulo = f"Top {n} comunas con {'mejor' if mejores else 'menor'} efectividad"
  fig = go.Figure(
  go.Bar(
  x=datos["Pct_Efectividad"], y=datos["Comuna"], orientation="h",
  marker_color=color, text=datos["Pct_Efectividad"].map(lambda v: f"{v:.1f}%"),
  textposition="outside",
  )
  )
  _aplicar_layout_base(fig, titulo)
  fig.update_xaxes(title="% Efectividad", range=[0, 105])
  return fig



def grafico_ranking_responsables(tabla_resp, n=10):
  datos = tabla_resp.sort_values("Pct_Productividad", ascending=False).head(n)
  datos = datos.sort_values("Pct_Productividad", ascending=True)
  fig = go.Figure(
  go.Bar(
  x=datos["Pct_Productividad"], y=datos["Responsable"], orientation="h",
  marker_color=COLOR_PRIMARIO, text=datos["Pct_Productividad"].map(lambda v: f"{v:.1f}%"),
  textposition="outside",
  )
  )
  _aplicar_layout_base(fig, f"Top {n} responsables por % de productividad")
  fig.update_xaxes(title="% Productividad", range=[0, 105])
  return fig


def grafico_comparativo_responsables(tabla_resp):
  datos = tabla_resp.sort_values("Total_Llamadas", ascending=True)
  fig = go.Figure()
  fig.add_bar(name="Efectivas (SI)", x=datos["Total_SI"], y=datos["Responsable"], orientation="h", marker_color=COLOR_SI)
  fig.add_bar(name="No efectivas (NO)", x=datos["Total_NO"], y=datos["Responsable"], orientation="h", marker_color=COLOR_NO)
  fig.update_layout(barmode="stack")
  _aplicar_layout_base(fig, "Comparativo de productividad por responsable")
  return fig



def grafico_tendencia(tendencia):
  if tendencia.empty:
    fig = go.Figure()
    fig.add_annotation(text="No hay columna de fecha disponible para mostrar tendencia", showarrow=False, font=dict(size=14, color=COLOR_NEUTRO))
    _aplicar_layout_base(fig, "Tendencia historica")
    fig.update_xaxes(visible=False)
    fig.update_yaxes(visible=False)
    return fig
  fig = go.Figure()
  fig.add_trace(go.Bar(x=tendencia["Fecha"], y=tendencia["Total_Llamadas"], name="Total llamadas", marker_color=COLOR_ACENTO, opacity=0.55, yaxis="y"))
  fig.add_trace(go.Scatter(x=tendencia["Fecha"], y=tendencia["Pct_Efectividad"], name="% Efectividad", mode="lines+markers", line=dict(color=COLOR_PRIMARIO, width=3), yaxis="y2"))
  fig.update_layout(
  yaxis=dict(title="Total llamadas"),
  yaxis2=dict(title="% Efectividad", overlaying="y", side="right", range=[0, 100]),
  )
  _aplicar_layout_base(fig, "Tendencia historica de volumen y efectividad")
  return fig



def grafico_pivote_heatmap(pivote_pct):
  if pivote_pct.empty:
    fig = go.Figure()
    _aplicar_layout_base(fig, "Relacion Responsable - Comuna")
    return fig
  fig = go.Figure(
    go.Heatmap(
    z=pivote_pct.values,
    x=pivote_pct.columns.tolist(),
    y=pivote_pct.index.tolist(),
    colorscale=[[0, ESCALA_CALOR[0]], [0.5, ESCALA_CALOR[1]], [1, ESCALA_CALOR[2]]],
    zmin=0, zmax=100,
    colorbar=dict(title="%"),
    )
    )
  _aplicar_layout_base(fig, "% de efectividad por Responsable y Comuna")
  fig.update_xaxes(tickangle=-35)
  fig.update_layout(height=max(260, 45 * max(len(pivote_pct.index), 1)))
  return fig


def kpi_card_color(pct):
  if pct >= 60:
    return COLOR_SI
  if pct >= 40:
    return "#F2A20C"
  return COLOR_NO
  


def grafico_cumplimiento_responsable(df_resumen):
  columna_valor = "% Cumplimiento de reporte"
  columna_nombre = "Responsable / Gestor"
  datos = df_resumen.sort_values(columna_valor, ascending=True)
  colores = [COLOR_SI if v >= 90 else (COLOR_ACENTO if v >= 70 else COLOR_NO) for v in datos[columna_valor]]
  fig = go.Figure(go.Bar(x=datos[columna_valor], y=datos[columna_nombre], orientation="h", marker_color=colores, text=datos[columna_valor].map(lambda v: f"{v:.1f}%"), textposition="outside"))
  fig.update_layout(title="% Cumplimiento de reporte por responsable", xaxis_title="% Cumplimiento", margin=dict(l=10, r=10, t=40, b=10), height=max(300, 40 * len(datos)))
  return fig
