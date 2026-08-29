"""
Calculo de indicadores (KPIs) y tablas agregadas para el dashboard de
productividad de gestion de llamadas.
"""

from __future__ import annotations

from dataclasses import dataclass

import pandas as pd


def _pct_efectividad(total: int, si: int) -> float:
  return round((si / total) * 100, 1) if total else 0.0


@dataclass
class KPIsGenerales:
  total_llamadas: int
  total_si: int
  total_no: int
  pct_efectividad: float
  total_responsables: int
  total_comunas: int
  promedio_efectividad_responsables: float
  mejor_comuna: str
  mejor_comuna_pct: float
  peor_comuna: str
  peor_comuna_pct: float
  mejor_responsable: str
  mejor_responsable_pct: float
  peor_responsable: str
  peor_responsable_pct: float


def calcular_tabla_comunas(df: pd.DataFrame) -> pd.DataFrame:
  """Tabla Comuna | Total | SI | NO | % Efectividad, ordenada por % desc."""
  agg = (
  df.groupby("Comuna")["Llamada_Efectiva"]
  .value_counts()
  .unstack(fill_value=0)
  .reindex(columns=["SI", "NO"], fill_value=0)
  )
  agg["Total_Llamadas"] = agg["SI"] + agg["NO"]
  agg["Pct_Efectividad"] = (agg["SI"] / agg["Total_Llamadas"] * 100).round(1)
  agg["Pct_Participacion"] = (agg["Total_Llamadas"] / agg["Total_Llamadas"].sum() * 100).round(1)
  agg = agg.rename(columns={"SI": "Total_SI", "NO": "Total_NO"})
  agg = agg.reset_index().sort_values("Pct_Efectividad", ascending=False)
  return agg[["Comuna", "Total_Llamadas", "Total_SI", "Total_NO", "Pct_Efectividad", "Pct_Participacion"]]


def calcular_tabla_responsables(df: pd.DataFrame) -> pd.DataFrame:
  """Tabla Responsable | Total | SI | NO | % Productividad, ordenada por % desc."""
  agg = (
  df.groupby("Responsable")["Llamada_Efectiva"]
  .value_counts()
  .unstack(fill_value=0)
  .reindex(columns=["SI", "NO"], fill_value=0)
  )
  agg["Total_Llamadas"] = agg["SI"] + agg["NO"]
  agg["Pct_Productividad"] = (agg["SI"] / agg["Total_Llamadas"] * 100).round(1)
  agg = agg.rename(columns={"SI": "Total_SI", "NO": "Total_NO"})
  agg = agg.reset_index().sort_values("Pct_Productividad", ascending=False)
  return agg[["Responsable", "Total_Llamadas", "Total_SI", "Total_NO", "Pct_Productividad"]]

def calcular_pivote_responsable_comuna(df: pd.DataFrame, metrica: str = "pct") -> pd.DataFrame:
  """Tabla dinamica Responsable x Comuna."""
  if metrica == "total":
    pivote = pd.pivot_table(
      df, index="Responsable", columns="Comuna", values="Llamada_Efectiva",
      aggfunc="count", fill_value=0,
    )
    return pivote

  def _pct(grupo: pd.DataFrame) -> float:
    total = len(grupo)
    si = int((grupo["Llamada_Efectiva"] == "SI").sum())
    return _pct_efectividad(total, si)

  pivote = (
  df.groupby(["Responsable", "Comuna"])
  .apply(_pct, include_groups=False)
  .unstack(fill_value=None)
  )
  return pivote

def calcular_kpis_generales(df: pd.DataFrame) -> KPIsGenerales:
  total = len(df)
  si = int((df["Llamada_Efectiva"] == "SI").sum())
  no = total - si

  tabla_comunas = calcular_tabla_comunas(df)
  tabla_resp = calcular_tabla_responsables(df)

  if not tabla_comunas.empty:
    fila_mejor_comuna = tabla_comunas.iloc[0]
    fila_peor_comuna = tabla_comunas.iloc[-1]
  else:
    fila_mejor_comuna = fila_peor_comuna = {"Comuna": "-", "Pct_Efectividad": 0.0}

  if not tabla_resp.empty:
    fila_mejor_resp = tabla_resp.iloc[0]
    fila_peor_resp = tabla_resp.iloc[-1]
  else:
    fila_mejor_resp = fila_peor_resp = {"Responsable": "-", "Pct_Productividad": 0.0}

  return KPIsGenerales(
  total_llamadas=total,
  total_si=si,
  total_no=no,
  pct_efectividad=_pct_efectividad(total, si),
  total_responsables=df["Responsable"].nunique(),
  total_comunas=df["Comuna"].nunique(),
  promedio_efectividad_responsables=round(tabla_resp["Pct_Productividad"].mean(), 1)
  if not tabla_resp.empty else 0.0,
  mejor_comuna=fila_mejor_comuna["Comuna"],
  mejor_comuna_pct=fila_mejor_comuna["Pct_Efectividad"],
  peor_comuna=fila_peor_comuna["Comuna"],
  peor_comuna_pct=fila_peor_comuna["Pct_Efectividad"],
  mejor_responsable=fila_mejor_resp["Responsable"],
  mejor_responsable_pct=fila_mejor_resp["Pct_Productividad"],
  peor_responsable=fila_peor_resp["Responsable"],
  peor_responsable_pct=fila_peor_resp["Pct_Productividad"],
  )

def calcular_tendencia_diaria(df: pd.DataFrame) -> pd.DataFrame:
  """Serie diaria de Total / SI / NO / % Efectividad, si hay columna Fecha."""
  if "Fecha" not in df.columns or df["Fecha"].isna().all():
    return pd.DataFrame(columns=["Fecha", "Total_Llamadas", "Total_SI", "Total_NO", "Pct_Efectividad"])
  tmp = df.dropna(subset=["Fecha"]).copy()
    tmp["Dia"] = tmp["Fecha"].dt.date
    agg = (
      tmp.groupby("Dia")["Llamada_Efectiva"]
      .value_counts()
      .unstack(fill_value=0)
      .reindex(columns=["SI", "NO"], fill_value=0)
    )
    agg["Total_Llamadas"] = agg["SI"] + agg["NO"]
    agg["Pct_Efectividad"] = (agg["SI"] / agg["Total_Llamadas"] * 100).round(1)
    agg = agg.rename(columns={"SI": "Total_SI", "NO": "Total_NO"}).reset_index()
    agg = agg.rename(columns={"Dia": "Fecha"}).sort_values("Fecha")
    return agg
    
