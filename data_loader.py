"""Carga y validacion de la fuente de datos del dashboard."""

from __future__ import annotations

import base64
import io
import unicodedata
from dataclasses import dataclass, field

import pandas as pd

COLUMNAS_REQUERIDAS = ["Responsable", "Comuna", "Llamada_Efectiva"]
COLUMNAS_OPCIONALES = ["Fecha", "Cliente", "Observaciones"]

ALIAS_COLUMNAS = {
  "Responsable": ["responsable", "agente", "asesor", "gestor"],
  "Comuna": ["comuna", "zona", "sector", "barrio"],
  "Llamada_Efectiva": ["llamadaefectiva", "efectiva", "esefectiva", "llamada_efectiva", "resultado", "efectividad"],
  "Fecha": ["fecha", "fecharegistro", "fecha_registro", "fechallamada"],
  "Cliente": ["cliente", "nombre", "usuario", "persona"],
  "Observaciones": ["observaciones", "observacion", "notas", "comentarios"],
}


def _normalizar_texto(txt):
  txt = str(txt).strip().lower()
  txt = "".join(c for c in unicodedata.normalize("NFD", txt) if unicodedata.category(c) != "Mn")
  return txt.replace(" ", "").replace("-", "").replace("_", "")


@dataclass
class ResultadoCarga:
  exito: bool
  df: pd.DataFrame | None = None
  errores: list = field(default_factory=list)
  advertencias: list = field(default_factory=list)
  nombre_archivo: str = ""


def _mapear_columnas(df):
  advertencias = []
  mapa_normalizado = {}
  for canonico, alias in ALIAS_COLUMNAS.items():
    objetivo = {_normalizar_texto(canonico)} | {_normalizar_texto(a) for a in alias}
    for col in df.columns:
      if _normalizar_texto(col) in objetivo:
        mapa_normalizado[col] = canonico
        break
  faltantes_canonicas = set(ALIAS_COLUMNAS) - set(mapa_normalizado.values())
  if faltantes_canonicas:
    advertencias.append("No se reconocieron automaticamente estas columnas: " + ", ".join(sorted(faltantes_canonicas)))
  return df.rename(columns=mapa_normalizado), advertencias


def _validar_estructura(df):
  errores = []
  faltantes = [c for c in COLUMNAS_REQUERIDAS if c not in df.columns]
  if faltantes:
    errores.append("Faltan columnas obligatorias: " + ", ".join(faltantes))
  if df.empty:
    errores.append("El archivo no contiene filas de datos.")
    return errores


def _normalizar_tipos(df):
  advertencias = []
  df = df.copy()
  for col in COLUMNAS_REQUERIDAS + COLUMNAS_OPCIONALES:
    if col in df.columns and df[col].dtype == object:
      df[col] = df[col].astype(str).str.strip()
  if "Llamada_Efectiva" in df.columns:
    df["Llamada_Efectiva"] = df["Llamada_Efectiva"].astype(str).str.strip().str.upper()
    equivalencias = {"SI": "SI", "SI ACENTO": "SI", "S": "SI", "1": "SI", "TRUE": "SI", "YES": "SI", "NO": "NO", "N": "NO", "0": "NO", "FALSE": "NO"}
    df["Llamada_Efectiva"] = df["Llamada_Efectiva"].map(lambda v: equivalencias.get(v, v))
    invalidas = ~df["Llamada_Efectiva"].isin(["SI", "NO"])
    n_invalidas = int(invalidas.sum())
    if n_invalidas:
      advertencias.append(f"Se descartaron {n_invalidas} fila(s) con valores no reconocidos en Llamada_Efectiva (se esperaba SI/NO).")
      df = df[~invalidas]
  if "Fecha" in df.columns:
    antes = len(df)
    df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce", dayfirst=False)
    n_malas = int(df["Fecha"].isna().sum())
    if n_malas and n_malas < antes:
      advertencias.append(f"{n_malas} fila(s) tienen fecha invalida o vacia; esas filas no apareceran en los analisis por fecha.")
  for col in ["Responsable", "Comuna"]:
    if col in df.columns:
      vacias = df[col].isna() | (df[col].astype(str).str.strip() == "")
      if vacias.any():
        n_vacias = int(vacias.sum())
        advertencias.append(f"Se descartaron {n_vacias} fila(s) sin valor en '{col}'.")
        df = df[~vacias]
    return df, advertencias


def cargar_desde_bytes(contenido, nombre_archivo):
  try:
    if nombre_archivo.lower().endswith(".csv"):
      df = pd.read_csv(io.BytesIO(contenido))
    elif nombre_archivo.lower().endswith((".xlsx", ".xls", ".xlsm")):
      df = pd.read_excel(io.BytesIO(contenido))
    else:
      return ResultadoCarga(exito=False, errores=["Formato no soportado. Sube un archivo .xlsx o .csv."], nombre_archivo=nombre_archivo)
  except Exception as exc:
    return ResultadoCarga(exito=False, errores=[f"No se pudo leer el archivo: {exc}"], nombre_archivo=nombre_archivo)
  return _procesar_dataframe(df, nombre_archivo)


def cargar_desde_ruta(ruta):
  try:
    if ruta.lower().endswith(".csv"):
      df = pd.read_csv(ruta)
    else:
      df = pd.read_excel(ruta)
  except Exception as exc:
    return ResultadoCarga(exito=False, errores=[f"No se pudo leer el archivo: {exc}"])
  return _procesar_dataframe(df, ruta)


def _procesar_dataframe(df, nombre_archivo):
  df, advertencias_mapeo = _mapear_columnas(df)
  errores = _validar_estructura(df)
  if errores:
    return ResultadoCarga(exito=False, errores=errores, advertencias=advertencias_mapeo, nombre_archivo=nombre_archivo)
    df, advertencias_tipos = _normalizar_tipos(df)
    errores_post = _validar_estructura(df)
    if errores_post:
      return ResultadoCarga(exito=False, errores=errores_post, advertencias=advertencias_mapeo + advertencias_tipos, nombre_archivo=nombre_archivo)
    return ResultadoCarga(exito=True, df=df.reset_index(drop=True), advertencias=advertencias_mapeo + advertencias_tipos, nombre_archivo=nombre_archivo)


def decodificar_contenido_upload(contents):
  _, contenido_b64 = contents.split(",", 1)
  return base64.b64decode(contenido_b64)
  
