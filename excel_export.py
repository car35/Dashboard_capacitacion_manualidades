"""Generacion del reporte ejecutivo en Excel (.xlsx) con XlsxWriter."""

from __future__ import annotations

import io
from datetime import datetime

import pandas as pd
import xlsxwriter

from kpi_engine import KPIsGenerales

COLOR_MARCA = "#1B4965"
COLOR_MARCA_CLARO = "#5FA8D3"
COLOR_SI = "#2BB673"
COLOR_NO = "#E4572E"
FUENTE = "Arial"


def _formatos(libro):
  return {
  "titulo": libro.add_format({"bold": True, "font_size": 18, "font_color": "white", "bg_color": COLOR_MARCA, "align": "left", "valign": "vcenter", "font_name": FUENTE}),
  "subtitulo": libro.add_format({"italic": True, "font_size": 10, "font_color": "#5A6472", "font_name": FUENTE}),
  "encabezado": libro.add_format({"bold": True, "font_color": "white", "bg_color": COLOR_MARCA, "align": "center", "valign": "vcenter", "border": 1, "font_name": FUENTE}),
  "celda": libro.add_format({"border": 1, "font_name": FUENTE, "font_size": 10}),
  "celda_pct": libro.add_format({"border": 1, "font_name": FUENTE, "font_size": 10, "num_format": "0.0%"}),
  "celda_num": libro.add_format({"border": 1, "font_name": FUENTE, "font_size": 10, "num_format": "#,##0"}),
  "kpi_label": libro.add_format({"font_size": 10, "font_color": "#5A6472", "font_name": FUENTE}),
  "kpi_valor": libro.add_format({"bold": True, "font_size": 20, "font_color": COLOR_MARCA, "font_name": FUENTE}),
  "kpi_valor_texto": libro.add_format({"bold": True, "font_size": 13, "font_color": COLOR_MARCA, "font_name": FUENTE}),
  }


def _hoja_resumen(libro, fmt, kpis, nombre_fuente):
  hoja = libro.add_worksheet("Resumen Ejecutivo")
  hoja.hide_gridlines(2)
  hoja.set_column("A:A", 3)
  hoja.set_column("B:I", 18)
  hoja.merge_range("B2:I2", "REPORTE EJECUTIVO DE PRODUCTIVIDAD DE LLAMADAS", fmt["titulo"])
  hoja.set_row(1, 32)
  hoja.write("B3", f"Generado el {datetime.now():%Y-%m-%d %H:%M}  -  Fuente: {nombre_fuente}", fmt["subtitulo"])
  tarjetas = [
  ("Total de llamadas", f"{kpis.total_llamadas:,}"),
  ("Llamadas efectivas (SI)", f"{kpis.total_si:,}"),
  ("Llamadas no efectivas (NO)", f"{kpis.total_no:,}"),
  ("% Efectividad general", f"{kpis.pct_efectividad:.1f}%"),
  ("Total responsables", f"{kpis.total_responsables:,}"),
  ("Total comunas", f"{kpis.total_comunas:,}"),
  ("Promedio efectividad (por responsable)", f"{kpis.promedio_efectividad_responsables:.1f}%"),
  ("Mejor comuna", f"{kpis.mejor_comuna} ({kpis.mejor_comuna_pct:.1f}%)"),
  ("Peor comuna", f"{kpis.peor_comuna} ({kpis.peor_comuna_pct:.1f}%)"),
  ("Responsable mas productivo", f"{kpis.mejor_responsable} ({kpis.mejor_responsable_pct:.1f}%)"),
  ("Responsable con menor productividad", f"{kpis.peor_responsable} ({kpis.peor_responsable_pct:.1f}%)"),
  ]
  fila, col = 5, 1
  for i, (etiqueta, valor) in enumerate(tarjetas):
    f = fila + (i // 3) * 3
    c = col + (i % 3) * 3
    hoja.merge_range(f, c, f, c + 1, etiqueta, fmt["kpi_label"])
    formato_valor = fmt["kpi_valor"] if len(valor) <= 8 else fmt["kpi_valor_texto"]
    hoja.merge_range(f + 1, c, f + 1, c + 1, valor, formato_valor)
  hoja.write(fila + 14, 1, "Este reporte se genero automaticamente desde el Dashboard de Productividad.", fmt["subtitulo"])


def _hoja_comunas(libro, fmt, tabla_comunas):
  hoja = libro.add_worksheet("Comunas")
  hoja.set_column("A:A", 22)
  hoja.set_column("B:F", 16)
  columnas = ["Comuna", "Total_Llamadas", "Total_SI", "Total_NO", "Pct_Efectividad", "Pct_Participacion"]
  encabezados = ["Comuna", "Total Llamadas", "Total SI", "Total NO", "% Efectividad", "% Participacion"]
  for c, encabezado in enumerate(encabezados):
    hoja.write(0, c, encabezado, fmt["encabezado"])
  for r, (_, fila) in enumerate(tabla_comunas.iterrows(), start=1):
    hoja.write(r, 0, fila["Comuna"], fmt["celda"])
    hoja.write(r, 1, int(fila["Total_Llamadas"]), fmt["celda_num"])
    hoja.write(r, 2, int(fila["Total_SI"]), fmt["celda_num"])
    hoja.write(r, 3, int(fila["Total_NO"]), fmt["celda_num"])
    hoja.write(r, 4, fila["Pct_Efectividad"] / 100, fmt["celda_pct"])
    hoja.write(r, 5, fila["Pct_Participacion"] / 100, fmt["celda_pct"])
  n = len(tabla_comunas)
  grafico = libro.add_chart({"type": "bar"})
  grafico.add_series({"name": "Total SI", "categories": ["Comunas", 1, 0, n, 0], "values": ["Comunas", 1, 2, n, 2], "fill": {"color": COLOR_SI}})
  grafico.add_series({"name": "Total NO", "categories": ["Comunas", 1, 0, n, 0], "values": ["Comunas", 1, 3, n, 3], "fill": {"color": COLOR_NO}})
  grafico.set_title({"name": "Llamadas efectivas vs no efectivas por comuna"})
  grafico.set_size({"width": 640, "height": 420})
  hoja.insert_chart("H2", grafico)


def _hoja_responsables(libro, fmt, tabla_resp):
  hoja = libro.add_worksheet("Responsables")
  hoja.set_column("A:A", 22)
  hoja.set_column("B:E", 16)
  encabezados = ["Responsable", "Total Llamadas", "Total SI", "Total NO", "% Productividad"]
  for c, encabezado in enumerate(encabezados):
    hoja.write(0, c, encabezado, fmt["encabezado"])
  for r, (_, fila) in enumerate(tabla_resp.iterrows(), start=1):
    hoja.write(r, 0, fila["Responsable"], fmt["celda"])
    hoja.write(r, 1, int(fila["Total_Llamadas"]), fmt["celda_num"])
    hoja.write(r, 2, int(fila["Total_SI"]), fmt["celda_num"])
    hoja.write(r, 3, int(fila["Total_NO"]), fmt["celda_num"])
    hoja.write(r, 4, fila["Pct_Productividad"] / 100, fmt["celda_pct"])
  n = len(tabla_resp)
  grafico = libro.add_chart({"type": "column"})
  grafico.add_series({"name": "% Productividad", "categories": ["Responsables", 1, 0, n, 0], "values": ["Responsables", 1, 4, n, 4], "fill": {"color": COLOR_MARCA_CLARO}, "data_labels": {"value": True, "num_format": "0.0%"}})
  grafico.set_title({"name": "% de productividad por responsable"})
  grafico.set_y_axis({"num_format": "0%"})
  grafico.set_size({"width": 640, "height": 420})
  hoja.insert_chart("G2", grafico)


def _hoja_pivote(libro, fmt, pivote_pct):
  hoja = libro.add_worksheet("Responsable-Comuna")
  hoja.set_column("A:A", 22)
  if pivote_pct.empty:
    hoja.write(0, 0, "Sin datos suficientes para la tabla dinamica.", fmt["celda"])
    return
  hoja.write(0, 0, "Responsable / Comuna (% efectividad)", fmt["encabezado"])
  for c, comuna in enumerate(pivote_pct.columns, start=1):
    hoja.write(0, c, comuna, fmt["encabezado"])
    hoja.set_column(c, c, 16)
  for r, (responsable, fila) in enumerate(pivote_pct.iterrows(), start=1):
    hoja.write(r, 0, responsable, fmt["celda"])
    for c, valor in enumerate(fila, start=1):
      if pd.isna(valor):
        hoja.write_blank(r, c, None, fmt["celda"])
      else:
        hoja.write(r, c, valor / 100, fmt["celda_pct"])


def generar_reporte_excel(df, kpis, tabla_comunas, tabla_resp, pivote_pct, nombre_fuente="Dashboard de Productividad"):
  """Devuelve los bytes de un .xlsx ejecutivo listo para descargar."""
  buffer = io.BytesIO()
  libro = xlsxwriter.Workbook(buffer, {"in_memory": True})
  fmt = _formatos(libro)
  _hoja_resumen(libro, fmt, kpis, nombre_fuente)
  _hoja_comunas(libro, fmt, tabla_comunas)
  _hoja_responsables(libro, fmt, tabla_resp)
  _hoja_pivote(libro, fmt, pivote_pct)
  libro.close()
  buffer.seek(0)
  return buffer.read()


COLUMNAS_BD_COMPONENTE = ["fecha_registro", "nacionalidad", "tipo_de_documento", "numero_de_documento", "nombre", "correo_electronico", "telefono_uno", "telefono_dos", "fecha_de_nacimiento", "edad", "sexo", "direccion", "comuna", "barrio", "estrato", "formacion_amautta", "club_vida_registro", "Actual_club_vida", "Comuna_Club_Vida", "RESPONSABLE LLAMADA", "LLAMADA EFECTIVA", "ACTUALIZACION DATOS", "INTERESADO EN INGRESAR", "DETALLE LLAMADA", "RESUMEN LLAMADA", "OBSERVACION", "GUION HERRAMIENTA", "Orden_Efectiva (auxiliar - no borrar)", "Orden_NoEfectiva (auxiliar - no borrar)", "Fecha_Hora_Registro_Efectiva (auxiliar - no borrar)"]


def _nombre_hoja_valido(texto):
  invalidos = ["/", "\\", "*", "?", "[", "]", ":"]
  limpio = str(texto)
  for caracter in invalidos:
    limpio = limpio.replace(caracter, "-")
    return limpio[:31] if limpio else "Responsable"


def _fila_bd_componente(registro):
  fecha = registro.get("Fecha")
  fecha_texto = fecha.strftime("%Y-%m-%d") if pd.notna(fecha) else ""
  efectiva = str(registro.get("Llamada_Efectiva", "") or "")
  fecha_efectiva = fecha.strftime("%d/%m/%Y") if pd.notna(fecha) and efectiva == "SI" else ""
  return [fecha_texto, "", "", "", registro.get("Cliente", ""), "", "", "", "", "", "", "", registro.get("Comuna", ""), "", "", "", "", "", "", registro.get("Responsable", ""), efectiva, "", "", "", "", registro.get("Observaciones", ""), "", "", "", fecha_efectiva]


def generar_exportacion_responsable(df):
  """Genera un Excel con una hoja por responsable, con las llamadas efectivas y no efectivas en el formato de columnas de BD_Componente #3, listo para subir al archivo maestro."""
  buffer = io.BytesIO()
  libro = xlsxwriter.Workbook(buffer, {"in_memory": True})
  fmt_encabezado = libro.add_format({"bold": True, "font_color": "white", "bg_color": COLOR_MARCA, "border": 1, "font_name": FUENTE, "text_wrap": True})
  fmt_celda = libro.add_format({"border": 1, "font_name": FUENTE, "font_size": 10})
  responsables = sorted(df["Responsable"].dropna().unique()) if "Responsable" in df.columns else []
  if not responsables:
    libro.add_worksheet("Sin datos")
  for responsable in responsables:
    datos_resp = df[df["Responsable"] == responsable]
    if "Fecha" in datos_resp.columns:
      datos_resp = datos_resp.sort_values("Fecha")
    hoja = libro.add_worksheet(_nombre_hoja_valido(responsable))
    for col, encabezado in enumerate(COLUMNAS_BD_COMPONENTE):
      hoja.write(0, col, encabezado, fmt_encabezado)
      hoja.set_column(col, col, 18)
    hoja.set_row(0, 32)
    for fila, (_, registro) in enumerate(datos_resp.iterrows(), start=1):
      valores = _fila_bd_componente(registro)
      for col, valor in enumerate(valores):
        hoja.write(fila, col, valor, fmt_celda)
  libro.close()
  buffer.seek(0)
    return buffer.read()
              
