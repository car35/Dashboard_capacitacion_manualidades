"""Dashboard Web de Productividad de Gestion de Llamadas."""

from __future__ import annotations

import io
import os
from datetime import datetime

import dash
from flask import request, redirect
from flask_login import LoginManager, current_user, login_required, login_user, logout_user
import dash_bootstrap_components as dbc
import pandas as pd
import plotly.graph_objects as go
from dash import Input, Output, State, dash_table, dcc, html
from dash.exceptions import PreventUpdate

import charts
import kpi_engine
import data_loader
from data_loader import cargar_desde_bytes, cargar_desde_ruta, cargar_validacion, decodificar_contenido_upload
from excel_export import generar_reporte_excel, generar_exportacion_responsable_zip

RUTA_DATOS_EJEMPLO = "data/sample_llamadas.xlsx"

app = dash.Dash(__name__, external_stylesheets=[dbc.themes.FLATLY, dbc.icons.BOOTSTRAP], title="Dashboard de Productividad de Llamadas", suppress_callback_exceptions=True)
server = app.server
server.secret_key = os.environ.get("SECRET_KEY", "clave-temporal-cambiar-en-produccion")

login_manager = LoginManager()
login_manager.init_app(server)
login_manager.login_view = "login"


@login_manager.user_loader
def cargar_usuario_sesion(user_id):
  try:
    engine_auth = data_loader.obtener_engine()
    data_loader.inicializar_bd(engine_auth)
    sesion_auth = data_loader.obtener_sesion(engine_auth)
    usuario = sesion_auth.get(data_loader.Usuario, int(user_id))
    sesion_auth.close()
    return usuario
  except Exception:
    return None


RUTAS_PUBLICAS = ("/login", "/logout", "/configuracion-inicial", "/verificar-bd", "/assets")


@server.before_request
def exigir_login():
  if any(request.path.startswith(ruta) for ruta in RUTAS_PUBLICAS):
    return None
  if not current_user.is_authenticated:
    return redirect("/login")


@server.route("/login", methods=["GET", "POST"])
def login():
  if request.method == "POST":
    correo = request.form.get("correo", "").strip()
    contrasena = request.form.get("contrasena", "")
    engine_auth = data_loader.obtener_engine()
    sesion_auth = data_loader.obtener_sesion(engine_auth)
    usuario = data_loader.verificar_credenciales(sesion_auth, correo, contrasena)
    sesion_auth.close()
    if usuario:
      login_user(usuario)
      return redirect("/")
    return """
    <html><body style="font-family: sans-serif; max-width: 400px; margin: 60px auto;">
    <h2>Iniciar sesion</h2>
    <p style="color: red;">Correo o contrasena incorrectos.</p>
    <form method="POST">
      <label>Correo:<br><input type="email" name="correo" required style="width:100%; padding:8px; margin-bottom:12px;"></label><br>
      <label>Contrasena:<br><input type="password" name="contrasena" required style="width:100%; padding:8px; margin-bottom:12px;"></label><br>
      <button type="submit" style="padding:10px 20px;">Entrar</button>
    </form>
    </body></html>
    """, 401
  return """
  <html><body style="font-family: sans-serif; max-width: 400px; margin: 60px auto;">
  <h2>Iniciar sesion</h2>
  <form method="POST">
    <label>Correo:<br><input type="email" name="correo" required style="width:100%; padding:8px; margin-bottom:12px;"></label><br>
    <label>Contrasena:<br><input type="password" name="contrasena" required style="width:100%; padding:8px; margin-bottom:12px;"></label><br>
    <button type="submit" style="padding:10px 20px;">Entrar</button>
  </form>
  </body></html>
  """


@server.route("/logout")
def logout():
  logout_user()
  return redirect("/login")


@server.route("/verificar-bd")
def verificar_bd():
  try:
    from sqlalchemy import text
    engine = data_loader.obtener_engine()
    data_loader.inicializar_bd(engine)
    sesion = data_loader.obtener_sesion(engine)
    sesion.execute(text("SELECT 1"))
    sesion.close()
    return "OK"
  except Exception:
    return "Error", 500



@server.route("/admin/respaldo")
def respaldo_completo():
  if not current_user.is_authenticated or not current_user.es_administrador:
    return "No autorizado", 403
  import zipfile
  engine_r = data_loader.obtener_engine()
  sesion_r = data_loader.obtener_sesion(engine_r)
  buffer_zip = io.BytesIO()
  with zipfile.ZipFile(buffer_zip, "w", zipfile.ZIP_DEFLATED) as zf:
    usuarios = sesion_r.query(data_loader.Usuario).all()
    filas_usuarios = [{"Nombre": u.nombre, "Correo": u.correo, "Administrador": u.es_administrador, "Fecha_creacion": str(u.fecha_creacion)} for u in usuarios]
    buffer_usuarios = io.BytesIO()
    pd.DataFrame(filas_usuarios).to_excel(buffer_usuarios, index=False)
    zf.writestr("usuarios.xlsx", buffer_usuarios.getvalue())
    proyectos = sesion_r.query(data_loader.Proyecto).all()
    for proyecto in proyectos:
      if not proyecto.datos_json:
        continue
      df_p = pd.read_json(io.StringIO(proyecto.datos_json), orient="split")
      buffer_proy = io.BytesIO()
      df_p.to_excel(buffer_proy, index=False)
      nombre_limpio = "".join(c if c.isalnum() or c in " -_" else "-" for c in proyecto.nombre)[:80]
      zf.writestr(f"proyectos/{nombre_limpio}.xlsx", buffer_proy.getvalue())
    permisos = sesion_r.query(data_loader.PermisoProyecto).all()
    filas_permisos = [{"Proyecto_id": pp.proyecto_id, "Usuario_id": pp.usuario_id, "Rol": pp.rol} for pp in permisos]
    buffer_permisos = io.BytesIO()
    pd.DataFrame(filas_permisos).to_excel(buffer_permisos, index=False)
    zf.writestr("permisos.xlsx", buffer_permisos.getvalue())
  sesion_r.close()
  buffer_zip.seek(0)
  nombre_zip = f"Respaldo_Dashboard_{datetime.now():%Y%m%d_%H%M}.zip"
  return buffer_zip.getvalue(), 200, {"Content-Type": "application/zip", "Content-Disposition": f"attachment; filename={nombre_zip}"}






@server.route("/configuracion-inicial", methods=["GET", "POST"])
def configuracion_inicial():
  engine = data_loader.obtener_engine()
  data_loader.inicializar_bd(engine)
  sesion = data_loader.obtener_sesion(engine)
  total_usuarios = sesion.query(data_loader.Usuario).count()
  if total_usuarios > 0:
    sesion.close()
    return "Ya existe al menos un usuario registrado. Esta pagina de configuracion inicial ya no esta disponible.", 403
  if request.method == "POST":
    correo = request.form.get("correo", "").strip()
    nombre = request.form.get("nombre", "").strip()
    contrasena = request.form.get("contrasena", "")
    if not correo or not nombre or len(contrasena) < 8:
      sesion.close()
      return "Datos invalidos. La contrasena debe tener al menos 8 caracteres. <a href=\"/configuracion-inicial\">Volver</a>", 400
    data_loader.crear_usuario(sesion, correo, nombre, contrasena, es_administrador=True)
    sesion.close()
    return "Cuenta de administrador creada correctamente. Ya puedes cerrar esta pagina e iniciar sesion."
  sesion.close()
  return """
  <html><body style="font-family: sans-serif; max-width: 400px; margin: 60px auto;">
  <h2>Configuracion inicial</h2>
  <p>Crea la primera cuenta de administrador del dashboard.</p>
  <form method="POST">
    <label>Nombre:<br><input type="text" name="nombre" required style="width:100%; padding:8px; margin-bottom:12px;"></label><br>
    <label>Correo:<br><input type="email" name="correo" required style="width:100%; padding:8px; margin-bottom:12px;"></label><br>
    <label>Contrasena (minimo 8 caracteres):<br><input type="password" name="contrasena" required minlength="8" style="width:100%; padding:8px; margin-bottom:12px;"></label><br>
    <button type="submit" style="padding:10px 20px;">Crear cuenta</button>
  </form>
  </body></html>
  """

def tarjeta_kpi(kpi_id, etiqueta, icono):
  return dbc.Card(dbc.CardBody([html.Div([html.I(className=f"bi {icono} kpi-icono"), html.Span(etiqueta, className="kpi-etiqueta")], className="kpi-encabezado"), html.Div(id=kpi_id, className="kpi-valor")]), className="tarjeta-kpi shadow-sm")


FILA_KPIS = [("kpi-total-llamadas", "Total de llamadas", "bi-telephone"), ("kpi-total-si", "Llamadas efectivas (SI)", "bi-check-circle"), ("kpi-total-no", "Llamadas no efectivas (NO)", "bi-x-circle"), ("kpi-pct-efectividad", "% Efectividad general", "bi-speedometer2"), ("kpi-meta-avance", "% Avance hacia la meta (490)", "bi-flag"), ("kpi-total-responsables", "Total responsables", "bi-people"), ("kpi-total-comunas", "Total comunas", "bi-geo-alt"), ("kpi-promedio-efectividad", "Promedio de efectividad", "bi-bar-chart-line"), ("kpi-mejor-comuna", "Mejor comuna", "bi-trophy"), ("kpi-peor-comuna", "Comuna a reforzar", "bi-exclamation-triangle"), ("kpi-mejor-responsable", "Responsable mas productivo", "bi-star"), ("kpi-peor-responsable", "Responsable a reforzar", "bi-arrow-down-circle")]


def barra_filtros():
  return dbc.Card(dbc.CardBody(dbc.Row([dbc.Col([html.Label("Rango de fechas", className="filtro-label"), dcc.DatePickerRange(id="filtro-fechas", display_format="YYYY-MM-DD", className="w-100")], md=4), dbc.Col([html.Label("Responsable(s)", className="filtro-label"), dcc.Dropdown(id="filtro-responsable", multi=True, placeholder="Todos")], md=4), dbc.Col([html.Label("Comuna(s)", className="filtro-label"), dcc.Dropdown(id="filtro-comuna", multi=True, placeholder="Todas")], md=4)], className="g-3")), className="shadow-sm mb-3")


def encabezado():
  return dbc.Navbar(dbc.Container([html.Div([html.I(className="bi bi-graph-up-arrow me-2"), html.Span("Productividad de Gestion de Llamadas", className="fw-bold")], className="navbar-brand d-flex align-items-center"), html.Div([dcc.Dropdown(id="selector-proyecto", placeholder="Selecciona un proyecto", clearable=False, style={"minWidth": "260px", "color": "#212529"}, className="me-2"), dcc.Input(id="nombre-proyecto-nuevo", type="text", placeholder="Nombre del proyecto nuevo", className="me-2 form-control form-control-sm", style={"minWidth": "220px", "width": "220px"}), dcc.Upload(id="upload-datos", children=dbc.Button([html.I(className="bi bi-upload me-2"), "Cargar archivo"], color="light", outline=True, size="sm"), multiple=False), dbc.Button([html.I(className="bi bi-file-earmark-excel me-2"), "Exportar reporte Excel"], id="btn-exportar-excel", color="success", size="sm", className="ms-2"), dcc.Download(id="descarga-excel"), html.Span(id="nombre-usuario-sesion", className="text-white ms-3 me-2 small"), html.A(dbc.Button([html.I(className="bi bi-box-arrow-right me-2"), "Cerrar sesion"], color="danger", outline=True, size="sm", className="ms-2"), href="/logout")], className="d-flex align-items-center")], fluid=True), color="dark", dark=True, className="mb-3 shadow-sm")


def panel_comuna():
  return html.Div([dbc.Row([dbc.Col(dcc.Graph(id="grafico-comuna-si", config={"displayModeBar": False}), md=6), dbc.Col(dcc.Graph(id="grafico-comuna-no", config={"displayModeBar": False}), md=6)], className="g-3 mb-3"), dbc.Row([dbc.Col(dcc.Graph(id="grafico-comuna-participacion", config={"displayModeBar": False}), md=6), dbc.Col(dcc.Graph(id="grafico-comuna-heatmap", config={"displayModeBar": False}), md=6)], className="g-3 mb-3"), dbc.Row([dbc.Col(dcc.Graph(id="grafico-comuna-top-mejores", config={"displayModeBar": False}), md=6), dbc.Col(dcc.Graph(id="grafico-comuna-top-peores", config={"displayModeBar": False}), md=6)], className="g-3 mb-3"), dbc.Card(dbc.CardBody([html.H6("Detalle por comuna", className="mb-3"), html.Div(id="tabla-comunas")]), className="shadow-sm")])


def panel_responsable():
  return html.Div([dbc.Row([dbc.Col(dcc.Graph(id="grafico-resp-ranking", config={"displayModeBar": False}), md=6), dbc.Col(dcc.Graph(id="grafico-resp-comparativo", config={"displayModeBar": False}), md=6)], className="g-3 mb-3"), dbc.Row(dbc.Col(dcc.Graph(id="grafico-resp-tendencia", config={"displayModeBar": False}), md=12), className="g-3 mb-3"), dbc.Card(dbc.CardBody([html.H6("Detalle por responsable", className="mb-3"), html.Div(id="tabla-responsables")]), className="shadow-sm")])


def panel_relacion():
  return html.Div([dbc.Row(dbc.Col(dcc.Graph(id="grafico-pivote-heatmap", config={"displayModeBar": False}), md=12), className="g-3 mb-3"), dbc.Card(dbc.CardBody([html.H6("Tabla dinamica: % de efectividad por Responsable y Comuna", className="mb-3"), html.Div(id="tabla-pivote")]), className="shadow-sm")])


def panel_exportar_responsable():
  return html.Div([dbc.Row(dbc.Col(dbc.Card(dbc.CardBody([html.H6("Exportar llamadas por responsable", className="mb-2"), html.P("Genera un archivo Excel con una hoja por cada responsable, con las llamadas efectivas y no efectivas en el formato de columnas de BD_Componente #3, listo para copiar y subir al archivo maestro.", className="text-muted small"), dbc.Button([html.I(className="bi bi-download me-2"), "Exportar por responsable"], id="btn-exportar-responsable", color="primary"), html.Div(id="mensaje-exportar-responsable", className="mt-2"), dcc.Download(id="descarga-excel-responsable")]), className="shadow-sm"), md=8), className="g-3 mb-3")])


def panel_validacion():
  return html.Div([dbc.Row(dbc.Col(dcc.Upload(id="upload-validacion", children=dbc.Button([html.I(className="bi bi-upload me-2"), "Cargar reporte de validacion"], color="primary", outline=True), multiple=False), md=4), className="mb-3"), html.Div(id="zona-alertas-validacion"), dbc.Row([dbc.Col(dbc.Card(dbc.CardBody([html.Div([html.I(className="bi bi-telephone me-2"), html.Span("Llamadas reales totales", className="kpi-etiqueta")], className="kpi-encabezado"), html.Div(id="val-kpi-total", className="kpi-valor")]), className="tarjeta-kpi shadow-sm"), md=3), dbc.Col(dbc.Card(dbc.CardBody([html.Div([html.I(className="bi bi-speedometer2 me-2"), html.Span("% Cumplimiento promedio", className="kpi-etiqueta")], className="kpi-encabezado"), html.Div(id="val-kpi-cumplimiento", className="kpi-valor")]), className="tarjeta-kpi shadow-sm"), md=3), dbc.Col(dbc.Card(dbc.CardBody([html.Div([html.I(className="bi bi-x-circle me-2"), html.Span("Llamadas faltantes", className="kpi-etiqueta")], className="kpi-encabezado"), html.Div(id="val-kpi-faltantes", className="kpi-valor")]), className="tarjeta-kpi shadow-sm"), md=3), dbc.Col(dbc.Card(dbc.CardBody([html.Div([html.I(className="bi bi-exclamation-triangle me-2"), html.Span("Mal tipificadas", className="kpi-etiqueta")], className="kpi-encabezado"), html.Div(id="val-kpi-mal-tipificadas", className="kpi-valor")]), className="tarjeta-kpi shadow-sm"), md=3)], className="g-3 mb-3"), dbc.Card(dbc.CardBody(dcc.Graph(id="grafico-validacion-cumplimiento", config={"displayModeBar": False})), className="shadow-sm mb-3"), dbc.Card(dbc.CardBody([html.H6("Llamadas faltantes (no encontradas en la base de datos)", className="mb-3"), html.Div(id="tabla-validacion-faltantes")]), className="shadow-sm mb-3"), dbc.Card(dbc.CardBody([html.H6("Llamadas mal tipificadas", className="mb-3"), html.Div(id="tabla-validacion-mal-tipificadas")]), className="shadow-sm")])


def panel_administracion():
  return html.Div(id="contenido-administracion")


app.layout = html.Div([dcc.Store(id="store-datos"), dcc.Store(id="store-nombre-archivo"), encabezado(), dbc.Container([html.Div(id="zona-alertas"), barra_filtros(), dbc.Row([dbc.Col(tarjeta_kpi(kpi_id, etiqueta, icono), lg=3, md=4, sm=6, xs=12, className="mb-3") for kpi_id, etiqueta, icono in FILA_KPIS], className="g-3 mb-2"), dbc.Tabs([dbc.Tab(panel_comuna(), label="Analisis por Comuna", tab_id="tab-comuna"), dbc.Tab(panel_responsable(), label="Analisis por Responsable", tab_id="tab-responsable"), dbc.Tab(panel_relacion(), label="Relacion Responsable - Comuna", tab_id="tab-relacion"), dbc.Tab(panel_exportar_responsable(), label="Exportar por Responsable", tab_id="tab-exportar"), dbc.Tab(panel_validacion(), label="Validacion de Llamadas", tab_id="tab-validacion"), dbc.Tab(panel_administracion(), label="Administracion", tab_id="tab-administracion")], id="tabs-principales", active_tab="tab-comuna", className="mt-2"), html.Footer("Dashboard de Productividad - generado con Dash, Plotly y Pandas", className="text-muted text-center small py-4")], fluid=True)])


@app.callback(Output("store-datos", "data"), Output("store-nombre-archivo", "data"), Output("zona-alertas", "children"), Output("selector-proyecto", "options"), Output("selector-proyecto", "value"), Input("upload-datos", "contents"), Input("selector-proyecto", "value"), State("upload-datos", "filename"), State("nombre-proyecto-nuevo", "value"), prevent_initial_call=False)
def gestionar_proyectos(contenido_upload, proyecto_seleccionado_id, nombre_upload, nombre_proyecto_nuevo):
  trigger_id = dash.ctx.triggered_id
  usuario_id = current_user.id if current_user.is_authenticated else None
  engine_p = data_loader.obtener_engine()
  data_loader.inicializar_bd(engine_p)
  sesion_p = data_loader.obtener_sesion(engine_p)
  proyectos_usuario = data_loader.proyectos_visibles_para(sesion_p, usuario_id) if usuario_id else []
  opciones_proyectos = [{"label": p.nombre, "value": p.id} for p in proyectos_usuario]
  if trigger_id == "upload-datos" and contenido_upload is not None:
    datos_bytes = decodificar_contenido_upload(contenido_upload)
    resultado = cargar_desde_bytes(datos_bytes, nombre_upload)
    if not resultado.exito:
      sesion_p.close()
      alerta = dbc.Alert([html.B("No se pudo cargar el archivo. "), *[html.Div(e) for e in resultado.errores]], color="danger", dismissable=True)
      return dash.no_update, dash.no_update, alerta, opciones_proyectos, dash.no_update
    nombre_final = (nombre_proyecto_nuevo or "").strip() or nombre_upload
    proyecto_nuevo = data_loader.crear_proyecto(sesion_p, nombre_final, usuario_id)
    proyecto_nuevo.datos_json = resultado.df.to_json(date_format="iso", orient="split")
    proyecto_nuevo.nombre_archivo_original = nombre_upload
    sesion_p.commit()
    opciones_proyectos = [{"label": p.nombre, "value": p.id} for p in data_loader.proyectos_visibles_para(sesion_p, usuario_id)]
    alertas = []
    if resultado.advertencias:
      alertas.append(dbc.Alert([html.B(f"Proyecto '{nombre_final}' guardado con observaciones: ")] + [html.Div(a) for a in resultado.advertencias], color="warning", dismissable=True))
    else:
      alertas.append(dbc.Alert(f"Proyecto '{nombre_final}' guardado correctamente ({len(resultado.df)} filas).", color="success", dismissable=True, duration=4000))
    sesion_p.close()
    return proyecto_nuevo.datos_json, nombre_final, alertas, opciones_proyectos, proyecto_nuevo.id
  if trigger_id == "selector-proyecto" and proyecto_seleccionado_id is not None:
    proyecto = sesion_p.query(data_loader.Proyecto).filter_by(id=proyecto_seleccionado_id).first()
    sesion_p.close()
    if proyecto is None:
      raise PreventUpdate
    alerta = dbc.Alert(f"Proyecto '{proyecto.nombre}' cargado.", color="success", dismissable=True, duration=3000)
    return proyecto.datos_json, proyecto.nombre_archivo_original, alerta, opciones_proyectos, dash.no_update
  if proyectos_usuario:
    primer_proyecto = proyectos_usuario[0]
    sesion_p.close()
    return primer_proyecto.datos_json, primer_proyecto.nombre_archivo_original, dash.no_update, opciones_proyectos, primer_proyecto.id
  resultado = cargar_desde_ruta(RUTA_DATOS_EJEMPLO)
  sesion_p.close()
  return resultado.df.to_json(date_format="iso", orient="split"), "Datos de ejemplo (sample_llamadas.xlsx)", dash.no_update, opciones_proyectos, dash.no_update


@app.callback(Output("filtro-fechas", "min_date_allowed"), Output("filtro-fechas", "max_date_allowed"), Output("filtro-fechas", "start_date"), Output("filtro-fechas", "end_date"), Input("store-datos", "data"))
def poblar_fechas(datos_json):
  if not datos_json:
    raise PreventUpdate
  df = pd.read_json(io.StringIO(datos_json), orient="split")
  if "Fecha" in df.columns and not df["Fecha"].isna().all():
    fechas = pd.to_datetime(df["Fecha"]).dropna()
    return fechas.min().date(), fechas.max().date(), fechas.min().date(), fechas.max().date()
  return None, None, None, None


@app.callback(Output("filtro-comuna", "options"), Input("store-datos", "data"), Input("filtro-responsable", "value"))
def actualizar_opciones_comuna(datos_json, responsables):
  if not datos_json:
    raise PreventUpdate
  df = pd.read_json(io.StringIO(datos_json), orient="split")
  if responsables:
    df = df[df["Responsable"].isin(responsables)]
  opciones = sorted(df["Comuna"].dropna().unique())
  return [{"label": c, "value": c} for c in opciones]


@app.callback(Output("filtro-responsable", "options"), Input("store-datos", "data"), Input("filtro-comuna", "value"))
def actualizar_opciones_responsable(datos_json, comunas):
  if not datos_json:
    raise PreventUpdate
  df = pd.read_json(io.StringIO(datos_json), orient="split")
  if comunas:
    df = df[df["Comuna"].isin(comunas)]
  opciones = sorted(df["Responsable"].dropna().unique())
  return [{"label": r, "value": r} for r in opciones]


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
      


OUTPUTS_PRINCIPALES = [Output("kpi-total-llamadas", "children"), Output("kpi-total-si", "children"), Output("kpi-total-no", "children"), Output("kpi-pct-efectividad", "children"), Output("kpi-meta-avance", "children"), Output("kpi-total-responsables", "children"), Output("kpi-total-comunas", "children"), Output("kpi-promedio-efectividad", "children"), Output("kpi-mejor-comuna", "children"), Output("kpi-peor-comuna", "children"), Output("kpi-mejor-responsable", "children"), Output("kpi-peor-responsable", "children"), Output("grafico-comuna-si", "figure"), Output("grafico-comuna-no", "figure"), Output("grafico-comuna-participacion", "figure"), Output("grafico-comuna-heatmap", "figure"), Output("grafico-comuna-top-mejores", "figure"), Output("grafico-comuna-top-peores", "figure"), Output("tabla-comunas", "children"), Output("grafico-resp-ranking", "figure"), Output("grafico-resp-comparativo", "figure"), Output("grafico-resp-tendencia", "figure"), Output("tabla-responsables", "children"), Output("grafico-pivote-heatmap", "figure"), Output("tabla-pivote", "children")]


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
    return "0", "0", "0", "0.0%", "0.0%", "0", "0", "0.0%", "-", "-", "-", "-", fig_vacia, fig_vacia, fig_vacia, fig_vacia, fig_vacia, fig_vacia, mensaje, fig_vacia, fig_vacia, fig_vacia, mensaje, fig_vacia, mensaje
  kpis = kpi_engine.calcular_kpis_generales(df_filtrado)
  tabla_comunas = kpi_engine.calcular_tabla_comunas(df_filtrado)
  tabla_resp = kpi_engine.calcular_tabla_responsables(df_filtrado)
  pivote_pct = kpi_engine.calcular_pivote_responsable_comuna(df_filtrado, metrica="pct")
  tendencia = kpi_engine.calcular_tendencia_diaria(df_filtrado)
  tabla_pivote_reset = pivote_pct.reset_index()
  return f"{kpis.total_llamadas:,}", f"{kpis.total_si:,}", f"{kpis.total_no:,}", f"{kpis.pct_efectividad:.1f}%", f"{(kpis.total_si / 490 * 100):.1f}%", f"{kpis.total_responsables:,}", f"{kpis.total_comunas:,}", f"{kpis.promedio_efectividad_responsables:.1f}%", f"{kpis.mejor_comuna} ({kpis.mejor_comuna_pct:.1f}%)", f"{kpis.peor_comuna} ({kpis.peor_comuna_pct:.1f}%)", f"{kpis.mejor_responsable} ({kpis.mejor_responsable_pct:.1f}%)", f"{kpis.peor_responsable} ({kpis.peor_responsable_pct:.1f}%)", charts.grafico_barras_comuna(tabla_comunas, "Total_SI", charts.COLOR_SI, "Llamadas efectivas (SI) por comuna"), charts.grafico_barras_comuna(tabla_comunas, "Total_NO", charts.COLOR_NO, "Llamadas no efectivas (NO) por comuna"), charts.grafico_participacion_comuna(tabla_comunas), charts.grafico_heatmap_comuna(tabla_comunas), charts.grafico_top_comunas(tabla_comunas, mejores=True), charts.grafico_top_comunas(tabla_comunas, mejores=False), _tabla_dash(tabla_comunas, columnas_pct=["Pct_Efectividad", "Pct_Participacion"]), charts.grafico_ranking_responsables(tabla_resp), charts.grafico_comparativo_responsables(tabla_resp), charts.grafico_tendencia(tendencia), _tabla_dash(tabla_resp, columnas_pct=["Pct_Productividad"]), charts.grafico_pivote_heatmap(pivote_pct), _tabla_dash(tabla_pivote_reset)
        


@app.callback(Output("descarga-excel", "data"), Input("btn-exportar-excel", "n_clicks"), State("store-datos", "data"), State("store-nombre-archivo", "data"), State("filtro-responsable", "value"), State("filtro-comuna", "value"), State("filtro-fechas", "start_date"), State("filtro-fechas", "end_date"), prevent_initial_call=True)
def exportar_excel(n_clicks, datos_json, nombre_archivo, responsables, comunas, fecha_ini, fecha_fin):
  if not datos_json:
    raise PreventUpdate
  df = pd.read_json(io.StringIO(datos_json), orient="split")
  if "Fecha" in df.columns:
    df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce")
  df_filtrado = _filtrar(df, responsables, comunas, fecha_ini, fecha_fin)
  if df_filtrado.empty:
    raise PreventUpdate
  kpis = kpi_engine.calcular_kpis_generales(df_filtrado)
  tabla_comunas = kpi_engine.calcular_tabla_comunas(df_filtrado)
  tabla_resp = kpi_engine.calcular_tabla_responsables(df_filtrado)
  pivote_pct = kpi_engine.calcular_pivote_responsable_comuna(df_filtrado, metrica="pct")
  contenido = generar_reporte_excel(df_filtrado, kpis, tabla_comunas, tabla_resp, pivote_pct, nombre_fuente=nombre_archivo or "Dashboard de Productividad")
  nombre_salida = f"Reporte_Productividad_{datetime.now():%Y%m%d_%H%M}.xlsx"
  return dcc.send_bytes(contenido, nombre_salida)


@app.callback(Output("descarga-excel-responsable", "data"), Output("mensaje-exportar-responsable", "children"), Input("btn-exportar-responsable", "n_clicks"), State("store-datos", "data"), State("filtro-responsable", "value"), State("filtro-comuna", "value"), State("filtro-fechas", "start_date"), State("filtro-fechas", "end_date"), prevent_initial_call=True)
def exportar_por_responsable(n_clicks, datos_json, responsables, comunas, fecha_ini, fecha_fin):
  if not datos_json:
    raise PreventUpdate
  df = pd.read_json(io.StringIO(datos_json), orient="split")
  if "Fecha" in df.columns:
    df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce")
  df_filtrado = _filtrar(df, responsables, comunas, fecha_ini, fecha_fin)
  if df_filtrado.empty:
    return dash.no_update, dbc.Alert("No hay llamadas para exportar con los filtros seleccionados.", color="warning", dismissable=True)
  contenido, nombre_salida = generar_exportacion_responsable_zip(df_filtrado)
  mensaje = dbc.Alert(f"Archivo generado con {df_filtrado['Responsable'].nunique()} responsable(s) y {len(df_filtrado)} llamada(s).", color="success", dismissable=True, duration=5000)
  return dcc.send_bytes(contenido, nombre_salida), mensaje
        


@app.callback(Output("zona-alertas-validacion", "children"), Output("val-kpi-total", "children"), Output("val-kpi-cumplimiento", "children"), Output("val-kpi-faltantes", "children"), Output("val-kpi-mal-tipificadas", "children"), Output("grafico-validacion-cumplimiento", "figure"), Output("tabla-validacion-faltantes", "children"), Output("tabla-validacion-mal-tipificadas", "children"), Input("upload-validacion", "contents"), State("upload-validacion", "filename"), prevent_initial_call=True)
def cargar_y_mostrar_validacion(contenido_upload, nombre_archivo):
  if contenido_upload is None:
    raise PreventUpdate
  fig_vacia = go.Figure()
  fig_vacia.add_annotation(text="Sin datos", showarrow=False, font=dict(size=14))
  fig_vacia.update_xaxes(visible=False)
  fig_vacia.update_yaxes(visible=False)
  mensaje_vacio = html.Div("Sin datos para mostrar.", className="text-muted p-3")
  try:
    datos_bytes = decodificar_contenido_upload(contenido_upload)
    resultado = cargar_validacion(io.BytesIO(datos_bytes))
    resumen = resultado["resumen"]
  except Exception as error:
    alerta = dbc.Alert(f"No se pudo procesar el archivo '{nombre_archivo}'. Verifica que sea un reporte de validacion valido.", color="danger", dismissable=True)
    return alerta, "0", "0.0%", "0", "0", fig_vacia, mensaje_vacio, mensaje_vacio
  if resumen.empty:
    alerta = dbc.Alert(f"El archivo '{nombre_archivo}' no tiene la hoja 'Resumen por Responsable' esperada.", color="danger", dismissable=True)
    return alerta, "0", "0.0%", "0", "0", fig_vacia, mensaje_vacio, mensaje_vacio
  total = int(resumen["Total llamadas reales (Llamadas 31)"].sum())
  cumplimiento = resumen["% Cumplimiento de reporte"].mean()
  faltantes_n = int(resumen["Total faltantes"].sum())
  mal_tip_n = int(resumen["Total mal tipificadas"].sum())
  fig = charts.grafico_cumplimiento_responsable(resumen)
  cols_faltantes = ["Responsable_Gestor", "Fecha_Llamada", "Nombre_Persona_Mayor", "Comuna", "Estado"]
  cols_mal_tip = ["Responsable_Gestor", "Fecha_Llamada", "Nombre_Persona_Mayor", "Comuna", "Diferencia_Encontrada"]
  df_falt = resultado["faltantes"]
  df_mal = resultado["mal_tipificadas"]
  tabla_falt = _tabla_dash(df_falt[cols_faltantes]) if not df_falt.empty else mensaje_vacio
  tabla_mal = _tabla_dash(df_mal[cols_mal_tip]) if not df_mal.empty else mensaje_vacio
  alerta = dbc.Alert(f"Archivo '{nombre_archivo}' procesado correctamente ({len(resumen)} responsable(s)).", color="success", dismissable=True, duration=4000)
  return alerta, f"{total:,}", f"{cumplimiento:.1f}%", f"{faltantes_n:,}", f"{mal_tip_n:,}", fig, tabla_falt, tabla_mal


@app.callback(Output("contenido-administracion", "children"), Input("tabs-principales", "active_tab"), prevent_initial_call=False)
def cargar_panel_administracion(tab_activo):
  if tab_activo != "tab-administracion":
    raise PreventUpdate
  if not current_user.is_authenticated or not current_user.es_administrador:
    return dbc.Alert("No tienes permisos para acceder a esta seccion.", color="danger")
  engine_a = data_loader.obtener_engine()
  try:
    sesion_a = data_loader.obtener_sesion(engine_a)
    usuarios = sesion_a.query(data_loader.Usuario).all()
    proyectos = sesion_a.query(data_loader.Proyecto).all()
    opciones_usuarios = [{"label": f"{u.nombre} ({u.correo})", "value": u.id} for u in usuarios]
    opciones_proyectos = [{"label": p.nombre, "value": p.id} for p in proyectos]
    filas_usuarios = [{"Nombre": u.nombre, "Correo": u.correo, "Administrador": "Si" if u.es_administrador else "No"} for u in usuarios]
    sesion_a.close()
    tabla_usuarios = _tabla_dash(pd.DataFrame(filas_usuarios)) if filas_usuarios else html.Div("No hay usuarios.", className="text-muted")
  except Exception as error_panel:
    return html.Div(f"Error al cargar el panel: {type(error_panel).__name__}: {error_panel}", className="text-danger p-3")
  return html.Div([
    dbc.Card(dbc.CardBody([html.H6("Crear nueva cuenta", className="mb-3"), dbc.Row([dbc.Col(dcc.Input(id="admin-nombre-nuevo", type="text", placeholder="Nombre", className="form-control"), md=3), dbc.Col(dcc.Input(id="admin-correo-nuevo", type="email", placeholder="Correo", className="form-control"), md=3), dbc.Col(dcc.Input(id="admin-contrasena-nueva", type="password", placeholder="Contrasena temporal", className="form-control"), md=3), dbc.Col(dbc.Button("Crear cuenta", id="btn-admin-crear-usuario", color="primary", className="w-100"), md=3)], className="g-2"), html.Div(id="admin-mensaje-usuario", className="mt-2")]), className="shadow-sm mb-3"),
    dbc.Card(dbc.CardBody([html.H6("Compartir proyecto", className="mb-3"), dbc.Row([dbc.Col(dcc.Dropdown(id="admin-selector-proyecto", options=opciones_proyectos, placeholder="Proyecto"), md=4), dbc.Col(dcc.Dropdown(id="admin-selector-usuario", options=opciones_usuarios, placeholder="Usuario"), md=4), dbc.Col(dcc.Dropdown(id="admin-selector-rol", options=[{"label": "Visualizador", "value": "visualizador"}, {"label": "Editor", "value": "editor"}], placeholder="Rol", value="visualizador"), md=2), dbc.Col(dbc.Button("Compartir", id="btn-admin-compartir", color="primary", className="w-100"), md=2)], className="g-2"), html.Div(id="admin-mensaje-compartir", className="mt-2")]), className="shadow-sm mb-3"),
    dbc.Card(dbc.CardBody([html.H6("Usuarios registrados", className="mb-3"), tabla_usuarios]), className="shadow-sm mb-3"),
    dbc.Card(dbc.CardBody([html.H6("Respaldo de datos", className="mb-2"), html.P("Descarga una copia completa de todos los proyectos, usuarios y permisos, por si necesitas recuperarlos.", className="text-muted small"), html.A(dbc.Button([html.I(className="bi bi-cloud-download me-2"), "Descargar respaldo completo"], color="secondary"), href="/admin/respaldo")]), className="shadow-sm"),
  ])


@app.callback(Output("admin-mensaje-usuario", "children"), Input("btn-admin-crear-usuario", "n_clicks"), State("admin-nombre-nuevo", "value"), State("admin-correo-nuevo", "value"), State("admin-contrasena-nueva", "value"), prevent_initial_call=True)
def crear_usuario_admin(n_clicks, nombre, correo, contrasena):
  if not current_user.is_authenticated or not current_user.es_administrador:
    return dbc.Alert("No autorizado.", color="danger")
  if not nombre or not correo or not contrasena or len(contrasena) < 8:
    return dbc.Alert("Completa todos los campos. La contrasena debe tener al menos 8 caracteres.", color="warning")
  engine_a = data_loader.obtener_engine()
  sesion_a = data_loader.obtener_sesion(engine_a)
  existente = sesion_a.query(data_loader.Usuario).filter_by(correo=correo.strip().lower()).first()
  if existente:
    sesion_a.close()
    return dbc.Alert("Ya existe una cuenta con ese correo.", color="warning")
  data_loader.crear_usuario(sesion_a, correo, nombre, contrasena)
  sesion_a.close()
  return dbc.Alert(f"Cuenta creada para {nombre} ({correo}).", color="success", dismissable=True, duration=5000)


@app.callback(Output("admin-mensaje-compartir", "children"), Input("btn-admin-compartir", "n_clicks"), State("admin-selector-proyecto", "value"), State("admin-selector-usuario", "value"), State("admin-selector-rol", "value"), prevent_initial_call=True)
def compartir_proyecto_admin(n_clicks, proyecto_id, usuario_id, rol):
  if not current_user.is_authenticated or not current_user.es_administrador:
    return dbc.Alert("No autorizado.", color="danger")
  if not proyecto_id or not usuario_id or not rol:
    return dbc.Alert("Selecciona proyecto, usuario y rol.", color="warning")
  engine_a = data_loader.obtener_engine()
  sesion_a = data_loader.obtener_sesion(engine_a)
  data_loader.compartir_proyecto(sesion_a, proyecto_id, usuario_id, rol)
  proyecto = sesion_a.query(data_loader.Proyecto).filter_by(id=proyecto_id).first()
  usuario = sesion_a.query(data_loader.Usuario).filter_by(id=usuario_id).first()
  nombre_proyecto = proyecto.nombre if proyecto else "?"
  nombre_usuario = usuario.nombre if usuario else "?"
  sesion_a.close()
  return dbc.Alert(f"Proyecto '{nombre_proyecto}' compartido con {nombre_usuario} como {rol}.", color="success", dismissable=True, duration=5000)


@app.callback(Output("nombre-usuario-sesion", "children"), Input("tabs-principales", "active_tab"), prevent_initial_call=False)
def mostrar_nombre_usuario(tab_activo):
  if current_user.is_authenticated:
    return f"Hola, {current_user.nombre}"
  return ""


if __name__ == "__main__":
  app.run(debug=True, host="0.0.0.0", port=8050)
  
