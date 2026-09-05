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
  "Responsable": ["responsable", "agente", "asesor", "gestor", "responsablellamada", "responsablellamada"],
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
  for col in df.columns:
    if _normalizar_texto(col).startswith("fechahoraregistroefectiva"):
      mapa_normalizado[col] = "Fecha"
      break
  for canonico, alias in ALIAS_COLUMNAS.items():
    if canonico in mapa_normalizado.values():
      continue
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


def _normalizar_fecha_mixta(serie):
  def convertir(valor):
    if pd.isna(valor):
      return pd.NaT
    if isinstance(valor, pd.Timestamp):
      return valor
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
      return pd.to_datetime(valor, unit="D", origin="1899-12-30")
    if isinstance(valor, str):
      texto = valor.strip()
      try:
        numero = float(texto)
        return pd.to_datetime(numero, unit="D", origin="1899-12-30")
      except ValueError:
        pass
    return pd.to_datetime(valor, errors="coerce")
  return serie.map(convertir)


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
    df["Fecha"] = _normalizar_fecha_mixta(df["Fecha"])
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


def _leer_excel_todas_hojas(fuente):
  hojas = pd.read_excel(fuente, sheet_name=None)
  marcos = [df for df in hojas.values() if not df.empty]
  if not marcos:
    return pd.DataFrame()
  return pd.concat(marcos, ignore_index=True, sort=False)


def cargar_desde_bytes(contenido, nombre_archivo):
  try:
    if nombre_archivo.lower().endswith(".csv"):
      df = pd.read_csv(io.BytesIO(contenido))
    elif nombre_archivo.lower().endswith((".xlsx", ".xls", ".xlsm")):
      df = _leer_excel_todas_hojas(io.BytesIO(contenido))
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
      df = _leer_excel_todas_hojas(ruta)
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
  


HOJAS_VALIDACION_DETALLE = {
  "faltantes": "Llamadas Faltantes",
  "mal_tipificadas": "Llamadas Mal Tipificadas",
  "inconsistentes": "Datos Inconsistentes",
  "requiere_revision": "Requiere Revision",
  "sin_datos": "Variables con Sin Datos",
}


def _tabla_resumen_responsable_validacion(hojas):
  resumen = hojas.get("Resumen por Responsable")
  if resumen is None or resumen.empty:
    return pd.DataFrame()
  col0 = resumen.columns[0]
  mascara_corte = resumen[col0].astype(str).str.contains("DESAGREGA", na=False)
  indices_corte = resumen[mascara_corte].index
  tabla = resumen.iloc[: indices_corte[0]].copy() if len(indices_corte) else resumen.copy()
  tabla = tabla.dropna(subset=[col0])
  return tabla.reset_index(drop=True)


def cargar_validacion(fuente):
  hojas = pd.read_excel(fuente, sheet_name=None)
  resultado = {"resumen": _tabla_resumen_responsable_validacion(hojas)}
  for clave, nombre_hoja in HOJAS_VALIDACION_DETALLE.items():
    resultado[clave] = hojas.get(nombre_hoja, pd.DataFrame())
  return resultado


# ==========================================================================
# Modulo de autenticacion: usuarios, proyectos y permisos (multi-proyecto)
# ==========================================================================

import os as _os
from datetime import datetime as _datetime

from flask_login import UserMixin as _UserMixin
from sqlalchemy import create_engine as _create_engine, Column as _Column, Integer as _Integer, String as _String, Text as _Text, Boolean as _Boolean, DateTime as _DateTime, ForeignKey as _ForeignKey, UniqueConstraint as _UniqueConstraint
from sqlalchemy.orm import declarative_base as _declarative_base, relationship as _relationship, sessionmaker as _sessionmaker
from werkzeug.security import generate_password_hash as _generate_password_hash, check_password_hash as _check_password_hash

BaseAuth = _declarative_base()


class Usuario(BaseAuth, _UserMixin):
  __tablename__ = "usuarios"
  id = _Column(_Integer, primary_key=True)
  correo = _Column(_String(255), unique=True, nullable=False)
  nombre = _Column(_String(255), nullable=False)
  contrasena_hash = _Column(_String(255), nullable=False)
  es_administrador = _Column(_Boolean, default=False)
  fecha_creacion = _Column(_DateTime, default=_datetime.utcnow)

  def set_password(self, contrasena):
    self.contrasena_hash = _generate_password_hash(contrasena)

  def check_password(self, contrasena):
    return _check_password_hash(self.contrasena_hash, contrasena)

  def get_id(self):
    return str(self.id)


class Proyecto(BaseAuth):
  __tablename__ = "proyectos"
  id = _Column(_Integer, primary_key=True)
  nombre = _Column(_String(255), nullable=False)
  descripcion = _Column(_Text)
  propietario_id = _Column(_Integer, _ForeignKey("usuarios.id"))
  datos_json = _Column(_Text)
  nombre_archivo_original = _Column(_String(255))
  fecha_carga = _Column(_DateTime, default=_datetime.utcnow)
  fecha_actualizacion = _Column(_DateTime, default=_datetime.utcnow)
  token_publico = _Column(_String(64), unique=True, nullable=True)
  propietario = _relationship("Usuario")


class PermisoProyecto(BaseAuth):
  __tablename__ = "permisos_proyecto"
  id = _Column(_Integer, primary_key=True)
  proyecto_id = _Column(_Integer, _ForeignKey("proyectos.id", ondelete="CASCADE"))
  usuario_id = _Column(_Integer, _ForeignKey("usuarios.id", ondelete="CASCADE"))
  rol = _Column(_String(20), nullable=False, default="visualizador")
  fecha_otorgado = _Column(_DateTime, default=_datetime.utcnow)
  __table_args__ = (_UniqueConstraint("proyecto_id", "usuario_id"),)


_engine_cache = None


def obtener_engine():
  global _engine_cache
  if _engine_cache is None:
    url_bd = _os.environ.get("DATABASE_URL", "sqlite:///dashboard_local.db")
    if url_bd.startswith("postgres://"):
      url_bd = url_bd.replace("postgres://", "postgresql://", 1)
    _engine_cache = _create_engine(url_bd, pool_size=3, max_overflow=2, pool_pre_ping=True, pool_recycle=280)
  return _engine_cache


def inicializar_bd(engine):
  BaseAuth.metadata.create_all(engine)


def obtener_sesion(engine):
  Sesion = _sessionmaker(bind=engine)
  return Sesion()


def crear_usuario(sesion, correo, nombre, contrasena, es_administrador=False):
  usuario = Usuario(correo=correo.strip().lower(), nombre=nombre, es_administrador=es_administrador)
  usuario.set_password(contrasena)
  sesion.add(usuario)
  sesion.commit()
  return usuario


def verificar_credenciales(sesion, correo, contrasena):
  usuario = sesion.query(Usuario).filter_by(correo=correo.strip().lower()).first()
  if usuario and usuario.check_password(contrasena):
    return usuario
  return None


def crear_proyecto(sesion, nombre, propietario_id, descripcion=None):
  proyecto = Proyecto(nombre=nombre, descripcion=descripcion, propietario_id=propietario_id)
  sesion.add(proyecto)
  sesion.commit()
  permiso = PermisoProyecto(proyecto_id=proyecto.id, usuario_id=propietario_id, rol="propietario")
  sesion.add(permiso)
  sesion.commit()
  return proyecto


def compartir_proyecto(sesion, proyecto_id, usuario_id, rol="visualizador"):
  existente = sesion.query(PermisoProyecto).filter_by(proyecto_id=proyecto_id, usuario_id=usuario_id).first()
  if existente:
    existente.rol = rol
  else:
    sesion.add(PermisoProyecto(proyecto_id=proyecto_id, usuario_id=usuario_id, rol=rol))
  sesion.commit()


def proyectos_visibles_para(sesion, usuario_id):
  return (sesion.query(Proyecto).join(PermisoProyecto, PermisoProyecto.proyecto_id == Proyecto.id).filter(PermisoProyecto.usuario_id == usuario_id).all())


def rol_de_usuario_en_proyecto(sesion, usuario_id, proyecto_id):
  permiso = sesion.query(PermisoProyecto).filter_by(usuario_id=usuario_id, proyecto_id=proyecto_id).first()
  return permiso.rol if permiso else None


def puede_editar(sesion, usuario_id, proyecto_id):
  rol = rol_de_usuario_en_proyecto(sesion, usuario_id, proyecto_id)
  return rol in ("propietario", "editor")

def obtener_o_crear_token_publico(sesion, proyecto):
  import secrets
  if proyecto.token_publico:
    return proyecto.token_publico
  proyecto.token_publico = secrets.token_urlsafe(16)
  sesion.commit()
  return proyecto.token_publico
