"""Generador de datos de ejemplo para el Dashboard de Productividad de Llamadas."""

from __future__ import annotations

import random
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

RESPONSABLES = [
  "Dahyana Ospina",
  "Daniela Restrepo",
  "Lorena Chanci",
  "Andres Munoz",
  "Camila Velez",
]

COMUNAS = [
  "1 Popular",
  "2 Santa Cruz",
  "3 Manrique",
  "4 Aranjuez",
  "5 Castilla",
  "6 Doce de Octubre",
  "7 Robledo",
  "8 Villa Hermosa",
  "9 Buenos Aires",
  "10 La Candelaria",
  "11 Laureles Estadio",
  "12 La America",
  "13 San Javier",
  "14 El Poblado",
  "50 Palmitas",
]

NOMBRES = [
  "Diana Obando", "Liliana Uribe", "Alba Aponte", "Eugenia Gomez",
  "Fabiola Carvajal", "Martha Obando", "Alejandro Ravelo", "Genny Burbano",
  "Maria Yepes", "Flor Cardona", "Beatriz Taborda", "Maria Taborda",
  "Nora Grajales", "Ana Plata", "Luisa Restrepo", "Mariela Morales",
  "Ruth Velandia", "Brisny Arroyave", "Maritza Coromoto", "Sandra Acevedo",
]

OBSERVACIONES_SI = [
  "Llamada efectiva, se realiza contacto con la persona mayor.",
  "Contesta el usuario, agenda visita de seguimiento.",
  "Contacto exitoso, confirma datos de residencia.",
]
OBSERVACIONES_NO = [
  "No contesta la llamada.",
  "Numero fuera de servicio.",
  "Buzon de voz, se reintentara.",
  "Usuario solicita no ser contactado.",
]


def generar_datos(n_filas=1800, semilla=42):
  """Genera un DataFrame sintetico con la estructura esperada por el dashboard."""
  random.seed(semilla)
  comunas_por_responsable = {resp: random.sample(COMUNAS, k=random.randint(3, 6)) for resp in RESPONSABLES}
  fecha_inicio = datetime.today() - timedelta(days=45)
  filas = []
  for _ in range(n_filas):
    responsable = random.choice(RESPONSABLES)
    comuna = random.choice(comunas_por_responsable[responsable])
    tasa_base = {"Dahyana Ospina": 0.62, "Daniela Restrepo": 0.48, "Lorena Chanci": 0.35, "Andres Munoz": 0.55, "Camila Velez": 0.70}[responsable]
    efectiva = "SI" if random.random() < tasa_base else "NO"
    fecha = fecha_inicio + timedelta(days=random.randint(0, 45), hours=random.randint(7, 18), minutes=random.randint(0, 59))
    observaciones = random.choice(OBSERVACIONES_SI if efectiva == "SI" else OBSERVACIONES_NO)
    filas.append({"Responsable": responsable, "Comuna": comuna, "Llamada_Efectiva": efectiva, "Fecha": fecha, "Cliente": random.choice(NOMBRES), "Observaciones": observaciones})
    df = pd.DataFrame(filas).sort_values("Fecha").reset_index(drop=True)
    return df


def main():
  df = generar_datos()
  destino = Path(__file__).parent / "data" / "sample_llamadas.xlsx"
  destino.parent.mkdir(parents=True, exist_ok=True)
  df.to_excel(destino, index=False, sheet_name="Llamadas")
  print(f"Archivo de ejemplo generado en: {destino}  ({len(df)} filas)")


if __name__ == "__main__":
  main()
  
