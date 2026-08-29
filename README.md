# Dashboard de Productividad de Gestion de Llamadas

Aplicacion web (Dash + Plotly + Pandas) para controlar, analizar y medir la
productividad de responsables asignados a comunas, a partir de llamadas
efectivas (SI) y no efectivas (NO). Incluye KPIs en tiempo real, analitica
por comuna y por responsable, la relacion Responsable-Comuna, y exportacion
de un reporte ejecutivo en Excel con graficos nativos.

## 1. Instalacion

Requiere Python 3.10+.

```bash
cd dashboard_productividad
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

## 2. Ejecutar

```bash
python app.py
```

Abre tu navegador en http://127.0.0.1:8050

Al abrir por primera vez, el dashboard carga automaticamente un dataset de
ejemplo (data/sample_llamadas.xlsx, 1800 registros sinteticos) para que
puedas explorar todas las funciones sin necesitar tu propio archivo.

Para regenerar o ampliar el dataset de ejemplo:

```bash
python sample_data_generator.py
```

## 3. Cargar tus propios datos

Haz clic en "Cargar archivo" (esquina superior derecha) y selecciona un
.xlsx o .csv con, como minimo, estas columnas: Responsable, Comuna,
Llamada_Efectiva (SI/NO), y opcionalmente Fecha, Cliente y Observaciones.

## 4. Que incluye el dashboard

- Barra de filtros por fecha, responsable y comuna.
- 11 tarjetas de KPI: total llamadas, SI/NO, porcentaje de efectividad,
  responsables, comunas, promedio de efectividad, mejor y peor comuna,
    responsable mas y menos productivo.
    - Analisis por Comuna: barras SI/NO, participacion, mapa de calor, Top 10.
    - Analisis por Responsable: ranking, comparativo, tendencia historica.
    - Relacion Responsable-Comuna: mapa de calor y tabla dinamica.
    - Exportar reporte Excel con graficos nativos, respetando los filtros.

    ## 5. Despliegue en produccion

    Este repositorio ya incluye Procfile y runtime.txt listos para plataformas
    como Render, Railway o Heroku. Comando de arranque: gunicorn app:server
    
