from flask import Flask, render_template_string, request, send_file, jsonify
import requests
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
import io
from datetime import datetime, timedelta
import random
import os
import json
import math
import pandas as pd

app = Flask(__name__)

# CONFIGURACIÓN
API_KEY = "7bd626cb4d3874faf995ec075af15d2cd35ec99d"
BASE_URL = "https://gps.idttecnologias.mx/api/v1"
COMPANY_ID = "87534"
TIMEZONE_OFFSET = -7
GEO_FILE = "kowi_principales.csv"

# CARGA DE GEOCERCAS MAESTRAS AL INICIO
def cargar_geocercas_maestras():
    try:
        df = pd.read_csv(GEO_FILE, encoding='utf-8')
        return df.to_dict('records')
    except:
        return []

GEOCERCAS_MAESTRAS = cargar_geocercas_maestras()

# ... (HTML_INTERFACE y rutas se mantienen igual) ...

@app.route('/generar_excel')
def generar_excel():
    # ... (código previo) ...
    # Sustituir la lógica de obtener_geocerca por esta:
    def obtener_geocerca(lat, lng, address=""):
        for g in GEOCERCAS_MAESTRAS:
            # Distancia euclidiana rápida
            dist_m = math.sqrt((lat - float(g['lat']))**2 + (lng - float(g['lng']))**2) * 111000
            if dist_m <= float(g.get('radius', 200)):
                return g['id'], g['name']
        return None, "Fuera de geocerca"
    
    # ... (resto del código de generación de Excel) ...
