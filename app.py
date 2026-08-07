from flask import Flask, request, jsonify, render_template_string, send_file
import requests
import openpyxl
import io
import math
from datetime import datetime, timedelta

app = Flask(__name__)

API_KEY = "7bd626cb4d3874faf995ec075af15d2cd35ec99d"
BASE_URL = "https://gps.idttecnologias.mx/api/v1"

# --- [HTML_INTERFACE permanece igual que antes] ---
# (Copia el mismo bloque de HTML_INTERFACE que tenías)

@app.route('/generar_excel')
def generar_excel():
    unit_id = request.args.get('unit_id')
    fecha_inicio = request.args.get('fecha_inicio')
    fecha_fin = request.args.get('fecha_fin')
    hora_inicio = request.args.get('hora_inicio', '00:00')
    hora_fin = request.args.get('hora_fin', '23:59')
    
    # URL manual para evitar codificación problemática de corchetes
    query = f"key={API_KEY}&unit_id={unit_id}&from={fecha_inicio}%20{hora_inicio}:00&till={fecha_fin}%20{hora_fin}:59&include[]=points&include[]=decoded_route&include[]=stops&include[]=summary"
    url = f"{BASE_URL}/route/list.json?{query}"
    
    try:
        res = requests.get(url, timeout=30)
        route_json = res.json()
        
        # Procesar los datos (igual a la lógica anterior)
        # ... (aquí iría el procesamiento que ya funcionaba) ...
    except Exception as e:
        return jsonify({'error': f'Error de comunicación con Mapon: {str(e)}'}), 500

# ... (resto de funciones de procesamiento) ...
