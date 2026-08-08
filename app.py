from flask import Flask, request, jsonify, render_template_string, send_file
import requests
import openpyxl
import io
import math
from datetime import datetime, timedelta

app = Flask(__name__)

API_KEY = "7bd626cb4d3874faf995ec075af15d2cd35ec99d"
BASE_URL = "https://gps.idttecnologias.mx/api/v1"

# [HTML_INTERFACE se mantiene igual, no lo modifiqué]
HTML_INTERFACE = """...""" # (Mantén tu HTML original aquí para no alterar el diseño)

def haversine(lat1, lon1, lat2, lon2):
    R = 6371.0
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = math.sin(dlat / 2)**2 + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(dlon / 2)**2
    return R * 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))

@app.route('/')
def index():
    return render_template_string(HTML_INTERFACE)

@app.route('/api_unidades')
def api_unidades():
    try:
        res = requests.get(f"{BASE_URL}/unit/list.json", params={'key': API_KEY}, timeout=15)
        return jsonify(res.json().get('data', {}).get('units', []))
    except: return jsonify([]), 500

@app.route('/generar_excel')
def generar_excel():
    unit_id = request.args.get('unit_id')
    unit_name = request.args.get('unit_name', 'Unidad')
    f_in = request.args.get('fecha_inicio')
    f_fin = request.args.get('fecha_fin')
    
    # Consulta directa a route/list.json
    url = f"{BASE_URL}/route/list.json"
    params = {
        'key': API_KEY,
        'unit_id': unit_id,
        'from': f"{f_in} 00:00:00",
        'till': f"{f_fin} 23:59:59",
        'include[]': ['points', 'decoded_route', 'stops', 'summary']
    }
    
    try:
        res = requests.get(url, params=params, timeout=45)
        data = res.json()
        points = data.get('data', {}).get('units', [{}])[0].get('points', [])
        
        if not points:
            return jsonify({'error': 'No se encontraron registros de movimiento para esta fecha.'}), 404
        
        # Generar Excel
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Vehículo", "Fecha", "Latitud", "Longitud", "Velocidad (km/h)"])
        
        for p in points:
            ws.append([
                unit_name,
                p.get('time'),
                p.get('lat'),
                p.get('lng'),
                p.get('speed', 0)
            ])
            
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return send_file(buf, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                         as_attachment=True, download_name=f"Reporte_{unit_name}.xlsx")
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
