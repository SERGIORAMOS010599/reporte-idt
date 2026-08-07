from flask import Flask, request, jsonify, render_template_string, send_file
import requests
import openpyxl
import io
import math
from datetime import datetime, timedelta

app = Flask(__name__)

API_KEY = "7bd626cb4d3874faf995ec075af15d2cd35ec99d"
BASE_URL = "https://gps.idttecnologias.mx/api/v1"

HTML_INTERFACE = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Reporte Minuto a Minuto - IDT</title>
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet" />
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <style>
        body { font-family: 'Segoe UI', sans-serif; background-color: #f4f6f9; padding: 20px; }
        .card { max-width: 820px; margin: 0 auto; background: #ffffff; padding: 25px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.08); }
        .btn-submit { width: 100%; background: #28a745; color: white; padding: 12px; border: none; border-radius: 6px; font-weight: bold; cursor: pointer; }
    </style>
</head>
<body>
    <div class="card">
        <h2>📊 Reporte Minuto a Minuto</h2>
        <select id="unit_select" style="width: 100%;"></select>
        <div style="display:flex; gap:10px; margin-top:15px;">
            <input type="date" id="fecha_inicio">
            <input type="date" id="fecha_fin">
        </div>
        <button class="btn-submit" id="btn_submit" onclick="generarReporte()">📥 Generar Excel</button>
        <div id="status_msg" style="text-align:center; margin-top:10px;"></div>
    </div>
    <script>
        $(document).ready(async () => {
            const res = await fetch('/api_unidades');
            const units = await res.json();
            const s = $('#unit_select');
            units.forEach(u => s.append(new Option(`${u.label} (ID: ${u.unit_id})`, u.unit_id)));
            s.select2();
        });
        async function generarReporte() {
            const btn = $('#btn_submit');
            btn.prop('disabled', true).text('⏳ Procesando...');
            const params = new URLSearchParams({
                unit_id: $('#unit_select').val(),
                fecha_inicio: $('#fecha_inicio').val(),
                fecha_fin: $('#fecha_fin').val()
            });
            const res = await fetch('/generar_excel?' + params.toString());
            if(res.ok) {
                const blob = await res.blob();
                const a = document.createElement('a');
                a.href = window.URL.createObjectURL(blob);
                a.download = "Reporte.xlsx";
                a.click();
            } else { alert("Error al generar: Verifica que la unidad tenga datos."); }
            btn.prop('disabled', false).text('📥 Generar Excel');
        }
    </script>
</body>
</html>
"""

# [MANTENEMOS TUS FUNCIONES haversine, parse_point_timestamp, format_sec_to_hm, extraer_puntos_y_resumen]
# (Asegúrate de no borrar las funciones que tenías debajo del HTML_INTERFACE)

@app.route('/')
def index(): return render_template_string(HTML_INTERFACE)

@app.route('/api_unidades')
def api_unidades():
    res = requests.get(f"{BASE_URL}/unit/list.json", params={'key': API_KEY})
    return jsonify(res.json().get('data', {}).get('units', []))

@app.route('/generar_excel')
def generar_excel():
    unit_id = request.args.get('unit_id')
    f_in = request.args.get('fecha_inicio')
    f_fin = request.args.get('fecha_fin')
    
    # La API de IDT espera formato YYYY-MM-DD. 
    # El input type="date" de HTML ya entrega ese formato.
    url = f"{BASE_URL}/route/list.json?key={API_KEY}&unit_id={unit_id}&from={f_in}%2000:00:00&till={f_fin}%2023:59:59&include[]=points&include[]=summary"
    
    try:
        res = requests.get(url, timeout=45)
        if not res.ok: return jsonify({'error': 'API error'}), 404
        data = res.json().get('data', {})
        units = data.get('units', [])
        if not units: return jsonify({'error': 'No hay datos'}), 404
        
        points = units[0].get('points', [])
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Vehículo", "Fecha", "Velocidad", "Evento", "Lat", "Lng"])
        for p in points:
            ws.append([unit_id, p.get('time'), p.get('speed', 0), "Movimiento" if p.get('speed', 0) > 5 else "Parado", p.get('lat'), p.get('lng')])
        
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return send_file(buf, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", as_attachment=True, download_name="Reporte.xlsx")
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
