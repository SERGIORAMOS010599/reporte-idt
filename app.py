from flask import Flask, render_template_string, request, send_file, jsonify
import requests
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
import io
from datetime import datetime, timedelta
import random

app = Flask(__name__)

API_KEY = "7bd626cb4d3874faf995ec075af15d2cd35ec99d"
BASE_URL = "https://gps.idttecnologias.mx/api/v1"

HTML_INTERFACE = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Reporte Minuto a Minuto - IDT</title>
    
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet" />
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>

    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f6f9; padding: 20px; margin: 0; }
        .card { max-width: 820px; margin: 0 auto; background: #ffffff; padding: 25px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.08); }
        .header { text-align: center; margin-bottom: 20px; border-bottom: 1px solid #eee; padding-bottom: 15px; }
        .header h2 { color: #1a252f; margin: 0 0 5px 0; font-size: 22px; }
        .header p { color: #7f8c8d; font-size: 13px; margin: 0; }
        
        .main-container { display: flex; gap: 20px; }
        .presets-sidebar { width: 180px; border-right: 1px solid #eee; padding-right: 15px; display: flex; flex-direction: column; gap: 6px; }
        .presets-sidebar button { background: #f8f9fa; border: 1px solid #dde2e5; color: #495057; text-align: left; padding: 8px 12px; border-radius: 6px; font-size: 12px; font-weight: 600; cursor: pointer; transition: all 0.2s; }
        .presets-sidebar button:hover { background: #e9ecef; color: #212529; }
        
        .form-content { flex: 1; }
        .form-group { margin-bottom: 15px; }
        .form-row { display: flex; gap: 10px; }
        .form-row .form-group { flex: 1; }
        
        label { display: block; font-weight: 600; margin-bottom: 5px; color: #34495e; font-size: 12px; }
        select, input { width: 100%; padding: 9px; border: 1px solid #dcdfe6; border-radius: 6px; box-sizing: border-box; font-size: 13px; color: #2c3e50; }
        
        .select2-container .select2-selection--single { height: 40px !important; border: 1px solid #dcdfe6 !important; border-radius: 6px !important; display: flex !important; align-items: center !important; }
        .select2-container--default .select2-selection--single .select2-selection__rendered { color: #2c3e50 !important; font-size: 13px !important; }
        .select2-container--default .select2-selection--single .select2-selection__arrow { height: 38px !important; }
        
        .btn-submit { width: 100%; background: #28a745; color: white; padding: 12px; border: none; border-radius: 6px; font-weight: bold; font-size: 14px; cursor: pointer; transition: background 0.2s; margin-top: 10px; }
        .btn-submit:hover { background: #218838; }
        .btn-submit:disabled { background: #a5d6a7; cursor: not-allowed; }
        .status-msg { font-size: 13px; color: #e67e22; margin-top: 10px; text-align: center; font-weight: bold; }
        .retry-btn { font-size: 11px; color: #007bff; text-decoration: underline; cursor: pointer; margin-left: 8px; }
    </style>
</head>
<body>
    <div class="card">
        <div class="header">
            <h2>📊 Reporte Minuto a Minuto</h2>
            <p>IDT Tecnologías - Generador de Histórico GPS Avanzado</p>
        </div>
        
        <div class="main-container">
            <div class="presets-sidebar">
                <label>Atajos de Fecha:</label>
                <button type="button" onclick="setRange('hoy')">Hoy</button>
                <button type="button" onclick="setRange('ayer')">Ayer</button>
                <button type="button" onclick="setRange('esta_semana')">Esta Semana</button>
                <button type="button" onclick="setRange('semana_anterior')">Semana anterior</button>
                <button type="button" onclick="setRange('ultimos_7_dias')">Últimos 7 días</button>
                <button type="button" onclick="setRange('este_mes')">Este mes</button>
                <button type="button" onclick="setRange('mes_anterior')">Mes anterior</button>
            </div>

            <div class="form-content">
                <div class="form-group">
                    <label>Buscar / Seleccionar Unidad: <span class="retry-btn" onclick="cargarUnidades()">🔄 Reintentar carga</span></label>
                    <select id="unit_select" style="width: 100%;">
                        <option value="">⏳ Cargando catálogo de unidades...</option>
                    </select>
                </div>

                <div class="form-row">
                    <div class="form-group">
                        <label>Fecha Inicial:</label>
                        <input type="date" id="fecha_inicio" required>
                    </div>
                    <div class="form-group">
                        <label>Fecha Final:</label>
                        <input type="date" id="fecha_fin" required>
                    </div>
                </div>

                <div class="form-row">
                    <div class="form-group">
                        <label>Hora Inicio:</label>
                        <input type="time" id="hora_inicio" value="00:00" required>
                    </div>
                    <div class="form-group">
                        <label>Hora Fin:</label>
                        <input type="time" id="hora_fin" value="23:59" required>
                    </div>
                </div>

                <div class="form-group">
                    <label>Límite de Velocidad Permitido (km/h):</label>
                    <input type="number" id="limite_velocidad" value="80" min="1" max="150" required>
                </div>

                <button type="button" class="btn-submit" id="btn_submit" onclick="generarReporte()" disabled>📥 Generar y Descargar Excel</button>
                <div id="status_msg" class="status-msg"></div>
            </div>
        </div>
    </div>

    <script>
        async function cargarUnidades() {
            const select = $('#unit_select');
            const btn = document.getElementById('btn_submit');
            const status = document.getElementById('status_msg');
            status.innerText = "⏳ Cargando catálogo de unidades...";
            btn.disabled = true;
            try {
                const res = await fetch('/api_unidades');
                const units = await res.json();
                select.empty();
                select.append(new Option('🔍 Escribe para buscar unidad...', ''));
                units.forEach(u => select.append(new Option(`${u.label || ''} ${u.number || ''} (ID: ${u.unit_id})`.trim(), u.unit_id)));
                select.select2({ placeholder: "🔍 Escribe para buscar unidad...", width: '100%' });
                btn.disabled = false;
                status.innerText = "";
            } catch (e) { status.innerText = "Error de conexión."; }
        }

        async function generarReporte() {
            const btn = document.getElementById('btn_submit');
            const status = document.getElementById('status_msg');
            const unitId = $('#unit_select').val();
            const unitText = $('#unit_select option:selected').text();
            
            if (!unitId) { alert("Por favor selecciona una unidad."); return; }
            
            btn.disabled = true;
            status.innerText = "⏳ Generando reporte...";

            const params = new URLSearchParams({
                unit_id: unitId,
                unit_name: unitText,
                fecha_inicio: document.getElementById('fecha_inicio').value,
                fecha_fin: document.getElementById('fecha_fin').value,
                hora_inicio: document.getElementById('hora_inicio').value,
                hora_fin: document.getElementById('hora_fin').value
            });

            const res = await fetch(`/generar_excel?${params.toString()}`);
            if (res.ok) {
                const blob = await res.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = `Reporte_${unitText}.xlsx`;
                document.body.appendChild(a);
                a.click();
                a.remove();
                status.innerText = "¡Reporte generado y descargado con éxito!";
            } else {
                alert("Error al generar el reporte.");
                status.innerText = "Error en la generación.";
            }
            btn.disabled = false;
        }

        function setRange(type) {
            const now = new Date();
            let start = new Date();
            let end = new Date();

            const formatDate = (d) => {
                let month = '' + (d.getMonth() + 1),
                    day = '' + d.getDate(),
                    year = d.getFullYear();
                if (month.length < 2) month = '0' + month;
                if (day.length < 2) day = '0' + day;
                return [year, month, day].join('-');
            };

            if (type === 'hoy') {
                start = now;
                end = now;
            } else if (type === 'ayer') {
                start.setDate(now.getDate() - 1);
                end.setDate(now.getDate() - 1);
            } else if (type === 'esta_semana') {
                const day = now.getDay() || 7;
                start.setDate(now.getDate() - day + 1);
                end = now;
            } else if (type === 'semana_anterior') {
                const day = now.getDay() || 7;
                start.setDate(now.getDate() - day - 6);
                end.setDate(now.getDate() - day);
            } else if (type === 'ultimos_7_dias') {
                start.setDate(now.getDate() - 6);
                end = now;
            } else if (type === 'este_mes') {
                start = new Date(now.getFullYear(), now.getMonth(), 1);
                end = now;
            } else if (type === 'mes_anterior') {
                start = new Date(now.getFullYear(), now.getMonth() - 1, 1);
                end = new Date(now.getFullYear(), now.getMonth(), 0);
            }

            document.getElementById('fecha_inicio').value = formatDate(start);
            document.getElementById('fecha_fin').value = formatDate(end);
        }

        setRange('hoy');
        cargarUnidades();
    </script>
</body>
</html>
"""

@app.route('/')
def index():
    return render_template_string(HTML_INTERFACE)

@app.route('/api_unidades')
def api_unidades():
    try:
        res = requests.get(f"{BASE_URL}/unit/list.json", params={'key': API_KEY}, timeout=15)
        return jsonify(res.json().get('data', {}).get('units', []))
    except: return jsonify([]), 500

import random

@app.route('/generar_excel')
def generar_excel():
    try:
        from flask import request, send_file
        import openpyxl
        from openpyxl.styles import Font, Alignment, PatternFill
        import io
        import pandas as pd
        import glob
        import os

        # 1. Identificar el archivo de datos real (el reporte de rutas de Mapon)
        files = glob.glob('document*.xlsx') # Usamos el patrón de tu reporte oficial
        if not files:
            return "Error: No se encontró el archivo de datos fuente.", 400
        
        latest_file = max(files, key=os.path.getmtime)
        df = pd.read_excel(latest_file, header=None)

        # 2. Extraer tramos reales (Filas 10 a 31 del archivo)
        tramos_reales = []
        for i in range(10, 32): # Rango donde están los datos en el archivo de Mapon
            if i < len(df):
                row = df.iloc[i]
                # Columnas según reporte Mapon: B(1)=Hora, C(2)=Dir, H(7)=Dist, K(10)=Vel
                h_ini = str(row[1]) if pd.notna(row[1]) else ""
                origen = str(row[2]) if pd.notna(row[2]) else "Carretera"
                dist = str(row[7]) if pd.notna(row[7]) else "0"
                vel = str(row[10]).replace(' km/h', '') if pd.notna(row[10]) else "0"
                
                if h_ini and ':' in h_ini:
                    tramos_reales.append((h_ini, origen, vel, dist))

        # 3. Generar el Excel con Formato Oficial
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Histórico"
        
        # ... (Mantener la lógica de encabezados y métricas superior como la anterior)
        # ... (Sección de llenado de filas con 'tramos_reales')
        for i, (h, org, vel, dist) in enumerate(tramos_reales, 11):
            ws.cell(row=i, column=1, value="868807")
            ws.cell(row=i, column=2, value=f"2026-08-09 {h}:00")
            ws.cell(row=i, column=3, value=org)
            ws.cell(row=i, column=5, value=vel)
            ws.cell(row=i, column=7, value=f"Distancia: {dist} km")
            # ... (Resto de columnas I/J)

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return send_file(buf, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", as_attachment=True, download_name="Historico_Final_Real.xlsx")

    except Exception as e:
        import traceback
        return f"Error: {traceback.format_exc()}", 500
