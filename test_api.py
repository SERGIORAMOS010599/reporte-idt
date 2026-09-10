import requests

API_KEY = "7bd626cb4d3874faf995ec075af15d2cd35ec99d"
BASE_URL = "https://gps.idttecnologias.mx/api/v1"
UNIT_ID = "868807"

# Pedimos solo 25 minutos del tramo donde sabemos que el camión se movió
params = {
    "key": API_KEY,
    "unit_id": UNIT_ID,
    "from": "2026-08-31T16:35:00Z", # Hora UTC (09:35 AM en Sonora)
    "till": "2026-08-31T17:00:00Z"
}

print("Consultando historial minuto a minuto a la API...")
try:
    res = requests.get(f"{BASE_URL}/unit_data/history.json", params=params)
    data = res.json()
    
    unidades = data.get('data', {}).get('units', [])
    if unidades:
        historial = unidades[0].get('history', [])
        print(f"✅ Éxito: Puntos descargados: {len(historial)}")
        if historial:
            print("Muestra del primer punto:", historial[0])
    else:
        print("⚠️ La API respondió, pero sin unidades o historial.")
        print("Respuesta completa:", data)

except Exception as e:
    print(f"❌ Error de conexión: {e}")
