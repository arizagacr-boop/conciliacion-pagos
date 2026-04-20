# Conciliación de Pagos — Servipag

Herramienta para reconciliar pagos de procesadores (Servipag, Transbank, etc.) entre el extracto bancario y la planilla interna.

## Deploy en Streamlit Cloud (sin código)

1. Subí estos archivos a un repo de GitHub
2. Entrá a [share.streamlit.io](https://share.streamlit.io)
3. Conectá tu cuenta de GitHub
4. Seleccioná el repo y el archivo `app.py`
5. Hacé clic en **Deploy** — en 2 minutos tenés tu URL

## Archivos necesarios
- `app.py` — la aplicación
- `requirements.txt` — dependencias (se instalan automáticamente)

## Uso
- **Extracto bancario**: PDF del banco
- **Planilla interna**: Excel con columnas de fecha, monto y procesador
- Configurá los nombres de columnas en el sidebar izquierdo
