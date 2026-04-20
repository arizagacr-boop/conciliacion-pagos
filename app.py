import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime

st.set_page_config(page_title="Check Payins", page_icon="📊", layout="wide")

st.markdown("""
<style>
    .stApp { background-color: #0d0d0d; }
    .main { background-color: #0d0d0d; }
    .block-container { padding: 2rem 2rem 2rem; }
    section[data-testid="stSidebar"] { background-color: #111111; border-right: 1px solid #1a4a8a; }
    section[data-testid="stSidebar"] * { color: #c8d8f0 !important; }
    section[data-testid="stSidebar"] input { background-color: #1a1a1a !important; border: 1px solid #1a4a8a !important; color: #ffffff !important; border-radius: 6px !important; }
    h1 { color: #4a9eff !important; font-size: 1.8rem; }
    h2, h3 { color: #4a9eff !important; }
    p, span, label, div { color: #e0e0e0; }
    div[data-testid="metric-container"] { background: #111111; border: 1px solid #1a4a8a; border-radius: 10px; padding: 1rem; }
    div[data-testid="metric-container"] label { color: #7aaeff !important; }
    div[data-testid="metric-container"] div { color: #ffffff !important; }
    .stRadio label { color: #c8d8f0 !important; }
    input[type="number"] { background-color: #1a1a1a !important; color: #ffffff !important; border: 1px solid #1a4a8a !important; }
    .stDownloadButton button { background-color: #1a4a8a !important; color: #ffffff !important; border: none !important; border-radius: 8px !important; font-weight: 600 !important; }
    .stDownloadButton button:hover { background-color: #4a9eff !important; }
    .stAlert { border-radius: 8px; background-color: #111111 !important; border-left: 4px solid #4a9eff !important; }
    .stDataFrame { border: 1px solid #1a4a8a; border-radius: 8px; }
    .stFileUploader { border: 1px dashed #1a4a8a !important; border-radius: 8px; background-color: #111111 !important; }
    .streamlit-expanderHeader { color: #4a9eff !important; }
    hr { border-color: #1a4a8a !important; }
</style>
""", unsafe_allow_html=True)

st.title("📊 Check Payins")
st.markdown("Cargá el **extracto bancario** (PDF) y el **AR Processors** (Excel) para reconciliar por día.")

# ── Sidebar config ─────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Configuración")
    processor_name = st.text_input("Procesador a analizar", value="Servipag")
    col_date = st.text_input("Columna de fecha (planilla)", value="Payment date")
    col_amount = st.text_input("Columna de monto (planilla)", value="LC Amount")
    col_processor = st.text_input("Columna de procesador (planilla)", value="Processor")
    st.markdown("---")
    st.markdown("**Conceptos bancarios a incluir:**")
    use_payment = st.checkbox("PAYMENT (0780537906)", value=True)
    use_cash = st.checkbox("SERVIPAG CASH COLLECTION", value=True)
    use_refund = st.checkbox("REFUND", value=False)
    custom_concepts = st.text_area("Otros conceptos (uno por línea)", value="")
    st.markdown("---")
    tolerance = st.number_input("Tolerancia sin alerta (%)", value=10, step=1, min_value=0, max_value=100)

active_concepts = []
if use_payment: active_concepts.append("PAYMENT")
if use_cash: active_concepts.append("SERVIPAG CASH COLLECTION")
if use_refund: active_concepts.append("REFUND")
for c in custom_concepts.strip().split("\n"):
    if c.strip(): active_concepts.append(c.strip().upper())

# ── File uploads ───────────────────────────────────────────────────────────────
col1, col2 = st.columns(2)
with col1:
    st.subheader("🏦 Extracto Bancario")
    banco_file = st.file_uploader("PDF del banco", type=["pdf"], key="banco")
with col2:
    st.subheader("📋 AR Processors")
    interno_file = st.file_uploader("AR Processors", type=["xlsx", "xls", "csv"], key="interno")

# ── Parse PDF ─────────────────────────────────────────────────────────────────
def parse_banco_pdf(file, concepts):
    data = {}
    try:
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                lines = text.split("\n")
                for line in lines:
                    matched = any(c.upper() in line.upper() for c in concepts)
                    if not matched:
                        continue
                    date_match = re.search(r'(\d{2}/\d{2}/\d{4})', line)
                    amount_match = re.findall(r'[\d,]+\.\d{2}', line)
                    if not date_match or not amount_match:
                        continue
                    date_str = date_match.group(1)
                    try:
                        dt = datetime.strptime(date_str, "%m/%d/%Y").strftime("%d/%m")
                    except:
                        continue
                    amounts = [float(a.replace(",", "")) for a in amount_match]
                    # pick the smallest positive amount (not balance)
                    positives = [a for a in amounts if 0 < a < 1e9]
                    if not positives:
                        continue
                    amt = min(positives)
                    data[dt] = data.get(dt, 0) + amt

        # fallback: table extraction
        if not data:
            with pdfplumber.open(file) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        for row in table:
                            if not row: continue
                            row_str = " ".join(str(c) for c in row if c)
                            matched = any(c.upper() in row_str.upper() for c in concepts)
                            if not matched: continue
                            date_match = re.search(r'(\d{2}/\d{2}/\d{4})', row_str)
                            amount_match = re.findall(r'[\d,]+\.\d{2}', row_str)
                            if not date_match or not amount_match: continue
                            date_str = date_match.group(1)
                            try:
                                dt = datetime.strptime(date_str, "%m/%d/%Y").strftime("%d/%m")
                            except:
                                continue
                            amounts = [float(a.replace(",", "")) for a in amount_match]
                            positives = [a for a in amounts if 0 < a < 1e9]
                            if not positives: continue
                            amt = min(positives)
                            data[dt] = data.get(dt, 0) + amt
    except Exception as e:
        st.error(f"Error leyendo PDF: {e}")
    return data

# ── Parse Internal Excel ───────────────────────────────────────────────────────
def parse_interno(file, proc_name, col_d, col_a, col_p):
    try:
        if file.name.endswith(".csv"):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file)

        df.columns = [str(c).strip() for c in df.columns]

        if col_p not in df.columns:
            st.error(f"No encontré la columna '{col_p}'. Columnas disponibles: {list(df.columns)}")
            return {}
        if col_d not in df.columns:
            st.error(f"No encontré la columna '{col_d}'. Columnas disponibles: {list(df.columns)}")
            return {}
        if col_a not in df.columns:
            st.error(f"No encontré la columna '{col_a}'. Columnas disponibles: {list(df.columns)}")
            return {}

        filtered = df[df[col_p].astype(str).str.lower().str.contains(proc_name.lower(), na=False)]

        data = {}
        for _, row in filtered.iterrows():
            dt = row[col_d]
            if pd.isnull(dt): continue
            if isinstance(dt, str):
                for fmt in ["%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d %H:%M:%S"]:
                    try: dt = datetime.strptime(dt.split(" ")[0], fmt); break
                    except: pass
            if hasattr(dt, 'strftime'):
                key = dt.strftime("%d/%m")
            else:
                continue
            amt = float(row[col_a]) if pd.notna(row[col_a]) else 0
            data[key] = data.get(key, 0) + amt
        return data
    except Exception as e:
        st.error(f"Error leyendo planilla: {e}")
        return {}

# ── Build Excel output ─────────────────────────────────────────────────────────
def build_excel(rows, processor, period):
    wb = Workbook()
    ws = wb.active
    ws.title = "Conciliación"

    thin = Side(style='thin', color='CCCCCC')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def hdr(r, c, v, bg="1F4E79", fg="FFFFFF"):
        cell = ws.cell(row=r, column=c, value=v)
        cell.font = Font(name='Arial', bold=True, color=fg, size=10)
        cell.fill = PatternFill('solid', start_color=bg)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border

    ws.merge_cells('A1:F1')
    t = ws.cell(row=1, column=1, value=f"Conciliación {processor} — {period}")
    t.font = Font(name='Arial', bold=True, size=13, color="1F4E79")
    t.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 26

    headers = ['Fecha', 'Banco (CLP)', 'Interno (CLP)', 'Diferencia (CLP)', 'Dif. %', 'Estado']
    for i, h in enumerate(headers, 1):
        hdr(2, i, h)
    ws.row_dimensions[2].height = 22

    for idx, row in enumerate(rows):
        r = idx + 3
        bg_alt = "F2F7FB" if idx % 2 == 0 else "FFFFFF"
        diff = row['diff']
        pct = abs(diff / row['interno'] * 100) if row['interno'] else None
        is_ok = pct is not None and pct <= 10
        if is_ok:
            st_bg, st_txt, st_lbl = "D9F7D9", "1A6B1A", "OK"
        elif diff < 0:
            st_bg, st_txt, st_lbl = "FFDAD9", "8B1A1A", "Falta en banco"
        else:
            st_bg, st_txt, st_lbl = "FFF3CD", "7A5A00", "Falta en interno"

        for c in range(1, 7):
            cell = ws.cell(row=r, column=c)
            cell.border = border
            cell.font = Font(name='Arial', size=10)
            if c != 5: cell.fill = PatternFill('solid', start_color=bg_alt)

        ws.cell(row=r, column=1).value = row['fecha']
        ws.cell(row=r, column=1).alignment = Alignment(horizontal='center')

        for c, key in [(2, 'banco'), (3, 'interno')]:
            cell = ws.cell(row=r, column=c)
            cell.value = row[key] if row[key] is not None else '-'
            cell.number_format = '#,##0;(#,##0);"-"'
            cell.alignment = Alignment(horizontal='right')

        diff_cell = ws.cell(row=r, column=4)
        if row['banco'] is not None and row['interno'] is not None:
            diff_cell.value = f"=B{r}-C{r}"
        else:
            diff_cell.value = '-'
        diff_cell.number_format = '#,##0;(#,##0);"-"'
        diff_cell.alignment = Alignment(horizontal='right')
        diff_cell.fill = PatternFill('solid', start_color=st_bg)

        pct_cell = ws.cell(row=r, column=5)
        if row['banco'] and row['interno']:
            pct_cell.value = f"=D{r}/C{r}"
            pct_cell.number_format = '0.0%;(0.0%);"-"'
        else:
            pct_cell.value = '-'
        pct_cell.alignment = Alignment(horizontal='center')
        pct_cell.fill = PatternFill('solid', start_color=st_bg)

        st_cell = ws.cell(row=r, column=6)
        st_cell.value = st_lbl
        st_cell.font = Font(name='Arial', bold=True, size=10, color=st_txt)
        st_cell.fill = PatternFill('solid', start_color=st_bg)
        st_cell.alignment = Alignment(horizontal='center')

    total_r = len(rows) + 3
    for c in range(1, 7):
        cell = ws.cell(row=total_r, column=c)
        cell.font = Font(name='Arial', bold=True, size=10, color="FFFFFF")
        cell.fill = PatternFill('solid', start_color="1F4E79")
        cell.border = border
        cell.alignment = Alignment(horizontal='center' if c in [1,5,6] else 'right')

    ws.cell(row=total_r, column=1).value = "TOTAL"
    ws.cell(row=total_r, column=2).value = f"=SUM(B3:B{total_r-1})"
    ws.cell(row=total_r, column=2).number_format = '#,##0;(#,##0);"-"'
    ws.cell(row=total_r, column=3).value = f"=SUM(C3:C{total_r-1})"
    ws.cell(row=total_r, column=3).number_format = '#,##0;(#,##0);"-"'
    ws.cell(row=total_r, column=4).value = f"=B{total_r}-C{total_r}"
    ws.cell(row=total_r, column=4).number_format = '#,##0;(#,##0);"-"'
    ws.cell(row=total_r, column=5).value = f"=D{total_r}/C{total_r}"
    ws.cell(row=total_r, column=5).number_format = '0.0%;(0.0%);"-"'
    ws.cell(row=total_r, column=6).value = ""

    ws.column_dimensions['A'].width = 10
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 18
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 10
    ws.column_dimensions['F'].width = 16
    ws.freeze_panes = 'A3'

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── Main logic ─────────────────────────────────────────────────────────────────
if banco_file and interno_file:
    with st.spinner("Procesando archivos..."):
        banco_data = parse_banco_pdf(banco_file, active_concepts)
        interno_data = parse_interno(interno_file, processor_name, col_date, col_amount, col_processor)

    if not banco_data:
        st.warning("⚠️ No se encontraron movimientos en el PDF con los conceptos configurados. Revisá los conceptos en el sidebar.")
    if not interno_data:
        st.warning("⚠️ No se encontraron movimientos en la planilla para el procesador configurado.")

    if banco_data or interno_data:
        all_dates = sorted(set(list(banco_data.keys()) + list(interno_data.keys())))
        rows = []
        for d in all_dates:
            b = banco_data.get(d)
            i = interno_data.get(d)
            diff = (b or 0) - (i or 0)
            pct = abs(diff / i * 100) if i else None
            is_ok = pct is not None and pct <= tolerance
            rows.append({'fecha': d, 'banco': b, 'interno': i, 'diff': diff, 'pct': pct, 'is_ok': is_ok})

        total_banco = sum(r['banco'] for r in rows if r['banco'])
        total_interno = sum(r['interno'] for r in rows if r['interno'])
        diff_neta = total_banco - total_interno
        dias_ok = sum(1 for r in rows if r['is_ok'])
        dias_diff = len(rows) - dias_ok

        st.markdown("---")
        st.subheader("📈 Resumen")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total banco (CLP)", f"{total_banco:,.0f}")
        m2.metric("Total interno (CLP)", f"{total_interno:,.0f}")
        m3.metric("Diferencia neta (CLP)", f"{diff_neta:,.0f}", delta=f"{diff_neta/total_interno*100:.1f}%" if total_interno else None)
        m4.metric("Días con diferencia", f"{dias_diff} / {len(rows)}")

        st.markdown("---")
        st.subheader("📅 Detalle por día")

        filter_opt = st.radio("Mostrar:", ["Todos", "Solo con diferencias", "Solo OK"], horizontal=True)

        display_rows = []
        for r in rows:
            if filter_opt == "Solo con diferencias" and r['is_ok']: continue
            if filter_opt == "Solo OK" and not r['is_ok']: continue

            diff = r['diff']
            pct = r['pct']
            if r['is_ok']:
                estado = "✅ OK"
            elif diff < 0:
                estado = "🔴 Falta en banco"
            else:
                estado = "🟡 Falta en interno"

            display_rows.append({
                "Fecha": r['fecha'],
                "Banco (CLP)": f"{r['banco']:,.0f}" if r['banco'] else "-",
                "Interno (CLP)": f"{r['interno']:,.0f}" if r['interno'] else "-",
                "Diferencia (CLP)": f"{diff:+,.0f}",
                "Dif. %": f"{pct:+.1f}%" if pct is not None else "-",
                "Estado": estado
            })

        # Fila de totales
        total_pct = diff_neta / total_interno * 100 if total_interno else 0
        display_rows.append({
            "Fecha": "**TOTAL**",
            "Banco (CLP)": f"**{total_banco:,.0f}**",
            "Interno (CLP)": f"**{total_interno:,.0f}**",
            "Diferencia (CLP)": f"**{diff_neta:+,.0f}**",
            "Dif. %": f"**{total_pct:+.1f}%**",
            "Estado": "✅ OK" if abs(total_pct) <= tolerance else ("🔴 Falta en banco" if diff_neta < 0 else "🟡 Falta en interno")
        })

        st.dataframe(pd.DataFrame(display_rows), use_container_width=True, hide_index=True)

        st.markdown("---")
        period = f"{all_dates[0]} al {all_dates[-1]}" if all_dates else ""
        excel_buf = build_excel(rows, processor_name, period)
        st.download_button(
            label="⬇️ Descargar Excel con conciliación",
            data=excel_buf,
            file_name=f"conciliacion_{processor_name.lower()}_{datetime.now().strftime('%Y%m')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("👆 Cargá los dos archivos para comenzar el análisis.")
    with st.expander("ℹ️ ¿Cómo usar esta herramienta?"):
        st.markdown("""
        1. **Extracto bancario**: PDF exportado de tu banco (Citi, BICE, etc.)
        2. **AR Processors**: Excel con los pagos registrados internamente, con columnas de fecha, monto y procesador
        3. **Configurá** en el sidebar izquierdo los nombres de columnas y conceptos a filtrar
        4. La herramienta reconcilia automáticamente por día y te muestra las diferencias
        5. Descargá el Excel con el detalle completo y colores por estado
        """)
