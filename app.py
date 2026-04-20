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
    .stApp { background-color: #05051a; }
    .main { background-color: #05051a; }
    .block-container { padding: 1rem 2rem 2rem; }
    section[data-testid="stSidebar"] { background-color: #07071f; border-right: 2px solid #1a1aff; }
    section[data-testid="stSidebar"] * { color: #a0b4ff !important; }
    section[data-testid="stSidebar"] input { background-color: #0a0a2e !important; border: 1px solid #1a1aff !important; color: #ffffff !important; border-radius: 6px !important; }
    section[data-testid="stSidebar"] h1, section[data-testid="stSidebar"] h2, section[data-testid="stSidebar"] h3 { color: #ffffff !important; }
    h2, h3 { color: #a0b4ff !important; }
    p, span, label, div { color: #c8d4ff; }
    div[data-testid="metric-container"] { background: #0a0a2e; border: 1px solid #1a1aff; border-radius: 10px; padding: 1rem; }
    div[data-testid="metric-container"] label { color: #7a94ff !important; }
    div[data-testid="metric-container"] div { color: #ffffff !important; }
    .stRadio label { color: #a0b4ff !important; }
    input[type="number"] { background-color: #0a0a2e !important; color: #ffffff !important; border: 1px solid #1a1aff !important; }
    .stDownloadButton button { background-color: #1a1aff !important; color: #ffffff !important; border: none !important; border-radius: 8px !important; font-weight: 700 !important; padding: 0.6rem 1.5rem !important; }
    .stDownloadButton button:hover { background-color: #3333ff !important; }
    .stAlert { border-radius: 8px; background-color: #0a0a2e !important; border-left: 4px solid #1a1aff !important; color: #c8d4ff !important; }
    .stDataFrame { border: 1px solid #1a1aff; border-radius: 8px; }
    .stFileUploader { border: 1px dashed #1a1aff !important; border-radius: 8px; background-color: #0a0a2e !important; }
    .streamlit-expanderHeader { color: #7a94ff !important; background-color: #0a0a2e !important; }
    .streamlit-expanderContent { background-color: #0a0a2e !important; }
    hr { border-color: #1a1aff !important; }
    .stRadio div { background: transparent !important; }
</style>
""", unsafe_allow_html=True)

# Header con logo dLocal
logo_b64 = "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"
st.markdown(f"""
<div style="background: #0a0a2e; padding: 2rem 2.5rem; border-radius: 14px; margin-bottom: 1.8rem; display: flex; align-items: center; justify-content: center; gap: 2.5rem; border-bottom: 3px solid #1a1aff; min-height: 110px;">
    <img src="data:image/webp;base64,{logo_b64}" style="height: 80px; filter: brightness(0) invert(1); display: block;">
    <div style="border-left: 2px solid #1a1aff; padding-left: 2rem; display: flex; align-items: center;">
        <div style="color: #ffffff; font-size: 2.2rem; font-weight: 800; letter-spacing: -1px; line-height: 1;">Check Payins</div>
    </div>
</div>
""", unsafe_allow_html=True)

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

    headers = ['Fecha', 'Banco (CLP)', 'AR Processors (CLP)', 'Diferencia (CLP)', 'Dif. %', 'Estado']
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
            st_bg, st_txt, st_lbl = "FFDAD9", "8B1A1A", "Cobro menor"
        else:
            st_bg, st_txt, st_lbl = "FFF3CD", "7A5A00", "Cobro mayor"

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
        m2.metric("Total AR Processors (CLP)", f"{total_interno:,.0f}")
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
                estado = "🔴 Cobro menor"
            else:
                estado = "🟡 Cobro mayor"

            display_rows.append({
                "Fecha": r['fecha'],
                "Banco (CLP)": f"{r['banco']:,.0f}" if r['banco'] else "-",
                "AR Processors (CLP)": f"{r['interno']:,.0f}" if r['interno'] else "-",
                "Diferencia (CLP)": f"{diff:+,.0f}",
                "Dif. %": f"{pct:+.1f}%" if pct is not None else "-",
                "Estado": estado
            })

        # Fila de totales
        total_pct = diff_neta / total_interno * 100 if total_interno else 0
        display_rows.append({
            "Fecha": "**TOTAL**",
            "Banco (CLP)": f"**{total_banco:,.0f}**",
            "AR Processors (CLP)": f"**{total_interno:,.0f}**",
            "Diferencia (CLP)": f"**{diff_neta:+,.0f}**",
            "Dif. %": f"**{total_pct:+.1f}%**",
            "Estado": "✅ OK" if abs(total_pct) <= tolerance else ("🔴 Cobro menor" if diff_neta < 0 else "🟡 Cobro mayor")
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
