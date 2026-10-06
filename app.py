# -*- coding: utf-8 -*-
"""
Ranking de ventas de feria — Kbas Office
Sube el Excel con las líneas de venta (referencia + cantidad, y precio si hay)
y un zip con las fotos (nombre = referencia). Devuelve el ranking de lo más y
menos vendido, con foto, en PDF y/o Excel, separado por marca (Kbas / Volum).
"""

import io, os, re, zipfile, tempfile
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from PIL import Image as PILImage

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle, Paragraph,
                                Image as RLImage, PageBreak)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# ---------- Paleta ----------
INK="26242A"; BURD="6E1F2A"; SAND="EFE9E1"; MUTE="8A8177"

# ---------- Detección flexible de columnas ----------
REF_CANDS   = ["REFERENCIA","REF.","REF","ARTICULO","ARTÍCULO","CODIGO","CÓDIGO","COD"]
QTY_CANDS   = ["CANTIDAD VENDIDA","CANTIDAD","CANT.","CANT","UDS","UNIDADES","QTY","VENDIDO"]
PRICE_CANDS = ["PRECIO","PVP","P.V.P","IMPORTE UNIT"]
DTO_CANDS   = ["DTO","DTO.","DESCUENTO","% DTO","DESC"]

def _norm(s): return re.sub(r"\s+", " ", str(s)).strip().upper()

def _find_col(cols, cands):
    m = {_norm(c): c for c in cols}
    for cand in cands:
        if _norm(cand) in m:
            return m[_norm(cand)]
    return None

def limpiar_ref(v):
    # quita TODO espacio/invisible (en cualquier posición) y pasa a mayúsculas
    return re.sub(r"\s+", "", str(v)).upper()

def leer_excel(file_like):
    """Busca la pestaña y columnas de referencia + cantidad (prefiere 'LINEAS')."""
    xls = pd.ExcelFile(file_like)
    orden = sorted(xls.sheet_names, key=lambda n: 0 if "LINEA" in _norm(n) else 1)
    for name in orden:
        df = xls.parse(name)
        rc = _find_col(df.columns, REF_CANDS)
        qc = _find_col(df.columns, QTY_CANDS)
        if rc and qc:
            return df, rc, qc, _find_col(df.columns, PRICE_CANDS), _find_col(df.columns, DTO_CANDS), name
    raise ValueError("No encuentro columnas de Referencia y Cantidad en ninguna pestaña del Excel.")

def construir_ranking(df, rc, qc, pc, dc, ordenar_por="unidades"):
    d = df.copy()
    d = d[d[rc].notna()]                        # fuera filas sin referencia (incluye la de TOTAL)
    d["_REF"] = d[rc].map(limpiar_ref)
    d = d[d["_REF"] != ""]
    d["_Q"] = pd.to_numeric(d[qc], errors="coerce").fillna(0)
    if pc:
        precio = pd.to_numeric(d[pc], errors="coerce").fillna(0)
        dto = pd.to_numeric(d[dc], errors="coerce").fillna(0) if dc else 0
        d["_IMP"] = precio * (1 - dto/100) * d["_Q"]
    else:
        d["_IMP"] = 0.0
    d["_MARCA"] = d["_REF"].str[:1].map(lambda x: "Volum" if x == "V" else "Kbas")
    out = {}
    for marca, sub in d.groupby("_MARCA"):
        g = (sub.groupby("_REF").agg(uds=("_Q","sum"), euros=("_IMP","sum"))
                 .reset_index().rename(columns={"_REF": "ref"}))
        col = "euros" if ordenar_por == "euros" else "uds"
        g = g.sort_values(col, ascending=False)
        out[marca] = [{"ref": r.ref, "uds": int(r.uds), "euros": round(float(r.euros), 2)} for r in g.itertuples()]
    return out

def _es_imagen(nombre):
    return nombre.lower().endswith((".jpg", ".jpeg", ".png"))

def recoger_fotos(archivos, tmpdir):
    """Acepta una lista de archivos subidos: imágenes sueltas y/o zips (mezcla).
    Devuelve {REF: ruta_foto}. Ignora ocultos del Mac."""
    m = {}
    for f in archivos:
        nombre = f.name
        datos = f.getvalue()
        if nombre.lower().endswith(".zip"):
            try:
                with zipfile.ZipFile(io.BytesIO(datos)) as z:
                    for n in z.namelist():
                        base = os.path.basename(n)
                        if not base or base.startswith("._") or "__MACOSX" in n or not _es_imagen(base):
                            continue
                        dst = os.path.join(tmpdir, base)
                        with z.open(n) as src, open(dst, "wb") as out:
                            out.write(src.read())
                        m[os.path.splitext(base)[0].upper()] = dst
            except zipfile.BadZipFile:
                continue
        elif _es_imagen(nombre):
            base = os.path.basename(nombre)
            if base.startswith("._"):
                continue
            dst = os.path.join(tmpdir, base)
            with open(dst, "wb") as out:
                out.write(datos)
            m[os.path.splitext(base)[0].upper()] = dst
    return m

def hacer_thumb(src, tmpdir, ref):
    try:
        im = PILImage.open(src).convert("RGB")
        im.thumbnail((130, 180))
        dst = os.path.join(tmpdir, f"thumb_{ref}.jpg")
        im.save(dst, quality=85)
        return dst
    except Exception:
        return None

def fmt_eur(v): return f"{v:,.2f} €".replace(",", "§").replace(".", ",").replace("§", ".")
def fmt_num(v): return f"{int(v):,}".replace(",", ".")

# ---------- Generar EXCEL ----------
def generar_excel(ranking, fotos, tmpdir, con_euros):
    wb = Workbook(); wb.remove(wb.active)
    thin = Side(style="thin", color="D9D2C7")
    for marca in ["Kbas", "Volum"]:
        if marca not in ranking: continue
        ws = wb.create_sheet(marca)
        ws.merge_cells("A1:D1")
        ws["A1"] = f"Ranking de ventas · {marca}"
        ws["A1"].font = Font(name="Arial", bold=True, size=13, color=INK)
        ws.row_dimensions[1].height = 24
        heads = ["Foto", "Referencia", "Cantidad vendida", "€ vendidos"]
        for c, h in enumerate(heads, 1):
            cell = ws.cell(row=2, column=c, value=h)
            cell.font = Font(name="Arial", bold=True, size=10, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor=BURD)
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions["A"].width = 11
        ws.column_dimensions["B"].width = 18
        ws.column_dimensions["C"].width = 16
        ws.column_dimensions["D"].width = 14
        ws.freeze_panes = "A3"
        r = 3
        for it in ranking[marca]:
            ws.row_dimensions[r].height = 56
            src = fotos.get(it["ref"])
            if src:
                th = hacer_thumb(src, tmpdir, it["ref"])
                if th:
                    img = XLImage(th); img.width = 52; img.height = 72
                    ws.add_image(img, f"A{r}")
            ws.cell(row=r, column=2, value=it["ref"]).font = Font(name="Arial", size=10)
            cc = ws.cell(row=r, column=3, value=it["uds"]); cc.font = Font(name="Arial", size=10); cc.alignment = Alignment(horizontal="center", vertical="center")
            cd = ws.cell(row=r, column=4, value=it["euros"] if con_euros else None)
            cd.font = Font(name="Arial", size=10); cd.number_format = '#,##0.00 €'; cd.alignment = Alignment(horizontal="right", vertical="center")
            for col in range(1, 5): ws.cell(row=r, column=col).border = Border(bottom=thin)
            r += 1
        ws.cell(row=r, column=2, value="TOTAL").font = Font(name="Arial", bold=True, size=10, color=INK)
        tc = ws.cell(row=r, column=3, value=f"=SUM(C3:C{r-1})"); tc.font = Font(name="Arial", bold=True, size=10); tc.alignment = Alignment(horizontal="center")
        if con_euros:
            td = ws.cell(row=r, column=4, value=f"=SUM(D3:D{r-1})"); td.font = Font(name="Arial", bold=True, size=10); td.number_format = '#,##0.00 €'; td.alignment = Alignment(horizontal="right")
        for col in range(1, 5): ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor=SAND)
    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf

# ---------- Generar PDF ----------
def generar_pdf(ranking, fotos, tmpdir, con_euros, titulo):
    styles = getSampleStyleSheet()
    h_st   = ParagraphStyle("h", parent=styles["Title"], fontName="Helvetica-Bold", fontSize=16, textColor=colors.HexColor("#"+INK), spaceAfter=2)
    sub_st = ParagraphStyle("s", parent=styles["Normal"], fontName="Helvetica", fontSize=9, textColor=colors.HexColor("#"+MUTE), spaceAfter=8)
    cell_st= ParagraphStyle("c", parent=styles["Normal"], fontName="Helvetica", fontSize=9, textColor=colors.HexColor("#"+INK))
    refb_st= ParagraphStyle("r", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=9, textColor=colors.HexColor("#"+INK))
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, topMargin=14*mm, bottomMargin=14*mm,
                            leftMargin=16*mm, rightMargin=16*mm, title=titulo)
    story = []
    marcas = [m for m in ["Kbas", "Volum"] if m in ranking]
    for mi, marca in enumerate(marcas):
        items = ranking[marca]
        tot_u = sum(x["uds"] for x in items); tot_e = sum(x["euros"] for x in items)
        story.append(Paragraph(titulo, h_st))
        resumen = f"{marca} · {len(items)} referencias · {fmt_num(tot_u)} uds"
        if con_euros: resumen += f" · {fmt_eur(tot_e)}"
        story.append(Paragraph(resumen, sub_st))
        data = [[Paragraph("<b>Foto</b>", cell_st), Paragraph("<b>Referencia</b>", cell_st),
                 Paragraph("<b>Cantidad</b>", cell_st), Paragraph("<b>€ vendidos</b>", cell_st)]]
        for it in items:
            src = fotos.get(it["ref"])
            th = hacer_thumb(src, tmpdir, it["ref"]) if src else None
            if th:
                im = RLImage(th); im.drawHeight = 14*mm; im.drawWidth = 14*mm*0.806
            else:
                im = Paragraph("<font color='#B0A99E'>— sin foto —</font>", cell_st)
            euro = fmt_eur(it["euros"]) if con_euros else "—"
            data.append([im, Paragraph(it["ref"], refb_st), Paragraph(fmt_num(it["uds"]), cell_st), Paragraph(euro, cell_st)])
        data.append([Paragraph("", cell_st), Paragraph("<b>TOTAL</b>", refb_st),
                     Paragraph(f"<b>{fmt_num(tot_u)}</b>", cell_st),
                     Paragraph(f"<b>{fmt_eur(tot_e) if con_euros else '—'}</b>", cell_st)])
        t = Table(data, colWidths=[20*mm, 60*mm, 35*mm, 43*mm], repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#"+BURD)),
            ("TEXTCOLOR",  (0,0), (-1,0), colors.white),
            ("ALIGN", (2,0), (3,-1), "RIGHT"), ("ALIGN", (0,0), (0,-1), "CENTER"),
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("ROWBACKGROUNDS", (0,1), (-1,-2), [colors.white, colors.HexColor("#F6F2EC")]),
            ("BACKGROUND", (0,-1), (-1,-1), colors.HexColor("#"+SAND)),
            ("LINEBELOW", (0,0), (-1,0), 0.5, colors.HexColor("#"+BURD)),
            ("TOPPADDING", (0,1), (-1,-1), 3), ("BOTTOMPADDING", (0,1), (-1,-1), 3),
            ("LEFTPADDING", (0,0), (-1,-1), 5),
        ]))
        story.append(t)
        if mi < len(marcas) - 1: story.append(PageBreak())
    doc.build(story); buf.seek(0); return buf

# ==================== INTERFAZ ====================
def check_password():
    def entered():
        if st.session_state.get("password") == st.secrets.get("dashboard_password"):
            st.session_state["ok"] = True; del st.session_state["password"]
        else:
            st.session_state["ok"] = False
    if st.session_state.get("ok"): return True
    st.title("Ranking de ventas")
    st.text_input("Contraseña", type="password", on_change=entered, key="password")
    if st.session_state.get("ok") is False: st.error("Contraseña incorrecta.")
    return False

st.set_page_config(page_title="Ranking de ventas · Kbas Office", page_icon="🏆")
if not ("dashboard_password" in st.secrets and not check_password()):

    st.title("🏆 Ranking de ventas")
    st.caption("Sube el Excel con las líneas (referencia + cantidad, y precio si lo hay) y el zip de fotos. "
               "Devuelve el ranking de lo más y menos vendido, con foto, separado por marca.")

    c1, c2 = st.columns(2)
    with c1:
        excel_file = st.file_uploader("1 · Excel con las líneas de venta", type=["xlsx", "xlsm"])
    with c2:
        fotos_files = st.file_uploader("2 · Fotos (sueltas o en ZIP · nombre = referencia)",
                                       type=["zip", "jpg", "jpeg", "png"], accept_multiple_files=True)

    o1, o2 = st.columns(2)
    with o1:
        ordenar = st.radio("Ordenar por", ["Unidades", "Euros"], horizontal=True)
    with o2:
        formatos = st.multiselect("Formato de salida", ["PDF", "Excel"], default=["PDF", "Excel"])
    titulo = st.text_input("Título del documento", value="Ranking de ventas")

    if excel_file and fotos_files and formatos:
        if st.button("Generar ranking", type="primary"):
            try:
                df, rc, qc, pc, dc, hoja = leer_excel(excel_file)
            except Exception as e:
                st.error(str(e)); st.stop()
            ranking = construir_ranking(df, rc, qc, pc, dc, "euros" if ordenar == "Euros" else "unidades")
            con_euros = pc is not None

            with tempfile.TemporaryDirectory() as tmp:
                fotos = recoger_fotos(fotos_files, tmp)
                # resumen de cobertura
                tot_ref = sum(len(v) for v in ranking.values())
                con_foto = sum(1 for v in ranking.values() for it in v if it["ref"] in fotos)
                st.success(f"Procesado: {tot_ref} referencias · {con_foto} con foto "
                           f"({round(100*con_foto/tot_ref) if tot_ref else 0}%). "
                           f"Columnas usadas: Ref='{rc}', Cantidad='{qc}'"
                           + (f", Precio='{pc}'" if pc else " (sin precio → sin euros)")
                           + (f", Dto='{dc}'" if dc else "") + f". Hoja: '{hoja}'.")
                for marca, items in ranking.items():
                    st.write(f"**{marca}**: {len(items)} referencias, {sum(i['uds'] for i in items)} uds")

                if "PDF" in formatos:
                    pdf = generar_pdf(ranking, fotos, tmp, con_euros, titulo)
                    st.download_button("⬇ Descargar PDF", pdf, file_name="ranking.pdf", mime="application/pdf")
                if "Excel" in formatos:
                    xls = generar_excel(ranking, fotos, tmp, con_euros)
                    st.download_button("⬇ Descargar Excel", xls, file_name="ranking.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.info("Sube el Excel y las fotos (sueltas o en ZIP), elige el formato y pulsa Generar.")
