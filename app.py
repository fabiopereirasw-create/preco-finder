# app.py
# ---------------------------------------------
# "Preço Finder" – Buscador de produtos em Excel + PDF (paisagem) e Excel (.xlsx)
# Fluxo: upload → busca → ordenar por preço → baixar PDF/Excel
# ---------------------------------------------

import io
from typing import List

import pandas as pd
import streamlit as st
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

st.set_page_config(page_title="Preço Finder", layout="wide")
st.title("🔎 Preço Finder – Buscador de Produtos + PDF/Excel")

# -------------------------
# Utilidades
# -------------------------

def _clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    # Remove colunas Unnamed:* e padroniza cabeçalhos (strip)
    df = df.rename(columns={c: str(c).strip() for c in df.columns})
    keep_cols = [c for c in df.columns if not str(c).strip().upper().startswith("UNNAMED")]
    return df[keep_cols]

@st.cache_data(show_spinner=False)
def load_excel(file) -> dict:
    """Lê o Excel e retorna dict{nome_aba: DataFrame limpo}."""
    xls = pd.ExcelFile(file)
    sheets = {}
    for name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=name, dtype=str)
        sheets[name] = _clean_columns(df)
    return sheets

def find_rows(df: pd.DataFrame, term: str) -> pd.DataFrame:
    term_up = term.upper().strip()
    mask = df.applymap(lambda x: term_up in str(x).upper() if pd.notnull(x) else False)
    result = df[mask.any(axis=1)].copy()
    return result

def try_to_number(series: pd.Series) -> pd.Series:
    # Converte 'R$ 1.234,56' -> 1234.56
    s = series.astype(str).str.replace(r"[^\d,.\-]", "", regex=True)
    s = s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    return pd.to_numeric(s, errors="coerce")

# -------------------------
# PDF (A4 paisagem)
# -------------------------

def build_pdf(df: pd.DataFrame, title: str, product_col: str, col_order: List[str]) -> bytes:
    styles = getSampleStyleSheet()
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=landscape(A4),
        leftMargin=20, rightMargin=20, topMargin=25, bottomMargin=20
    )

    columns = [c for c in col_order if c in df.columns] or list(df.columns)
    data_df = df[columns].fillna("")

    header = [columns]
    rows = []
    for _, r in data_df.astype(str).iterrows():
        row = []
        for c, val in zip(columns, r):
            row.append(Paragraph(val, styles["Normal"]) if c == product_col else val)
        rows.append(row)

    table = Table(header + rows, colWidths=[360 if c == product_col else 90 for c in columns], repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#555555")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.black),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
    ]))

    elements = [Paragraph(f"Produtos encontrados: {title}", styles["Heading1"]), Spacer(1, 8), table]
    doc.build(elements)
    buffer.seek(0)
    return buffer.read()

# -------------------------
# Excel (.xlsx)
# -------------------------

def build_excel(df: pd.DataFrame, col_order: list[str]) -> bytes:
    """Gera um .xlsx em memória respeitando a ordem das colunas."""
    from pandas import ExcelWriter
    cols = [c for c in col_order if c in df.columns] or list(df.columns)
    data = df[cols].copy()

    buf = io.BytesIO()
    with ExcelWriter(buf, engine="openpyxl") as writer:
        data.to_excel(writer, index=False, sheet_name="resultado")
        ws = writer.sheets["resultado"]
        # largura básica para ficar legível
        for i, c in enumerate(cols, start=1):
            try:
                max_len = max([len(str(c))] + [len(str(v)) for v in data[c].astype(str).head(200)])
                ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = min(max(10, max_len + 2), 60)
            except Exception:
                pass
    buf.seek(0)
    return buf.read()

# -------------------------
# Interface
# -------------------------

uploaded = st.file_uploader("Envie sua planilha .xlsx (colunas padronizadas)", type=["xlsx"]) 

if uploaded:
    sheets = load_excel(uploaded)
    sheet_names = list(sheets.keys())

    # Selecionar aba (se houver mais de uma)
    sheet_sel = sheet_names[0] if len(sheet_names) == 1 else st.selectbox("Escolha a aba", sheet_names)
    df = sheets[sheet_sel]

    # Colunas conforme arquivo (limpas)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.str.contains(r"^Unnamed", case=False)]
    all_cols = list(df.columns)

    # Mapeamento leve (só se existir na planilha)
    map_try = {
        "CÓD.": next((c for c in all_cols if c.upper().startswith("CÓD")), None),
        "PRODUTOS": next((c for c in all_cols if c.upper() == "PRODUTOS"), None),
        "MARCA": next((c for c in all_cols if c.upper() == "MARCA"), None),
        "PESO KG": next((c for c in all_cols if c.upper().startswith("PESO")), None),
        "De 01 à 10 CXS": next((c for c in all_cols if ("01" in c and "10" in c)), None),
        "De 11 à 25 CXS2": next((c for c in all_cols if ("11" in c and "25" in c)), None),
        "Acima 25 CXS": next((c for c in all_cols if ("ACIMA" in c.upper() or "Acima" in c)), None),
    }

    st.markdown("---")
    term = st.text_input("Digite o nome do produto para buscar (ex.: BATATA)", value="")

    # Coluna de preço para ordenar (default inteligente)
    price_candidates = [map_try["De 01 à 10 CXS"], map_try["De 11 à 25 CXS2"], map_try["Acima 25 CXS"]]
    price_default = next((c for c in price_candidates if c in all_cols), (all_cols[0] if all_cols else None))
    price_col = st.selectbox("Coluna de preço para ordenar (decrescente)", options=all_cols,
                             index=(all_cols.index(price_default) if price_default in all_cols else 0))

    # Colunas para exportação (PDF/Excel) — por padrão todas, na ordem da planilha
    pdf_cols = st.multiselect("Colunas que vão para o PDF/Excel (ordem preservada)",
                              options=all_cols, default=all_cols)

    # Buscar
    if st.button("🔍 Buscar"):
        if not term.strip():
            st.warning("Digite um termo para pesquisar.")
        else:
            found = find_rows(df, term)
            if found.empty:
                st.info("Nenhuma linha encontrada para esse termo.")
            else:
                # Ordenar por preço, se possível
                if price_col in found.columns:
                    try:
                        found[price_col] = try_to_number(found[price_col])
                    except Exception:
                        pass
                    found = found.sort_values(by=price_col, ascending=False, na_position="last")

                st.success(f"{len(found)} linha(s) encontrada(s) na aba '{sheet_sel}'.")
                st.dataframe(found, use_container_width=True)

                # Colunas finais (ordem) a usar nos exports
                export_cols = [c for c in pdf_cols if c in found.columns] or list(found.columns)
                prod_col = map_try["PRODUTOS"] or (next((c for c in all_cols if c.upper()=="PRODUTOS"), all_cols[0]))

                # Gera bytes
                pdf_bytes = build_pdf(found, term.upper(), prod_col, export_cols)
                xlsx_bytes = build_excel(found, export_cols)

                # Dois botões lado a lado
                c1, c2 = st.columns(2)
                with c1:
                    st.download_button(
                        label="⬇️ Baixar PDF (paisagem)",
                        data=pdf_bytes,
                        file_name=f"resultado_{term.lower()}_{sheet_sel}.pdf",
                        mime="application/pdf",
                        key="btn_pdf",
                    )
                with c2:
                    st.download_button(
                        label="⬇️ Baixar Excel (.xlsx)",
                        data=xlsx_bytes,
                        file_name=f"resultado_{term.lower()}_{sheet_sel}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="btn_xlsx",
                    )
else:
    st.info("Envie sua planilha .xlsx para começar.")
