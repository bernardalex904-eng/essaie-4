import io
... from datetime import datetime
... 
... import numpy as np
... import pandas as pd
... import streamlit as st
... 
... # Dépendances optionnelles
... try:
...     from docx import Document
...     from docx.shared import Inches
...     DOCX_OK = True
... except Exception:
...     DOCX_OK = False
... 
... try:
...     import fitz  # PyMuPDF
...     PYMUPDF_OK = True
... except Exception:
...     PYMUPDF_OK = False
... 
... st.set_page_config(page_title="Contrôle dimensionnel – PV", layout="wide")
... st.title("🧪 Contrôle dimensionnel – Génération de PV")
... 
... # --- Sidebar: Infos PV ---
... st.sidebar.header("Paramètres du PV")
... operateur = st.sidebar.text_input("Opérateur", value="")
... ref_piece = st.sidebar.text_input("Référence pièce", value="")
... ref_commande = st.sidebar.text_input("Référence commande", value="")
... commentaire = st.sidebar.text_area("Commentaires", value="")
... 
... # --- Upload du plan (image/PDF) ---
... st.subheader("1) Plan de référence (image ou PDF)")
... plan_file = st.file_uploader(
...     "Importer un plan (PNG/JPG/PDF)", type=["png", "jpg", "jpeg", "pdf"], key="plan"
)
plan_preview = None   # tuple (type, bytes) où type ∈ {"image", "pdf_image", "pdf_no_preview"}
plan_filename = None

if plan_file is not None:
    plan_filename = plan_file.name
    name_lower = plan_file.name.lower()
    mime = plan_file.type or ""

    if mime in ("image/png", "image/jpeg") or name_lower.endswith((".png", ".jpg", ".jpeg")):
        plan_bytes = plan_file.read()
        st.image(plan_bytes, caption=f"Aperçu du plan – {plan_filename}", use_container_width=True)
        plan_preview = ("image", plan_bytes)

    elif mime == "application/pdf" or name_lower.endswith(".pdf"):
        pdf_bytes = plan_file.read()
        if PYMUPDF_OK:
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            if doc.page_count > 0:
                page = doc.load_page(0)
                pix = page.get_pixmap(dpi=150)
                img_bytes = pix.tobytes("png")  # image PNG de la page 1
                st.image(img_bytes, caption=f"Aperçu du plan (page 1) – {plan_filename}", use_container_width=True)
                plan_preview = ("pdf_image", img_bytes)
            else:
                st.warning("Le PDF semble vide.")
        else:
            st.info("Prévisualisation PDF indisponible (PyMuPDF non installé). Le PDF sera référencé dans le PV.")
            plan_preview = ("pdf_no_preview", b"")

# --- Données de contrôle ---
st.subheader("2) Données de contrôle – Import ou saisie")
mode = st.radio("Choisir la méthode :", ["Importer feuille Excel", "Saisie manuelle"], horizontal=True)

REQUIRED_COLS = ["Caractéristique", "Nominal", "Tolérance -", "Tolérance +", "Mesuré"]

@st.cache_data(show_spinner=False)
def template_excel_bytes():
    """Modèle Excel en mémoire avec colonnes requises."""
    df_template = pd.DataFrame([
        {"Caractéristique": "Ø10 H7", "Nominal": 10.0, "Tolérance -": -0.015, "Tolérance +": 0.0, "Mesuré": 9.988},
        {"Caractéristique": "Longueur A", "Nominal": 100.0, "Tolérance -": -0.2, "Tolérance +": 0.2, "Mesuré": 100.12},
    ])
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df_template.to_excel(writer, index=False, sheet_name="Controle")
    bio.seek(0)
    return bio.read()

st.download_button(
    label="📥 Télécharger modèle Excel (colonnes requises)",
    data=template_excel_bytes(),
    file_name="modele_controle.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

if mode == "Importer feuille Excel":
    xlsx = st.file_uploader("Importer le fichier Excel de contrôle", type=["xlsx"], key="excel")
    if xlsx is not None:
        try:
            df = pd.read_excel(xlsx, engine="openpyxl")
        except Exception as e:
            st.error(f"Erreur de lecture Excel : {e}")
            df = None
    else:
        df = None

    if df is not None:
        st.write("Aperçu des données importées :")
        st.dataframe(df, use_container_width=True)
        missing = [c for c in REQUIRED_COLS if c not in df.columns]
        if missing:
            st.warning(f"Colonnes manquantes : {missing}. Renommez vos colonnes ou basculez en saisie manuelle.")
            data_df = pd.DataFrame(columns=REQUIRED_COLS)
        else:
            data_df = df.copy()
    else:
        data_df = pd.DataFrame(columns=REQUIRED_COLS)

else:
    # Saisie manuelle via data_editor
    default_rows = [
        {"Caractéristique": "Ø10 H7", "Nominal": 10.0, "Tolérance -": -0.015, "Tolérance +": 0.0, "Mesuré": 9.988},
    ]
    data_df = st.data_editor(
        pd.DataFrame(default_rows, columns=REQUIRED_COLS),
        num_rows="dynamic",
        use_container_width=True,
    )

# --- Calcul de conformité ---
st.subheader("3) Calcul de conformité")

def compute_conformity(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df.copy()

    out = df.copy()

    # Si l'Excel fournit une seule colonne "Tolérance" (symétrique), on la scinde
    if "Tolérance -" not in out.columns and "Tolérance" in out.columns:
        out["Tolérance -"] = -pd.to_numeric(out["Tolérance"], errors="coerce")
    if "Tolérance +" not in out.columns and "Tolérance" in out.columns:
        out["Tolérance +"] = pd.to_numeric(out["Tolérance"], errors="coerce")

    for col in ["Nominal", "Tolérance -", "Tolérance +", "Mesuré"]:
        if col in out.columns:
            out[col] = pd.to_numeric(out[col], errors="coerce")

    out["Borne min"] = out["Nominal"] + out["Tolérance -"]
    out["Borne max"] = out["Nominal"] + out["Tolérance +"]

    ok_mask = (out["Mesuré"] >= out["Borne min"]) & (out["Mesuré"] <= out["Borne max"])
    out["Conforme"] = ok_mask.map(lambda x: "✅ Conforme" if x else "❌ Non conforme")
    return out

if not data_df.empty and set(["Nominal", "Tolérance -", "Tolérance +", "Mesuré"]).issubset(data_df.columns):
    res_df = compute_conformity(data_df)
    n_total = len(res_df)
    n_ok = int((res_df["Conforme"] == "✅ Conforme").sum())
    n_ko = n_total - n_ok

    st.success(f"Résultat: {n_ok}/{n_total} conformes • {n_ko} non conformes")
    st.dataframe(res_df, use_container_width=True)
else:
    res_df = pd.DataFrame()
    st.info("Renseignez les colonnes requises puis lancez le calcul.")

# --- Export PV ---
st.subheader("4) Génération du PV (Word / PDF)")
colw, colp = st.columns(2)

def build_docx(df: pd.DataFrame) -> bytes:
    if not DOCX_OK:
        raise RuntimeError("python-docx non installé.")
    doc = Document()
    doc.add_heading("Procès-Verbal de Contrôle Dimensionnel", level=1)

    # Métadonnées
    p = doc.add_paragraph()
    p.add_run("Date : ").bold = True
    p.add_run(datetime.now().strftime("%Y-%m-%d %H:%M"))
    p.add_run("\nOpérateur : ").bold = True
    p.add_run(operateur or "—")
    p.add_run("\nRéf. pièce : ").bold = True
    p.add_run(ref_piece or "—")
    p.add_run("\nRéf. commande : ").bold = True
    p.add_run(ref_commande or "—")

    if plan_filename:
        p = doc.add_paragraph()
        p.add_run("\nPlan : ").bold = True
        p.add_run(plan_filename)

    if commentaire:
        doc.add_paragraph(f"Commentaire : {commentaire}")

    # Aperçu plan (si image disponible)
    try:
        if plan_preview and plan_preview[0] in ("image", "pdf_image"):
            img_bytes = plan_preview[1]
            tmp = io.BytesIO(img_bytes)
            doc.add_picture(tmp, width=Inches(5.5))
    except Exception:
        pass

    # Tableau résultats
    table_cols = ["Caractéristique", "Nominal", "Tolérance -", "Tolérance +", "Mesuré",
                  "Borne min", "Borne max", "Conforme"]
    df_show = df[table_cols].copy()

    table = doc.add_table(rows=1, cols=len(table_cols))
    hdr_cells = table.rows[0].cells
    for j, col in enumerate(table_cols):
        hdr_cells[j].text = str(col)

    for _, row in df_show.iterrows():
        cells = table.add_row().cells
        for j, col in enumerate(table_cols):
            cells[j].text = str(row[col])

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.read()

def build_pdf(df: pd.DataFrame) -> bytes:
    if not PYMUPDF_OK:
        raise RuntimeError("PyMuPDF (pymupdf) non installé.")
    doc = fitz.open()
    page = doc.new_page(width=595, height=842)  # A4 portrait (points)

    y = 40
    x_margin = 40

    def write_line(txt, size=11, bold=False):
        nonlocal y
        font = "helv-B" if bold else "helv"
        page.insert_text((x_margin, y), txt, fontsize=size, fontname=font)
        y += size + 6

    # En-tête
    write_line("Procès-Verbal de Contrôle Dimensionnel", size=16, bold=True)
    write_line(f"Date : {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    write_line(f"Opérateur : {operateur or '—'}")
    write_line(f"Réf. pièce : {ref_piece or '—'}")
    write_line(f"Réf. commande : {ref_commande or '—'}")
    if plan_filename:
        write_line(f"Plan : {plan_filename}")
    if commentaire:
        write_line(f"Commentaire : {commentaire}")

    y += 10
    write_line("Résultats :", bold=True)

    # Tableau simple (texte aligné)
    headers = ["Caractéristique", "Nominal", "Tol-", "Tol+", "Mesuré", "Bmin", "Bmax", "Conf."]
    col_x = [x_margin, 240, 305, 340, 380, 430, 470, 520]

    for hx, hx_x in zip(headers, col_x):
        page.insert_text((hx_x, y), hx, fontsize=10, fontname="helv-B")
    y += 14

    for _, r in df.iterrows():
        vals = [
            str(r.get("Caractéristique", ""))[:22],
            f"{r.get('Nominal', '')}",
            f"{r.get('Tolérance -', '')}",
            f"{r.get('Tolérance +', '')}",
            f"{r.get('Mesuré', '')}",
            f"{r.get('Borne min', '')}",
            f"{r.get('Borne max', '')}",
            "OK" if str(r.get("Conforme", "")).startswith("✅") else "NOK"
        ]
        for v, vx in zip(vals, col_x):
            page.insert_text((vx, y), str(v), fontsize=10, fontname="helv")
        y += 14
        if y > 780:
            page = doc.new_page(width=595, height=842)
            y = 40

    # Aperçu plan (si image dispo) sur une nouvelle page
    if plan_preview and plan_preview[0] in ("image", "pdf_image"):
        page = doc.new_page(width=595, height=842)
        img_bytes = plan_preview[1]
        rect = fitz.Rect(40, 80, 555, 800)
        page.insert_image(rect, stream=img_bytes, keep_proportion=True)

    pdf_bytes = doc.tobytes()
    doc.close()
    return pdf_bytes

# Boutons d'export
if not res_df.empty:
    with colw:
        if DOCX_OK:
            docx_bytes = build_docx(res_df)
            st.download_button(
                label="💾 Télécharger le PV (Word .docx)",
                data=docx_bytes,
                file_name=f"PV_controle_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
        else:
            st.warning("Export Word indisponible (python-docx non installé). Ajoutez 'python-docx' dans requirements.txt")

    with colp:
        if PYMUPDF_OK:
            try:
                pdf_bytes = build_pdf(res_df)
                st.download_button(
                    label="📄 Télécharger le PV (PDF)",
                    data=pdf_bytes,
                    file_name=f"PV_controle_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                    mime="application/pdf",
                )
            except Exception as e:
                st.error(f"Erreur génération PDF: {e}")
        else:
            st.info("Pour exporter en PDF, ajoutez 'pymupdf' dans requirements.txt")
else:
    st.info("Aucun résultat à exporter pour l'instant.")

st.caption("© 2025 – Application de démonstration pour contrôle dimensionnel et génération de PV.")
