import streamlit as st
import pandas as pd
from pyxlsb import open_workbook
from pathlib import Path
import base64

# ------------------------
# 🎨 Personnalisation CSS
# ------------------------
def inject_custom_css():
    st.markdown("""
        <style>
            .main {
                background-image: url('background.jpg');
                background-size: cover;
                background-position: center;
                background-repeat: no-repeat;
                padding: 2rem;
            }
            .title {
                text-align: center;
                font-size: 2.5rem;
                font-weight: bold;
                color: #003366;
                margin-bottom: 0.5rem;
            }
            .subtitle {
                text-align: center;
                font-size: 1.5rem;
                color: #555555;
                margin-bottom: 2rem;
            }
            .report-box {
                background-color: rgba(255, 255, 255, 0.8);
                padding: 1.5rem;
                border-radius: 12px;
                box-shadow: 0 0 10px rgba(0,0,0,0.2);
            }
        </style>
    """, unsafe_allow_html=True)

# ------------------------
# 🧠 Vérification du fichier
# ------------------------
def check_referentiel(file):
    errors = []

    try:
        with open_workbook(file) as wb:
            if "Référentiel" not in wb.sheets:
                errors.append("❌ L'onglet 'Référentiel' est manquant.")
                return errors

            with wb.get_sheet("Référentiel") as sheet:
                data = []
                for row in sheet.rows():
                    data.append([item.v for item in row])

                df = pd.DataFrame(data)
                df.columns = df.iloc[0]
                df = df[1:]

                # Vérifie que la ligne 1 (en-têtes) n'a pas de cellule vide
                if any(pd.isna(data[0])):
                    colonnes_vides = [i for i, val in enumerate(data[0]) if pd.isna(val)]
                    errors.append(f"❌ Ligne 1 (entêtes) contient des colonnes vides aux positions : {colonnes_vides}")

                # Vérifie colonne CODEPSS
                if 'CODEPSS' not in df.columns:
                    errors.append("❌ Colonne 'CODEPSS' manquante.")
                else:
                    nb_vides = df['CODEPSS'].isna().sum()
                    if nb_vides > 0:
                        errors.append(f"❌ {nb_vides} cellule(s) vide(s) dans la colonne 'CODEPSS'.")

                # Vérifie colonne CODECLIENT
                if 'CODECLIENT' not in df.columns:
                    errors.append("❌ Colonne 'CODECLIENT' manquante ou mal orthographiée.")
                else:
                    vides_codeclient = df[df['CODECLIENT'].isna()]
                    if not vides_codeclient.empty:
                        lignes_vides = vides_codeclient.index.tolist()
                        lignes_affichage = [i + 2 for i in lignes_vides]  # +2 car index commence à 0 et ligne 1 = entêtes
                        errors.append(f"❌ Cellules vides dans la colonne 'CODECLIENT' aux lignes : {lignes_affichage}")

    except Exception as e:
        errors.append(f"Erreur lors de l'analyse : {e}")

    return errors

# ------------------------
# 🖼️ Logo Carrefour Centré
# ------------------------
def add_logo_centered_faded(image_path, width=150, opacity=0.6):
    with open(image_path, "rb") as img_file:
        encoded = base64.b64encode(img_file.read()).decode()
        st.markdown(
            f"""
            <div style="display: flex; justify-content: center; align-items: center; margin-bottom: 20px;">
                <img src="data:image/png;base64,{encoded}" style="width: {width}px; opacity: {opacity};" />
            </div>
            """,
            unsafe_allow_html=True
        )

# ------------------------
# 🚀 App principale
# ------------------------
inject_custom_css()

logo_path = Path("carrefour_logo.png")
if logo_path.exists():
    add_logo_centered_faded(logo_path)
else:
    st.error("Logo Carrefour introuvable.")

st.markdown('<div class="title">Vérification LinePlan</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Service Textile - Carrefour</div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("📥 Uploadez un fichier LinePlan (.xlsb)", type="xlsb")

if uploaded_file:
    st.markdown('<div class="report-box">', unsafe_allow_html=True)
    st.write(f"📂 **Fichier sélectionné :** {uploaded_file.name}")
    errors = check_referentiel(uploaded_file)

    if errors:
        st.error("🛑 Problèmes détectés :")
        for err in errors:
            st.write("•", err)
    else:
        st.success("✅ Aucune erreur détectée dans l’onglet 'Référentiel'.")
    st.markdown('</div>', unsafe_allow_html=True)
