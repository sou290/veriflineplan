import streamlit as st
import pandas as pd
from pyxlsb import open_workbook
from pathlib import Path
import base64

# ------------------------
# 🖼️ Personnalisation CSS
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
# 📊 Vérification du LinePlan
# ------------------------
def check_referentiel(file):
    errors = []
    colonnes_attendues = ["CODEPSS", "CODECLIENT"]  # Ajoute ici les colonnes obligatoires si besoin

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

                # 🔍 Vérif 1 : colonnes vides dans la ligne d'entête
                header_row = data[0]
                for idx, val in enumerate(header_row):
                    if pd.isna(val) or str(val).strip() == "":
                        errors.append(f"❌ En-tête vide détectée dans la colonne n°{idx+1}.")

                # 🔍 Vérif 2 : noms de colonnes attendus
                for col in colonnes_attendues:
                    if col not in df.columns:
                        errors.append(f"❌ Colonne obligatoire '{col}' manquante.")

                # 🔍 Vérif 3 : cellules vides dans 'CODEPSS'
                if 'CODEPSS' in df.columns:
                    nb_vides = df['CODEPSS'].isna().sum()
                    if nb_vides > 0:
                        lignes_vides = df[df['CODEPSS'].isna()].index + 2  # +2 car DataFrame commence à 0 + 1 ligne d'en-tête
                        errors.append(f"❌ {nb_vides} cellule(s) vide(s) dans 'CODEPSS' (lignes : {list(lignes_vides)})")

                # 🔍 Vérif 4 : cellules vides dans 'CODECLIENT'
                if 'CODECLIENT' in df.columns:
                    nb_vide_client = df['CODECLIENT'].isna().sum()
                    if nb_vide_client > 0:
                        lignes_vides_client = df[df['CODECLIENT'].isna()].index + 2
                        errors.append(f"❌ {nb_vide_client} cellule(s) vide(s) dans 'CODECLIENT' (lignes : {list(lignes_vides_client)})")

    except Exception as e:
        errors.append(f"Erreur lors de l'analyse : {e}")

    return errors

# ------------------------
# 🚀 Logo Carrefour affiché au centre et estompé
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
# 🚀 Page principale
# ------------------------

inject_custom_css()

# Affiche le logo Carrefour
logo_path = Path("carrefour_logo.png")
if logo_path.exists():
    add_logo_centered_faded(logo_path)
else:
    st.error("Logo Carrefour introuvable.")

# Titre principal
st.markdown('<div class="title">Vérification LinePlan</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Service Textile - Carrefour</div>', unsafe_allow_html=True)

# Upload du fichier
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
