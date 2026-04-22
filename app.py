import streamlit as st
from pptx import Presentation
import io
from datetime import datetime
from PIL import Image

# --- CONFIGURATION DU MOT DE PASSE ---
MOT_DE_PASSE_ENTREPRISE = "ETECHNOLOGIE"


# --- NORMALISATION IMAGE (IMPORTANT POUR MOBILE) ---
def normaliser_image(file):
    img = Image.open(file)
    img = img.convert("RGB")
    output = io.BytesIO()
    img.save(output, format="JPEG", quality=95)
    output.seek(0)
    return output


# --- 1. MOTEUR DE REMPLISSAGE ---
def generer_rapport(donnees, photos, template_path):
    try:
        prs = Presentation(template_path)
    except Exception as e:
        st.error(f"Erreur : Impossible de charger 'template.pptx'. ({e})")
        return None

    for slide in prs.slides:
        for shape in slide.shapes:

            # --- TEXTE (VIDE SI AUCUNE VALEUR) ---
            if shape.has_text_frame:
                for balise, valeur in donnees.items():

                    if valeur is None or str(valeur).strip() == "":
                        replacement = ""
                    else:
                        replacement = str(valeur).strip()

                    if balise in shape.text:
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                if balise in run.text:
                                    run.text = run.text.replace(balise, replacement)

            # --- IMAGES (POSITION FIXE TEMPLATE) ---
            if shape.name in photos and photos[shape.name] is not None:
                photo_file = normaliser_image(photos[shape.name])

                left, top, width, height = shape.left, shape.top, shape.width, shape.height

                img = slide.shapes.add_picture(photo_file, left, top)

                # FORCER POSITION EXACTE DU TEMPLATE
                img.left = left
                img.top = top
                img.width = width
                img.height = height

                # supprimer placeholder
                sp = shape._element
                sp.getparent().remove(sp)

    pptx_io = io.BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)
    return pptx_io


# --- 2. AUTHENTIFICATION ---
def est_authentifie():
    if "auth" not in st.session_state:
        st.session_state["auth"] = False

    if not st.session_state["auth"]:
        st.markdown("### 🔒 Accès Sécurisé - Audit Environnemental")
        mdp = st.text_input("Veuillez entrer le mot de passe entreprise :", type="password")

        if st.button("Se connecter"):
            if mdp == MOT_DE_PASSE_ENTREPRISE:
                st.session_state["auth"] = True
                st.rerun()
            else:
                st.error("Mot de passe incorrect.")
        return False

    return True


# --- 3. INTERFACE ---
if est_authentifie():

    st.set_page_config(page_title="Audit Site Telecom", layout="wide")
    st.title("📱 Rapport d'Audit Environnemental Automatisé")

    donnees = {}
    photos = {}
    date_du_jour = datetime.now().strftime("%d-%m-%Y")

    tabs = st.tabs([
        "📍 INFOS SITE", "📸 VUE GENERALE", "🛡️ SECURITE",
        "🧹 PROPRETE", "🏗️ PYLÔNE & EQUIP.", "⚡ ELECTRICITE",
        "⚙️ GE", "❄️ CLIMATISATION", "⚠️ ANOMALIES"
    ])

    # --- INFOS SITE ---
    with tabs[0]:
        donnees["INS_NOM"] = st.text_input("Nom du site")
        donnees["INS_CODE"] = st.text_input("Code du site")
        donnees["INS_ZONE"] = st.text_input("Zone")
        donnees["INS_TRAVAUX"] = st.text_area("Travaux exercés (INS_TRAVAUX)")
        donnees["INS_CHEF"] = st.text_input("Chef d'équipes")
        donnees["INS_INSCONTACT"] = st.text_input("Contact")

    # --- VUE GENERALE ---
    with tabs[1]:
        photos["INS_VUE_DU_SITE"] = st.file_uploader("VUE DU SITE", type=['jpg','png','jpeg'])
        photos["INS_PLAQUE"] = st.file_uploader("PLAQUE", type=['jpg','png','jpeg'])

        c1, c2, c3 = st.columns(3)
        photos["INS_SITE1"] = c1.file_uploader("Site 1", type=['jpg'])
        photos["INS_SITE2"] = c2.file_uploader("Site 2", type=['jpg'])
        photos["INS_SITE3"] = c3.file_uploader("Site 3", type=['jpg'])

    # --- SECURITE ---
    with tabs[2]:
        photos["INS_ACCES"] = st.file_uploader("ACCES", type=['jpg'])
        photos["INS_PORTAIL"] = st.file_uploader("PORTAIL", type=['jpg'])
        photos["INS_SERRURE"] = st.file_uploader("SERRURE", type=['jpg'])

    # --- PROPRETE ---
    with tabs[3]:
        photos["INS_PROPRE"] = st.file_uploader("PROPRETE", type=['jpg'])

    # --- PYLONE ---
    with tabs[4]:
        photos["INS_FONDA1"] = st.file_uploader("Fondation", type=['jpg'])

    # --- ELECTRICITE ---
    with tabs[5]:
        photos["INS_CIE1"] = st.file_uploader("CIE", type=['jpg'])
        photos["INS_TGBT"] = st.file_uploader("TGBT", type=['jpg'])

    # --- GE ---
    with tabs[6]:
        photos["INS_MOTEUR"] = st.file_uploader("GE Face", type=['jpg'])

        donnees["INS_MARQUE_GE"] = st.text_input("Marque GE")
        donnees["INS_VAL_H"] = st.text_input("Heures Compteur")
        donnees["INS_VAL_CARB"] = st.text_input("Carburant (%)")

    # --- CLIM ---
    with tabs[7]:
        photos["INS_CLIM1"] = st.file_uploader("Clim", type=['jpg'])

    # --- ANOMALIES + GENERATION ---
    with tabs[8]:
        photos["INS_AN1"] = st.file_uploader("Anomalie 1", type=['jpg'])
        donnees["INS_REMARQUES"] = st.text_area("Remarques Finales")

        if st.button("🚀 GÉNÉRER LE RAPPORT", use_container_width=True):

            if not donnees["INS_CODE"] or not donnees["INS_NOM"]:
                st.error("Nom et Code du site obligatoires")
            else:
                with st.spinner("Génération en cours..."):
                    output = generer_rapport(donnees, photos, "template.pptx")

                    if output:
                        nom_fichier = f"RAPPORT AUDIT {donnees['INS_CODE']} {donnees['INS_NOM']} {date_du_jour}.pptx"

                        st.success("Rapport généré !")

                        st.download_button(
                            "📥 Télécharger",
                            data=output,
                            file_name=nom_fichier,
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
