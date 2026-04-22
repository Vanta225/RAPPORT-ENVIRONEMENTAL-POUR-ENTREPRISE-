import streamlit as st
from pptx import Presentation
import io
from datetime import datetime
from PIL import Image

# =========================================================
# 🔐 MOT DE PASSE
# =========================================================
MOT_DE_PASSE_ENTREPRISE = "ETECHNOLOGIE"


# =========================================================
# 📱 FIX MOBILE IMAGE (ANDROID / IOS STABLE)
# =========================================================
def normaliser_image(file):
    try:
        img = Image.open(file)

        # correction rotation mobile
        try:
            exif = img._getexif()
            if exif is not None:
                orientation = 274
                if orientation in exif:
                    if exif[orientation] == 3:
                        img = img.rotate(180, expand=True)
                    elif exif[orientation] == 6:
                        img = img.rotate(270, expand=True)
                    elif exif[orientation] == 8:
                        img = img.rotate(90, expand=True)
        except:
            pass

        img = img.convert("RGB")

        buffer = io.BytesIO()
        img.save(buffer, format="JPEG", quality=85, optimize=True)
        buffer.seek(0)

        return buffer

    except:
        return None


# =========================================================
# 📊 GENERATION RAPPORT PPT
# =========================================================
def generer_rapport(donnees, photos, template_path):

    prs = Presentation(template_path)

    for slide in prs.slides:
        for shape in slide.shapes:

            # =================================================
            # 🧾 TEXTE PROPRE
            # =================================================
            if shape.has_text_frame:
                for balise, valeur in donnees.items():

                    replacement = str(valeur).strip() if valeur else ""

                    if balise in shape.text:
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                if balise in run.text:
                                    run.text = run.text.replace(balise, replacement)

            # =================================================
            # 🖼️ IMAGES SUR CASE (SANS SUPPRESSION TEMPLATE)
            # =================================================
            if shape.name in photos and photos[shape.name] is not None:

                photo_file = normaliser_image(photos[shape.name])

                if photo_file:

                    try:
                        # ✔️ remplit la case existante (meilleur rendu PPT)
                        shape.fill.user_picture(photo_file)

                    except Exception:

                        # fallback sécurisé
                        left = shape.left
                        top = shape.top
                        width = shape.width
                        height = shape.height

                        slide.shapes.add_picture(photo_file, left, top, width, height)

    pptx_io = io.BytesIO()
    prs.save(pptx_io)

    return io.BytesIO(pptx_io.getvalue())


# =========================================================
# 🔐 AUTH
# =========================================================
def est_authentifie():

    if "auth" not in st.session_state:
        st.session_state["auth"] = False

    if not st.session_state["auth"]:
        st.markdown("### 🔒 Accès Sécurisé - Audit Environnemental")

        mdp = st.text_input("Mot de passe entreprise :", type="password")

        if st.button("Se connecter"):
            if mdp == MOT_DE_PASSE_ENTREPRISE:
                st.session_state["auth"] = True
                st.rerun()
            else:
                st.error("Mot de passe incorrect")

        return False

    return True


# =========================================================
# 🖥️ INTERFACE COMPLETE (9 ONGLETS RESTAURÉS)
# =========================================================
if est_authentifie():

    st.set_page_config(page_title="Audit Site Telecom", layout="wide")
    st.title("📱 Rapport d'Audit Environnemental Automatisé")

    donnees = {}
    photos = {}
    date_du_jour = datetime.now().strftime("%d-%m-%Y")

    # =====================================================
    # 🧭 9 ONGLETS COMPLETS
    # =====================================================
    tabs = st.tabs([
        "📍 INFOS SITE",
        "📸 VUE GENERALE",
        "🛡️ SECURITE",
        "🧹 PROPRETE",
        "🏗️ PYLÔNE & EQUIP.",
        "⚡ ELECTRICITE",
        "⚙️ GE",
        "❄️ CLIMATISATION",
        "⚠️ ANOMALIES"
    ])

    # =====================================================
    # 1. INFOS SITE
    # =====================================================
    with tabs[0]:
        donnees["INS_NOM"] = st.text_input("Nom du site")
        donnees["INS_CODE"] = st.text_input("Code du site")
        donnees["INS_ZONE"] = st.text_input("Zone")
        donnees["INS_TRAVAUX"] = st.text_area("Travaux exercés")
        donnees["INS_CHEF"] = st.text_input("Chef d'équipes")
        donnees["INS_INSCONTACT"] = st.text_input("Contact")

    # =====================================================
    # 2. VUE GENERALE
    # =====================================================
    with tabs[1]:
        photos["INS_VUE_DU_SITE"] = st.file_uploader("Vue du site", type=['jpg','jpeg','png'])
        photos["INS_PLAQUE"] = st.file_uploader("Plaque", type=['jpg','jpeg','png'])

        c1, c2, c3 = st.columns(3)
        photos["INS_SITE1"] = c1.file_uploader("Site 1", type=['jpg','jpeg','png'])
        photos["INS_SITE2"] = c2.file_uploader("Site 2", type=['jpg','jpeg','png'])
        photos["INS_SITE3"] = c3.file_uploader("Site 3", type=['jpg','jpeg','png'])

    # =====================================================
    # 3. SECURITE
    # =====================================================
    with tabs[2]:
        photos["INS_ACCES"] = st.file_uploader("Accès site", type=['jpg','jpeg','png'])
        photos["INS_PORTAIL"] = st.file_uploader("Portail", type=['jpg','jpeg','png'])
        photos["INS_SERRURE"] = st.file_uploader("Serrure", type=['jpg','jpeg','png'])

    # =====================================================
    # 4. PROPRETE
    # =====================================================
    with tabs[3]:
        photos["INS_PROPRE"] = st.file_uploader("Propreté générale", type=['jpg','jpeg','png'])
        photos["INS_DRAIN1"] = st.file_uploader("Drainage", type=['jpg','jpeg','png'])

    # =====================================================
    # 5. PYLÔNE
    # =====================================================
    with tabs[4]:
        photos["INS_FONDA1"] = st.file_uploader("Fondation", type=['jpg','jpeg','png'])
        photos["EQUIP1"] = st.file_uploader("Équipement", type=['jpg','jpeg','png'])

    # =====================================================
    # 6. ELECTRICITE
    # =====================================================
    with tabs[5]:
        photos["INS_CIE1"] = st.file_uploader("CIE", type=['jpg','jpeg','png'])
        photos["INS_TGBT"] = st.file_uploader("TGBT", type=['jpg','jpeg','png'])

    # =====================================================
    # 7. GE
    # =====================================================
    with tabs[6]:
        photos["INS_MOTEUR"] = st.file_uploader("GE", type=['jpg','jpeg','png'])

        donnees["INS_MARQUE_GE"] = st.text_input("Marque GE")
        donnees["INS_VAL_H"] = st.text_input("Heures")

    # =====================================================
    # 8. CLIM
    # =====================================================
    with tabs[7]:
        photos["INS_CLIM1"] = st.file_uploader("Climatisation", type=['jpg','jpeg','png'])

    # =====================================================
    # 9. ANOMALIES + GENERATION
    # =====================================================
    with tabs[8]:
        photos["INS_AN1"] = st.file_uploader("Anomalie 1", type=['jpg','jpeg','png'])
        donnees["INS_REMARQUES"] = st.text_area("Remarques finales")

        st.markdown("---")

        if st.button("🚀 GÉNÉRER LE RAPPORT", use_container_width=True):

            if not donnees["INS_CODE"] or not donnees["INS_NOM"]:
                st.error("Nom et Code obligatoires")
            else:

                with st.spinner("Génération en cours..."):

                    output = generer_rapport(donnees, photos, "template.pptx")

                    if output:

                        nom_fichier = f"AUDIT {donnees['INS_CODE']} {donnees['INS_NOM']} {date_du_jour}.pptx"

                        st.success("Rapport généré avec succès")

                        st.download_button(
                            "📥 Télécharger le rapport",
                            data=output.getvalue(),
                            file_name=nom_fichier,
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
