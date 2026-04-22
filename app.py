import streamlit as st
from pptx import Presentation
import io
from datetime import datetime
from PIL import Image

# --- CONFIGURATION DU MOT DE PASSE ---
MOT_DE_PASSE_ENTREPRISE = "ETECHNOLOGIE"


# =========================================================
# 📱 FIX MOBILE IMAGES (ANDROID / IOS SAFE)
# =========================================================
def normaliser_image(file):
    try:
        img = Image.open(file)

        # correction rotation mobile (Android surtout)
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

        output = io.BytesIO()
        img.save(output, format="JPEG", quality=85, optimize=True)
        output.seek(0)

        return output

    except Exception as e:
        st.error(f"Erreur image mobile : {e}")
        return None


# =========================================================
# 📊 GENERATION RAPPORT POWERPOINT
# =========================================================
def generer_rapport(donnees, photos, template_path):

    try:
        prs = Presentation(template_path)
    except Exception as e:
        st.error(f"Erreur template : {e}")
        return None

    for slide in prs.slides:
        for shape in slide.shapes:

            # ---------------- TEXTE CLEAN ----------------
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

            # ---------------- IMAGES FIX MOBILE ----------------
            if shape.name in photos and photos[shape.name] is not None:

                photo_file = normaliser_image(photos[shape.name])

                if photo_file is not None:

                    left, top, width, height = shape.left, shape.top, shape.width, shape.height

                    img = slide.shapes.add_picture(photo_file, left, top)

                    # verrouillage position EXACTE template
                    img.left = left
                    img.top = top
                    img.width = width
                    img.height = height

                    # suppression placeholder safe
                    try:
                        sp = shape._element
                        sp.getparent().remove(sp)
                    except:
                        pass

    # ---------------- FIX DOWNLOAD MOBILE ----------------
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
        st.markdown("### 🔒 Accès Sécurisé")
        mdp = st.text_input("Mot de passe :", type="password")

        if st.button("Connexion"):
            if mdp == MOT_DE_PASSE_ENTREPRISE:
                st.session_state["auth"] = True
                st.rerun()
            else:
                st.error("Mot de passe incorrect")

        return False

    return True


# =========================================================
# 🖥️ INTERFACE
# =========================================================
if est_authentifie():

    st.set_page_config(page_title="Audit Site Telecom", layout="wide")
    st.title("📱 Rapport d'Audit Automatisé")

    donnees = {}
    photos = {}
    date_du_jour = datetime.now().strftime("%d-%m-%Y")

    tabs = st.tabs([
        "📍 INFOS SITE", "📸 VUE GENERALE", "🛡️ SECURITE",
        "🧹 PROPRETE", "🏗️ PYLÔNE", "⚡ ELECTRICITE",
        "⚙️ GE", "❄️ CLIM", "⚠️ ANOMALIES"
    ])

    # ================= INFOS =================
    with tabs[0]:
        donnees["INS_NOM"] = st.text_input("Nom site")
        donnees["INS_CODE"] = st.text_input("Code site")
        donnees["INS_ZONE"] = st.text_input("Zone")
        donnees["INS_TRAVAUX"] = st.text_area("Travaux exécutés")
        donnees["INS_CHEF"] = st.text_input("Chef équipe")
        donnees["INS_INSCONTACT"] = st.text_input("Contact")

    # ================= IMAGES =================
    with tabs[1]:
        photos["INS_VUE_DU_SITE"] = st.file_uploader(
            "Vue site",
            type=['jpg','jpeg','png']
        )

        photos["INS_PLAQUE"] = st.file_uploader(
            "Plaque",
            type=['jpg','jpeg','png']
        )

    # ================= GENERATION =================
    with tabs[8]:
        photos["INS_AN1"] = st.file_uploader("Anomalie", type=['jpg','jpeg','png'])
        donnees["INS_REMARQUES"] = st.text_area("Remarques")

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
                            label="📥 Télécharger le rapport",
                            data=output.getvalue(),
                            file_name=nom_fichier,
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            use_container_width=True
                        )
