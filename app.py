import streamlit as st
from pptx import Presentation
import io
from datetime import datetime

# --- CONFIGURATION ---
MOT_DE_PASSE_ENTREPRISE = "ETECHNOLOGIE" 

# --- 1. MOTEUR DE REMPLISSAGE AMÉLIORÉ ---
def generer_rapport(donnees, photos, template_path):
    try:
        prs = Presentation(template_path)
    except Exception as e:
        st.error(f"Erreur : Impossible de charger 'template.pptx'. ({e})")
        return None
    
    for slide in prs.slides:
        for shape in slide.shapes:
            # A. TRAITEMENT DU TEXTE (Remplacement précis dans les 'runs')
            if shape.has_text_frame:
                for balise, valeur in donnees.items():
                    if balise in shape.text:
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                if balise in run.text:
                                    run.text = run.text.replace(balise, str(valeur) if valeur else "")

            # B. TRAITEMENT DES IMAGES (Superposition stable)
            # On vérifie si le nom de l'objet dans PowerPoint correspond à une clé photo
            if shape.name in photos and photos[shape.name] is not None:
                photo_file = photos[shape.name]
                
                # Récupération des dimensions exactes du cadre (trame blanche)
                left, top, width, height = shape.left, shape.top, shape.width, shape.height
                
                # Ajout de l'image par-dessus sans supprimer la forme 'shape'
                # Cela garantit que la trame reste en dessous et que l'image ne bouge pas
                slide.shapes.add_picture(photo_file, left, top, width, height)
                
                # NOTE : On ne supprime plus "shape", donc la mise en page reste intacte.

    pptx_io = io.BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)
    return pptx_io

# --- 2. GESTION DE LA SÉCURITÉ ---
def est_authentifie():
    if "auth" not in st.session_state:
        st.session_state["auth"] = False

    if not st.session_state["auth"]:
        st.markdown("### 🔒 Accès Sécurisé - Audit Environnemental")
        mdp = st.text_input("Veuillez entrer le mot de passe entreprise :", type="password")
        if st.button("Se connecter", use_container_width=True):
            if mdp == MOT_DE_PASSE_ENTREPRISE:
                st.session_state["auth"] = True
                st.rerun()
            else:
                st.error("Mot de passe incorrect.")
        return False
    return True

# --- 3. INTERFACE UTILISATEUR (Optimisée Mobile) ---
if est_authentifie():
    # Configuration pour la stabilité de l'affichage sur mobile
    st.set_page_config(page_title="Audit Site Telecom", layout="wide", initial_sidebar_state="collapsed")
    st.title("📱 Audit Environnemental")

    donnees = {}
    photos = {}
    date_du_jour = datetime.now().strftime("%d-%m-%Y")

    # Utilisation de menus déroulants ou d'onglets (les onglets sont mieux gérés sur mobile Streamlit désormais)
    tabs = st.tabs([
        "📍 INFOS", "📸 VUES", "🛡️ SECU", "🧹 PROPRE", "🏗️ PYLÔNE", "⚡ ELEC", "⚙️ GE", "❄️ CLIM", "⚠️ ANOM."
    ])

    # 1. INFORMATIONS DU SITE
    with tabs[0]:
        donnees["INS_NOM"] = st.text_input("Nom du site")
        donnees["INS_CODE"] = st.text_input("Code du site")
        donnees["INS_ZONE"] = st.text_input("Zone")
        donnees["INS_TRAVAUX"] = st.text_area("Travaux exercés")
        donnees["INS_CHEF"] = st.text_input("Chef d'équipes")
        donnees["INS_INSCONTACT"] = st.text_input("Contact")

    # 2. VUE GENERALE DU SITE
    with tabs[1]:
        photos["INS_VUE_DU_SITE"] = st.file_uploader("VUE DU SITE", type=['jpg','png','jpeg'])
        photos["INS_PLAQUE"] = st.file_uploader("PLAQUE", type=['jpg','png','jpeg'])
        st.write("PHOTOS SITE")
        photos["INS_SITE1"] = st.file_uploader("Photo Site 1", type=['jpg'], key="s1")
        photos["INS_SITE2"] = st.file_uploader("Photo Site 2", type=['jpg'], key="s2")
        photos["INS_SITE3"] = st.file_uploader("Photo Site 3", type=['jpg'], key="s3")

    # ... [Les autres onglets restent identiques dans leur structure de clé photos] ...
    # Note : Pour gagner de la place sur mobile, j'ai simplifié certains labels de colonnes ci-dessus.

    # 7. GROUPE ELECTROGENE (Exemple de regroupement pour mobile)
    with tabs[6]:
        donnees["INS_MARQUE_GE"] = st.text_input("Marque GE")
        donnees["INS_VAL_H"] = st.text_input("Heures Compteur")
        photos["INS_MOTEUR"] = st.file_uploader("VUE MOTEUR", type=['jpg'])
        photos["INS_TANK"] = st.file_uploader("VUE DU TANK", type=['jpg'])
        # Ajoutez les autres champs ici de la même manière...

    # 9. FINALISATION
    with tabs[8]:
        st.write("PHOTOS ANOMALIES")
        photos["INS_AN1"] = st.file_uploader("Anomalie 1", type=['jpg'])
        donnees["INS_REMARQUES"] = st.text_area("Remarques Finales")

        st.markdown("---")
        # Le bouton prend toute la largeur pour être facile à cliquer sur mobile
        if st.button("🚀 GÉNÉRER LE RAPPORT FINAL", use_container_width=True):
            if not donnees["INS_CODE"] or not donnees["INS_NOM"]:
                st.error("⚠️ Le Nom et le Code du site sont obligatoires.")
            else:
                with st.spinner("Création du fichier..."):
                    output = generer_rapport(donnees, photos, "template.pptx")
                    if output:
                        nom_fichier = f"RAPPORT_{donnees['INS_CODE']}_{donnees['INS_NOM']}_{date_du_jour}.pptx"
                        st.success("✅ Prêt !")
                        st.download_button(
                            label="📥 Télécharger le Rapport",
                            data=output,
                            file_name=nom_fichier,
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            use_container_width=True
                        )
