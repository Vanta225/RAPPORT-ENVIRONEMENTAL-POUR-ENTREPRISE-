import streamlit as st
from pptx import Presentation
import io
from datetime import datetime

# --- CONFIGURATION ---
MOT_DE_PASSE_ENTREPRISE = "ETECHNOLOGIE" 

# --- 1. MOTEUR DE REMPLISSAGE (Superposition stable) ---
def generer_rapport(donnees, photos, template_path):
    try:
        prs = Presentation(template_path)
    except Exception as e:
        st.error(f"Erreur : Impossible de charger 'template.pptx'. ({e})")
        return None
    
    for slide in prs.slides:
        for shape in slide.shapes:
            # TRAITEMENT DU TEXTE (Conservation du style par les runs)
            if shape.has_text_frame:
                for balise, valeur in donnees.items():
                    if balise in shape.text:
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                if balise in run.text:
                                    run.text = run.text.replace(balise, str(valeur) if valeur else "")

            # TRAITEMENT DES IMAGES (Superposition sans suppression de la trame)
            if shape.name in photos and photos[shape.name] is not None:
                photo_file = photos[shape.name]
                # On récupère les coordonnées de la forme (la case blanche)
                left, top, width, height = shape.left, shape.top, shape.width, shape.height
                
                # On ajoute la photo par-dessus aux mêmes dimensions
                slide.shapes.add_picture(photo_file, left, top, width, height)
                
                # IMPORTANT : On ne supprime pas "shape" pour garder la trame blanche dessous
                # et assurer la stabilité de la mise en page.

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

# --- 3. INTERFACE UTILISATEUR COMPLETE ---
if est_authentifie():
    st.set_page_config(page_title="Audit Site Telecom", layout="wide")
    st.title("📱 Rapport d'Audit Environnemental")

    donnees = {}
    photos = {}
    date_du_jour = datetime.now().strftime("%d-%m-%Y")

    # Création des 9 onglets complets
    tabs = st.tabs([
        "📍 INFOS", "📸 VUE GEN.", "🛡️ SÉCU", 
        "🧹 PROPRETÉ", "🏗️ PYLÔNE", "⚡ ELEC", 
        "⚙️ GE", "❄️ CLIM", "⚠️ ANOMALIES"
    ])

    # 1. INFORMATIONS DU SITE
    with tabs[0]:
        st.subheader("Informations Générales")
        donnees["INS_NOM"] = st.text_input("Nom du site")
        donnees["INS_CODE"] = st.text_input("Code du site")
        donnees["INS_ZONE"] = st.text_input("Zone")
        donnees["INS_TRAVAUX"] = st.text_area("Travaux exercés (INS_TRAVAUX)")
        donnees["INS_CHEF"] = st.text_input("Chef d'équipes")
        donnees["INS_INSCONTACT"] = st.text_input("Contact")

    # 2. VUE GENERALE DU SITE
    with tabs[1]:
        st.subheader("Vues Générales")
        photos["INS_VUE_DU_SITE"] = st.file_uploader("VUE DU SITE", type=['jpg','png','jpeg'])
        photos["INS_PLAQUE"] = st.file_uploader("PLAQUE D'IMMATRICULATION", type=['jpg','png','jpeg'])
        st.divider()
        c1, c2, c3 = st.columns(3)
        photos["INS_SITE1"] = c1.file_uploader("Photo Site 1", type=['jpg'], key="s1")
        photos["INS_SITE2"] = c2.file_uploader("Photo Site 2", type=['jpg'], key="s2")
        photos["INS_SITE3"] = c3.file_uploader("Photo Site 3", type=['jpg'], key="s3")

    # 3. SECURITE DU SITE
    with tabs[2]:
        st.subheader("Accès et Clôture")
        photos["INS_ACCES"] = st.file_uploader("ACCES AU SITE", type=['jpg'])
        photos["INS_PORTAIL"] = st.file_uploader("PORTAIL", type=['jpg'])
        photos["INS_SERRURE"] = st.file_uploader("SERRURE DU PORTAIL", type=['jpg'])
        st.divider()
        st.write("CLOTURE & GUERITE")
        c1, c2, c3 = st.columns(3)
        photos["INS_CLO1"] = c1.file_uploader("Clôture 1", type=['jpg'], key="cl1")
        photos["INS_CLO2"] = c2.file_uploader("Clôture 2", type=['jpg'], key="cl2")
        photos["INS_CLO3"] = c3.file_uploader("Clôture 3", type=['jpg'], key="cl3")
        photos["INS_GUERITE1"] = c1.file_uploader("Vue Ext. Guérite", type=['jpg'], key="g1")
        photos["INS_GUERITE2"] = c2.file_uploader("Vue Int. Guérite", type=['jpg'], key="g2")
        photos["INS_GUERITE3"] = c3.file_uploader("Système Verrouillage", type=['jpg'], key="g3")

    # 4. PROPRETE GENERALE
    with tabs[3]:
        st.subheader("État de Propreté")
        photos["INS_PROPRE"] = st.file_uploader("VUE PROPRETE GENERALE", type=['jpg'])
        photos["INS_DRAIN1"] = st.file_uploader("POINT DE DRAINAGE", type=['jpg'])
        photos["INS_DRAIN2"] = st.file_uploader("VUE POINT EVACUATION EAU", type=['jpg'])
        photos["INS_DRAIN3"] = st.file_uploader("VUE EPANDAGE", type=['jpg'])

    # 5. PYLÔNE ET EQUIPEMENTS
    with tabs[4]:
        st.subheader("Infrastructures")
        st.write("FONDATIONS & ANCRAGES")
        c1, c2, c3 = st.columns(3)
        photos["INS_FONDA1"] = c1.file_uploader("Fondation 1", type=['jpg'], key="f1")
        photos["INS_FONDA2"] = c2.file_uploader("Fondation 2", type=['jpg'], key="f2")
        photos["INS_FONDA3"] = c3.file_uploader("Fondation 3", type=['jpg'], key="f3")
        st.divider()
        st.write("EQUIPEMENTS & RACKS")
        photos["EQUIP1"] = c1.file_uploader("Equipement 1", type=['jpg'], key="e1")
        photos["EQUIP2"] = c2.file_uploader("Equipement 2", type=['jpg'], key="e2")
        photos["EQUIP3"] = c3.file_uploader("Equipement 3", type=['jpg'], key="e3")

    # 6. ELECTRICITE
    with tabs[5]:
        st.subheader("Installation Électrique")
        c1, c2, c3 = st.columns(3)
        photos["INS_CIE1"] = c1.file_uploader("Intérieur Niche", type=['jpg'])
        photos["INS_CIE2"] = c2.file_uploader("Compteur CIE", type=['jpg'])
        photos["INS_CIE3"] = c3.file_uploader("Extérieur Niche", type=['jpg'])
        st.divider()
        photos["INS_TGBT"] = st.file_uploader("TGBT", type=['jpg'])
        photos["INS_COF"] = st.file_uploader("COFFRET INVERSEUR", type=['jpg'])

    # 7. GROUPE ELECTROGENE
    with tabs[6]:
        st.subheader("Groupe Électrogène")
        c1, c2 = st.columns(2)
        photos["INS_MOTEUR"] = c1.file_uploader("DE FACE DU GE", type=['jpg'], key="ge1")
        photos["INS_GE_LAT"] = c2.file_uploader("LATERAL DU GE", type=['jpg'], key="ge2")
        photos["INS_GE_FACE"] = c1.file_uploader("VUE MOTEUR", type=['jpg'], key="ge3")
        photos["INS_TANK"] = c2.file_uploader("VUE DU TANK", type=['jpg'], key="ge4")
        photos["INS_COMPT"] = c1.file_uploader("VUE COMPTEUR HORAIRE", type=['jpg'], key="ge5")
        photos["INS_RAD"] = c2.file_uploader("VUE RADIATEUR", type=['jpg'], key="ge6")
        photos["INS_FILTRE"] = c1.file_uploader("VUE FILTRES", type=['jpg'], key="ge7")
        photos["INS_INT_GE"] = c2.file_uploader("VUE INTERIEUR GE", type=['jpg'], key="ge8")
        
        st.divider()
        donnees["INS_MARQUE_GE"] = st.text_input("Marque GE")
        donnees["INS_VAL_H"] = st.text_input("Heures Compteur")
        donnees["INS_VAL_CARB"] = st.text_input("Niveau Carburant (%)")
        donnees["INS_VAL_PUISS"] = st.text_input("Puissance GE")
        donnees["INS_MARQUE_MOT"] = st.text_input("Marque Moteur")
        donnees["INS_RES"] = st.text_input("Capacité Réservoir")

    # 8. CLIMATISATION
    with tabs[7]:
        st.subheader("Maintenance Climatisation")
        c1, c2, c3 = st.columns(3)
        photos["INS_CLIM1"] = c1.file_uploader("AVANT LAVAGE", type=['jpg'], key="clm1")
        photos["INS_CLIM2"] = c2.file_uploader("PENDANT LAVAGE", type=['jpg'], key="clm2")
        photos["INS_CLIM3"] = c3.file_uploader("APRES LAVAGE", type=['jpg'], key="clm3")

    # 9. ANOMALIES & BOUTON FINAL
    with tabs[8]:
        st.subheader("Observations Finales")
        c1, c2, c3 = st.columns(3)
        photos["INS_AN1"] = c1.file_uploader("Anomalie 1", type=['jpg'])
        photos["INS_AN2"] = c2.file_uploader("Anomalie 2", type=['jpg'])
        photos["INS_AN3"] = c3.file_uploader("Anomalie 3", type=['jpg'])
        
        donnees["INS_REMARQUES"] = st.text_area("Remarques Finales / Recommandations")

        st.markdown("---")
        if st.button("🚀 GÉNÉRER LE RAPPORT FINAL", use_container_width=True):
            if not donnees["INS_CODE"] or not donnees["INS_NOM"]:
                st.error("⚠️ Erreur : Le Nom et le Code du site sont obligatoires pour nommer le fichier.")
            else:
                with st.spinner("Génération du PowerPoint en cours..."):
                    output = generer_rapport(donnees, photos, "template.pptx")
                    if output:
                        nom_fichier = f"RAPPORT_AUDIT_{donnees['INS_CODE']}_{donnees['INS_NOM']}_{date_du_jour}.pptx"
                        
                        st.success("✅ Rapport généré avec succès !")
                        st.download_button(
                            label="📥 Télécharger le Rapport (.pptx)",
                            data=output,
                            file_name=nom_fichier,
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            use_container_width=True
                        )
