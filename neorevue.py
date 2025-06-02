import json
import pandas as pd
import streamlit as st
from io import BytesIO
import openpyxl
from datetime import datetime

# Configuration Streamlit
st.set_page_config(
    page_title="IFS NEO Data Extractor",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS personnalisé pour améliorer l'apparence
def apply_custom_css():
    st.markdown("""
        <style>
        .main-header {
            font-size: 2.5rem;
            color: #1f77b4;
            text-align: center;
            margin-bottom: 2rem;
        }
        .section-header {
            font-size: 1.5rem;
            color: #2e8b57;
            border-bottom: 2px solid #2e8b57;
            padding-bottom: 0.5rem;
            margin: 1rem 0;
        }
        .info-box {
            background-color: #f0f8ff;
            border-left: 4px solid #1f77b4;
            padding: 1rem;
            margin: 1rem 0;
        }
        .warning-box {
            background-color: #fff3cd;
            border-left: 4px solid #ffc107;
            padding: 1rem;
            margin: 1rem 0;
        }
        .success-box {
            background-color: #d4edda;
            border-left: 4px solid #28a745;
            padding: 1rem;
            margin: 1rem 0;
        }
        </style>
    """, unsafe_allow_html=True)

def flatten_json_safe(nested_json, parent_key='', sep='_'):
    """Aplatit une structure JSON imbriquée de manière sécurisée."""
    items = []
    if isinstance(nested_json, dict):
        for k, v in nested_json.items():
            new_key = f'{parent_key}{sep}{k}' if parent_key else k
            if isinstance(v, dict):
                items.extend(flatten_json_safe(v, new_key, sep=sep).items())
            elif isinstance(v, list):
                for i, item in enumerate(v):
                    items.extend(flatten_json_safe(item, f'{new_key}{sep}{i}', sep=sep).items())
            else:
                items.append((new_key, v))
    else:
        items.append((parent_key, nested_json))
    return dict(items)

def extract_from_flattened(flattened_data, mapping, selected_fields):
    """Extrait les données du JSON aplati selon le mapping fourni."""
    extracted_data = {}
    for label, flat_path in mapping.items():
        if label in selected_fields:
            extracted_data[label] = flattened_data.get(flat_path, 'N/A')
    return extracted_data

def get_user_comments():
    """Récupère tous les commentaires utilisateur depuis la session state."""
    comments = {}
    for key, value in st.session_state.items():
        if key.startswith(('profile_comment_', 'checklist_comment_', 'non_conformity_comment_')):
            comments[key] = value
    return comments

def initialize_session_state():
    """Initialise les variables de session state nécessaires."""
    if 'json_data' not in st.session_state:
        st.session_state.json_data = None
    if 'profile_data' not in st.session_state:
        st.session_state.profile_data = {}
    if 'checklist_data' not in st.session_state:
        st.session_state.checklist_data = []
    if 'non_conformities' not in st.session_state:
        st.session_state.non_conformities = []

# Mapping complet des champs
FLATTENED_FIELD_MAPPING = {
    "Nom du site à auditer": "data_modules_food_8_questions_companyName_answer",
    "N° COID du portail": "data_modules_food_8_questions_companyCoid_answer",
    "Code GLN": "data_modules_food_8_questions_companyGln_answer_0_rootQuestions_companyGlnNumber_answer",
    "Rue": "data_modules_food_8_questions_companyStreetNo_answer",
    "Code postal": "data_modules_food_8_questions_companyZip_answer",
    "Nom de la ville": "data_modules_food_8_questions_companyCity_answer",
    "Pays": "data_modules_food_8_questions_companyCountry_answer",
    "Téléphone": "data_modules_food_8_questions_companyTelephone_answer",
    "Latitude": "data_modules_food_8_questions_companyGpsLatitude_answer",
    "Longitude": "data_modules_food_8_questions_companyGpsLongitude_answer",
    "Email": "data_modules_food_8_questions_companyEmail_answer",
    "Nom du siège social": "data_modules_food_8_questions_headquartersName_answer",
    "Rue (siège social)": "data_modules_food_8_questions_headquartersStreetNo_answer",
    "Nom de la ville (siège social)": "data_modules_food_8_questions_headquartersCity_answer",
    "Code postal (siège social)": "data_modules_food_8_questions_headquartersZip_answer",
    "Pays (siège social)": "data_modules_food_8_questions_headquartersCountry_answer",
    "Téléphone (siège social)": "data_modules_food_8_questions_headquartersTelephone_answer",
    "Surface couverte de l'entreprise (m²)": "data_modules_food_8_questions_productionAreaSize_answer",
    "Nombre de bâtiments": "data_modules_food_8_questions_numberOfBuildings_answer",
    "Nombre de lignes de production": "data_modules_food_8_questions_numberOfProductionLines_answer",
    "Nombre d'étages": "data_modules_food_8_questions_numberOfFloors_answer",
    "Nombre maximum d'employés dans l'année, au pic de production": "data_modules_food_8_questions_numberOfEmployeesForTimeCalculation_answer",
    "Langue parlée et écrite sur le site": "data_modules_food_8_questions_workingLanguage_answer",
    "Périmètre de l'audit": "data_modules_food_8_questions_scopeCertificateScopeDescription_en_answer",
    "Process et activités": "data_modules_food_8_questions_scopeProductGroupsDescription_answer",
    "Activité saisonnière ? (O/N)": "data_modules_food_8_questions_seasonalProduction_answer",
    "Une partie du procédé de fabrication est-elle sous traitée? (OUI/NON)": "data_modules_food_8_questions_partlyOutsourcedProcesses_answer",
    "Si oui lister les procédés sous-traités": "data_modules_food_8_questions_partlyOutsourcedProcessesDescription_answer",
    "Avez-vous des produits totalement sous-traités? (OUI/NON)": "data_modules_food_8_questions_fullyOutsourcedProducts_answer",
    "Si oui, lister les produits totalement sous-traités": "data_modules_food_8_questions_fullyOutsourcedProductsDescription_answer",
    "Avez-vous des produits de négoce? (OUI/NON)": "data_modules_food_8_questions_tradedProductsBrokerActivity_answer",
    "Si oui, lister les produits de négoce": "data_modules_food_8_questions_tradedProductsBrokerActivityDescription_answer",
    "Produits à exclure du champ d'audit (OUI/NON)": "data_modules_food_8_questions_exclusions_answer",
    "Préciser les produits à exclure": "data_modules_food_8_questions_exclusionsDescription_answer"
}

def process_json_file(uploaded_file):
    """Traite le fichier JSON uploadé et extrait les données."""
    try:
        json_data = json.load(uploaded_file)
        st.session_state.json_data = json_data
        
        # Aplatir les données JSON
        flattened_json_data = flatten_json_safe(json_data)
        
        # Extraire les données de profil
        profile_data = extract_from_flattened(
            flattened_json_data, 
            FLATTENED_FIELD_MAPPING, 
            list(FLATTENED_FIELD_MAPPING.keys())
        )
        st.session_state.profile_data = profile_data
        
        # Extraire les données de checklist
        checklist_data = []
        if 'data' in json_data and 'modules' in json_data['data']:
            modules = json_data['data']['modules']
            if 'food_8' in modules and 'checklists' in modules['food_8']:
                checklists = modules['food_8']['checklists']
                if 'checklistFood8' in checklists and 'resultScorings' in checklists['checklistFood8']:
                    for uuid, scoring in checklists['checklistFood8']['resultScorings'].items():
                        checklist_data.append({
                            "Num": uuid,
                            "Explanation": scoring['answers'].get('englishExplanationText', 'N/A'),
                            "Detailed Explanation": scoring['answers'].get('explanationText', 'N/A'),
                            "Score": scoring['score']['label'],
                            "Response": scoring['answers'].get('fieldAnswers', 'N/A')
                        })
        
        st.session_state.checklist_data = checklist_data
        
        # Extraire les non-conformités
        non_conformities = [item for item in checklist_data if item['Score'] != 'A']
        st.session_state.non_conformities = non_conformities
        
        return True, "Fichier traité avec succès!"
        
    except json.JSONDecodeError as e:
        return False, f"Erreur lors du décodage JSON: {str(e)}"
    except Exception as e:
        return False, f"Erreur lors du traitement du fichier: {str(e)}"

def display_profile_section():
    """Affiche la section du profil avec possibilité d'ajout de commentaires."""
    st.markdown('<div class="section-header">📋 Profil de l\'entreprise</div>', unsafe_allow_html=True)
    
    if not st.session_state.profile_data:
        st.warning("Aucune donnée de profil disponible. Veuillez d'abord charger un fichier IFS.")
        return
    
    # Organiser les données en colonnes pour une meilleure présentation
    col1, col2 = st.columns(2)
    
    profile_items = list(st.session_state.profile_data.items())
    mid_point = len(profile_items) // 2
    
    with col1:
        for field, value in profile_items[:mid_point]:
            st.text_input(f"**{field}**", value=str(value), key=f"profile_field_{field}", disabled=True)
            # Zone de commentaire pour chaque champ
            st.text_area(
                f"Commentaire - {field}",
                key=f"profile_comment_{field}",
                height=60,
                placeholder="Ajoutez vos commentaires ici..."
            )
    
    with col2:
        for field, value in profile_items[mid_point:]:
            st.text_input(f"**{field}**", value=str(value), key=f"profile_field_{field}", disabled=True)
            # Zone de commentaire pour chaque champ
            st.text_area(
                f"Commentaire - {field}",
                key=f"profile_comment_{field}",
                height=60,
                placeholder="Ajoutez vos commentaires ici..."
            )

def display_checklist_section():
    """Affiche la section de la checklist complète."""
    st.markdown('<div class="section-header">✅ Checklist complète</div>', unsafe_allow_html=True)
    
    if not st.session_state.checklist_data:
        st.warning("Aucune donnée de checklist disponible. Veuillez d'abord charger un fichier IFS.")
        return
    
    # Filtre par score
    score_filter = st.selectbox(
        "Filtrer par score:",
        ["Tous", "A", "B", "C", "D", "Non applicable"]
    )
    
    # Appliquer le filtre
    filtered_data = st.session_state.checklist_data
    if score_filter != "Tous":
        filtered_data = [item for item in st.session_state.checklist_data if item['Score'] == score_filter]
    
    st.info(f"Affichage de {len(filtered_data)} éléments sur {len(st.session_state.checklist_data)} au total")
    
    # Afficher les éléments de la checklist
    for i, item in enumerate(filtered_data):
        with st.expander(f"Exigence {item['Num']} - Score: {item['Score']}", expanded=False):
            col1, col2 = st.columns([3, 1])
            
            with col1:
                st.write(f"**Explication:** {item['Explanation']}")
                st.write(f"**Explication détaillée:** {item['Detailed Explanation']}")
                st.write(f"**Réponse:** {item['Response']}")
                
                # Zone de commentaire pour chaque élément
                st.text_area(
                    "Commentaire de l'auditeur:",
                    key=f"checklist_comment_{item['Num']}",
                    height=100,
                    placeholder="Ajoutez vos observations, commentaires ou actions à prendre..."
                )
            
            with col2:
                # Affichage du score avec couleur
                score_color = {
                    'A': '#28a745',
                    'B': '#ffc107', 
                    'C': '#fd7e14',
                    'D': '#dc3545',
                    'Non applicable': '#6c757d'
                }.get(item['Score'], '#6c757d')
                
                st.markdown(f"""
                    <div style="background-color: {score_color}; color: white; 
                                padding: 10px; border-radius: 5px; text-align: center; 
                                font-weight: bold; font-size: 18px;">
                        {item['Score']}
                    </div>
                """, unsafe_allow_html=True)

def display_non_conformities_section():
    """Affiche la section des non-conformités."""
    st.markdown('<div class="section-header">⚠️ Non-conformités</div>', unsafe_allow_html=True)
    
    if not st.session_state.non_conformities:
        st.success("Aucune non-conformité détectée ! Toutes les exigences sont notées A.")
        return
    
    st.warning(f"Nombre de non-conformités détectées: {len(st.session_state.non_conformities)}")
    
    # Statistiques des non-conformités
    scores_count = {}
    for item in st.session_state.non_conformities:
        score = item['Score']
        scores_count[score] = scores_count.get(score, 0) + 1
    
    col1, col2, col3, col4 = st.columns(4)
    for i, (score, count) in enumerate(scores_count.items()):
        with [col1, col2, col3, col4][i % 4]:
            st.metric(f"Score {score}", count)
    
    # Afficher chaque non-conformité
    for item in st.session_state.non_conformities:
        with st.container():
            st.markdown(f"""
                <div style="border-left: 4px solid #dc3545; padding: 15px; margin: 10px 0; 
                           background-color: #f8f9fa;">
                    <h4>🔍 Exigence {item['Num']} - Score: {item['Score']}</h4>
                </div>
            """, unsafe_allow_html=True)
            
            col1, col2 = st.columns([3, 1])
            
            with col1:
                st.write(f"**Explication:** {item['Explanation']}")
                st.write(f"**Explication détaillée:** {item['Detailed Explanation']}")
                st.write(f"**Réponse:** {item['Response']}")
                
                # Plan d'action
                st.text_area(
                    "Plan d'action corrective:",
                    key=f"non_conformity_action_{item['Num']}",
                    height=100,
                    placeholder="Décrivez les actions correctives à mettre en place..."
                )
                
                # Commentaire de l'auditeur
                st.text_area(
                    "Commentaire de l'auditeur:",
                    key=f"non_conformity_comment_{item['Num']}",
                    height=80,
                    placeholder="Observations de l'auditeur..."
                )
            
            with col2:
                # Sélection de la priorité
                priority = st.selectbox(
                    "Priorité:",
                    ["Haute", "Moyenne", "Basse"],
                    key=f"priority_{item['Num']}"
                )
                
                # Date limite
                deadline = st.date_input(
                    "Date limite:",
                    key=f"deadline_{item['Num']}"
                )
                
                # Responsable
                responsible = st.text_input(
                    "Responsable:",
                    key=f"responsible_{item['Num']}",
                    placeholder="Nom du responsable"
                )

def create_enhanced_excel_export():
    """Crée un fichier Excel enrichi avec toutes les données et commentaires."""
    if not st.session_state.profile_data:
        st.error("Aucune donnée à exporter. Veuillez d'abord charger un fichier IFS.")
        return None
    
    # Récupérer tous les commentaires
    comments = get_user_comments()
    
    # Créer le fichier Excel en mémoire
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Onglet Profil
        profile_rows = []
        for field, value in st.session_state.profile_data.items():
            comment_key = f"profile_comment_{field}"
            comment = comments.get(comment_key, "")
            profile_rows.append({
                "Champ": field,
                "Valeur": value,
                "Commentaire": comment,
                "Réponse auditeur": ""
            })
        
        df_profile = pd.DataFrame(profile_rows)
        df_profile.to_excel(writer, index=False, sheet_name="Profil")
        
        # Onglet Checklist complète
        checklist_rows = []
        for item in st.session_state.checklist_data:
            comment_key = f"checklist_comment_{item['Num']}"
            comment = comments.get(comment_key, "")
            checklist_rows.append({
                "Numéro": item['Num'],
                "Explication": item['Explanation'],
                "Explication détaillée": item['Detailed Explanation'],
                "Score": item['Score'],
                "Réponse": item['Response'],
                "Commentaire auditeur": comment,
                "Action requise": ""
            })
        
        df_checklist = pd.DataFrame(checklist_rows)
        df_checklist.to_excel(writer, index=False, sheet_name="Checklist")
        
        # Onglet Non-conformités avec plan d'action
        nc_rows = []
        for item in st.session_state.non_conformities:
            comment_key = f"non_conformity_comment_{item['Num']}"
            action_key = f"non_conformity_action_{item['Num']}"
            priority_key = f"priority_{item['Num']}"
            deadline_key = f"deadline_{item['Num']}"
            responsible_key = f"responsible_{item['Num']}"
            
            comment = comments.get(comment_key, "")
            action = st.session_state.get(action_key, "")
            priority = st.session_state.get(priority_key, "")
            deadline = st.session_state.get(deadline_key, "")
            responsible = st.session_state.get(responsible_key, "")
            
            nc_rows.append({
                "Numéro": item['Num'],
                "Score": item['Score'],
                "Explication": item['Explanation'],
                "Explication détaillée": item['Detailed Explanation'],
                "Réponse": item['Response'],
                "Commentaire auditeur": comment,
                "Plan d'action": action,
                "Priorité": priority,
                "Date limite": deadline,
                "Responsable": responsible,
                "Statut": "En attente"
            })
        
        df_nc = pd.DataFrame(nc_rows)
        df_nc.to_excel(writer, index=False, sheet_name="Non-conformités")
        
        # Onglet Résumé
        summary_data = {
            "Indicateur": [
                "Nombre total d'exigences",
                "Exigences conformes (A)",
                "Non-conformités mineures (B)",
                "Non-conformités majeures (C)",
                "Non-conformités critiques (D)",
                "Taux de conformité (%)"
            ],
            "Valeur": [
                len(st.session_state.checklist_data),
                len([x for x in st.session_state.checklist_data if x['Score'] == 'A']),
                len([x for x in st.session_state.checklist_data if x['Score'] == 'B']),
                len([x for x in st.session_state.checklist_data if x['Score'] == 'C']),
                len([x for x in st.session_state.checklist_data if x['Score'] == 'D']),
                round((len([x for x in st.session_state.checklist_data if x['Score'] == 'A']) / 
                      len(st.session_state.checklist_data)) * 100, 2) if st.session_state.checklist_data else 0
            ]
        }
        
        df_summary = pd.DataFrame(summary_data)
        df_summary.to_excel(writer, index=False, sheet_name="Résumé")
        
        # Ajuster la largeur des colonnes
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
    
    output.seek(0)
    return output

def main():
    """Fonction principale de l'application."""
    # Initialiser la session state
    initialize_session_state()
    
    # Appliquer le CSS personnalisé
    apply_custom_css()
    
    # En-tête principal
    st.markdown('<div class="main-header">🔍 IFS NEO Data Extractor</div>', unsafe_allow_html=True)
    st.markdown('<div class="info-box">Application d\'extraction et d\'analyse des données d\'audit IFS</div>', unsafe_allow_html=True)
    
    # Navigation dans la sidebar
    st.sidebar.title("📋 Navigation")
    
    # Upload du fichier IFS
    st.sidebar.markdown("### 📁 Chargement des fichiers")
    uploaded_json_file = st.sidebar.file_uploader(
        "Charger le fichier IFS (.ifs)",
        type="ifs",
        help="Sélectionnez le fichier d'audit IFS exporté depuis NEO"
    )
    
    # Traitement du fichier JSON
    if uploaded_json_file and st.session_state.json_data is None:
        with st.spinner("Traitement du fichier IFS en cours..."):
            success, message = process_json_file(uploaded_json_file)
            if success:
                st.sidebar.success(message)
            else:
                st.sidebar.error(message)
                return
    
    # Menu de navigation principal
    if st.session_state.json_data:
        st.sidebar.markdown("### 🎯 Sections disponibles")
        page = st.sidebar.radio(
            "Choisissez une section:",
            ["📋 Profil de l'entreprise", "✅ Checklist complète", "⚠️ Non-conformités", "📊 Tableau de bord", "📄 Export Excel"]
        )
        
        # Affichage des sections selon la navigation
        if page == "📋 Profil de l'entreprise":
            display_profile_section()
            
        elif page == "✅ Checklist complète":
            display_checklist_section()
            
        elif page == "⚠️ Non-conformités":
            display_non_conformities_section()
            
        elif page == "📊 Tableau de bord":
            st.markdown('<div class="section-header">📊 Tableau de bord de l\'audit</div>', unsafe_allow_html=True)
            
            # Métriques principales
            col1, col2, col3, col4 = st.columns(4)
            
            total_items = len(st.session_state.checklist_data)
            conformes = len([x for x in st.session_state.checklist_data if x['Score'] == 'A'])
            non_conformites = len(st.session_state.non_conformities)
            taux_conformite = (conformes / total_items * 100) if total_items > 0 else 0
            
            with col1:
                st.metric("Total exigences", total_items)
            with col2:
                st.metric("Conformes (A)", conformes)
            with col3:
                st.metric("Non-conformités", non_conformites)
            with col4:
                st.metric("Taux conformité", f"{taux_conformite:.1f}%")
            
            # Répartition des scores
            if st.session_state.checklist_data:
                scores_count = {}
                for item in st.session_state.checklist_data:
                    score = item['Score']
                    scores_count[score] = scores_count.get(score, 0) + 1
                
                # Graphique de répartition
                st.subheader("Répartition des scores")
                chart_data = pd.DataFrame(list(scores_count.items()), columns=['Score', 'Nombre'])
                st.bar_chart(chart_data.set_index('Score'))
            
        elif page == "📄 Export Excel":
            st.markdown('<div class="section-header">📄 Export des données</div>', unsafe_allow_html=True)
            
            st.info("Exportez toutes les données collectées avec vos commentaires dans un fichier Excel structuré.")
            
            if st.button("🔄 Générer le fichier Excel", type="primary"):
                with st.spinner("Génération du fichier Excel..."):
                    excel_file = create_enhanced_excel_export()
                    
                    if excel_file:
                        # Nom du fichier avec COID et date
                        coid = st.session_state.profile_data.get("N° COID du portail", "inconnu")
                        date_str = datetime.now().strftime("%Y%m%d_%H%M")
                        filename = f"audit_IFS_{coid}_{date_str}.xlsx"
                        
                        st.download_button(
                            label="📥 Télécharger le rapport Excel",
                            data=excel_file,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        
                        st.success("Fichier Excel généré avec succès!")
                        
                        # Informations sur le contenu du fichier
                        st.markdown("### 📋 Contenu du fichier Excel:")
                        st.markdown("""
                        - **Profil**: Informations sur l'entreprise avec commentaires
                        - **Checklist**: Liste complète des exigences avec scores et commentaires
                        - **Non-conformités**: Plan d'action détaillé pour chaque non-conformité
                        - **Résumé**: Statistiques et indicateurs de performance
                        """)
    
    else:
        # Page d'accueil si aucun fichier n'est chargé
        st.markdown("""
        ### 🚀 Bienvenue dans l'extracteur de données IFS NEO
        
        Cette application vous permet de:
        - 📊 Extraire et analyser les données d'audit IFS
        - 💬 Ajouter vos commentaires et observations
        - 📋 Créer des plans d'action pour les non-conformités
        - 📄 Exporter tout dans un rapport Excel structuré
        
        **Pour commencer:**
        1. Chargez votre fichier d'audit IFS (.ifs) dans la barre latérale
        2. Naviguez entre les différentes sections
        3. Ajoutez vos commentaires et plans d'action
        4. Exportez le rapport final
        """)
        
        st.markdown('<div class="warning-box">⚠️ Veuillez charger un fichier IFS pour commencer l\'analyse.</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
