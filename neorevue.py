import json
import pandas as pd
import streamlit as st
from io import BytesIO
import requests
from datetime import datetime

# Set Streamlit to wide mode
st.set_page_config(layout="wide")

# Function to flatten the nested JSON structure
def flatten_json_safe(nested_json, parent_key='', sep='_'):
    """Flatten a nested JSON dictionary, safely handling strings and primitives."""
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

# Function to extract data from the flattened JSON
def extract_from_flattened(flattened_data, mapping, selected_fields):
    extracted_data = {}
    for label, flat_path in mapping.items():
        if label in selected_fields:
            extracted_data[label] = flattened_data.get(flat_path, 'N/A')
    return extracted_data

# Custom CSS for the table display
def apply_table_css():
    st.markdown(
        """
        <style>
        table {
            width: 100%;
            border-collapse: collapse;
            background-color: #f9f9f9;
        }
        th, td {
            border: 1px solid #ddd;
            padding: 10px;
            text-align: left;
        }
        th {
            background-color: #f2f2f2;
        }
        </style>
        """, unsafe_allow_html=True
    )

# Load the CSV mapping for UUIDs corresponding to NUM from a URL
def load_uuid_mapping_from_url(url):
    try:
        response = requests.get(url)
        if response.status_code == 200:
            from io import StringIO
            csv_data = StringIO(response.text)
            uuid_mapping_df = pd.read_csv(csv_data)
            
            # Check if the columns 'UUID', 'Num', 'Chapitre', 'Theme', and 'SSTheme' exist and have non-empty values
            required_columns = ['UUID', 'Num', 'Chapitre', 'Theme', 'SSTheme']
            for column in required_columns:
                if column not in uuid_mapping_df.columns:
                    st.error(f"Le fichier CSV doit contenir une colonne '{column}' avec des valeurs valides.")
                    return pd.DataFrame()
            
            uuid_mapping_df = uuid_mapping_df.dropna(subset=['UUID', 'Num'])  # Drop rows with empty 'UUID' or 'Num' values
            uuid_mapping_df['Chapitre'] = uuid_mapping_df['Chapitre'].astype(str).str.strip()
            uuid_mapping_df = uuid_mapping_df.drop_duplicates(subset=['Chapitre', 'Num'])  # Remove duplicate rows based on 'Chapitre' and 'Num'
            return uuid_mapping_df
        else:
            st.error("Impossible de charger le fichier CSV des UUID depuis l'URL fourni.")
            return pd.DataFrame()
    except Exception as e:
        st.error(f"Erreur lors du chargement du mapping UUID: {str(e)}")
        return pd.DataFrame()

# Complete mapping based on your provided field names and JSON structure
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

# URL for the UUID CSV
UUID_MAPPING_URL = "https://raw.githubusercontent.com/M00N69/Gemini-Knowledge/refs/heads/main/IFSV8listUUID.csv"

# Functions for work save/load
def save_work_to_excel(profile_data, checklist_data, non_conformities, edited_profile=None, edited_checklist=None, edited_nc=None, coid="inconnu"):
    """Save current work with comments to Excel for later resume"""
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Metadata sheet to identify this as a work file
        metadata = pd.DataFrame({
            "Type": ["IFS_WORK_SAVE"],
            "COID": [coid],
            "Date": [pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")]
        })
        metadata.to_excel(writer, index=False, sheet_name="METADATA")
        
        # Profile with comments
        if edited_profile is not None:
            profile_save_df = edited_profile.copy()
        else:
            profile_save_df = pd.DataFrame([
                {"Champ": k, "Valeur": v, "Commentaire": ""} 
                for k, v in profile_data.items()
            ])
        profile_save_df.to_excel(writer, index=False, sheet_name="Profile_Work")

        # Checklist with comments
        if edited_checklist is not None:
            checklist_save_df = edited_checklist.copy()
        else:
            checklist_save_df = pd.DataFrame(checklist_data)
            checklist_save_df['Commentaire'] = ''
        checklist_save_df.to_excel(writer, index=False, sheet_name="Checklist_Work")

        # Non-conformities with comments and action plans
        if non_conformities:
            if edited_nc is not None:
                nc_save_df = edited_nc.copy()
            else:
                nc_save_df = pd.DataFrame(non_conformities)
                nc_save_df['Commentaire'] = ''
                nc_save_df['Plan d\'action'] = ''
            nc_save_df.to_excel(writer, index=False, sheet_name="NonConformities_Work")

        # Adjust column widths
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            for col in worksheet.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min((max_length + 2) * 1.2, 50)
                worksheet.column_dimensions[column].width = adjusted_width

    output.seek(0)
    return output

def load_work_from_excel(uploaded_file):
    """Load previously saved work from Excel"""
    try:
        # Read all sheets
        excel_data = pd.read_excel(uploaded_file, sheet_name=None)
        
        # Check if this is a work file
        if "METADATA" not in excel_data:
            return None, "Ce fichier n'est pas un fichier de travail sauvegardé."
        
        metadata = excel_data["METADATA"]
        if metadata.iloc[0]["Type"] != "IFS_WORK_SAVE":
            return None, "Ce fichier n'est pas un fichier de travail IFS valide."
        
        work_data = {}
        
        # Load profile work
        if "Profile_Work" in excel_data:
            work_data['profile'] = excel_data["Profile_Work"]
        
        # Load checklist work
        if "Checklist_Work" in excel_data:
            work_data['checklist'] = excel_data["Checklist_Work"]
        
        # Load non-conformities work
        if "NonConformities_Work" in excel_data:
            work_data['nc'] = excel_data["NonConformities_Work"]
        
        coid = metadata.iloc[0]["COID"]
        save_date = metadata.iloc[0]["Date"]
        
        return work_data, f"Travail chargé avec succès (COID: {coid}, sauvegardé le {save_date})"
        
    except Exception as e:
        return None, f"Erreur lors du chargement du fichier de travail: {str(e)}"

# Streamlit app
st.sidebar.title("Menu de Navigation")
main_option = st.sidebar.radio("Choisissez une fonctionnalité:", ["Traitement des rapports IFS", "Gestion des fichiers Excel", "Reprendre un travail sauvegardé"])

st.title("IFS NEO Form Data Extractor")

if main_option == "Traitement des rapports IFS":
    # Step 1: Upload the JSON (.ifs) file
    uploaded_json_file = st.file_uploader("Charger le fichier IFS de NEO", type="ifs")

    if uploaded_json_file:
        try:
            # Step 2: Load the uploaded JSON file
            json_data = json.load(uploaded_json_file)

            # Step 3: Flatten the JSON data
            flattened_json_data_safe = flatten_json_safe(json_data)

            # Extract profile data
            profile_data = extract_from_flattened(flattened_json_data_safe, FLATTENED_FIELD_MAPPING, list(FLATTENED_FIELD_MAPPING.keys()))

            # Extract checklist data - ALWAYS use direct method first
            checklist_data = []
            
            # Debug: Check if checklist data exists in JSON
            if 'data' in json_data and 'modules' in json_data['data'] and 'food_8' in json_data['data']['modules']:
                if 'checklists' in json_data['data']['modules']['food_8'] and 'checklistFood8' in json_data['data']['modules']['food_8']['checklists']:
                    if 'resultScorings' in json_data['data']['modules']['food_8']['checklists']['checklistFood8']:
                        result_scorings = json_data['data']['modules']['food_8']['checklists']['checklistFood8']['resultScorings']
                        st.info(f"✅ Données de checklist trouvées : {len(result_scorings)} exigences détectées")
                        
                        # Load UUID mapping for better display
                        with st.spinner("Chargement du mapping des exigences..."):
                            uuid_mapping_df = load_uuid_mapping_from_url(UUID_MAPPING_URL)
                        
                        # Create a mapping dict for faster lookup
                        uuid_to_info = {}
                        if not uuid_mapping_df.empty:
                            for _, row in uuid_mapping_df.iterrows():
                                uuid_to_info[row['UUID']] = {
                                    'Num': row['Num'],
                                    'Chapitre': row.get('Chapitre', 'N/A'),
                                    'Theme': row.get('Theme', 'N/A'),
                                    'SSTheme': row.get('SSTheme', 'N/A')
                                }
                            st.success(f"✅ Mapping UUID chargé : {len(uuid_to_info)} correspondances")
                        else:
                            st.warning("⚠️ Mapping UUID non disponible, utilisation des UUID bruts")
                        
                        # Extract data for each UUID found in the JSON
                        for uuid, scoring in result_scorings.items():
                            # Get mapped info if available, otherwise use UUID
                            if uuid in uuid_to_info:
                                mapped_info = uuid_to_info[uuid]
                                display_num = mapped_info['Num']
                                chapitre = mapped_info['Chapitre']
                                theme = mapped_info['Theme']
                                sstheme = mapped_info['SSTheme']
                            else:
                                display_num = uuid  # Use UUID as fallback
                                chapitre = "N/A"
                                theme = "N/A"
                                sstheme = "N/A"
                            
                            # Extract scoring data
                            explanation_text = scoring['answers'].get('englishExplanationText', 'N/A')
                            detailed_explanation = scoring['answers'].get('explanationText', 'N/A')
                            score_label = scoring['score']['label']
                            response = scoring['answers'].get('fieldAnswers', 'N/A')
                            
                            checklist_data.append({
                                "Num": display_num,
                                "UUID": uuid,
                                "Chapitre": chapitre,
                                "Theme": theme,
                                "SSTheme": sstheme,
                                "Explanation": explanation_text,
                                "Detailed Explanation": detailed_explanation,
                                "Score": score_label,
                                "Response": response
                            })
                        
                        st.success(f"✅ Extraction réussie : {len(checklist_data)} exigences extraites")
                    else:
                        st.error("❌ Aucun 'resultScorings' trouvé dans le fichier JSON")
                else:
                    st.error("❌ Structure de checklist non trouvée dans le fichier JSON")
            else:
                st.error("❌ Structure JSON invalide - données manquantes")

            # Extract non-conformities (points not rated A and exclude NA - Non Applicable)
            non_conformities = [item for item in checklist_data if item['Score'] not in ['A', 'N/A', 'NA', 'Non applicable']]

            # Store data in session state for work save/load
            st.session_state['current_profile_data'] = profile_data
            st.session_state['current_checklist_data'] = checklist_data
            st.session_state['current_non_conformities'] = non_conformities

            # Create tabs for Profile, Checklist, and Non-conformities
            tab = st.radio("Sélectionnez un onglet:", ["Profile", "Checklist", "Non-conformities"])

            if tab == "Profile":
                st.subheader("Profile de l'entreprise")
                
                # Create DataFrame for profile with comments column
                profile_list = []
                for field, value in profile_data.items():
                    profile_list.append({
                        "Champ": field,
                        "Valeur": str(value),
                        "Commentaire": ""
                    })
                
                profile_df = pd.DataFrame(profile_list)
                
                # Affichage en deux colonnes : lecture et édition
                col1, col2 = st.columns([3, 2])
                
                with col1:
                    st.write("**📋 Données du profil (lecture seule)**")
                    # Affichage optimisé pour la lecture avec retour à la ligne
                    st.dataframe(
                        profile_df[['Champ', 'Valeur']],
                        column_config={
                            "Champ": st.column_config.TextColumn("Champ", width="medium"),
                            "Valeur": st.column_config.TextColumn("Valeur", width="large")
                        },
                        hide_index=True,
                        use_container_width=True,
                        height=400
                    )
                
                with col2:
                    st.write("**✏️ Zone de commentaires**")
                    # Zone d'édition des commentaires
                    edited_profile_df = st.data_editor(
                        profile_df,
                        column_config={
                            "Champ": st.column_config.TextColumn("Champ", disabled=True, width="small"),
                            "Valeur": st.column_config.TextColumn("Valeur", disabled=True, width="small"),
                            "Commentaire": st.column_config.TextColumn("Commentaire", width="large")
                        },
                        hide_index=True,
                        use_container_width=True,
                        height=400
                    )
                
                # Store edited data in session state
                st.session_state['edited_profile'] = edited_profile_df

            elif tab == "Checklist":
                st.subheader("Checklist des exigences")
                
                # Add filters
                col1, col2, col3 = st.columns(3)
                with col1:
                    score_filter = st.selectbox("Filtrer par score:", ["Tous", "A", "B", "C", "D", "NA"])
                with col2:
                    if not uuid_mapping_df.empty:
                        chapitre_options = ["Tous"] + sorted([str(x) for x in uuid_mapping_df['Chapitre'].dropna().unique()])
                        chapitre_filter = st.selectbox("Filtrer par chapitre:", chapitre_options)
                    else:
                        chapitre_filter = "Tous"
                with col3:
                    show_responses = st.checkbox("Afficher les réponses", value=True)
                
                # Filter data
                filtered_checklist = checklist_data.copy()
                if score_filter != "Tous":
                    filtered_checklist = [item for item in filtered_checklist if item['Score'] == score_filter]
                if chapitre_filter != "Tous":
                    filtered_checklist = [item for item in filtered_checklist if str(item['Chapitre']) == chapitre_filter]
                
                st.info(f"Affichage de {len(filtered_checklist)} exigences sur {len(checklist_data)} au total")
                
                if filtered_checklist:
                    # Create DataFrame for checklist
                    checklist_list = []
                    for item in filtered_checklist:
                        row = {
                            "Num": item['Num'],
                            "Score": item['Score'],
                            "Chapitre": item.get('Chapitre', 'N/A'),
                            "Explication": item['Explanation'],
                            "Explication détaillée": item['Detailed Explanation'],
                            "Commentaire": ""
                        }
                        if show_responses:
                            row["Réponse"] = item['Response']
                        checklist_list.append(row)
                    
                    checklist_df = pd.DataFrame(checklist_list)
                    
                    # Affichage amélioré avec expandeurs pour les longues descriptions
                    st.write("**📋 Exigences avec commentaires éditables**")
                    
                    # Utilisation d'un mode d'affichage hybride
                    display_mode = st.radio("Mode d'affichage:", ["Tableau compact", "Vue détaillée"])
                    
                    if display_mode == "Tableau compact":
                        # Tableau éditable compact
                        column_config = {
                            "Num": st.column_config.TextColumn("N°", disabled=True, width="small"),
                            "Score": st.column_config.TextColumn("Score", disabled=True, width="small"),
                            "Chapitre": st.column_config.TextColumn("Chapitre", disabled=True, width="small"),
                            "Explication": st.column_config.TextColumn("Explication", disabled=True, width="medium"),
                            "Explication détaillée": st.column_config.TextColumn("Explication détaillée", disabled=True, width="medium"),
                            "Commentaire": st.column_config.TextColumn("Commentaire", width="large")
                        }
                        if show_responses:
                            column_config["Réponse"] = st.column_config.TextColumn("Réponse", disabled=True, width="medium")
                        
                        edited_checklist_df = st.data_editor(
                            checklist_df,
                            column_config=column_config,
                            hide_index=True,
                            use_container_width=True,
                            height=600
                        )
                        st.session_state['edited_checklist'] = edited_checklist_df
                        
                    else:  # Vue détaillée
                        # Affichage avec expandeurs pour chaque exigence
                        comments_dict = {}
                        
                        for i, item in enumerate(filtered_checklist):
                            score_emoji = {"A": "✅", "B": "⚠️", "C": "🟠", "D": "🔴", "NA": "⚫"}.get(item['Score'], "❓")
                            
                            with st.expander(f"{score_emoji} Exigence {item['Num']} - Score: {item['Score']}", expanded=False):
                                col_detail1, col_detail2 = st.columns([2, 1])
                                
                                with col_detail1:
                                    st.write(f"**📖 Explication:** {item['Explanation']}")
                                    st.write(f"**📋 Explication détaillée:** {item['Detailed Explanation']}")
                                    if show_responses:
                                        st.write(f"**💬 Réponse:** {item['Response']}")
                                
                                with col_detail2:
                                    st.write(f"**N°:** {item['Num']}")
                                    st.write(f"**Score:** {item['Score']}")
                                    st.write(f"**Chapitre:** {item.get('Chapitre', 'N/A')}")
                                
                                # Zone de commentaire pour chaque exigence
                                comment_key = f"comment_detail_{item['Num']}"
                                comment = st.text_area(
                                    "💬 Votre commentaire:",
                                    key=comment_key,
                                    height=100,
                                    placeholder="Ajoutez vos observations pour cette exigence..."
                                )
                                comments_dict[item['Num']] = comment
                        
                        # Reconstituer le DataFrame avec les commentaires
                        for i, row in enumerate(checklist_df.itertuples()):
                            num = row.Num
                            if num in comments_dict:
                                checklist_df.at[i, 'Commentaire'] = comments_dict[num]
                        
                        st.session_state['edited_checklist'] = checklist_df
                else:
                    st.warning("Aucune exigence trouvée avec ces filtres.")

            elif tab == "Non-conformities":
                st.subheader("Non-conformités")
                
                if non_conformities:
                    st.warning(f"**{len(non_conformities)} non-conformité(s) détectée(s)** (scores B, C, D uniquement)")
                    
                    # Affichage avec expandeurs pour les non-conformités
                    nc_comments = {}
                    nc_actions = {}
                    
                    for item in non_conformities:
                        score_emoji = {"B": "⚠️", "C": "🟠", "D": "🔴"}.get(item['Score'], "❓")
                        score_color = {"B": "orange", "C": "orange", "D": "red"}.get(item['Score'], "gray")
                        
                        with st.expander(f"{score_emoji} **NC {item['Num']}** - Score: {item['Score']}", expanded=True):
                            # Informations détaillées
                            col_nc1, col_nc2 = st.columns([3, 1])
                            
                            with col_nc1:
                                st.markdown(f"**📖 Explication:** {item['Explanation']}")
                                st.markdown(f"**📋 Explication détaillée:** {item['Detailed Explanation']}")
                                st.markdown(f"**💬 Réponse:** {item['Response']}")
                            
                            with col_nc2:
                                st.markdown(f"**N°:** {item['Num']}")
                                st.markdown(f"**Score:** :color[{score_color}][{item['Score']}]")
                                st.markdown(f"**Chapitre:** {item.get('Chapitre', 'N/A')}")
                            
                            # Zone de commentaires et plan d'action
                            col_action1, col_action2 = st.columns(2)
                            
                            with col_action1:
                                comment_key = f"nc_comment_{item['Num']}"
                                comment = st.text_area(
                                    "💬 Commentaire de l'auditeur:",
                                    key=comment_key,
                                    height=120,
                                    placeholder="Observations, causes identifiées, éléments de preuve..."
                                )
                                nc_comments[item['Num']] = comment
                            
                            with col_action2:
                                action_key = f"nc_action_{item['Num']}"
                                action = st.text_area(
                                    "🎯 Plan d'action corrective:",
                                    key=action_key,
                                    height=120,
                                    placeholder="Actions à mettre en place, responsable, échéance..."
                                )
                                nc_actions[item['Num']] = action
                    
                    # Reconstituer le DataFrame avec commentaires et actions
                    nc_list = []
                    for item in non_conformities:
                        nc_list.append({
                            "Num": item['Num'],
                            "Score": item['Score'],
                            "Chapitre": item.get('Chapitre', 'N/A'),
                            "Explication": item['Explanation'],
                            "Explication détaillée": item['Detailed Explanation'],
                            "Réponse": item['Response'],
                            "Commentaire": nc_comments.get(item['Num'], ""),
                            "Plan d'action": nc_actions.get(item['Num'], "")
                        })
                    
                    nc_df = pd.DataFrame(nc_list)
                    st.session_state['edited_nc'] = nc_df
                    
                else:
                    st.success("🎉 Aucune non-conformité détectée ! Toutes les exigences sont conformes (A) ou non applicables (NA).")

            # Work save and export options
            st.markdown("---")
            col1, col2 = st.columns(2)
            
            with col1:
                if st.button("💾 Sauvegarder le travail en cours", type="secondary"):
                    numero_coid = profile_data.get("N° COID du portail", "inconnu")
                    work_file = save_work_to_excel(
                        profile_data, 
                        checklist_data, 
                        non_conformities,
                        st.session_state.get('edited_profile'),
                        st.session_state.get('edited_checklist'),
                        st.session_state.get('edited_nc'),
                        numero_coid
                    )
                    
                    from datetime import datetime
                    date_str = datetime.now().strftime("%Y%m%d_%H%M")
                    work_filename = f"travail_IFS_{numero_coid}_{date_str}.xlsx"
                    
                    st.download_button(
                        label="📥 Télécharger la sauvegarde de travail",
                        data=work_file,
                        file_name=work_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        help="Sauvegarde tous vos commentaires pour reprendre plus tard"
                    )
                    st.success("💾 Sauvegarde de travail générée ! Vous pourrez reprendre là où vous vous êtes arrêté.")
            
            with col2:
                if st.button("📊 Exporter le rapport final", type="primary"):
                    numero_coid = profile_data.get("N° COID du portail", "inconnu")
                    
                    # Create the Excel file with column formatting
                    output = BytesIO()
                    
                    # Create Excel writer and adjust column widths
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        # Profile tab
                        if 'edited_profile' in st.session_state:
                            profile_export_df = st.session_state['edited_profile']
                        else:
                            profile_export_df = pd.DataFrame([
                                {"Champ": k, "Valeur": v, "Commentaire": ""} 
                                for k, v in profile_data.items()
                            ])
                        
                        profile_export_df.to_excel(writer, index=False, sheet_name="Profile")

                        # Checklist tab
                        if 'edited_checklist' in st.session_state:
                            checklist_export_df = st.session_state['edited_checklist']
                        else:
                            checklist_export_df = pd.DataFrame(checklist_data)
                            checklist_export_df['Commentaire'] = ''
                        
                        checklist_export_df.to_excel(writer, index=False, sheet_name="Checklist")

                        # Non-conformities tab
                        if non_conformities:
                            if 'edited_nc' in st.session_state:
                                nc_export_df = st.session_state['edited_nc']
                            else:
                                nc_export_df = pd.DataFrame(non_conformities)
                                nc_export_df['Commentaire'] = ''
                                nc_export_df['Plan d\'action'] = ''
                            
                            nc_export_df.to_excel(writer, index=False, sheet_name="Non-conformities")

                        # Adjust column widths for all sheets
                        for sheet_name in writer.sheets:
                            worksheet = writer.sheets[sheet_name]
                            for col in worksheet.columns:
                                max_length = 0
                                column = col[0].column_letter # Get the column name
                                for cell in col:
                                    try: # Necessary to avoid error on empty cells
                                        if len(str(cell.value)) > max_length:
                                            max_length = len(str(cell.value))
                                    except:
                                        pass
                                adjusted_width = min((max_length + 2) * 1.2, 50)  # Limit max width
                                worksheet.column_dimensions[column].width = adjusted_width

                    # Reset the position of the output to the start
                    output.seek(0)

                    # Provide the download button with the COID number in the filename
                    st.download_button(
                        label="📥 Télécharger le rapport final",
                        data=output,
                        file_name=f'rapport_final_IFS_{numero_coid}.xlsx',
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

        except json.JSONDecodeError:
            st.error("Erreur lors du décodage du fichier JSON. Veuillez vous assurer qu'il est au format correct.")
        except Exception as e:
            st.error(f"Erreur lors du traitement du fichier: {str(e)}")
    else:
        st.info("Veuillez charger un fichier IFS de NEO pour commencer.")

elif main_option == "Reprendre un travail sauvegardé":
    st.subheader("📂 Reprendre un travail sauvegardé")
    st.info("Chargez un fichier Excel de travail sauvegardé pour reprendre là où vous vous êtes arrêté.")
    
    uploaded_work_file = st.file_uploader(
        "Charger un fichier de travail Excel", 
        type="xlsx",
        help="Sélectionnez un fichier Excel généré avec 'Sauvegarder le travail en cours'"
    )
    
    if uploaded_work_file:
        work_data, message = load_work_from_excel(uploaded_work_file)
        
        if work_data is None:
            st.error(message)
        else:
            st.success(message)
            
            # Create tabs for the loaded work
            tab_work = st.radio("Sélectionnez un onglet:", ["Profile sauvegardé", "Checklist sauvegardée", "Non-conformités sauvegardées"])
            
            if tab_work == "Profile sauvegardé" and 'profile' in work_data:
                st.subheader("Profile de l'entreprise (travail repris)")
                
                # Display and allow editing of the saved work
                edited_profile_work = st.data_editor(
                    work_data['profile'],
                    column_config={
                        "Champ": st.column_config.TextColumn("Champ", disabled=True),
                        "Valeur": st.column_config.TextColumn("Valeur", disabled=True),
                        "Commentaire": st.column_config.TextColumn("Commentaire", width="medium")
                    },
                    hide_index=True,
                    use_container_width=True
                )
                
                st.session_state['edited_profile_work'] = edited_profile_work
            
            elif tab_work == "Checklist sauvegardée" and 'checklist' in work_data:
                st.subheader("Checklist des exigences (travail repris)")
                
                # Display and allow editing of the saved work
                edited_checklist_work = st.data_editor(
                    work_data['checklist'],
                    hide_index=True,
                    use_container_width=True,
                    height=600
                )
                
                st.session_state['edited_checklist_work'] = edited_checklist_work
            
            elif tab_work == "Non-conformités sauvegardées" and 'nc' in work_data:
                st.subheader("Non-conformités (travail repris)")
                
                # Display and allow editing of the saved work
                edited_nc_work = st.data_editor(
                    work_data['nc'],
                    hide_index=True,
                    use_container_width=True,
                    height=600
                )
                
                st.session_state['edited_nc_work'] = edited_nc_work
            
            # Export updated work
            st.markdown("---")
            if st.button("💾 Sauvegarder les modifications", type="primary"):
                output = BytesIO()
                
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    # Metadata
                    metadata = pd.DataFrame({
                        "Type": ["IFS_WORK_SAVE"],
                        "COID": ["travail_repris"],
                        "Date": [pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")]
                    })
                    metadata.to_excel(writer, index=False, sheet_name="METADATA")
                    
                    # Save updated work
                    if 'edited_profile_work' in st.session_state:
                        st.session_state['edited_profile_work'].to_excel(writer, index=False, sheet_name="Profile_Work")
                    elif 'profile' in work_data:
                        work_data['profile'].to_excel(writer, index=False, sheet_name="Profile_Work")
                    
                    if 'edited_checklist_work' in st.session_state:
                        st.session_state['edited_checklist_work'].to_excel(writer, index=False, sheet_name="Checklist_Work")
                    elif 'checklist' in work_data:
                        work_data['checklist'].to_excel(writer, index=False, sheet_name="Checklist_Work")
                    
                    if 'edited_nc_work' in st.session_state:
                        st.session_state['edited_nc_work'].to_excel(writer, index=False, sheet_name="NonConformities_Work")
                    elif 'nc' in work_data:
                        work_data['nc'].to_excel(writer, index=False, sheet_name="NonConformities_Work")
                
                output.seek(0)
                
                from datetime import datetime
                date_str = datetime.now().strftime("%Y%m%d_%H%M")
                
                st.download_button(
                    label="📥 Télécharger la nouvelle sauvegarde",
                    data=output,
                    file_name=f"travail_IFS_mis_a_jour_{date_str}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    else:
        st.markdown("""
        ### 📋 Comment utiliser cette fonctionnalité :
        
        1. **Travaillez sur un audit IFS** dans l'onglet "Traitement des rapports IFS"
        2. **Ajoutez vos commentaires** dans les tableaux
        3. **Cliquez sur "💾 Sauvegarder le travail en cours"** pour créer un fichier Excel de sauvegarde
        4. **Téléchargez le fichier** de sauvegarde
        5. **Plus tard, revenez ici** et chargez votre fichier de sauvegarde
        6. **Continuez votre travail** là où vous vous étiez arrêté !
        
        ✅ **Avantages :**
        - Travail sauvegardé avec tous vos commentaires
        - Possibilité de reprendre l'audit plus tard
        - Collaboration possible (partage du fichier de travail)
        - Pas de perte de données
        """)

elif main_option == "Gestion des fichiers Excel":
    st.subheader("Gestion des fichiers Excel")
    uploaded_excel_file = st.file_uploader("Charger le fichier Excel pour compléter les commentaires", type="xlsx")

    if uploaded_excel_file:
        try:
            # Load the uploaded Excel file
            excel_data = pd.read_excel(uploaded_excel_file, sheet_name=None)

            # Display the Excel content for editing
            st.subheader("Contenu du fichier Excel")
            sheet = st.selectbox("Sélectionnez une feuille:", list(excel_data.keys()))
            df = excel_data[sheet]

            # Display the DataFrame
            edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True)

            # Save and provide a download link for the edited Excel file
            if st.button("Enregistrer les modifications"):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    # Update the specific sheet with edited data
                    excel_data[sheet] = edited_df
                    for sheet_name, df_sheet in excel_data.items():
                        df_sheet.to_excel(writer, index=False, sheet_name=sheet_name)
                output.seek(0)
                st.download_button(
                    label="Télécharger le fichier Excel modifié",
                    data=output,
                    file_name="modified_excel.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"Erreur lors du traitement du fichier Excel: {e}")
    else:
        st.info("Veuillez charger un fichier Excel pour commencer.")
