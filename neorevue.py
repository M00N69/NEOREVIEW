import json
import pandas as pd
import streamlit as st
from io import BytesIO
import requests
from datetime import datetime
import re

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
    """Save current work with comments to Excel for later resume and reviewer/auditor communication"""
    output = BytesIO()
    
    # Extract company name for filename
    company_name = profile_data.get("Nom du site à auditer", "entreprise_inconnue")
    # Clean company name for filename (remove special characters)
    clean_company_name = re.sub(r'[^\w\s-]', '', company_name).strip()
    clean_company_name = re.sub(r'[-\s]+', '_', clean_company_name)
    
    # Use xlsxwriter for better formatting in communication files
    with pd.ExcelWriter(output, engine='xlsxwriter', options={'strings_to_urls': False}) as writer:
        workbook = writer.book
        
        # Define professional formats for reviewer/auditor communication
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#4472C4',  # Professional blue
            'font_color': 'white',
            'border': 1,
            'font_size': 11
        })
        
        cell_format = workbook.add_format({
            'text_wrap': True,
            'valign': 'top',
            'border': 1,
            'font_size': 10
        })
        
        reviewer_format = workbook.add_format({
            'text_wrap': True,
            'valign': 'top',
            'border': 1,
            'fg_color': '#E7F3FF',  # Light blue for reviewer comments
            'font_size': 10
        })
        
        auditor_format = workbook.add_format({
            'text_wrap': True,
            'valign': 'top',
            'border': 1,
            'fg_color': '#F0F8E7',  # Light green for auditor responses
            'font_size': 10
        })
        
        # Metadata sheet to identify this as a work file
        metadata = pd.DataFrame({
            "Type": ["IFS_WORK_SAVE"],
            "COID": [coid],
            "Company_Name": [company_name],
            "Date": [pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")],
            "Purpose": ["Communication Reviewer/Auditeur"]
        })
        metadata.to_excel(writer, index=False, sheet_name="METADATA")
        
        # Profile with enhanced communication structure
        if edited_profile is not None:
            profile_save_df = edited_profile.copy()
        else:
            profile_save_df = pd.DataFrame([
                {"Champ": k, "Valeur": v, "Commentaire du reviewer": "", "Réponse de l'auditeur": ""} 
                for k, v in profile_data.items()
            ])
        
        # Ensure communication columns exist
        if "Commentaire du reviewer" not in profile_save_df.columns:
            profile_save_df["Commentaire du reviewer"] = profile_save_df.get("Commentaire", "")
        if "Réponse de l'auditeur" not in profile_save_df.columns:
            profile_save_df["Réponse de l'auditeur"] = ""
        
        profile_save_df.to_excel(writer, index=False, sheet_name="Profile_Communication")
        
        # Format Profile sheet
        profile_ws = writer.sheets["Profile_Communication"]
        profile_ws.set_column('A:A', 25)  # Champ
        profile_ws.set_column('B:B', 30)  # Valeur
        profile_ws.set_column('C:C', 40)  # Commentaire reviewer
        profile_ws.set_column('D:D', 40)  # Réponse auditeur
        
        # Apply formatting
        for col_num, value in enumerate(profile_save_df.columns.values):
            profile_ws.write(0, col_num, value, header_format)
        
        # Checklist with enhanced communication structure
        if edited_checklist is not None:
            checklist_save_df = edited_checklist.copy()
        else:
            checklist_save_df = pd.DataFrame(checklist_data)
            checklist_save_df['Commentaire du reviewer'] = ''
            checklist_save_df['Réponse de l\'auditeur'] = ''
        
        # Ensure communication columns exist
        if "Commentaire du reviewer" not in checklist_save_df.columns:
            checklist_save_df["Commentaire du reviewer"] = checklist_save_df.get("Commentaire", "")
        if "Réponse de l'auditeur" not in checklist_save_df.columns:
            checklist_save_df["Réponse de l'auditeur"] = ""
        
        checklist_save_df.to_excel(writer, index=False, sheet_name="Checklist_Communication")
        
        # Format Checklist sheet
        checklist_ws = writer.sheets["Checklist_Communication"]
        checklist_ws.set_column('A:A', 12)  # Num
        checklist_ws.set_column('B:B', 8)   # Score
        checklist_ws.set_column('C:C', 15)  # Chapitre
        checklist_ws.set_column('D:D', 50)  # Explication
        checklist_ws.set_column('E:E', 50)  # Explication détaillée
        checklist_ws.set_column('F:F', 40)  # Commentaire reviewer
        checklist_ws.set_column('G:G', 40)  # Réponse auditeur
        if 'Réponse' in checklist_save_df.columns:
            checklist_ws.set_column('H:H', 30)  # Réponse système
        
        # Apply formatting
        for col_num, value in enumerate(checklist_save_df.columns.values):
            checklist_ws.write(0, col_num, value, header_format)

        # Non-conformities with enhanced communication and action tracking
        if non_conformities:
            if edited_nc is not None:
                nc_save_df = edited_nc.copy()
            else:
                nc_save_df = pd.DataFrame(non_conformities)
                nc_save_df['Commentaire du reviewer'] = ''
                nc_save_df['Plan d\'action proposé'] = ''
                nc_save_df['Réponse de l\'auditeur'] = ''
                nc_save_df['Actions mises en place'] = ''
                nc_save_df['Date limite'] = ''
                nc_save_df['Responsable'] = ''
                nc_save_df['Statut'] = 'En attente'
            
            # Ensure all communication and tracking columns exist
            communication_columns = [
                "Commentaire du reviewer", "Plan d'action proposé", "Réponse de l'auditeur", 
                "Actions mises en place", "Date limite", "Responsable", "Statut"
            ]
            
            for col in communication_columns:
                if col not in nc_save_df.columns:
                    if col == "Statut":
                        nc_save_df[col] = "En attente"
                    else:
                        nc_save_df[col] = ""
            
            nc_save_df.to_excel(writer, index=False, sheet_name="NonConformities_ActionPlan")
            
            # Format Non-conformities sheet
            nc_ws = writer.sheets["NonConformities_ActionPlan"]
            nc_ws.set_column('A:A', 12)  # Num
            nc_ws.set_column('B:B', 8)   # Score
            nc_ws.set_column('C:C', 15)  # Chapitre
            nc_ws.set_column('D:D', 40)  # Explication
            nc_ws.set_column('E:E', 40)  # Explication détaillée
            nc_ws.set_column('F:F', 30)  # Réponse système
            nc_ws.set_column('G:G', 35)  # Commentaire reviewer
            nc_ws.set_column('H:H', 35)  # Plan d'action proposé
            nc_ws.set_column('I:I', 35)  # Réponse auditeur
            nc_ws.set_column('J:J', 35)  # Actions mises en place
            nc_ws.set_column('K:K', 15)  # Date limite
            nc_ws.set_column('L:L', 20)  # Responsable
            nc_ws.set_column('M:M', 15)  # Statut
            
            # Apply formatting with color coding
            for col_num, value in enumerate(nc_save_df.columns.values):
                nc_ws.write(0, col_num, value, header_format)

    output.seek(0)
    return output, clean_company_name

def clean_dataframe_for_editor(df, required_columns):
    """Clean and prepare DataFrame for st.data_editor"""
    if df is None or df.empty:
        return pd.DataFrame(columns=required_columns)
    
    # Ensure all required columns exist
    for col in required_columns:
        if col not in df.columns:
            df[col] = ''
    
    # Clean data types and handle NaN values
    for col in df.columns:
        if col in required_columns:
            # Convert to string and replace NaN with empty strings
            df[col] = df[col].astype(str).fillna('').replace('nan', '').replace('None', '')
    
    # Reset index to avoid issues
    df = df.reset_index(drop=True)
    
    return df[required_columns]  # Return only required columns in correct order
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
        company_name = metadata.iloc[0].get("Company_Name", "Entreprise inconnue") if "Company_Name" in metadata.columns else "Entreprise inconnue"
        save_date = metadata.iloc[0]["Date"]
        
        return work_data, f"Travail chargé avec succès\n**COID:** {coid} | **Entreprise:** {company_name}\n**Sauvegardé le:** {save_date}"
        
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
                
                # Create DataFrame for profile with reviewer/auditor communication columns
                profile_list = []
                for field, value in profile_data.items():
                    profile_list.append({
                        "Champ": field,
                        "Valeur": str(value),
                        "Commentaire du reviewer": "",
                        "Réponse de l'auditeur": ""
                    })
                
                profile_df = pd.DataFrame(profile_list)
                
                # Tableau éditable avec commentaires reviewer/auditeur
                st.write("**✏️ Profil de l'entreprise - Communication reviewer/auditeur**")
                st.info("💡 **Workflow:** Reviewer ajoute commentaires → Auditeur répond → Communication Excel")
                
                edited_profile_df = st.data_editor(
                    profile_df,
                    column_config={
                        "Champ": st.column_config.TextColumn("Champ", disabled=True, width="medium"),
                        "Valeur": st.column_config.TextColumn("Valeur", disabled=True, width="medium"),
                        "Commentaire du reviewer": st.column_config.TextColumn("📝 Commentaire du reviewer", width="large"),
                        "Réponse de l'auditeur": st.column_config.TextColumn("💬 Réponse de l'auditeur", width="large")
                    },
                    hide_index=True,
                    use_container_width=True,
                    height=500
                )
                
                # Store edited data in session state
                st.session_state['edited_profile'] = edited_profile_df

            elif tab == "Checklist":
                st.subheader("Checklist des exigences")
                
                # Add filters
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    score_filter = st.selectbox("Filtrer par score:", ["Tous", "A", "B", "C", "D", "NA"])
                with col2:
                    if not uuid_mapping_df.empty:
                        chapitre_options = ["Tous"] + sorted([str(x) for x in uuid_mapping_df['Chapitre'].dropna().unique()])
                        chapitre_filter = st.selectbox("Filtrer par chapitre:", chapitre_options)
                    else:
                        chapitre_filter = "Tous"
                with col3:
                    content_filter = st.selectbox(
                        "Filtrer le contenu:", 
                        ["Explications non vides", "Tous", "Explications vides"],
                        index=0  # Par défaut sur "Explications non vides"
                    )
                with col4:
                    show_responses = st.checkbox("Afficher les réponses", value=True)
                
                # Filter data
                filtered_checklist = checklist_data.copy()
                
                # Filtre par score
                if score_filter != "Tous":
                    filtered_checklist = [item for item in filtered_checklist if item['Score'] == score_filter]
                
                # Filtre par chapitre
                if chapitre_filter != "Tous":
                    filtered_checklist = [item for item in filtered_checklist if str(item['Chapitre']) == chapitre_filter]
                
                # Filtre par contenu des explications
                if content_filter == "Explications non vides":
                    filtered_checklist = [item for item in filtered_checklist 
                                        if (item['Explanation'] != 'N/A' and item['Explanation'].strip() != '') 
                                        or (item['Detailed Explanation'] != 'N/A' and item['Detailed Explanation'].strip() != '')]
                elif content_filter == "Explications vides":
                    filtered_checklist = [item for item in filtered_checklist 
                                        if (item['Explanation'] == 'N/A' or item['Explanation'].strip() == '') 
                                        and (item['Detailed Explanation'] == 'N/A' or item['Detailed Explanation'].strip() == '')]
                
                st.info(f"Affichage de {len(filtered_checklist)} exigences sur {len(checklist_data)} au total")
                
                if filtered_checklist:
                    # Vue détaillée par défaut avec expandeurs ouverts
                    st.write("**📋 Vue détaillée des exigences - Communication reviewer/auditeur**")
                    st.write("💡 *Workflow: Reviewer commente → Auditeur répond → Communication via Excel*")
                    
                    # Affichage avec expandeurs pour chaque exigence (OUVERTS par défaut)
                    comments_reviewer_dict = {}
                    comments_auditor_dict = {}
                    
                    for i, item in enumerate(filtered_checklist):
                        score_emoji = {"A": "✅", "B": "⚠️", "C": "🟠", "D": "🔴", "NA": "⚫"}.get(item['Score'], "❓")
                        
                        # TOUS LES EXPANDEURS OUVERTS PAR DÉFAUT (expanded=True)
                        with st.expander(f"{score_emoji} Exigence {item['Num']} - Score: {item['Score']}", expanded=True):
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
                            
                            # Communication reviewer/auditeur
                            st.markdown("---")
                            col_comm1, col_comm2 = st.columns(2)
                            
                            with col_comm1:
                                # Zone de commentaire reviewer
                                comment_reviewer_key = f"comment_reviewer_{item['Num']}"
                                comment_reviewer = st.text_area(
                                    "📝 Commentaire du reviewer:",
                                    key=comment_reviewer_key,
                                    height=80,
                                    placeholder="Observations du reviewer pour cette exigence..."
                                )
                                comments_reviewer_dict[item['Num']] = comment_reviewer
                            
                            with col_comm2:
                                # Zone de réponse auditeur
                                comment_auditor_key = f"comment_auditor_{item['Num']}"
                                comment_auditor = st.text_area(
                                    "💬 Réponse de l'auditeur:",
                                    key=comment_auditor_key,
                                    height=80,
                                    placeholder="Réponse de l'auditeur..."
                                )
                                comments_auditor_dict[item['Num']] = comment_auditor
                    
                    # Reconstituer le DataFrame avec les commentaires de communication
                    checklist_list = []
                    for item in filtered_checklist:
                        row = {
                            "Num": item['Num'],
                            "Score": item['Score'],
                            "Chapitre": item.get('Chapitre', 'N/A'),
                            "Explication": item['Explanation'],
                            "Explication détaillée": item['Detailed Explanation'],
                            "Commentaire du reviewer": comments_reviewer_dict.get(item['Num'], ""),
                            "Réponse de l'auditeur": comments_auditor_dict.get(item['Num'], "")
                        }
                        if show_responses:
                            row["Réponse"] = item['Response']
                        checklist_list.append(row)
                    
                    checklist_df = pd.DataFrame(checklist_list)
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
                st.write("**💾 Sauvegarder le travail en cours**")
                st.write("*Permet de reprendre plus tard*")
                if st.button("💾 Générer la sauvegarde", type="secondary", key="save_work_btn"):
                    numero_coid = profile_data.get("N° COID du portail", "inconnu")
                    
                    with st.spinner("Génération de la sauvegarde..."):
                        work_file, clean_company_name = save_work_to_excel(
                            profile_data, 
                            checklist_data, 
                            non_conformities,
                            st.session_state.get('edited_profile'),
                            st.session_state.get('edited_checklist'),
                            st.session_state.get('edited_nc'),
                            numero_coid
                        )
                    
                    date_str = datetime.now().strftime("%Y%m%d_%H%M")
                    work_filename = f"travail_IFS_{numero_coid}_{clean_company_name}_{date_str}.xlsx"
                    
                    st.success("💾 Sauvegarde générée avec succès !")
                    st.download_button(
                        label="📥 Télécharger la sauvegarde",
                        data=work_file,
                        file_name=work_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        help="Fichier Excel contenant tous vos commentaires pour reprendre plus tard",
                        key="download_work_btn"
                    )
            
            with col2:
                st.write("**📊 Exporter le rapport final**")
                st.write("*Rapport final pour présentation*")
                if st.button("📊 Générer le rapport", type="primary", key="export_final_btn"):
                    numero_coid = profile_data.get("N° COID du portail", "inconnu")
                    
                    with st.spinner("Génération du rapport final..."):
                        # Create the Excel file with xlsxwriter for better performance and formatting
                        output = BytesIO()
                        
                        # Use xlsxwriter for final export (better performance and formatting)
                        with pd.ExcelWriter(output, engine='xlsxwriter', options={'strings_to_urls': False}) as writer:
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

                            # Enhanced formatting with xlsxwriter
                            workbook = writer.book
                            
                            # Define formats
                            header_format = workbook.add_format({
                                'bold': True,
                                'text_wrap': True,
                                'valign': 'top',
                                'fg_color': '#D7E4BC',
                                'border': 1
                            })
                            
                            cell_format = workbook.add_format({
                                'text_wrap': True,
                                'valign': 'top',
                                'border': 1
                            })
                            
                            # Apply formatting and adjust column widths for all sheets
                            for sheet_name in writer.sheets:
                                worksheet = writer.sheets[sheet_name]
                                
                                # Get the dataframe for this sheet to know the number of columns
                                if sheet_name == "Profile":
                                    df = profile_export_df
                                elif sheet_name == "Checklist":
                                    df = checklist_export_df
                                elif sheet_name == "Non-conformities" and non_conformities:
                                    df = nc_export_df
                                else:
                                    continue
                                
                                # Format headers
                                for col_num, value in enumerate(df.columns.values):
                                    worksheet.write(0, col_num, value, header_format)
                                
                                # Set column widths and wrap text
                                for col_num, col_name in enumerate(df.columns):
                                    # Calculate optimal width
                                    max_len = max(
                                        df[col_name].astype(str).map(len).max(),  # Max length in column
                                        len(str(col_name))  # Header length
                                    )
                                    
                                    # Set width with limits
                                    if col_name in ['Commentaire', 'Plan d\'action', 'Explication', 'Explication détaillée', 'Réponse']:
                                        worksheet.set_column(col_num, col_num, min(50, max(20, max_len)))
                                    else:
                                        worksheet.set_column(col_num, col_num, min(30, max(10, max_len)))
                                
                                # Apply cell formatting to data rows
                                for row_num in range(1, len(df) + 1):
                                    for col_num in range(len(df.columns)):
                                        worksheet.write(row_num, col_num, df.iloc[row_num-1, col_num], cell_format)

                        # Reset the position of the output to the start
                        output.seek(0)
                    
                    st.success("📊 Rapport final généré avec succès !")
                    
                    # Create filename with company name
                    company_name = profile_data.get("Nom du site à auditer", "entreprise_inconnue")
                    clean_company_name = re.sub(r'[^\w\s-]', '', company_name).strip()
                    clean_company_name = re.sub(r'[-\s]+', '_', clean_company_name)
                    final_filename = f'rapport_final_IFS_{numero_coid}_{clean_company_name}.xlsx'
                    
                    # Provide the download button with the COID number in the filename
                    st.download_button(
                        label="📥 Télécharger le rapport final",
                        data=output,
                        file_name=final_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_final_btn"
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
            
            # Store work data in session state like normal mode
            st.session_state['work_mode'] = True
            st.session_state['loaded_work_data'] = work_data
            
            # Create tabs like in normal mode - SAME DESIGN
            tab_work = st.radio("Sélectionnez un onglet:", ["Profile", "Checklist", "Non-conformities"])
            
            if tab_work == "Profile" and 'profile' in work_data:
                st.subheader("Profile de l'entreprise (travail repris)")
                
                # Clean and prepare DataFrame for st.data_editor
                required_profile_columns = ["Champ", "Valeur", "Commentaire du reviewer", "Réponse de l'auditeur"]
                work_profile_df = clean_dataframe_for_editor(work_data['profile'], required_profile_columns)
                
                # Handle legacy format (convert old "Commentaire" to "Commentaire du reviewer")
                if "Commentaire" in work_data['profile'].columns and "Commentaire du reviewer" not in work_data['profile'].columns:
                    work_profile_df["Commentaire du reviewer"] = work_data['profile']['Commentaire']
                
                # SAME DESIGN as normal Profile tab
                st.write("**✏️ Profil de l'entreprise - Communication reviewer/auditeur (travail repris)**")
                st.info("💡 **Workflow repris:** Modifiez les commentaires → Continuez la communication")
                
                edited_profile_work = st.data_editor(
                    work_profile_df,
                    column_config={
                        "Champ": st.column_config.TextColumn("Champ", disabled=True, width="medium"),
                        "Valeur": st.column_config.TextColumn("Valeur", disabled=True, width="medium"),
                        "Commentaire du reviewer": st.column_config.TextColumn("📝 Commentaire du reviewer", width="large"),
                        "Réponse de l'auditeur": st.column_config.TextColumn("💬 Réponse de l'auditeur", width="large")
                    },
                    hide_index=True,
                    use_container_width=True,
                    height=500,
                    key="work_profile_editor"
                )
                
                st.session_state['edited_profile_work'] = edited_profile_work
            
            elif tab_work == "Checklist" and 'checklist' in work_data:
                st.subheader("Checklist des exigences (travail repris)")
                
                # Ensure consistent DataFrame structure
                work_checklist_df = work_data['checklist'].copy()
                
                # Verify column structure
                if 'Commentaire' not in work_checklist_df.columns:
                    work_checklist_df['Commentaire'] = ''
                
                # SAME FILTERS as normal Checklist tab
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    score_filter_work = st.selectbox("Filtrer par score:", ["Tous", "A", "B", "C", "D", "NA"], key="work_score_filter")
                with col2:
                    if 'Chapitre' in work_checklist_df.columns:
                        chapitre_options_work = ["Tous"] + sorted([str(x) for x in work_checklist_df['Chapitre'].dropna().unique()])
                        chapitre_filter_work = st.selectbox("Filtrer par chapitre:", chapitre_options_work, key="work_chapitre_filter")
                    else:
                        chapitre_filter_work = "Tous"
                with col3:
                    content_filter_work = st.selectbox(
                        "Filtrer le contenu:", 
                        ["Explications non vides", "Tous", "Explications vides"],
                        index=0,
                        key="work_content_filter"
                    )
                with col4:
                    show_responses_work = st.checkbox("Afficher les réponses", value=True, key="work_show_responses")
                
                # Apply filters
                filtered_work_checklist = work_checklist_df.copy()
                
                if score_filter_work != "Tous":
                    filtered_work_checklist = filtered_work_checklist[filtered_work_checklist['Score'] == score_filter_work]
                
                if chapitre_filter_work != "Tous" and 'Chapitre' in filtered_work_checklist.columns:
                    filtered_work_checklist = filtered_work_checklist[filtered_work_checklist['Chapitre'].astype(str) == chapitre_filter_work]
                
                if content_filter_work == "Explications non vides":
                    filtered_work_checklist = filtered_work_checklist[
                        ((filtered_work_checklist['Explication'] != 'N/A') & (filtered_work_checklist['Explication'].str.strip() != '')) |
                        ((filtered_work_checklist['Explication détaillée'] != 'N/A') & (filtered_work_checklist['Explication détaillée'].str.strip() != ''))
                    ]
                elif content_filter_work == "Explications vides":
                    filtered_work_checklist = filtered_work_checklist[
                        ((filtered_work_checklist['Explication'] == 'N/A') | (filtered_work_checklist['Explication'].str.strip() == '')) &
                        ((filtered_work_checklist['Explication détaillée'] == 'N/A') | (filtered_work_checklist['Explication détaillée'].str.strip() == ''))
                    ]
                
                st.info(f"Affichage de {len(filtered_work_checklist)} exigences sur {len(work_checklist_df)} au total")
                
                if len(filtered_work_checklist) > 0:
                    # SAME DETAILED VIEW as normal Checklist tab
                    st.write("**📋 Vue détaillée des exigences avec commentaires**")
                    st.write("💡 *Conseil : Modifiez vos commentaires puis refermez les expandeurs au fur et à mesure*")
                    
                    # Expandeurs avec commentaires
                    comments_work_dict = {}
                    
                    for i, row in filtered_work_checklist.iterrows():
                        score_emoji = {"A": "✅", "B": "⚠️", "C": "🟠", "D": "🔴", "NA": "⚫"}.get(row['Score'], "❓")
                        
                        # Expandeurs OUVERTS par défaut pour travail repris aussi
                        with st.expander(f"{score_emoji} Exigence {row['Num']} - Score: {row['Score']}", expanded=True):
                            col_detail1, col_detail2 = st.columns([2, 1])
                            
                            with col_detail1:
                                st.write(f"**📖 Explication:** {row['Explication']}")
                                st.write(f"**📋 Explication détaillée:** {row['Explication détaillée']}")
                                if show_responses_work and 'Réponse' in row:
                                    st.write(f"**💬 Réponse:** {row['Réponse']}")
                            
                            with col_detail2:
                                st.write(f"**N°:** {row['Num']}")
                                st.write(f"**Score:** {row['Score']}")
                                if 'Chapitre' in row:
                                    st.write(f"**Chapitre:** {row['Chapitre']}")
                            
                            # Zone de commentaire avec valeur existante
                            comment_key_work = f"comment_work_{row['Num']}"
                            existing_comment = row.get('Commentaire', '') if pd.notna(row.get('Commentaire', '')) else ''
                            comment_work = st.text_area(
                                "💬 Votre commentaire:",
                                value=existing_comment,
                                key=comment_key_work,
                                height=100,
                                placeholder="Modifiez vos observations pour cette exigence..."
                            )
                            comments_work_dict[row['Num']] = comment_work
                    
                    # Update DataFrame with modified comments
                    for index, row in filtered_work_checklist.iterrows():
                        if row['Num'] in comments_work_dict:
                            work_checklist_df.loc[work_checklist_df['Num'] == row['Num'], 'Commentaire'] = comments_work_dict[row['Num']]
                    
                    st.session_state['edited_checklist_work'] = work_checklist_df
                else:
                    st.warning("Aucune exigence trouvée avec ces filtres.")
            
            elif tab_work == "Non-conformities" and 'nc' in work_data:
                st.subheader("Non-conformités (travail repris)")
                
                # Ensure consistent DataFrame structure
                work_nc_df = work_data['nc'].copy()
                
                # Verify column structure
                if 'Commentaire' not in work_nc_df.columns:
                    work_nc_df['Commentaire'] = ''
                if 'Plan d\'action' not in work_nc_df.columns:
                    work_nc_df['Plan d\'action'] = ''
                
                if len(work_nc_df) > 0:
                    st.warning(f"**{len(work_nc_df)} non-conformité(s) en cours de traitement** (scores B, C, D uniquement)")
                    
                    # SAME DETAILED VIEW as normal NC tab
                    nc_work_comments = {}
                    nc_work_actions = {}
                    
                    for index, row in work_nc_df.iterrows():
                        score_emoji = {"B": "⚠️", "C": "🟠", "D": "🔴"}.get(row['Score'], "❓")
                        score_color = {"B": "orange", "C": "orange", "D": "red"}.get(row['Score'], "gray")
                        
                        with st.expander(f"{score_emoji} **NC {row['Num']}** - Score: {row['Score']}", expanded=True):
                            # Informations détaillées
                            col_nc1, col_nc2 = st.columns([3, 1])
                            
                            with col_nc1:
                                st.markdown(f"**📖 Explication:** {row['Explication']}")
                                st.markdown(f"**📋 Explication détaillée:** {row['Explication détaillée']}")
                                if 'Réponse' in row:
                                    st.markdown(f"**💬 Réponse:** {row['Réponse']}")
                            
                            with col_nc2:
                                st.markdown(f"**N°:** {row['Num']}")
                                st.markdown(f"**Score:** :color[{score_color}][{row['Score']}]")
                                if 'Chapitre' in row:
                                    st.markdown(f"**Chapitre:** {row['Chapitre']}")
                            
                            # Zone de commentaires et plan d'action avec valeurs existantes
                            col_action1, col_action2 = st.columns(2)
                            
                            with col_action1:
                                comment_key_nc = f"nc_work_comment_{row['Num']}"
                                existing_comment_nc = row.get('Commentaire', '') if pd.notna(row.get('Commentaire', '')) else ''
                                comment_nc = st.text_area(
                                    "💬 Commentaire de l'auditeur:",
                                    value=existing_comment_nc,
                                    key=comment_key_nc,
                                    height=120,
                                    placeholder="Observations, causes identifiées, éléments de preuve..."
                                )
                                nc_work_comments[row['Num']] = comment_nc
                            
                            with col_action2:
                                action_key_nc = f"nc_work_action_{row['Num']}"
                                existing_action_nc = row.get('Plan d\'action', '') if pd.notna(row.get('Plan d\'action', '')) else ''
                                action_nc = st.text_area(
                                    "🎯 Plan d'action corrective:",
                                    value=existing_action_nc,
                                    key=action_key_nc,
                                    height=120,
                                    placeholder="Actions à mettre en place, responsable, échéance..."
                                )
                                nc_work_actions[row['Num']] = action_nc
                    
                    # Update DataFrame with modified comments and actions
                    for index, row in work_nc_df.iterrows():
                        if row['Num'] in nc_work_comments:
                            work_nc_df.loc[index, 'Commentaire'] = nc_work_comments[row['Num']]
                        if row['Num'] in nc_work_actions:
                            work_nc_df.loc[index, 'Plan d\'action'] = nc_work_actions[row['Num']]
                    
                    st.session_state['edited_nc_work'] = work_nc_df
                else:
                    st.success("🎉 Aucune non-conformité dans ce travail sauvegardé !")
            
            # Export updated work - SAME BUTTONS as normal mode
            st.markdown("---")
            col1, col2 = st.columns(2)
            
            with col1:
                st.write("**💾 Sauvegarder les modifications**")
                st.write("*Sauvegarde mise à jour*")
                if st.button("💾 Générer nouvelle sauvegarde", type="secondary", key="save_updated_work_btn"):
                    with st.spinner("Génération de la sauvegarde mise à jour..."):
                        output = BytesIO()
                        
                        with pd.ExcelWriter(output, engine='xlsxwriter', options={'strings_to_urls': False}) as writer:
                            workbook = writer.book
                            
                            # Use same professional formatting
                            header_format = workbook.add_format({
                                'bold': True,
                                'text_wrap': True,
                                'valign': 'top',
                                'fg_color': '#4472C4',
                                'font_color': 'white',
                                'border': 1,
                                'font_size': 11
                            })
                            
                            # Metadata
                            metadata = pd.DataFrame({
                                "Type": ["IFS_WORK_SAVE"],
                                "COID": ["travail_repris"],
                                "Company_Name": ["travail_modifie"],
                                "Date": [pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")],
                                "Purpose": ["Communication Reviewer/Auditeur - Mise à jour"]
                            })
                            metadata.to_excel(writer, index=False, sheet_name="METADATA")
                            
                            # Save updated work with new sheet names
                            if 'edited_profile_work' in st.session_state:
                                st.session_state['edited_profile_work'].to_excel(writer, index=False, sheet_name="Profile_Communication")
                            elif 'profile' in work_data:
                                work_data['profile'].to_excel(writer, index=False, sheet_name="Profile_Communication")
                            
                            if 'edited_checklist_work' in st.session_state:
                                st.session_state['edited_checklist_work'].to_excel(writer, index=False, sheet_name="Checklist_Communication")
                            elif 'checklist' in work_data:
                                work_data['checklist'].to_excel(writer, index=False, sheet_name="Checklist_Communication")
                            
                            if 'edited_nc_work' in st.session_state:
                                st.session_state['edited_nc_work'].to_excel(writer, index=False, sheet_name="NonConformities_ActionPlan")
                            elif 'nc' in work_data:
                                work_data['nc'].to_excel(writer, index=False, sheet_name="NonConformities_ActionPlan")
                            
                            # Apply formatting to headers
                            for sheet_name in ["Profile_Communication", "Checklist_Communication", "NonConformities_ActionPlan"]:
                                if sheet_name in writer.sheets:
                                    worksheet = writer.sheets[sheet_name]
                                    # Format headers
                                    for col_num in range(10):  # Assume max 10 columns
                                        try:
                                            worksheet.write(0, col_num, worksheet.cell(0, col_num).value, header_format)
                                        except:
                                            break
                        
                        output.seek(0)
                    
                    date_str = datetime.now().strftime("%Y%m%d_%H%M")
                    
                    st.success("💾 Sauvegarde mise à jour générée !")
                    st.download_button(
                        label="📥 Télécharger la sauvegarde mise à jour",
                        data=output,
                        file_name=f"travail_IFS_mis_a_jour_{date_str}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_updated_work_btn"
                    )
            
            with col2:
                st.write("**📊 Exporter en rapport final**")
                st.write("*Rapport final à partir du travail repris*")
                if st.button("📊 Générer rapport final", type="primary", key="export_from_work_btn"):
                    st.info("Fonctionnalité d'export final depuis travail repris - à implémenter si besoin")
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
        - **Interface identique** à l'audit normal
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
