import json
import pandas as pd
import streamlit as st
from io import BytesIO
import requests

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

# Streamlit app
st.sidebar.title("Menu de Navigation")
main_option = st.sidebar.radio("Choisissez une fonctionnalité:", ["Traitement des rapports IFS", "Gestion des fichiers Excel"])

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

            # Load UUID mapping
            with st.spinner("Chargement du mapping des exigences..."):
                uuid_mapping_df = load_uuid_mapping_from_url(UUID_MAPPING_URL)

            # Extract checklist data with proper NUM mapping
            checklist_data = []
            if not uuid_mapping_df.empty:
                for _, row in uuid_mapping_df.iterrows():
                    uuid = row['UUID']
                    num = row['Num']
                    
                    # Check if this UUID exists in the JSON data
                    if f"data_modules_food_8_checklists_checklistFood8_resultScorings_{uuid}" in flattened_json_data_safe:
                        prefix = f"data_modules_food_8_checklists_checklistFood8_resultScorings_{uuid}"
                        
                        explanation_text = flattened_json_data_safe.get(f"{prefix}_answers_englishExplanationText", "N/A")
                        detailed_explanation = flattened_json_data_safe.get(f"{prefix}_answers_explanationText", "N/A")
                        score_label = flattened_json_data_safe.get(f"{prefix}_score_label", "N/A")
                        response = flattened_json_data_safe.get(f"{prefix}_answers_fieldAnswers", "N/A")
                        
                        checklist_data.append({
                            "Num": num,
                            "UUID": uuid,
                            "Chapitre": row.get('Chapitre', 'N/A'),
                            "Theme": row.get('Theme', 'N/A'),
                            "SSTheme": row.get('SSTheme', 'N/A'),
                            "Explanation": explanation_text,
                            "Detailed Explanation": detailed_explanation,
                            "Score": score_label,
                            "Response": response
                        })
            else:
                # Fallback: use original method if UUID mapping fails
                st.warning("Utilisation du mapping par défaut (UUID mapping non disponible)")
                for uuid, scoring in json_data['data']['modules']['food_8']['checklists']['checklistFood8']['resultScorings'].items():
                    checklist_data.append({
                        "Num": uuid,
                        "UUID": uuid,
                        "Chapitre": "N/A",
                        "Theme": "N/A",
                        "SSTheme": "N/A",
                        "Explanation": scoring['answers'].get('englishExplanationText', 'N/A'),
                        "Detailed Explanation": scoring['answers'].get('explanationText', 'N/A'),
                        "Score": scoring['score']['label'],
                        "Response": scoring['answers'].get('fieldAnswers', 'N/A')
                    })

            # Extract non-conformities (points not rated A)
            non_conformities = [item for item in checklist_data if item['Score'] != 'A' and item['Score'] != 'N/A']

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
                
                # Display editable table
                st.write("Vous pouvez ajouter vos commentaires dans la colonne 'Commentaire':")
                edited_profile_df = st.data_editor(
                    profile_df,
                    column_config={
                        "Champ": st.column_config.TextColumn("Champ", disabled=True),
                        "Valeur": st.column_config.TextColumn("Valeur", disabled=True),
                        "Commentaire": st.column_config.TextColumn("Commentaire", width="medium")
                    },
                    hide_index=True,
                    use_container_width=True
                )
                
                # Store edited data in session state
                st.session_state['edited_profile'] = edited_profile_df

            elif tab == "Checklist":
                st.subheader("Checklist des exigences")
                
                # Add filters
                col1, col2, col3 = st.columns(3)
                with col1:
                    score_filter = st.selectbox("Filtrer par score:", ["Tous", "A", "B", "C", "D"])
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
                
                # Create DataFrame for checklist with comments column
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
                
                if checklist_list:
                    checklist_df = pd.DataFrame(checklist_list)
                    
                    # Column configuration
                    column_config = {
                        "Num": st.column_config.TextColumn("N°", disabled=True, width="small"),
                        "Score": st.column_config.TextColumn("Score", disabled=True, width="small"),
                        "Chapitre": st.column_config.TextColumn("Chapitre", disabled=True, width="small"),
                        "Explication": st.column_config.TextColumn("Explication", disabled=True, width="medium"),
                        "Explication détaillée": st.column_config.TextColumn("Explication détaillée", disabled=True, width="medium"),
                        "Commentaire": st.column_config.TextColumn("Commentaire", width="medium")
                    }
                    if show_responses:
                        column_config["Réponse"] = st.column_config.TextColumn("Réponse", disabled=True, width="medium")
                    
                    # Display editable table
                    st.write("Vous pouvez ajouter vos commentaires dans la colonne 'Commentaire':")
                    edited_checklist_df = st.data_editor(
                        checklist_df,
                        column_config=column_config,
                        hide_index=True,
                        use_container_width=True,
                        height=600
                    )
                    
                    # Store edited data in session state
                    st.session_state['edited_checklist'] = edited_checklist_df
                else:
                    st.warning("Aucune exigence trouvée avec ces filtres.")

            elif tab == "Non-conformities":
                st.subheader("Non-conformités")
                
                if non_conformities:
                    st.warning(f"**{len(non_conformities)} non-conformité(s) détectée(s)**")
                    
                    # Create DataFrame for non-conformities with comments column
                    nc_list = []
                    for item in non_conformities:
                        nc_list.append({
                            "Num": item['Num'],
                            "Score": item['Score'],
                            "Chapitre": item.get('Chapitre', 'N/A'),
                            "Explication": item['Explanation'],
                            "Explication détaillée": item['Detailed Explanation'],
                            "Réponse": item['Response'],
                            "Commentaire": "",
                            "Plan d'action": ""
                        })
                    
                    nc_df = pd.DataFrame(nc_list)
                    
                    # Display editable table
                    st.write("Vous pouvez ajouter vos commentaires et plans d'action:")
                    edited_nc_df = st.data_editor(
                        nc_df,
                        column_config={
                            "Num": st.column_config.TextColumn("N°", disabled=True, width="small"),
                            "Score": st.column_config.TextColumn("Score", disabled=True, width="small"),
                            "Chapitre": st.column_config.TextColumn("Chapitre", disabled=True, width="small"),
                            "Explication": st.column_config.TextColumn("Explication", disabled=True, width="medium"),
                            "Explication détaillée": st.column_config.TextColumn("Explication détaillée", disabled=True, width="medium"),
                            "Réponse": st.column_config.TextColumn("Réponse", disabled=True, width="medium"),
                            "Commentaire": st.column_config.TextColumn("Commentaire", width="medium"),
                            "Plan d'action": st.column_config.TextColumn("Plan d'action", width="medium")
                        },
                        hide_index=True,
                        use_container_width=True,
                        height=600
                    )
                    
                    # Store edited data in session state
                    st.session_state['edited_nc'] = edited_nc_df
                else:
                    st.success("🎉 Aucune non-conformité détectée ! Toutes les exigences sont conformes.")

            # Option to download the extracted data as an Excel file with formatting and COID in the name
            st.markdown("---")
            if st.button("Exporter en Excel", type="primary"):
                # Extract the COID number to use in the file name
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
                    label="Télécharger le fichier Excel",
                    data=output,
                    file_name=f'extraction_{numero_coid}.xlsx',
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except json.JSONDecodeError:
            st.error("Erreur lors du décodage du fichier JSON. Veuillez vous assurer qu'il est au format correct.")
        except Exception as e:
            st.error(f"Erreur lors du traitement du fichier: {str(e)}")
    else:
        st.info("Veuillez charger un fichier IFS de NEO pour commencer.")

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
