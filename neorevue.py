import json
import pandas as pd
import streamlit as st
from io import BytesIO

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

            # Extract checklist data
            checklist_data = []
            for uuid, scoring in json_data['data']['modules']['food_8']['checklists']['checklistFood8']['resultScorings'].items():
                checklist_data.append({
                    "Num": uuid,
                    "Explanation": scoring['answers'].get('englishExplanationText', 'N/A'),
                    "Detailed Explanation": scoring['answers'].get('explanationText', 'N/A'),
                    "Score": scoring['score']['label'],
                    "Response": scoring['answers'].get('fieldAnswers', 'N/A')
                })

            # Extract non-conformities (points not rated A)
            non_conformities = [item for item in checklist_data if item['Score'] != 'A']

            # Create tabs for Profile, Checklist, and Non-conformities
            tab = st.radio("Sélectionnez un onglet:", ["Profile", "Checklist", "Non-conformities"])

            if tab == "Profile":
                st.subheader("Profile")
                for field, value in profile_data.items():
                    st.text_input(f"{field}", value=value, key=f"profile_{field}")

            elif tab == "Checklist":
                st.subheader("Checklist")
                for item in checklist_data:
                    st.write(f"Numéro d'exigence: {item['Num']}")
                    st.write(f"Explication: {item['Explanation']}")
                    st.write(f"Explication Détaillée: {item['Detailed Explanation']}")
                    st.write(f"Note: {item['Score']}")
                    st.write(f"Réponse: {item['Response']}")
                    st.text_area("Commentaire de l'utilisateur", key=f"checklist_comment_{item['Num']}")

            elif tab == "Non-conformities":
                st.subheader("Non-conformities")
                for item in non_conformities:
                    st.write(f"Numéro d'exigence: {item['Num']}")
                    st.write(f"Explication: {item['Explanation']}")
                    st.write(f"Explication Détaillée: {item['Detailed Explanation']}")
                    st.write(f"Note: {item['Score']}")
                    st.write(f"Réponse: {item['Response']}")
                    st.text_area("Commentaire de l'utilisateur", key=f"non_conformity_comment_{item['Num']}")

            # Option to download the extracted data as an Excel file with formatting and COID in the name
            if st.button("Exporter en Excel"):
                # Extract the COID number to use in the file name
                numero_coid = profile_data.get("N° COID du portail", "inconnu")

                # Create the Excel file with column formatting
                output = BytesIO()

                # Create Excel writer and adjust column widths
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    # Profile tab
                    df_profile = pd.DataFrame(list(profile_data.items()), columns=["Field", "Value"])
                    df_profile['Commentaire de l\'utilisateur'] = ''
                    df_profile['Réponse de l\'auditeur'] = ''
                    df_profile.to_excel(writer, index=False, sheet_name="Profile")

                    # Checklist tab
                    df_checklist = pd.DataFrame(checklist_data)
                    df_checklist['Commentaire de l\'utilisateur'] = ''
                    df_checklist['Réponse de l\'auditeur'] = ''
                    df_checklist.to_excel(writer, index=False, sheet_name="Checklist")

                    # Non-conformities tab
                    df_non_conformities = pd.DataFrame(non_conformities)
                    df_non_conformities['Commentaire de l\'utilisateur'] = ''
                    df_non_conformities['Réponse de l\'auditeur'] = ''
                    df_non_conformities.to_excel(writer, index=False, sheet_name="Non-conformities")

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
                            adjusted_width = (max_length + 2) * 1.2
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
    else:
        st.write("Veuillez charger un fichier IFS de NEO pour commencer.")

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
            edited_df = st.data_editor(df, num_rows="dynamic")

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
        st.write("Veuillez charger un fichier Excel pour commencer.")
