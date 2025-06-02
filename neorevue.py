import json
import pandas as pd
import streamlit as st
from io import BytesIO
import requests
from datetime import datetime
import re

# ================================
# CONFIGURATION STREAMLIT
# ================================
st.set_page_config(
    page_title="IFS NEO Data Extractor",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ================================
# FONCTIONS UTILITAIRES
# ================================
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

def load_uuid_mapping_from_url(url):
    """Charger le mapping UUID depuis l'URL"""
    try:
        response = requests.get(url)
        if response.status_code == 200:
            from io import StringIO
            csv_data = StringIO(response.text)
            uuid_mapping_df = pd.read_csv(csv_data)
            
            # Vérifier les colonnes requises
            required_columns = ['UUID', 'Num', 'Chapitre', 'Theme', 'SSTheme']
            for column in required_columns:
                if column not in uuid_mapping_df.columns:
                    st.error(f"Le fichier CSV doit contenir une colonne '{column}' avec des valeurs valides.")
                    return pd.DataFrame()
            
            uuid_mapping_df = uuid_mapping_df.dropna(subset=['UUID', 'Num'])
            uuid_mapping_df['Chapitre'] = uuid_mapping_df['Chapitre'].astype(str).str.strip()
            uuid_mapping_df = uuid_mapping_df.drop_duplicates(subset=['Chapitre', 'Num'])
            return uuid_mapping_df
        else:
            st.error("Impossible de charger le fichier CSV des UUID depuis l'URL fourni.")
            return pd.DataFrame()
    except Exception as e:
        st.error(f"Erreur lors du chargement du mapping UUID: {str(e)}")
        return pd.DataFrame()

def generate_audit_summary(checklist_df, nc_df, profile_df):
    """Génère un résumé d'audit détaillé"""
    summary_data = []
    
    # Summary statistics from checklist
    if not checklist_df.empty:
        total_requirements = len(checklist_df)
        score_counts = checklist_df['Score'].value_counts()
        
        summary_data.extend([
            {"Métrique": "Total des exigences", "Valeur": total_requirements, "Description": "Nombre total d'exigences auditées"},
            {"Métrique": "Score A (Conforme)", "Valeur": score_counts.get('A', 0), "Description": "Exigences entièrement conformes"},
            {"Métrique": "Score B (Écart mineur)", "Valeur": score_counts.get('B', 0), "Description": "Écarts mineurs détectés"},
            {"Métrique": "Score C (Écart majeur)", "Valeur": score_counts.get('C', 0), "Description": "Écarts majeurs détectés"},
            {"Métrique": "Score D (Critique)", "Valeur": score_counts.get('D', 0), "Description": "Non-conformités critiques"},
            {"Métrique": "Score NA (Non applicable)", "Valeur": score_counts.get('NA', 0), "Description": "Exigences non applicables"},
        ])
        
        # Calculate compliance percentage
        conformes = score_counts.get('A', 0) + score_counts.get('NA', 0)
        if total_requirements > 0:
            compliance_rate = (conformes / total_requirements) * 100
            summary_data.append({
                "Métrique": "Taux de conformité (%)", 
                "Valeur": f"{compliance_rate:.1f}%", 
                "Description": "Pourcentage d'exigences conformes ou non applicables"
            })

    # Non-conformities summary
    if not nc_df.empty:
        nc_total = len(nc_df)
        status_counts = nc_df['Statut'].value_counts() if 'Statut' in nc_df.columns else {}
        
        summary_data.extend([
            {"Métrique": "Non-conformités totales", "Valeur": nc_total, "Description": "Nombre total de non-conformités (B, C, D)"},
            {"Métrique": "Actions terminées", "Valeur": status_counts.get('Terminé', 0), "Description": "Actions correctives terminées"},
            {"Métrique": "Actions en cours", "Valeur": status_counts.get('En cours', 0), "Description": "Actions correctives en cours"},
            {"Métrique": "Actions en attente", "Valeur": status_counts.get('En attente', 0), "Description": "Actions correctives en attente"},
        ])
        
        if nc_total > 0:
            completion_rate = (status_counts.get('Terminé', 0) / nc_total) * 100
            summary_data.append({
                "Métrique": "Taux de résolution (%)", 
                "Valeur": f"{completion_rate:.1f}%", 
                "Description": "Pourcentage d'actions correctives terminées"
            })

    # Add audit info if available from profile
    if not profile_df.empty:
        try:
            if 'Champ' in profile_df.columns and 'Valeur' in profile_df.columns:
                company_info = profile_df.set_index('Champ')['Valeur'].to_dict()
            else:
                company_info = {}
        except:
            company_info = {}
            
        summary_data.extend([
            {"Métrique": "Entreprise auditée", "Valeur": company_info.get("Nom du site à auditer", "N/A"), "Description": "Nom de l'entreprise"},
            {"Métrique": "COID", "Valeur": company_info.get("N° COID du portail", "N/A"), "Description": "Numéro COID"},
            {"Métrique": "Pays", "Valeur": company_info.get("Pays", "N/A"), "Description": "Pays du site audité"},
            {"Métrique": "Date d'export", "Valeur": datetime.now().strftime("%Y-%m-%d %H:%M"), "Description": "Date et heure de génération du rapport"},
        ])

    return pd.DataFrame(summary_data)

# ================================
# FONCTIONS SAUVEGARDE/CHARGEMENT
# ================================
def save_work_to_excel(profile_data, checklist_data, non_conformities, edited_profile=None, edited_checklist=None, edited_nc=None, coid="inconnu"):
    """Save current work with comments to Excel for later resume and reviewer/auditor communication"""
    output = BytesIO()
    
    # Extract company name for filename
    company_name = profile_data.get("Nom du site à auditer", "entreprise_inconnue")
    # Clean company name for filename (remove special characters)
    clean_company_name = re.sub(r'[^\w\s-]', '', company_name).strip()
    clean_company_name = re.sub(r'[-\s]+', '_', clean_company_name)
    
    # Use xlsxwriter for better formatting in communication files
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
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
                nc_save_df['Réponse de l\'auditeur'] = ''
                nc_save_df['Plan d\'action proposé'] = ''
                nc_save_df['Actions mises en place'] = ''
                nc_save_df['Date limite'] = ''
                nc_save_df['Responsable'] = ''
                nc_save_df['Statut'] = 'En attente'
            
            # Ensure all communication and tracking columns exist
            communication_columns = [
                "Commentaire du reviewer", "Réponse de l'auditeur", "Plan d'action proposé", 
                "Actions mises en place", "Date limite", "Responsable", "Statut"
            ]
            
            for col in communication_columns:
                if col not in nc_save_df.columns:
                    if col == "Statut":
                        nc_save_df[col] = "En attente"
                    else:
                        nc_save_df[col] = ""
            
            nc_save_df.to_excel(writer, index=False, sheet_name="NonConformities_ActionPlan")
            
            # Format Non-conformities sheet with proper column order
            nc_ws = writer.sheets["NonConformities_ActionPlan"]
            nc_ws.set_column('A:A', 12)  # Num
            nc_ws.set_column('B:B', 8)   # Score
            nc_ws.set_column('C:C', 15)  # Chapitre
            nc_ws.set_column('D:D', 40)  # Explication
            nc_ws.set_column('E:E', 40)  # Explication détaillée
            nc_ws.set_column('F:F', 30)  # Réponse système
            nc_ws.set_column('G:G', 35)  # Commentaire reviewer
            nc_ws.set_column('H:H', 35)  # Réponse auditeur
            nc_ws.set_column('I:I', 35)  # Plan d'action proposé
            nc_ws.set_column('J:J', 35)  # Actions mises en place
            nc_ws.set_column('K:K', 15)  # Date limite
            nc_ws.set_column('L:L', 20)  # Responsable
            nc_ws.set_column('M:M', 15)  # Statut
            
            # Apply formatting with color coding
            for col_num, value in enumerate(nc_save_df.columns.values):
                nc_ws.write(0, col_num, value, header_format)

    output.seek(0)
    return output, clean_company_name

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
        
        # Load profile work (try both old and new sheet names)
        if "Profile_Communication" in excel_data:
            work_data['profile'] = excel_data["Profile_Communication"]
        elif "Profile_Work" in excel_data:
            work_data['profile'] = excel_data["Profile_Work"]
        
        # Load checklist work (try both old and new sheet names)
        if "Checklist_Communication" in excel_data:
            work_data['checklist'] = excel_data["Checklist_Communication"]
        elif "Checklist_Work" in excel_data:
            work_data['checklist'] = excel_data["Checklist_Work"]
        
        # Load non-conformities work (try both old and new sheet names)
        if "NonConformities_ActionPlan" in excel_data:
            work_data['nc'] = excel_data["NonConformities_ActionPlan"]
        elif "NonConformities_Work" in excel_data:
            work_data['nc'] = excel_data["NonConformities_Work"]
        
        coid = metadata.iloc[0]["COID"]
        company_name = metadata.iloc[0].get("Company_Name", "Entreprise inconnue") if "Company_Name" in metadata.columns else "Entreprise inconnue"
        save_date = metadata.iloc[0]["Date"]
        purpose = metadata.iloc[0].get("Purpose", "Travail standard") if "Purpose" in metadata.columns else "Travail standard"
        
        return work_data, f"Travail chargé avec succès\n**COID:** {coid} | **Entreprise:** {company_name}\n**Sauvegardé le:** {save_date}\n**Type:** {purpose}"
        
    except Exception as e:
        return None, f"Erreur lors du chargement du fichier de travail: {str(e)}"

def create_final_report_excel(profile_df, checklist_df, nc_df, is_from_work=False, work_data=None):
    """Crée un rapport final Excel complet avec formatage professionnel"""
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # Define professional formats for final report
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D7E4BC',  # Green theme for final report
            'border': 1,
            'font_size': 11
        })
        
        cell_format = workbook.add_format({
            'text_wrap': True,
            'valign': 'top',
            'border': 1,
            'font_size': 10
        })
        
        # Professional formatting for different sections
        nc_format = workbook.add_format({
            'text_wrap': True,
            'valign': 'top',
            'border': 1,
            'fg_color': '#FFE6E6',  # Light red for non-conformities
            'font_size': 10
        })
        
        completed_format = workbook.add_format({
            'text_wrap': True,
            'valign': 'top',
            'border': 1,
            'fg_color': '#E6FFE6',  # Light green for completed actions
            'font_size': 10
        })
        
        warning_format = workbook.add_format({
            'text_wrap': True,
            'valign': 'top',
            'border': 1,
            'fg_color': '#FFF2CC',  # Light yellow for warnings
            'font_size': 10
        })
        
        # Create summary dashboard first
        summary_df = generate_audit_summary(checklist_df, nc_df, profile_df)
        if not summary_df.empty:
            summary_df.to_excel(writer, index=False, sheet_name="Résumé_Audit")
        
        # Profile sheet
        if not profile_df.empty:
            profile_df.to_excel(writer, index=False, sheet_name="Profile")

        # Checklist sheet
        if not checklist_df.empty:
            checklist_df.to_excel(writer, index=False, sheet_name="Checklist")

        # Non-conformities sheet
        if not nc_df.empty:
            nc_df.to_excel(writer, index=False, sheet_name="Non-conformities")

        # Apply enhanced formatting to all sheets
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            
            # Get the dataframe for this sheet
            if sheet_name == "Profile":
                df = profile_df
            elif sheet_name == "Checklist":
                df = checklist_df
            elif sheet_name == "Non-conformities":
                df = nc_df
            elif sheet_name == "Résumé_Audit":
                df = summary_df
            else:
                continue
            
            if df.empty:
                continue
            
            # Format headers
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Set column widths dynamically
            for col_num, col_name in enumerate(df.columns):
                if isinstance(col_name, str):
                    if col_name in ['Commentaire du reviewer', 'Réponse de l\'auditeur', 'Plan d\'action proposé', 'Actions mises en place']:
                        worksheet.set_column(col_num, col_num, 40)  # Wide columns for comments
                    elif col_name in ['Explication', 'Explication détaillée', 'Description']:
                        worksheet.set_column(col_num, col_num, 35)  # Medium-wide for explanations
                    elif col_name in ['Réponse']:
                        worksheet.set_column(col_num, col_num, 30)  # Medium for responses
                    elif col_name in ['Valeur', 'Champ']:
                        worksheet.set_column(col_num, col_num, 25)  # Medium for values
                    elif col_name in ['Num', 'Score', 'Statut']:
                        worksheet.set_column(col_num, col_num, 12)  # Narrow for short data
                    else:
                        worksheet.set_column(col_num, col_num, 20)  # Default width
            
            # Apply conditional formatting for non-conformities sheet
            if sheet_name == "Non-conformities" and not df.empty:
                for row_num in range(1, len(df) + 1):
                    # Get the score and status for conditional formatting
                    score_col = None
                    status_col = None
                    for col_num, col_name in enumerate(df.columns):
                        if col_name == 'Score':
                            score_col = col_num
                        elif col_name == 'Statut':
                            status_col = col_num
                    
                    for col_num in range(len(df.columns)):
                        cell_value = df.iloc[row_num-1, col_num]
                        
                        # Apply formatting based on score and status
                        if score_col is not None:
                            score = df.iloc[row_num-1, score_col]
                            status = df.iloc[row_num-1, status_col] if status_col is not None else None
                            
                            if status == 'Terminé':
                                worksheet.write(row_num, col_num, cell_value, completed_format)
                            elif score == 'D':
                                worksheet.write(row_num, col_num, cell_value, nc_format)
                            elif score in ['B', 'C']:
                                worksheet.write(row_num, col_num, cell_value, warning_format)
                            else:
                                worksheet.write(row_num, col_num, cell_value, cell_format)
                        else:
                            worksheet.write(row_num, col_num, cell_value, cell_format)
            else:
                # Apply standard formatting for other sheets
                for row_num in range(1, len(df) + 1):
                    for col_num in range(len(df.columns)):
                        cell_value = df.iloc[row_num-1, col_num]
                        worksheet.write(row_num, col_num, cell_value, cell_format)

    output.seek(0)
    return output

# ================================
# CONSTANTES ET CONFIGURATION
# ================================
UUID_MAPPING_URL = "https://raw.githubusercontent.com/M00N69/Gemini-Knowledge/refs/heads/main/IFSV8listUUID.csv"

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

# ================================
# FONCTION PRINCIPALE
# ================================
def main():
    """Fonction principale de l'application"""
    
    # Header avec style amélioré
    st.markdown("""
    <div style="background: linear-gradient(90deg, #4472C4, #70AD47); padding: 20px; border-radius: 10px; margin-bottom: 30px;">
        <h1 style="color: white; text-align: center; margin: 0;">🔍 IFS NEO Data Extractor</h1>
        <p style="color: white; text-align: center; margin: 10px 0 0 0; font-size: 18px;">
            Application de communication reviewer ↔ auditeur pour audits IFS
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # Navigation dans la sidebar
    st.sidebar.title("📋 Menu de Navigation")
    st.sidebar.markdown("---")
    main_option = st.sidebar.radio("Choisissez une fonctionnalité:", [
        "Traitement des rapports IFS", 
        "Gestion des fichiers Excel", 
        "Reprendre un travail sauvegardé"
    ])

    # ================================
    # TRAITEMENT DES RAPPORTS IFS
    # ================================
    if main_option == "Traitement des rapports IFS":
        st.subheader("📊 Traitement d'un nouveau rapport IFS")
        
        # Upload du fichier IFS
        uploaded_json_file = st.file_uploader(
            "📁 Charger le fichier IFS de NEO", 
            type="ifs",
            help="Sélectionnez le fichier d'audit IFS (.ifs) exporté depuis NEO"
        )

        if uploaded_json_file:
            try:
                # Charger et traiter le fichier JSON
                with st.spinner("Traitement du fichier IFS..."):
                    json_data = json.load(uploaded_json_file)
                    flattened_json_data = flatten_json_safe(json_data)
                
                # Extraire les données de profil
                profile_data = extract_from_flattened(
                    flattened_json_data, 
                    FLATTENED_FIELD_MAPPING, 
                    list(FLATTENED_FIELD_MAPPING.keys())
                )
                
                # Extract checklist data - ALWAYS use direct method first
                checklist_data = []
                
                # Debug: Check if checklist data exists in JSON
                if 'data' in json_data and 'modules' in json_data['data'] and 'food_8' in json_data['data']['modules']:
                    if 'checklists' in json_data['data']['modules']['food_8'] and 'checklistFood8' in json_data['data']['modules']['food_8']['checklists']:
                        if 'resultScorings' in json_data['data']['modules']['food_8']['checklists']['checklistFood8']:
                            result_scorings = json_data['data']['modules']['food_8']['checklists']['checklistFood8']['resultScorings']
                            st.success(f"✅ Données de checklist trouvées : {len(result_scorings)} exigences détectées")
                            
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
                                    "Explication": explanation_text,
                                    "Explication détaillée": detailed_explanation,
                                    "Score": score_label,
                                    "Réponse": response
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

                # Quick stats
                if checklist_data:
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("Total exigences", len(checklist_data))
                    with col2:
                        conformes = len([item for item in checklist_data if item['Score'] in ['A', 'NA']])
                        st.metric("Conformes", conformes)
                    with col3:
                        st.metric("Non-conformités", len(non_conformities))
                    with col4:
                        if len(checklist_data) > 0:
                            taux = (conformes / len(checklist_data)) * 100
                            st.metric("Taux conformité", f"{taux:.1f}%")

                # Work save and export options in SIDEBAR
                st.sidebar.markdown("---")
                st.sidebar.markdown("### 💾 Sauvegarde et Export")
                
                # Sauvegarde du travail
                st.sidebar.write("**💾 Sauvegarde travail**")
                st.sidebar.write("*Pour reprendre plus tard*")
                if st.sidebar.button("💾 Générer la sauvegarde", type="secondary", key="save_work_btn"):
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
                    
                    st.sidebar.success("💾 Sauvegarde générée !")
                    st.sidebar.download_button(
                        label="📥 Télécharger sauvegarde",
                        data=work_file,
                        file_name=work_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        help="Fichier Excel contenant tous vos commentaires pour reprendre plus tard",
                        key="download_work_btn"
                    )
                
                # Export rapport final
                st.sidebar.write("**📊 Rapport final**")
                st.sidebar.write("*Pour présentation*")
                if st.sidebar.button("📊 Générer le rapport", type="primary", key="export_final_btn"):
                    numero_coid = profile_data.get("N° COID du portail", "inconnu")
                    
                    with st.spinner("Génération du rapport final..."):
                        # Prepare DataFrames for final report
                        if 'edited_profile' in st.session_state:
                            profile_export_df = st.session_state['edited_profile']
                        else:
                            profile_export_df = pd.DataFrame([
                                {"Champ": k, "Valeur": v, "Commentaire du reviewer": "", "Réponse de l'auditeur": ""} 
                                for k, v in profile_data.items()
                            ])
                        
                        if 'edited_checklist' in st.session_state:
                            checklist_export_df = st.session_state['edited_checklist']
                        else:
                            checklist_export_df = pd.DataFrame(checklist_data)
                            checklist_export_df['Commentaire du reviewer'] = ''
                            checklist_export_df['Réponse de l\'auditeur'] = ''
                        
                        if non_conformities:
                            if 'edited_nc' in st.session_state:
                                nc_export_df = st.session_state['edited_nc']
                            else:
                                nc_export_df = pd.DataFrame(non_conformities)
                                nc_export_df['Commentaire du reviewer'] = ''
                                nc_export_df['Réponse de l\'auditeur'] = ''
                                nc_export_df['Plan d\'action proposé'] = ''
                                nc_export_df['Actions mises en place'] = ''
                                nc_export_df['Date limite'] = ''
                                nc_export_df['Responsable'] = ''
                                nc_export_df['Statut'] = 'En attente'
                        else:
                            nc_export_df = pd.DataFrame()
                        
                        # Create final report
                        final_report = create_final_report_excel(profile_export_df, checklist_export_df, nc_export_df)
                    
                    st.sidebar.success("📊 Rapport final généré !")
                    
                    # Create filename with company name
                    company_name = profile_data.get("Nom du site à auditer", "entreprise_inconnue")
                    clean_company_name = re.sub(r'[^\w\s-]', '', company_name).strip()
                    clean_company_name = re.sub(r'[-\s]+', '_', clean_company_name)
                    final_filename = f'rapport_final_IFS_{numero_coid}_{clean_company_name}.xlsx'
                    
                    # Provide the download button with the COID number in the filename
                    st.sidebar.download_button(
                        label="📥 Télécharger rapport",
                        data=final_report,
                        file_name=final_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        help="Rapport final complet avec résumé d'audit",
                        key="download_final_btn"
                    )

                # Create tabs for Profile, Checklist, and Non-conformities
                tab = st.radio("Sélectionnez un onglet:", ["Profile", "Checklist", "Non-conformities"])

                if tab == "Profile":
                    st.subheader("🏢 Profile de l'entreprise")
                    
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
                    st.subheader("📋 Checklist des exigences")
                    
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
                                            if (item['Explication'] != 'N/A' and item['Explication'].strip() != '') 
                                            or (item['Explication détaillée'] != 'N/A' and item['Explication détaillée'].strip() != '')]
                    elif content_filter == "Explications vides":
                        filtered_checklist = [item for item in filtered_checklist 
                                            if (item['Explication'] == 'N/A' or item['Explication'].strip() == '') 
                                            and (item['Explication détaillée'] == 'N/A' or item['Explication détaillée'].strip() == '')]
                    
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
                                    st.write(f"**📖 Explication:** {item['Explication']}")
                                    st.write(f"**📋 Explication détaillée:** {item['Explication détaillée']}")
                                    if show_responses:
                                        st.write(f"**💬 Réponse:** {item['Réponse']}")
                                
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
                                "Explication": item['Explication'],
                                "Explication détaillée": item['Explication détaillée'],
                                "Commentaire du reviewer": comments_reviewer_dict.get(item['Num'], ""),
                                "Réponse de l'auditeur": comments_auditor_dict.get(item['Num'], "")
                            }
                            if show_responses:
                                row["Réponse"] = item['Réponse']
                            checklist_list.append(row)
                        
                        checklist_df = pd.DataFrame(checklist_list)
                        st.session_state['edited_checklist'] = checklist_df
                    else:
                        st.warning("Aucune exigence trouvée avec ces filtres.")

                elif tab == "Non-conformities":
                    st.subheader("❗ Non-conformités - Plan d'actions et suivi")
                    
                    if non_conformities:
                        st.warning(f"**{len(non_conformities)} non-conformité(s) détectée(s)** (scores B, C, D uniquement)")
                        st.info("💡 **Communication reviewer ↔ auditeur:** Commentaires → Plans d'action → Réponses → Suivi")
                        
                        # Affichage avec expandeurs pour les non-conformités
                        nc_comments = {}
                        nc_auditor_responses = {}
                        nc_action_plans = {}
                        nc_implemented_actions = {}
                        nc_deadlines = {}
                        nc_responsibles = {}
                        nc_status = {}
                        
                        for item in non_conformities:
                            score_emoji = {"B": "⚠️", "C": "🟠", "D": "🔴"}.get(item['Score'], "❓")
                            score_color = {"B": "orange", "C": "orange", "D": "red"}.get(item['Score'], "gray")
                            
                            with st.expander(f"{score_emoji} **NC {item['Num']}** - Score: {item['Score']}", expanded=True):
                                # Informations détaillées
                                col_nc1, col_nc2 = st.columns([3, 1])
                                
                                with col_nc1:
                                    st.markdown(f"**📖 Explication:** {item['Explication']}")
                                    st.markdown(f"**📋 Explication détaillée:** {item['Explication détaillée']}")
                                    st.markdown(f"**💬 Réponse système:** {item['Réponse']}")
                                
                                with col_nc2:
                                    st.markdown(f"**N°:** {item['Num']}")
                                    st.markdown(f"**Score:** :color[{score_color}][{item['Score']}]")
                                    if 'Chapitre' in item:
                                        st.markdown(f"**Chapitre:** {item['Chapitre']}")
                                
                                # Communication reviewer/auditeur
                                st.markdown("---")
                                col_comm1, col_comm2 = st.columns(2)
                                
                                with col_comm1:
                                    # Commentaires du reviewer
                                    comment_reviewer_key = f"nc_reviewer_comment_{item['Num']}"
                                    comment_reviewer = st.text_area(
                                        "📝 Commentaire du reviewer:",
                                        key=comment_reviewer_key,
                                        height=100,
                                        placeholder="Observations du reviewer, éléments à améliorer, contexte..."
                                    )
                                    nc_comments[item['Num']] = comment_reviewer
                                
                                with col_comm2:
                                    # Réponse de l'auditeur
                                    auditor_response_key = f"nc_auditor_response_{item['Num']}"
                                    auditor_response = st.text_area(
                                        "💬 Réponse de l'auditeur:",
                                        key=auditor_response_key,
                                        height=100,
                                        placeholder="Réponse et justifications de l'auditeur..."
                                    )
                                    nc_auditor_responses[item['Num']] = auditor_response
                                
                                # Plan d'action et suivi
                                st.markdown("**🎯 Plan d'action et suivi**")
                                col_action1, col_action2 = st.columns(2)
                                
                                with col_action1:
                                    # Plan d'action proposé
                                    action_plan_key = f"nc_action_plan_{item['Num']}"
                                    action_plan = st.text_area(
                                        "📋 Plan d'action proposé:",
                                        key=action_plan_key,
                                        height=80,
                                        placeholder="Actions recommandées..."
                                    )
                                    nc_action_plans[item['Num']] = action_plan
                                
                                with col_action2:
                                    # Actions mises en place
                                    implemented_key = f"nc_implemented_{item['Num']}"
                                    implemented = st.text_area(
                                        "✅ Actions mises en place:",
                                        key=implemented_key,
                                        height=80,
                                        placeholder="Actions réalisées..."
                                    )
                                    nc_implemented_actions[item['Num']] = implemented
                                
                                # Suivi des actions
                                st.markdown("**📊 Suivi des actions**")
                                col_track1, col_track2, col_track3 = st.columns(3)
                                
                                with col_track1:
                                    deadline_key = f"nc_deadline_{item['Num']}"
                                    deadline = st.date_input(
                                        "📅 Date limite:",
                                        key=deadline_key,
                                        help="Date limite pour la mise en œuvre"
                                    )
                                    nc_deadlines[item['Num']] = deadline.strftime("%Y-%m-%d") if deadline else ""
                                
                                with col_track2:
                                    responsible_key = f"nc_responsible_{item['Num']}"
                                    responsible = st.text_input(
                                        "👤 Responsable:",
                                        key=responsible_key,
                                        placeholder="Nom du responsable"
                                    )
                                    nc_responsibles[item['Num']] = responsible
                                
                                with col_track3:
                                    status_key = f"nc_status_{item['Num']}"
                                    status = st.selectbox(
                                        "📈 Statut:",
                                        ["En attente", "En cours", "Terminé", "Reporté", "Annulé"],
                                        key=status_key
                                    )
                                    nc_status[item['Num']] = status
                        
                        # Reconstituer le DataFrame avec toutes les colonnes de communication
                        nc_list = []
                        for item in non_conformities:
                            nc_list.append({
                                "Num": item['Num'],
                                "Score": item['Score'],
                                "Chapitre": item.get('Chapitre', 'N/A'),
                                "Explication": item['Explication'],
                                "Explication détaillée": item['Explication détaillée'],
                                "Réponse": item['Réponse'],
                                "Commentaire du reviewer": nc_comments.get(item['Num'], ""),
                                "Réponse de l'auditeur": nc_auditor_responses.get(item['Num'], ""),
                                "Plan d'action proposé": nc_action_plans.get(item['Num'], ""),
                                "Actions mises en place": nc_implemented_actions.get(item['Num'], ""),
                                "Date limite": nc_deadlines.get(item['Num'], ""),
                                "Responsable": nc_responsibles.get(item['Num'], ""),
                                "Statut": nc_status.get(item['Num'], "En attente")
                            })
                        
                        nc_df = pd.DataFrame(nc_list)
                        st.session_state['edited_nc'] = nc_df
                        
                    else:
                        st.success("🎉 Aucune non-conformité détectée ! Toutes les exigences sont conformes (A) ou non applicables (NA).")

            except json.JSONDecodeError:
                st.error("Erreur lors du décodage du fichier JSON. Veuillez vous assurer qu'il est au format correct.")
            except Exception as e:
                st.error(f"Erreur lors du traitement du fichier: {str(e)}")
        else:
            st.info("Veuillez charger un fichier IFS de NEO pour commencer.")
            
            # Guide d'utilisation
            st.markdown("""
            ### 📋 Guide d'utilisation :
            
            1. **Exportez** votre audit depuis NEO au format `.ifs`
            2. **Chargez** le fichier ci-dessus
            3. **Explorez** les données extraites dans les onglets
            4. **Ajoutez** vos commentaires reviewer/auditeur
            5. **Sauvegardez** votre travail ou exportez le rapport final
            
            ✨ **Fonctionnalités principales :**
            - Extraction automatique des données IFS
            - Communication bidirectionnelle reviewer ↔ auditeur
            - Suivi des non-conformités avec plans d'action
            - Exports Excel formatés professionnellement
            - Sauvegarde de travail pour reprendre plus tard
            """)

    # ================================
    # REPRENDRE UN TRAVAIL SAUVEGARDÉ
    # ================================
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
                    st.subheader("🏢 Profile de l'entreprise (travail repris)")
                    
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
                    st.subheader("📋 Checklist des exigences (travail repris)")
                    
                    # Ensure consistent DataFrame structure
                    work_checklist_df = work_data['checklist'].copy()
                    
                    # Verify column structure for reviewer/auditor communication
                    if "Commentaire du reviewer" not in work_checklist_df.columns:
                        if "Commentaire" in work_checklist_df.columns:
                            work_checklist_df["Commentaire du reviewer"] = work_checklist_df["Commentaire"]
                        else:
                            work_checklist_df["Commentaire du reviewer"] = ""
                    
                    if "Réponse de l'auditeur" not in work_checklist_df.columns:
                        work_checklist_df["Réponse de l'auditeur"] = ""
                    
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
                        st.write("**📋 Vue détaillée des exigences - Communication reviewer/auditeur (travail repris)**")
                        st.write("💡 *Modifiez les commentaires et réponses pour continuer la communication*")
                        
                        # Expandeurs avec commentaires reviewer/auditeur
                        comments_work_reviewer_dict = {}
                        comments_work_auditor_dict = {}
                        
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
                                
                                # Communication reviewer/auditeur avec valeurs existantes
                                st.markdown("---")
                                col_comm1, col_comm2 = st.columns(2)
                                
                                with col_comm1:
                                    # Zone de commentaire reviewer avec valeur existante
                                    comment_reviewer_key_work = f"comment_work_reviewer_{row['Num']}"
                                    existing_reviewer_comment = row.get('Commentaire du reviewer', '') if pd.notna(row.get('Commentaire du reviewer', '')) else ''
                                    comment_reviewer_work = st.text_area(
                                        "📝 Commentaire du reviewer:",
                                        value=existing_reviewer_comment,
                                        key=comment_reviewer_key_work,
                                        height=80,
                                        placeholder="Modifiez les observations du reviewer..."
                                    )
                                    comments_work_reviewer_dict[row['Num']] = comment_reviewer_work
                                
                                with col_comm2:
                                    # Zone de réponse auditeur avec valeur existante
                                    comment_auditor_key_work = f"comment_work_auditor_{row['Num']}"
                                    existing_auditor_comment = row.get('Réponse de l\'auditeur', '') if pd.notna(row.get('Réponse de l\'auditeur', '')) else ''
                                    comment_auditor_work = st.text_area(
                                        "💬 Réponse de l'auditeur:",
                                        value=existing_auditor_comment,
                                        key=comment_auditor_key_work,
                                        height=80,
                                        placeholder="Modifiez la réponse de l'auditeur..."
                                    )
                                    comments_work_auditor_dict[row['Num']] = comment_auditor_work
                        
                        # Update DataFrame with modified comments
                        for index, row in filtered_work_checklist.iterrows():
                            if row['Num'] in comments_work_reviewer_dict:
                                work_checklist_df.loc[work_checklist_df['Num'] == row['Num'], 'Commentaire du reviewer'] = comments_work_reviewer_dict[row['Num']]
                            if row['Num'] in comments_work_auditor_dict:
                                work_checklist_df.loc[work_checklist_df['Num'] == row['Num'], 'Réponse de l\'auditeur'] = comments_work_auditor_dict[row['Num']]
                        
                        st.session_state['edited_checklist_work'] = work_checklist_df
                    else:
                        st.warning("Aucune exigence trouvée avec ces filtres.")
                
                elif tab_work == "Non-conformities" and 'nc' in work_data:
                    st.subheader("❗ Non-conformités (travail repris)")
                    
                    # Ensure consistent DataFrame structure
                    work_nc_df = work_data['nc'].copy()
                    
                    # Verify column structure for reviewer/auditor communication
                    required_nc_columns = [
                        'Commentaire du reviewer', 'Réponse de l\'auditeur', 'Plan d\'action proposé', 
                        'Actions mises en place', 'Date limite', 'Responsable', 'Statut'
                    ]
                    
                    for col in required_nc_columns:
                        if col not in work_nc_df.columns:
                            if col == "Statut":
                                work_nc_df[col] = "En attente"
                            else:
                                work_nc_df[col] = ""
                    
                    # Handle legacy format (convert old "Commentaire" and "Plan d'action")
                    if "Commentaire" in work_nc_df.columns and "Commentaire du reviewer" not in work_nc_df.columns:
                        work_nc_df["Commentaire du reviewer"] = work_nc_df['Commentaire']
                    if "Plan d'action" in work_nc_df.columns and "Plan d'action proposé" not in work_nc_df.columns:
                        work_nc_df["Plan d'action proposé"] = work_nc_df["Plan d'action"]
                    
                    if len(work_nc_df) > 0:
                        st.warning(f"**{len(work_nc_df)} non-conformité(s) en cours de traitement** (scores B, C, D uniquement)")
                        
                        # SAME DETAILED VIEW as normal NC tab with reviewer/auditor communication
                        nc_work_reviewer_comments = {}
                        nc_work_auditor_responses = {}
                        nc_work_action_plans = {}
                        nc_work_implemented = {}
                        nc_work_deadlines = {}
                        nc_work_responsibles = {}
                        nc_work_status = {}
                        
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
                                        st.markdown(f"**💬 Réponse système:** {row['Réponse']}")
                                
                                with col_nc2:
                                    st.markdown(f"**N°:** {row['Num']}")
                                    st.markdown(f"**Score:** :color[{score_color}][{row['Score']}]")
                                    if 'Chapitre' in row:
                                        st.markdown(f"**Chapitre:** {row['Chapitre']}")
                                
                                # Communication reviewer/auditeur avec valeurs existantes
                                st.markdown("---")
                                col_comm1, col_comm2 = st.columns(2)
                                
                                with col_comm1:
                                    # Commentaire reviewer
                                    comment_reviewer_key = f"nc_work_reviewer_{row['Num']}"
                                    existing_reviewer_comment = row.get('Commentaire du reviewer', '') if pd.notna(row.get('Commentaire du reviewer', '')) else ''
                                    reviewer_comment = st.text_area(
                                        "📝 Commentaire du reviewer:",
                                        value=existing_reviewer_comment,
                                        key=comment_reviewer_key,
                                        height=100,
                                        placeholder="Observations du reviewer..."
                                    )
                                    nc_work_reviewer_comments[row['Num']] = reviewer_comment
                                
                                with col_comm2:
                                    # Réponse auditeur
                                    auditor_response_key = f"nc_work_auditor_{row['Num']}"
                                    existing_auditor_response = row.get('Réponse de l\'auditeur', '') if pd.notna(row.get('Réponse de l\'auditeur', '')) else ''
                                    auditor_response = st.text_area(
                                        "💬 Réponse de l'auditeur:",
                                        value=existing_auditor_response,
                                        key=auditor_response_key,
                                        height=100,
                                        placeholder="Réponse de l'auditeur..."
                                    )
                                    nc_work_auditor_responses[row['Num']] = auditor_response
                                
                                # Plan d'action et suivi
                                st.markdown("**🎯 Plan d'action et suivi**")
                                col_action1, col_action2 = st.columns(2)
                                
                                with col_action1:
                                    # Plan d'action proposé
                                    action_plan_key = f"nc_work_plan_{row['Num']}"
                                    existing_plan = row.get('Plan d\'action proposé', '') if pd.notna(row.get('Plan d\'action proposé', '')) else ''
                                    action_plan = st.text_area(
                                        "📋 Plan d'action proposé:",
                                        value=existing_plan,
                                        key=action_plan_key,
                                        height=80,
                                        placeholder="Plan d'action proposé..."
                                    )
                                    nc_work_action_plans[row['Num']] = action_plan
                                
                                with col_action2:
                                    # Actions mises en place
                                    implemented_key = f"nc_work_implemented_{row['Num']}"
                                    existing_implemented = row.get('Actions mises en place', '') if pd.notna(row.get('Actions mises en place', '')) else ''
                                    implemented = st.text_area(
                                        "✅ Actions mises en place:",
                                        value=existing_implemented,
                                        key=implemented_key,
                                        height=80,
                                        placeholder="Actions réalisées..."
                                    )
                                    nc_work_implemented[row['Num']] = implemented
                                
                                # Suivi
                                col_track1, col_track2, col_track3 = st.columns(3)
                                
                                with col_track1:
                                    deadline_key = f"nc_work_deadline_{row['Num']}"
                                    existing_deadline = row.get('Date limite', '')
                                    if existing_deadline and existing_deadline != '':
                                        try:
                                            deadline_value = pd.to_datetime(existing_deadline).date()
                                        except:
                                            deadline_value = None
                                    else:
                                        deadline_value = None
                                    
                                    deadline = st.date_input(
                                        "📅 Date limite:",
                                        value=deadline_value,
                                        key=deadline_key
                                    )
                                    nc_work_deadlines[row['Num']] = deadline.strftime("%Y-%m-%d") if deadline else ""
                                
                                with col_track2:
                                    responsible_key = f"nc_work_responsible_{row['Num']}"
                                    existing_responsible = row.get('Responsable', '') if pd.notna(row.get('Responsable', '')) else ''
                                    responsible = st.text_input(
                                        "👤 Responsable:",
                                        value=existing_responsible,
                                        key=responsible_key,
                                        placeholder="Nom du responsable"
                                    )
                                    nc_work_responsibles[row['Num']] = responsible
                                
                                with col_track3:
                                    status_key = f"nc_work_status_{row['Num']}"
                                    existing_status = row.get('Statut', 'En attente') if pd.notna(row.get('Statut', '')) else 'En attente'
                                    status_options = ["En attente", "En cours", "Terminé", "Reporté", "Annulé"]
                                    try:
                                        status_index = status_options.index(existing_status)
                                    except:
                                        status_index = 0
                                    
                                    status = st.selectbox(
                                        "📈 Statut:",
                                        status_options,
                                        index=status_index,
                                        key=status_key
                                    )
                                    nc_work_status[row['Num']] = status
                        
                        # Update DataFrame with all modifications
                        for index, row in work_nc_df.iterrows():
                            if row['Num'] in nc_work_reviewer_comments:
                                work_nc_df.loc[index, 'Commentaire du reviewer'] = nc_work_reviewer_comments[row['Num']]
                            if row['Num'] in nc_work_auditor_responses:
                                work_nc_df.loc[index, 'Réponse de l\'auditeur'] = nc_work_auditor_responses[row['Num']]
                            if row['Num'] in nc_work_action_plans:
                                work_nc_df.loc[index, 'Plan d\'action proposé'] = nc_work_action_plans[row['Num']]
                            if row['Num'] in nc_work_implemented:
                                work_nc_df.loc[index, 'Actions mises en place'] = nc_work_implemented[row['Num']]
                            if row['Num'] in nc_work_deadlines:
                                work_nc_df.loc[index, 'Date limite'] = nc_work_deadlines[row['Num']]
                            if row['Num'] in nc_work_responsibles:
                                work_nc_df.loc[index, 'Responsable'] = nc_work_responsibles[row['Num']]
                            if row['Num'] in nc_work_status:
                                work_nc_df.loc[index, 'Statut'] = nc_work_status[row['Num']]
                        
                        st.session_state['edited_nc_work'] = work_nc_df
                    else:
                        st.success("🎉 Aucune non-conformité dans ce travail sauvegardé !")
                
                # Export updated work - BUTTONS IN SIDEBAR
                st.sidebar.markdown("---")
                st.sidebar.markdown("### 💾 Mise à jour travail")
                
                # Sauvegarde mise à jour
                st.sidebar.write("**💾 Sauvegarde mise à jour**")
                if st.sidebar.button("💾 Générer nouvelle sauvegarde", type="secondary", key="save_updated_work_btn"):
                    with st.spinner("Génération de la sauvegarde mise à jour..."):
                        output = BytesIO()
                        
                        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
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
                    
                    st.sidebar.success("💾 Sauvegarde mise à jour générée !")
                    st.sidebar.download_button(
                        label="📥 Télécharger sauvegarde",
                        data=output,
                        file_name=f"travail_IFS_mis_a_jour_{date_str}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_updated_work_btn"
                    )
                
                # Export rapport final depuis travail repris - IMPLÉMENTATION COMPLÈTE
                st.sidebar.write("**📊 Rapport final**")
                if st.sidebar.button("📊 Générer rapport final", type="primary", key="export_from_work_btn"):
                    with st.spinner("Génération du rapport final depuis travail repris..."):
                        # Prepare DataFrames for final report from work data
                        if 'edited_profile_work' in st.session_state:
                            profile_export_df = st.session_state['edited_profile_work']
                        elif 'profile' in work_data:
                            profile_export_df = work_data['profile'].copy()
                            # Ensure communication columns exist
                            if "Commentaire du reviewer" not in profile_export_df.columns:
                                profile_export_df["Commentaire du reviewer"] = ""
                            if "Réponse de l'auditeur" not in profile_export_df.columns:
                                profile_export_df["Réponse de l'auditeur"] = ""
                        else:
                            profile_export_df = pd.DataFrame()
                        
                        if 'edited_checklist_work' in st.session_state:
                            checklist_export_df = st.session_state['edited_checklist_work']
                        elif 'checklist' in work_data:
                            checklist_export_df = work_data['checklist'].copy()
                            # Ensure communication columns exist
                            if "Commentaire du reviewer" not in checklist_export_df.columns:
                                checklist_export_df["Commentaire du reviewer"] = ""
                            if "Réponse de l'auditeur" not in checklist_export_df.columns:
                                checklist_export_df["Réponse de l'auditeur"] = ""
                        else:
                            checklist_export_df = pd.DataFrame()
                        
                        if 'edited_nc_work' in st.session_state:
                            nc_export_df = st.session_state['edited_nc_work']
                        elif 'nc' in work_data:
                            nc_export_df = work_data['nc'].copy()
                            # Ensure all required columns exist
                            required_nc_columns = [
                                'Commentaire du reviewer', 'Réponse de l\'auditeur', 'Plan d\'action proposé', 
                                'Actions mises en place', 'Date limite', 'Responsable', 'Statut'
                            ]
                            for col in required_nc_columns:
                                if col not in nc_export_df.columns:
                                    if col == "Statut":
                                        nc_export_df[col] = "En attente"
                                    else:
                                        nc_export_df[col] = ""
                        else:
                            nc_export_df = pd.DataFrame()
                        
                        # Create final report using the enhanced function
                        final_report = create_final_report_excel(profile_export_df, checklist_export_df, nc_export_df, is_from_work=True, work_data=work_data)
                    
                    st.sidebar.success("📊 Rapport final généré depuis travail repris !")
                    
                    # Create filename - try to extract company info from work data
                    try:
                        if 'profile' in work_data and not work_data['profile'].empty:
                            if 'Champ' in work_data['profile'].columns and 'Valeur' in work_data['profile'].columns:
                                profile_dict = work_data['profile'].set_index('Champ')['Valeur'].to_dict()
                                company_name = profile_dict.get("Nom du site à auditer", "entreprise_reprise")
                                coid = profile_dict.get("N° COID du portail", "coid_repris")
                            else:
                                company_name = "entreprise_reprise"
                                coid = "travail_repris"
                        else:
                            company_name = "entreprise_reprise"
                            coid = "travail_repris"
                        
                        # Clean company name for filename
                        clean_company_name = re.sub(r'[^\w\s-]', '', str(company_name)).strip()
                        clean_company_name = re.sub(r'[-\s]+', '_', clean_company_name)
                        
                        final_filename = f'rapport_final_IFS_{coid}_{clean_company_name}_repris.xlsx'
                    except:
                        # Fallback filename
                        date_str = datetime.now().strftime("%Y%m%d_%H%M")
                        final_filename = f'rapport_final_IFS_travail_repris_{date_str}.xlsx'
                    
                    # Provide the download button
                    st.sidebar.download_button(
                        label="📥 Télécharger rapport final",
                        data=final_report,
                        file_name=final_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        help="Rapport final complet avec résumé d'audit, formatage professionnel et suivi des actions",
                        key="download_final_from_work_btn"
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
            - **Interface identique** à l'audit normal
            - **Export rapport final** directement depuis le travail repris
            """)

    # ================================
    # GESTION DES FICHIERS EXCEL
    # ================================
    elif main_option == "Gestion des fichiers Excel":
        st.subheader("📊 Gestion des fichiers Excel")
        st.info("Chargez et éditez des fichiers Excel existants pour les compléter ou les modifier.")
        
        uploaded_excel_file = st.file_uploader("Charger le fichier Excel pour compléter les commentaires", type="xlsx")

        if uploaded_excel_file:
            try:
                # Load the uploaded Excel file
                with st.spinner("Chargement du fichier Excel..."):
                    excel_data = pd.read_excel(uploaded_excel_file, sheet_name=None)

                st.success(f"✅ Fichier Excel chargé avec {len(excel_data)} feuille(s)")

                # Display the Excel content for editing
                st.subheader("Contenu du fichier Excel")
                sheet = st.selectbox("Sélectionnez une feuille:", list(excel_data.keys()))
                df = excel_data[sheet]
                
                # Show sheet info
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("Lignes", len(df))
                with col2:
                    st.metric("Colonnes", len(df.columns))

                # Display the DataFrame
                st.write(f"**Édition de la feuille : {sheet}**")
                edited_df = st.data_editor(
                    df, 
                    num_rows="dynamic", 
                    use_container_width=True,
                    height=600
                )

                # Save and provide a download link for the edited Excel file
                col1, col2 = st.columns(2)
                with col1:
                    if st.button("💾 Enregistrer les modifications", type="primary"):
                        with st.spinner("Sauvegarde des modifications..."):
                            output = BytesIO()
                            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                                # Update the specific sheet with edited data
                                excel_data[sheet] = edited_df
                                
                                # Write all sheets to the new file
                                for sheet_name, df_sheet in excel_data.items():
                                    df_sheet.to_excel(writer, index=False, sheet_name=sheet_name)
                                    
                                    # Apply basic formatting
                                    workbook = writer.book
                                    worksheet = writer.sheets[sheet_name]
                                    
                                    # Header formatting
                                    header_format = workbook.add_format({
                                        'bold': True,
                                        'text_wrap': True,
                                        'valign': 'top',
                                        'fg_color': '#D7E4BC',
                                        'border': 1
                                    })
                                    
                                    # Apply header format
                                    for col_num, value in enumerate(df_sheet.columns.values):
                                        worksheet.write(0, col_num, value, header_format)
                                        
                                        # Auto-adjust column width
                                        worksheet.set_column(col_num, col_num, 20)
                            
                            output.seek(0)
                        
                        st.success("✅ Modifications sauvegardées!")
                        
                        # Generate filename
                        original_name = uploaded_excel_file.name.replace('.xlsx', '')
                        date_str = datetime.now().strftime("%Y%m%d_%H%M")
                        modified_filename = f"{original_name}_modifie_{date_str}.xlsx"
                        
                        st.download_button(
                            label="📥 Télécharger le fichier Excel modifié",
                            data=output,
                            file_name=modified_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            help="Fichier Excel avec vos modifications appliquées"
                        )

                with col2:
                    if st.button("🔄 Réinitialiser", type="secondary"):
                        st.rerun()

            except Exception as e:
                st.error(f"❌ Erreur lors du traitement du fichier Excel: {e}")
        else:
            st.markdown("""
            ### 📊 Fonctionnalités de gestion Excel :
            
            ✅ **Ce que vous pouvez faire :**
            - Charger n'importe quel fichier Excel (.xlsx)
            - Éditer le contenu des cellules directement
            - Ajouter ou supprimer des lignes
            - Naviguer entre les différentes feuilles
            - Sauvegarder avec formatage professionnel
            - Télécharger le fichier modifié
            
            💡 **Idéal pour :**
            - Compléter des rapports d'audit existants
            - Modifier des templates IFS
            - Ajouter des commentaires à des fichiers de travail
            - Corriger des données avant finalisation
            """)
    
    # Footer
    st.sidebar.markdown("---")
    st.sidebar.markdown("""
    <div style="text-align: center; color: #666; font-size: 12px;">
        <p>🔍 <strong>IFS NEO Data Extractor</strong></p>
        <p>Application spécialisée pour audits IFS</p>
        <p>Communication reviewer ↔ auditeur</p>
    </div>
    """, unsafe_allow_html=True)

# ================================
# EXÉCUTION PRINCIPALE
# ================================
if __name__ == "__main__":
    main()
