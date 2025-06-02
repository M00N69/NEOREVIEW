import json
import pandas as pd
import streamlit as st
from io import BytesIO
import requests
from datetime import datetime

# Configuration Streamlit
st.set_page_config(
    page_title="IFS NEO Data Extractor",
    layout="wide"
)

# CSS personnalisé pour les tableaux
def apply_table_css():
    st.markdown(
        """
        <style>
        .main-header {
            font-size: 2.5rem;
            color: #1f77b4;
            text-align: center;
            margin-bottom: 2rem;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            background-color: #f9f9f9;
            margin: 10px 0;
        }
        th, td {
            border: 1px solid #ddd;
            padding: 12px;
            text-align: left;
            vertical-align: top;
        }
        th {
            background-color: #2e8b57;
            color: white;
            font-weight: bold;
        }
        tr:nth-child(even) {
            background-color: #f2f2f2;
        }
        tr:hover {
            background-color: #e8f4f8;
        }
        .score-A { background-color: #d4edda !important; color: #155724; font-weight: bold; }
        .score-B { background-color: #fff3cd !important; color: #856404; font-weight: bold; }
        .score-C { background-color: #ffeaa7 !important; color: #856404; font-weight: bold; }
        .score-D { background-color: #f8d7da !important; color: #721c24; font-weight: bold; }
        .info-box {
            background-color: #e3f2fd;
            border-left: 4px solid #2196f3;
            padding: 15px;
            margin: 15px 0;
            border-radius: 4px;
        }
        .success-box {
            background-color: #e8f5e8;
            border-left: 4px solid #4caf50;
            padding: 15px;
            margin: 15px 0;
            border-radius: 4px;
        }
        .warning-box {
            background-color: #fff8e1;
            border-left: 4px solid #ff9800;
            padding: 15px;
            margin: 15px 0;
            border-radius: 4px;
        }
        </style>
        """, unsafe_allow_html=True
    )

# Fonction pour aplatir le JSON
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

# Fonction pour extraire les données du JSON aplati
def extract_from_flattened(flattened_data, mapping, selected_fields):
    extracted_data = {}
    for label, flat_path in mapping.items():
        if label in selected_fields:
            extracted_data[label] = flattened_data.get(flat_path, 'N/A')
    return extracted_data

# Charger le mapping UUID depuis l'URL
def load_uuid_mapping_from_url(url):
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

# URL pour le mapping UUID
UUID_MAPPING_URL = "https://raw.githubusercontent.com/M00N69/Gemini-Knowledge/refs/heads/main/IFSV8listUUID.csv"

# Mapping des champs
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

def create_profile_table(profile_data):
    """Crée un tableau HTML pour les données de profil."""
    table_html = """
    <table>
    <thead>
        <tr>
            <th style="width: 40%;">Champ</th>
            <th style="width: 60%;">Valeur</th>
        </tr>
    </thead>
    <tbody>
    """
    
    for field, value in profile_data.items():
        # Escape HTML characters
        value_str = str(value).replace('<', '&lt;').replace('>', '&gt;')
        table_html += f"""
        <tr>
            <td><strong>{field}</strong></td>
            <td>{value_str}</td>
        </tr>
        """
    
    table_html += "</tbody></table>"
    return table_html

def create_checklist_table(checklist_data, show_filters=True):
    """Crée un tableau HTML pour les données de checklist avec filtres optionnels."""
    
    if show_filters:
        # Filtres
        col1, col2, col3 = st.columns(3)
        with col1:
            score_filter = st.selectbox(
                "Filtrer par score:",
                ["Tous", "A", "B", "C", "D", "Non applicable"],
                key="checklist_score_filter"
            )
        
        with col2:
            search_term = st.text_input(
                "Rechercher dans les exigences:",
                key="checklist_search"
            )
        
        with col3:
            show_responses = st.checkbox("Afficher les réponses", value=True, key="show_responses")
        
        # Appliquer les filtres
        filtered_data = checklist_data.copy()
        if score_filter != "Tous":
            filtered_data = [item for item in filtered_data if item['Score'] == score_filter]
        
        if search_term:
            filtered_data = [item for item in filtered_data 
                           if search_term.lower() in item['Explanation'].lower() 
                           or search_term.lower() in item['Detailed Explanation'].lower()
                           or search_term.lower() in str(item['Num']).lower()]
        
        st.info(f"Affichage de {len(filtered_data)} exigences sur {len(checklist_data)} au total")
    else:
        filtered_data = checklist_data
        show_responses = True
    
    # Créer le tableau
    table_html = """
    <table>
    <thead>
        <tr>
            <th style="width: 8%;">N°</th>
            <th style="width: 8%;">Score</th>
            <th style="width: 30%;">Explication</th>
            <th style="width: 30%;">Explication détaillée</th>
    """
    
    if show_responses:
        table_html += '<th style="width: 24%;">Réponse</th>'
    
    table_html += """
        </tr>
    </thead>
    <tbody>
    """
    
    for item in filtered_data:
        score_class = f"score-{item['Score']}" if item['Score'] in ['A', 'B', 'C', 'D'] else ""
        
        explanation = str(item['Explanation']).replace('<', '&lt;').replace('>', '&gt;')
        detailed_explanation = str(item['Detailed Explanation']).replace('<', '&lt;').replace('>', '&gt;')
        response = str(item['Response']).replace('<', '&lt;').replace('>', '&gt;')
        
        table_html += f"""
        <tr>
            <td><strong>{item['Num']}</strong></td>
            <td class="{score_class}">{item['Score']}</td>
            <td>{explanation}</td>
            <td>{detailed_explanation}</td>
        """
        
        if show_responses:
            table_html += f"<td>{response}</td>"
        
        table_html += "</tr>"
    
    table_html += "</tbody></table>"
    return table_html

def create_excel_export(profile_data, checklist_data, non_conformities, coid):
    """Crée un fichier Excel avec toutes les données."""
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Onglet Profil
        profile_df = pd.DataFrame([
            {"Champ": k, "Valeur": v} for k, v in profile_data.items()
        ])
        profile_df.to_excel(writer, index=False, sheet_name="Profil")
        
        # Onglet Checklist complète
        checklist_df = pd.DataFrame(checklist_data)
        checklist_df.to_excel(writer, index=False, sheet_name="Checklist")
        
        # Onglet Non-conformités
        if non_conformities:
            nc_df = pd.DataFrame(non_conformities)
            nc_df.to_excel(writer, index=False, sheet_name="Non-conformités")
        
        # Onglet Statistiques
        stats_data = []
        if checklist_data:
            total = len(checklist_data)
            scores_count = {}
            for item in checklist_data:
                score = item['Score']
                scores_count[score] = scores_count.get(score, 0) + 1
            
            stats_data = [
                {"Indicateur": "Total exigences", "Valeur": total},
                {"Indicateur": "Score A", "Valeur": scores_count.get('A', 0)},
                {"Indicateur": "Score B", "Valeur": scores_count.get('B', 0)},
                {"Indicateur": "Score C", "Valeur": scores_count.get('C', 0)},
                {"Indicateur": "Score D", "Valeur": scores_count.get('D', 0)},
                {"Indicateur": "Taux conformité (%)", 
                 "Valeur": round((scores_count.get('A', 0) / total) * 100, 2) if total > 0 else 0}
            ]
        
        stats_df = pd.DataFrame(stats_data)
        stats_df.to_excel(writer, index=False, sheet_name="Statistiques")
        
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
    # Appliquer le CSS
    apply_table_css()
    
    # En-tête
    st.markdown('<div class="main-header">🔍 IFS NEO Data Extractor</div>', unsafe_allow_html=True)
    
    # Upload du fichier
    uploaded_json_file = st.file_uploader(
        "📁 Charger le fichier IFS de NEO", 
        type="ifs",
        help="Sélectionnez le fichier d'audit IFS (.ifs) exporté depuis NEO"
    )
    
    if uploaded_json_file:
        try:
            # Charger et traiter le fichier JSON
            json_data = json.load(uploaded_json_file)
            flattened_json_data = flatten_json_safe(json_data)
            
            # Extraire les données de profil
            profile_data = extract_from_flattened(
                flattened_json_data, 
                FLATTENED_FIELD_MAPPING, 
                list(FLATTENED_FIELD_MAPPING.keys())
            )
            
            # Charger le mapping UUID
            with st.spinner("Chargement du mapping des exigences..."):
                uuid_mapping_df = load_uuid_mapping_from_url(UUID_MAPPING_URL)
            
            # Extraire les données de checklist
            checklist_data = []
            if not uuid_mapping_df.empty:
                for _, row in uuid_mapping_df.iterrows():
                    uuid = row['UUID']
                    prefix = f"data_modules_food_8_checklists_checklistFood8_resultScorings_{uuid}"
                    
                    explanation_text = flattened_json_data.get(f"{prefix}_answers_englishExplanationText", "N/A")
                    detailed_explanation = flattened_json_data.get(f"{prefix}_answers_explanationText", "N/A")
                    score_label = flattened_json_data.get(f"{prefix}_score_label", "N/A")
                    response = flattened_json_data.get(f"{prefix}_answers_fieldAnswers", "N/A")
                    
                    checklist_data.append({
                        "Num": row['Num'],
                        "Chapitre": row['Chapitre'],
                        "Theme": row['Theme'],
                        "SSTheme": row['SSTheme'],
                        "Explanation": explanation_text,
                        "Detailed Explanation": detailed_explanation,
                        "Score": score_label,
                        "Response": response
                    })
            
            # Extraire les non-conformités
            non_conformities = [item for item in checklist_data if item['Score'] not in ['A', 'N/A']]
            
            # Interface à onglets
            tab1, tab2, tab3, tab4 = st.tabs([
                "📋 Profil de l'entreprise", 
                "✅ Exigences de la checklist", 
                "⚠️ Non-conformités",
                "📊 Export & Statistiques"
            ])
            
            with tab1:
                st.markdown("### 📋 Informations sur l'entreprise auditée")
                
                if profile_data:
                    # Afficher les informations importantes en haut
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Nom du site", profile_data.get("Nom du site à auditer", "N/A"))
                    with col2:
                        st.metric("COID", profile_data.get("N° COID du portail", "N/A"))
                    with col3:
                        st.metric("Ville", profile_data.get("Nom de la ville", "N/A"))
                    
                    # Tableau complet des données
                    st.markdown("#### Données complètes du profil")
                    profile_table = create_profile_table(profile_data)
                    st.markdown(profile_table, unsafe_allow_html=True)
                else:
                    st.warning("Aucune donnée de profil trouvée dans le fichier.")
            
            with tab2:
                st.markdown("### ✅ Exigences de la checklist IFS")
                
                if checklist_data:
                    # Statistiques rapides
                    total_req = len(checklist_data)
                    scores_count = {}
                    for item in checklist_data:
                        score = item['Score']
                        scores_count[score] = scores_count.get(score, 0) + 1
                    
                    col1, col2, col3, col4, col5 = st.columns(5)
                    with col1:
                        st.metric("Total", total_req)
                    with col2:
                        st.metric("Score A", scores_count.get('A', 0), 
                                delta_color="normal")
                    with col3:
                        st.metric("Score B", scores_count.get('B', 0),
                                delta_color="off")
                    with col4:
                        st.metric("Score C", scores_count.get('C', 0),
                                delta_color="off")
                    with col5:
                        st.metric("Score D", scores_count.get('D', 0),
                                delta_color="inverse")
                    
                    # Filtres par chapitre/thème
                    if not uuid_mapping_df.empty:
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            chapitre_options = ["Tous"] + sorted(uuid_mapping_df['Chapitre'].dropna().unique())
                            chapitre_filter = st.selectbox("Filtrer par Chapitre", options=chapitre_options)
                        
                        filtered_mapping = uuid_mapping_df
                        if chapitre_filter != "Tous":
                            filtered_mapping = filtered_mapping[filtered_mapping['Chapitre'] == chapitre_filter]
                        
                        with col2:
                            theme_options = ["Tous"] + sorted(filtered_mapping['Theme'].dropna().unique())
                            theme_filter = st.selectbox("Filtrer par Thème", options=theme_options)
                        
                        if theme_filter != "Tous":
                            filtered_mapping = filtered_mapping[filtered_mapping['Theme'] == theme_filter]
                        
                        with col3:
                            sstheme_options = ["Tous"] + sorted(filtered_mapping['SSTheme'].dropna().unique())
                            sstheme_filter = st.selectbox("Filtrer par Sous-Thème", options=sstheme_options)
                        
                        # Appliquer les filtres de chapitre/thème
                        filtered_checklist = checklist_data
                        if chapitre_filter != "Tous":
                            filtered_checklist = [item for item in filtered_checklist if item['Chapitre'] == chapitre_filter]
                        if theme_filter != "Tous":
                            filtered_checklist = [item for item in filtered_checklist if item['Theme'] == theme_filter]
                        if sstheme_filter != "Tous":
                            filtered_checklist = [item for item in filtered_checklist if item['SSTheme'] == sstheme_filter]
                    else:
                        filtered_checklist = checklist_data
                    
                    # Tableau des exigences
                    checklist_table = create_checklist_table(filtered_checklist, show_filters=True)
                    st.markdown(checklist_table, unsafe_allow_html=True)
                else:
                    st.warning("Aucune donnée de checklist trouvée.")
            
            with tab3:
                st.markdown("### ⚠️ Non-conformités détectées")
                
                if non_conformities:
                    st.error(f"**{len(non_conformities)} non-conformité(s) détectée(s)**")
                    
                    # Répartition par score
                    nc_scores = {}
                    for nc in non_conformities:
                        score = nc['Score']
                        nc_scores[score] = nc_scores.get(score, 0) + 1
                    
                    if nc_scores:
                        cols = st.columns(len(nc_scores))
                        for i, (score, count) in enumerate(nc_scores.items()):
                            with cols[i]:
                                color = {"B": "🟡", "C": "🟠", "D": "🔴"}.get(score, "⚫")
                                st.metric(f"{color} Score {score}", count)
                    
                    # Tableau des non-conformités
                    nc_table = create_checklist_table(non_conformities, show_filters=False)
                    st.markdown(nc_table, unsafe_allow_html=True)
                    
                    st.markdown("""
                    <div class="warning-box">
                    <strong>Actions requises :</strong><br>
                    • Analyser chaque non-conformité<br>
                    • Établir un plan d'action corrective<br>
                    • Définir les responsabilités et échéances<br>
                    • Programmer le suivi des actions
                    </div>
                    """, unsafe_allow_html=True)
                    
                else:
                    st.markdown("""
                    <div class="success-box">
                    🎉 <strong>Félicitations !</strong><br>
                    Aucune non-conformité détectée. Toutes les exigences sont conformes (Score A).
                    </div>
                    """, unsafe_allow_html=True)
            
            with tab4:
                st.markdown("### 📊 Statistiques et Export")
                
                # Statistiques globales
                if checklist_data:
                    total = len(checklist_data)
                    conformes = len([x for x in checklist_data if x['Score'] == 'A'])
                    taux_conformite = (conformes / total * 100) if total > 0 else 0
                    
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.markdown("#### 📈 Indicateurs de performance")
                        st.metric("Taux de conformité", f"{taux_conformite:.1f}%")
                        st.metric("Exigences conformes", f"{conformes}/{total}")
                        st.metric("Non-conformités", len(non_conformities))
                    
                    with col2:
                        st.markdown("#### 📊 Répartition des scores")
                        scores_count = {}
                        for item in checklist_data:
                            score = item['Score']
                            scores_count[score] = scores_count.get(score, 0) + 1
                        
                        chart_data = pd.DataFrame(list(scores_count.items()), columns=['Score', 'Nombre'])
                        st.bar_chart(chart_data.set_index('Score'))
                
                # Export Excel
                st.markdown("#### 📄 Export des données")
                st.info("Générez un rapport Excel complet avec toutes les données extraites.")
                
                if st.button("🔄 Générer le rapport Excel", type="primary"):
                    with st.spinner("Génération du fichier Excel..."):
                        coid = profile_data.get("N° COID du portail", "inconnu")
                        excel_file = create_excel_export(profile_data, checklist_data, non_conformities, coid)
                        
                        date_str = datetime.now().strftime("%Y%m%d_%H%M")
                        filename = f"rapport_IFS_{coid}_{date_str}.xlsx"
                        
                        st.download_button(
                            label="📥 Télécharger le rapport Excel",
                            data=excel_file,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        
                        st.success("✅ Rapport Excel généré avec succès!")
                        
                        # Contenu du fichier
                        st.markdown("""
                        **Contenu du rapport :**
                        - 📋 **Profil** : Informations complètes sur l'entreprise
                        - ✅ **Checklist** : Toutes les exigences avec scores et réponses
                        - ⚠️ **Non-conformités** : Liste détaillée des points à corriger
                        - 📊 **Statistiques** : Indicateurs de performance et taux de conformité
                        """)
        
        except json.JSONDecodeError:
            st.error("❌ Erreur lors du décodage du fichier JSON. Veuillez vérifier le format.")
        except Exception as e:
            st.error(f"❌ Erreur lors du traitement du fichier : {str(e)}")
    
    else:
        # Page d'accueil
        st.markdown("""
        <div class="info-box">
        <h3>🚀 Bienvenue dans l'extracteur de données IFS NEO</h3>
        <p>Cette application vous permet d'analyser facilement vos rapports d'audit IFS :</p>
        <ul>
            <li>📊 <strong>Extraction automatique</strong> des données d'entreprise et d'audit</li>
            <li>✅ <strong>Analyse complète</strong> de toutes les exigences IFS</li>
            <li>⚠️ <strong>Identification</strong> des non-conformités</li>
            <li>📄 <strong>Export Excel</strong> structuré pour vos rapports</li>
        </ul>
        
        <p><strong>Pour commencer :</strong> Chargez votre fichier d'audit IFS (.ifs) ci-dessus</p>
        </div>
        """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
