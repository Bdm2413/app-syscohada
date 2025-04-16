import streamlit as st
import pandas as pd
import io
from fpdf import FPDF
import base64

st.set_page_config(page_title="App SYSCOHADA", page_icon="🏳️‍🌈", layout="wide")

# Menu latéral
st.sidebar.success("Menu de navigation")
menu = st.sidebar.selectbox("", ["Import Fichier", "Plan de comptes", "Grand Livre", "Balance", "Bilan", "Compte de résultat", "Flux de trésorerie"])
st.title("📊 :rainbow[Etats financiers SYSCOHADA]")

# Initialisation session
if "data_loaded" not in st.session_state:
    st.session_state.data_loaded = False

# Import
if menu == "Import Fichier":
    uploaded_file = st.file_uploader("📂 Importer le fichier Excel contenant le plan comptable et le grand livre", type=["xlsx"])
    if uploaded_file:
        try:
            plan_df = pd.read_excel(uploaded_file, sheet_name="Plan de comptes", header=0, usecols="A:G")
            gl_df = pd.read_excel(uploaded_file, sheet_name="Grand Livre", header=0, usecols="A:J")
            gl_df['Date'] = pd.to_datetime(gl_df['Date'], errors='coerce').dt.strftime('%d/%m/%Y')
            
            # Nettoyage des colonnes
            gl_df.columns = gl_df.columns.str.strip()
            plan_df.columns = plan_df.columns.str.strip()

            # Standardisation des comptes
            gl_df['Compte'] = gl_df['Compte'].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
            plan_df['Compte'] = plan_df['Compte'].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)

            st.session_state.plan_df = plan_df
            st.session_state.gl_df = gl_df
            st.session_state.data_loaded = True
            st.success("✅ Fichier importé avec succès.")
        except Exception as e:
            st.error(f"❌ Erreur lors de la lecture du fichier : {e}")

# Plan de comptes
elif menu == "Plan de comptes":
    st.title("Plan de comptes")
    if not st.session_state.data_loaded:
        st.warning("📂 Veuillez importer un fichier d'abord.")
    else:
        st.subheader("📚 Plan de comptes - Liste des comptes")
        st.dataframe(st.session_state.plan_df, use_container_width=True)

# Grand Livre
elif menu == "Grand Livre":
    st.title("Grand Livre")
    if not st.session_state.data_loaded:
        st.warning("📂 Veuillez importer un fichier d'abord.")
    else:
        st.subheader("📚 Grand Livre - Écritures comptables")

        # Copier le DataFrame
        gl_df = st.session_state.gl_df.copy()

        # Convertir la colonne Date
        if "Date" in gl_df.columns:
            gl_df["Date"] = pd.to_datetime(gl_df["Date"], format='%d/%m/%Y', errors='coerce')
            gl_df["Date_formatee"] = gl_df["Date"].dt.strftime('%d/%m/%Y')
            gl_df["Année"] = gl_df["Date"].dt.year
            gl_df["Mois"] = gl_df["Date"].dt.strftime("%Y%m")

        # Convertir la colonne 'Année' en entier pour enlever la partie décimale
        gl_df["Année"] = gl_df["Année"].fillna(0).astype(int)

        # Remplir les valeurs manquantes dans Débit et Crédit
        for col in ["Débit", "Crédit"]:
            if col in gl_df.columns:
                gl_df[col] = pd.to_numeric(gl_df[col], errors="coerce").fillna(0)

        # Filtres
        st.sidebar.header("🧮 Filtres")

        journal_filter = st.sidebar.multiselect("Journal", options=gl_df["Journal"].dropna().unique())
        an_filter = st.sidebar.multiselect("AN", options=gl_df["AN"].dropna().unique())
        compte_filter = st.sidebar.multiselect("Compte", options=gl_df["Compte"].dropna().unique())
        annee_filter = st.sidebar.multiselect("Année", options=sorted(gl_df["Année"].dropna().unique()))
        mois_filter = st.sidebar.multiselect("Mois", options=sorted(gl_df["Mois"].dropna().unique()))

        if journal_filter:
            gl_df = gl_df[gl_df["Journal"].isin(journal_filter)]
        if an_filter:
            gl_df = gl_df[gl_df["AN"].isin(an_filter)]
        if compte_filter:
            gl_df = gl_df[gl_df["Compte"].isin(compte_filter)]
        if annee_filter:
            gl_df = gl_df[gl_df["Année"].isin(annee_filter)]
        if mois_filter:
            gl_df = gl_df[gl_df["Mois"].isin(mois_filter)]

        # Calculs
        total_debit = gl_df["Débit"].sum()
        total_credit = gl_df["Crédit"].sum()
        difference = total_debit - total_credit

        # Observation
        if difference == 0:
            interpretation = "RAS"
            bg_color = "#FF595E"
        elif difference > 0:
            interpretation = "Solde Débiteur"
            bg_color = "#FF595E"
        else:
            interpretation = "Solde Créditeur"
            bg_color = "#FF595E"

        # Styles
        styles = {
            "debit": "#1982C4",
            "credit": "#013026",
            "diff": "#162A2C",
            "obs": "#FF595E"
        }

        def format_int(val):
            return f"{int(val):,}".replace(",", " ")

        # Affichage des cartes
        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.markdown(f"""
                <div style="background-color:{styles['debit']}; padding:20px; border-radius:10px; text-align:center; height:110px;">
                    <div style="color:white; font-size:16px;">Total Débit</div>
                    <div style="color:white; font-size:24px; font-weight:bold; margin-top:10px;">{format_int(total_debit)}</div>
                </div>
            """, unsafe_allow_html=True)

        with col2:
            st.markdown(f"""
                <div style="background-color:{styles['credit']}; padding:20px; border-radius:10px; text-align:center; height:110px;">
                    <div style="color:white; font-size:16px;">Total Crédit</div>
                    <div style="color:white; font-size:24px; font-weight:bold; margin-top:10px;">{format_int(total_credit)}</div>
                </div>
            """, unsafe_allow_html=True)

        with col3:
            st.markdown(f"""
                <div style="background-color:{styles['diff']}; padding:20px; border-radius:10px; text-align:center; height:110px;">
                    <div style="color:white; font-size:16px;">Solde</div>
                    <div style="color:white; font-size:24px; font-weight:bold; margin-top:10px;">{format_int(difference)}</div>
                </div>
            """, unsafe_allow_html=True)

        with col4:
            st.markdown(f"""
                <div style="background-color:{styles['obs']}; padding:20px; border-radius:10px; text-align:center; height:110px;">
                    <div style="color:white; font-size:16px;">Observations</div>
                    <div style="color:white; font-size:20px; font-weight:bold; margin-top:10px;">{interpretation}</div>
                </div>
            """, unsafe_allow_html=True)

        # Espacement
        st.markdown("<br>", unsafe_allow_html=True)

        # Formater les colonnes Débit / Crédit pour affichage
        gl_df["Débit"] = gl_df["Débit"].apply(lambda x: format_int(x))
        gl_df["Crédit"] = gl_df["Crédit"].apply(lambda x: format_int(x))

        colonnes_affichage = ["Date_formatee", "Journal", "AN", "Compte", "Libellé", "Débit", "Crédit"]
        colonnes_presentes = [col for col in colonnes_affichage if col in gl_df.columns]

        # Tableau
        st.dataframe(gl_df[colonnes_presentes], use_container_width=True)

# Balance
elif menu == "Balance":
    st.title("Balance à 8 colonnes")
    if not st.session_state.data_loaded:
        st.warning("📂 Veuillez d'abord importer un fichier Excel via le menu 'Import Fichier'.")
    else:
        plan_df = st.session_state.plan_df
        gl_df = st.session_state.gl_df

        # Vérification colonne Année
        if "Année" not in gl_df.columns:
            st.warning("❗ La colonne 'Année' n'a pas été trouvée automatiquement.")
            annee_col = st.selectbox("Sélectionnez la colonne correspondant à l'année :", gl_df.columns)
        else:
            annee_col = "Année"

        # Nettoyage et conversion
        gl_df['Débit'] = pd.to_numeric(gl_df['Débit'], errors='coerce').fillna(0)
        gl_df['Crédit'] = pd.to_numeric(gl_df['Crédit'], errors='coerce').fillna(0)
        gl_df['AN'] = gl_df['AN'].fillna("NON")

        # Sidebar : Filtres
        annees = sorted([int(a) for a in gl_df[annee_col].dropna().unique()])
        annee_choisie = st.sidebar.selectbox("📅 Choisir l'année", annees)

        tableaux = sorted(plan_df['Tableau'].dropna().unique())
        tableaux_choisis = st.sidebar.multiselect("🏷️ Choisir les tableaux", tableaux, default=tableaux)

        classes = sorted(plan_df['Compte'].astype(str).str[0].unique())
        classes_choisies = st.sidebar.multiselect("🏷️ Choisir les classes de comptes", classes, default=classes)

        # Filtres appliqués
        comptes_classes = plan_df[plan_df['Compte'].astype(str).str[0].isin(classes_choisies)]
        comptes_tableaux = comptes_classes[comptes_classes['Tableau'].isin(tableaux_choisis)]
        gl_df = gl_df[gl_df['Compte'].isin(comptes_tableaux['Compte'])]

        gl_annee = gl_df[gl_df[annee_col] == annee_choisie]
        gl_si = gl_annee[gl_annee['AN'].str.upper() == 'OUI']
        gl_mouv = gl_annee[gl_annee['AN'].str.upper() != 'OUI']

        def aggregate(df, col_name):
            return df.groupby('Compte')[[col_name]].sum().rename(columns={col_name: col_name})

        si_debit = aggregate(gl_si, 'Débit')
        si_credit = aggregate(gl_si, 'Crédit')
        mouv_debit = aggregate(gl_mouv, 'Débit')
        mouv_credit = aggregate(gl_mouv, 'Crédit')

        balance = comptes_tableaux.set_index('Compte').copy()
        balance = balance.join(si_debit.rename(columns={"Débit": "SI Débit"}), how="left")
        balance = balance.join(si_credit.rename(columns={"Crédit": "SI Crédit"}), how="left")
        balance = balance.join(mouv_debit.rename(columns={"Débit": "Mouv Débit"}), how="left")
        balance = balance.join(mouv_credit.rename(columns={"Crédit": "Mouv Crédit"}), how="left")

        for col in ["SI Débit", "SI Crédit", "Mouv Débit", "Mouv Crédit"]:
            balance[col] = balance[col].fillna(0)

        balance["SF Débit"] = (balance["SI Débit"] + balance["Mouv Débit"] - balance["SI Crédit"] - balance["Mouv Crédit"]).apply(lambda x: x if x > 0 else 0)
        balance["SF Crédit"] = (balance["SI Crédit"] + balance["Mouv Crédit"] - balance["SI Débit"] - balance["Mouv Débit"]).apply(lambda x: x if x > 0 else 0)

        # Colonnes BD, BC, RD, RC doivent exister même si vides
        for col in ["BD", "BC", "RD", "RC"]:
            if col not in balance.columns:
                balance[col] = ""

        # Nouvelle colonne : Code Bilan
        balance["Code Bilan"] = balance.apply(
            lambda row: row["BD"] if row["Tableau"] == "Bilan" and row["SF Débit"] > 0 else
                        row["BC"] if row["Tableau"] == "Bilan" and row["SF Crédit"] > 0 else "", axis=1
        )

        # Nouvelle colonne : Code Résultat
        balance["Code Résultat"] = balance.apply(
            lambda row: row["RD"] if row["Tableau"] == "Résultat" and row["SF Débit"] > 0 else
                        row["RC"] if row["Tableau"] == "Résultat" and row["SF Crédit"] > 0 else "", axis=1
        )

        balance = balance.reset_index()

        colonnes = ["Compte", "Intitulé", "Tableau", "BD", "BC", "RD", "RC",
                    "SI Débit", "SI Crédit", "Mouv Débit", "Mouv Crédit", "SF Débit", "SF Crédit",
                    "Code Bilan", "Code Résultat"]

        # Totaux
        totaux = balance[colonnes[7:13]].sum(numeric_only=True).to_dict()
        total_row = {
            "Compte": "Total", "Intitulé": "", "Tableau": "", "BD": "", "BC": "", "RD": "", "RC": "",
            "Code Bilan": "", "Code Résultat": ""
        }
        total_row.update({col: round(totaux.get(col, 0), 2) for col in colonnes[7:13]})
        balance_with_total = balance[colonnes].copy()
        balance_with_total.loc[len(balance_with_total)] = total_row

        # Format montant
        def format_int(val):
            try:
                return f"{int(val):,}".replace(",", " ")
            except:
                return val

        for col in ["SI Débit", "SI Crédit", "Mouv Débit", "Mouv Crédit", "SF Débit", "SF Crédit"]:
            balance_with_total[col] = balance_with_total[col].apply(lambda x: format_int(x))

        # Affichage
        st.dataframe(balance_with_total, use_container_width=True)

        # Export Excel : toutes classes
        output_excel_all_classes = io.BytesIO()
        with pd.ExcelWriter(output_excel_all_classes, engine='xlsxwriter') as writer:
            balance_with_total.to_excel(writer, index=False, sheet_name='Balance_Toutes_Classes')

        st.download_button(
            label="📥 Télécharger Excel (toutes les classes)",
            data=output_excel_all_classes.getvalue(),
            file_name=f"balance_toutes_classes_{annee_choisie}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Export Excel : séparé par classes
        output_excel_separated_classes = io.BytesIO()
        with pd.ExcelWriter(output_excel_separated_classes, engine='xlsxwriter') as writer:
            for classe in classes_choisies:
                classe_df = balance_with_total[balance_with_total['Compte'].str.startswith(classe)]
                classe_df.to_excel(writer, index=False, sheet_name=f'Classe_{classe}')

        st.download_button(
            label="📥 Télécharger Excel (séparé par classes)",
            data=output_excel_separated_classes.getvalue(),
            file_name=f"balance_separee_classes_{annee_choisie}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        # À la fin du bloc "Balance"
        if "balance_par_annee" not in st.session_state:
            st.session_state.balance_par_annee = {}

        # On enregistre la balance de l'année sélectionnée dans la session
        st.session_state.balance_par_annee[annee_choisie] = balance_with_total.copy()


# Structures des bilans
structure_bilan_actif = [
    {"code": "AD", "libelle": "IMMOBILISATIONS INCORPORELLES"},
    {"code": "AE", "libelle": "Frais de développement et de prospection"},
    {"code": "AF", "libelle": "Brevets, licences, logiciels et droits similaires"},
    {"code": "AG", "libelle": "Fonds commercial et droit au bail"},
    {"code": "AH", "libelle": "Autres immobilisations incorporelles"},
    {"code": "AI", "libelle": "IMMOBILISATIONS CORPORELLES"},
    {"code": "AJ", "libelle": "Terrains"},
    {"code": "AK", "libelle": "Bâtiments"},
    {"code": "AL", "libelle": "Aménagements, agencements et installations"},
    {"code": "AM", "libelle": "Matériel, mobilier et actifs biologiques"},
    {"code": "AN", "libelle": "Matériel de transport"},
    {"code": "AP", "libelle": "AVANCES ET ACOMPTES VERSES SUR IMMOBILISATIONS"},
    {"code": "AQ", "libelle": "IMMOBILISATIONS FINANCIÈRES"},
    {"code": "AR", "libelle": "Titres de participation"},
    {"code": "AS", "libelle": "Autres immobilisations financières"},
    {"code": "AZ", "libelle": "TOTAL ACTIF IMMOBILISÉ"},
    {"code": "BA", "libelle": "ACTIF CIRCULANT HAO"},
    {"code": "BB", "libelle": "STOCKS ET ENCOURS"},
    {"code": "BG", "libelle": "CRÉANCES ET EMPLOIS ASSIMILÉS"},
    {"code": "BH", "libelle": "Fournisseurs avances versées"},
    {"code": "BI", "libelle": "Clients"},
    {"code": "BJ", "libelle": "Autres créances"},
    {"code": "BK", "libelle": "TOTAL ACTIF CIRCULANT"},
    {"code": "BQ", "libelle": "Titres de placement"},
    {"code": "BR", "libelle": "Valeurs à encaisser"},
    {"code": "BS", "libelle": "Banques, chèques postaux, caisse et assimilés"},
    {"code": "BT", "libelle": "TOTAL TRÉSORERIE-ACTIF"},
    {"code": "BU", "libelle": "Écart de conversion-Actif"},
    {"code": "BZ", "libelle": "TOTAL ACTIF"}
]

structure_bilan_passif = [
    {"code": "CA", "libelle": "Capital"},
    {"code": "CB", "libelle": "Apporteurs capital non appelé (-)"},
    {"code": "CD", "libelle": "Primes liées au capital social"},
    {"code": "CE", "libelle": "Écarts de réévaluation"},
    {"code": "CF", "libelle": "Réserves indisponibles"},
    {"code": "CG", "libelle": "Réserves libres"},
    {"code": "CH", "libelle": "Report à nouveau (+ ou -)"},
    {"code": "CJ", "libelle": "Résultat net de l'exercice (bénéfice + ou perte -)"},
    {"code": "CL", "libelle": "Subventions d'investissement"},
    {"code": "CM", "libelle": "Provisions réglementées"},
    {"code": "CP", "libelle": "TOTAL CAPITAUX PROPRES ET RESSOURCES ASSIMILÉES"},
    {"code": "DA", "libelle": "Emprunts et dettes financières diverses"},
    {"code": "DB", "libelle": "Dettes de location-acquisition"},
    {"code": "DC", "libelle": "Provisions pour risques et charges"},
    {"code": "DD", "libelle": "TOTAL DETTES FINANCIÈRES ET RESSOURCES ASSIMILÉES"},
    {"code": "DF", "libelle": "TOTAL RESSOURCES STABLES"},
    {"code": "DH", "libelle": "Dettes circulantes HAO"},
    {"code": "DI", "libelle": "Clients, avances reçues"},
    {"code": "DJ", "libelle": "Fournisseurs d'exploitation"},
    {"code": "DK", "libelle": "Dettes fiscales et sociales"},
    {"code": "DM", "libelle": "Autres dettes"},
    {"code": "DN", "libelle": "Provisions pour risques et charges à court terme"},
    {"code": "DP", "libelle": "TOTAL PASSIF CIRCULANT"},
    {"code": "DQ", "libelle": "Banques, crédits d'escompte"},
    {"code": "DR", "libelle": "Banques, établissements financiers et crédits de trésorerie"},
    {"code": "DT", "libelle": "TOTAL TRÉSORERIE-PASSIF"},
    {"code": "DV", "libelle": "Écart de conversion-Passif"},
    {"code": "DZ", "libelle": "TOTAL PASSIF"}
]

def bilan():
    st.sidebar.subheader("Filtres Bilan")
    annees_balance = st.session_state.gl_df["Année"].dropna().unique().astype(int)
    annee_n = st.sidebar.selectbox("Année N", sorted(annees_balance, reverse=True))
    annee_n1 = st.sidebar.selectbox("Année N-1", sorted(annees_balance, reverse=True), index=1 if len(annees_balance) > 1 else 0)

    st.title("Bilan")

    # Fonction pour construire le bilan à partir de la structure et de la balance
    def build_bilan_df(structure, balance, annee, colonne_label):
        df = pd.DataFrame(structure)
        df[colonne_label] = ""

        for idx, row in df.iterrows():
            code = row["code"]
            # Filtrer les lignes de la balance correspondant à ce code bilan
            lignes_code = balance[balance["Code Bilan"] == code]

            # Convertir les colonnes SF en numérique (si ce n'est pas déjà fait)
            lignes_code["SF Débit"] = pd.to_numeric(lignes_code["SF Débit"], errors="coerce").fillna(0)
            lignes_code["SF Crédit"] = pd.to_numeric(lignes_code["SF Crédit"], errors="coerce").fillna(0)

            montant = lignes_code["SF Débit"].sum() + lignes_code["SF Crédit"].sum()
            if montant != 0:
                df.at[idx, colonne_label] = f"{int(montant):,}".replace(",", " ")

        return df

    if "balance_par_annee" not in st.session_state:
        st.warning("Veuillez d'abord générer une balance pour que les données soient disponibles.")
        return

    # Récupérer la balance complète (filtrée par classes et tableaux) depuis la session
    balance_dict = st.session_state.balance_par_annee  # Clé : année, Valeur : DataFrame balance
    balance_n = balance_dict.get(annee_n)
    balance_n1 = balance_dict.get(annee_n1)

    if balance_n is None or balance_n1 is None:
        st.warning("Balance non disponible pour l'une des années sélectionnées.")
        return

    # Générer les deux bilans
    df_actif = build_bilan_df(structure_bilan_actif, balance_n, annee_n, "Année N")
    df_actif = build_bilan_df(structure_bilan_actif, balance_n1, annee_n1, "Année N-1")

    df_passif = build_bilan_df(structure_bilan_passif, balance_n, annee_n, "Année N")
    df_passif = build_bilan_df(structure_bilan_passif, balance_n1, annee_n1, "Année N-1")

    # Affichage
    st.subheader("Bilan Actif")
    st.dataframe(df_actif, hide_index=True, use_container_width=True)

    st.subheader("Bilan Passif")
    st.dataframe(df_passif, hide_index=True, use_container_width=True)