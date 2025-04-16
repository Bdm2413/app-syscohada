import streamlit as st
import pandas as pd
import io
from fpdf import FPDF
import base64
from PIL import Image
import tempfile

st.set_page_config(page_title="App SYSCOHADA", page_icon="🏳️‍🌈", layout="wide")

# Menu latéral
st.sidebar.success("Menu de navigation")
menu = st.sidebar.selectbox("", ["Import Fichier", "Plan de comptes", "Grand Livre", "Balance"])
st.title("📊 :rainbow[Générateur de Balance Comptable]")

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
    if not st.session_state.data_loaded:
        st.warning("📂 Veuillez importer un fichier d'abord.")
    else:
        st.subheader("📚 Plan de comptes - Liste des comptes")
        st.dataframe(st.session_state.plan_df, use_container_width=True)

# Grand Livre
elif menu == "Grand Livre":
    if not st.session_state.data_loaded:
        st.warning("📂 Veuillez importer un fichier d'abord.")
    else:
        st.subheader("📚 Grand Livre - Écritures comptables")

        gl_df = st.session_state.gl_df.copy()

        if "Date" in gl_df.columns:
            gl_df["Date"] = pd.to_datetime(gl_df["Date"], format='%d/%m/%Y', errors='coerce')
            gl_df["Date "] = gl_df["Date"].dt.strftime('%d/%m/%Y')
            gl_df["Année"] = gl_df["Date"].dt.year
            gl_df["Mois"] = gl_df["Date"].dt.strftime("%Y%m")

        gl_df["Année"] = gl_df["Année"].fillna(0).astype(int)

        for col in ["Débit", "Crédit"]:
            if col in gl_df.columns:
                gl_df[col] = pd.to_numeric(gl_df[col], errors="coerce").fillna(0)

        # Filtres dynamiques
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

        total_debit = gl_df["Débit"].sum()
        total_credit = gl_df["Crédit"].sum()
        difference = total_debit - total_credit

        interpretation = "RAS" if difference == 0 else "Solde Débiteur" if difference > 0 else "Solde Créditeur"

        styles = {"debit": "#1982C4", "credit": "#013026", "diff": "#162A2C", "obs": "#FF595E"}

        def format_int(val): return f"{int(val):,}".replace(",", " ")

        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.markdown(f"""<div style="background-color:{styles['debit']}; padding:20px; border-radius:10px; text-align:center;">
                <div style="color:white; font-size:16px;">Total Débit</div>
                <div style="color:white; font-size:24px; font-weight:bold;">{format_int(total_debit)}</div></div>""", unsafe_allow_html=True)
        with col2:
            st.markdown(f"""<div style="background-color:{styles['credit']}; padding:20px; border-radius:10px; text-align:center;">
                <div style="color:white; font-size:16px;">Total Crédit</div>
                <div style="color:white; font-size:24px; font-weight:bold;">{format_int(total_credit)}</div></div>""", unsafe_allow_html=True)
        with col3:
            st.markdown(f"""<div style="background-color:{styles['diff']}; padding:20px; border-radius:10px; text-align:center;">
                <div style="color:white; font-size:16px;">Solde</div>
                <div style="color:white; font-size:24px; font-weight:bold;">{format_int(difference)}</div></div>""", unsafe_allow_html=True)
        with col4:
            st.markdown(f"""<div style="background-color:{styles['obs']}; padding:20px; border-radius:10px; text-align:center;">
                <div style="color:white; font-size:16px;">Observation</div>
                <div style="color:white; font-size:20px; font-weight:bold;">{interpretation}</div></div>""", unsafe_allow_html=True)

        st.markdown("### 📋 Tableau des écritures")

        # Colonnes à afficher
        colonnes_affichage = ["Date ", "Journal", "AN", "Compte", "Libellé", "Débit", "Crédit"]
        colonnes_presentes = [col for col in colonnes_affichage if col in gl_df.columns]
        gl_df["Débit"] = gl_df["Débit"].round(0)
        gl_df["Crédit"] = gl_df["Crédit"].round(0)

        st.dataframe(gl_df[colonnes_presentes], use_container_width=True)

        # Export Excel
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
            gl_df[colonnes_presentes].to_excel(writer, index=False, sheet_name="Grand Livre")
            writer.save()
        st.download_button(
            label="📥 Exporter en Excel",
            data=excel_buffer.getvalue(),
            file_name="grand_livre_filtré.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Export PDF avec logo
        uploaded_logo = st.file_uploader("🖼️ Importer un logo pour le PDF", type=["png", "jpg", "jpeg"])
        if uploaded_logo:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                tmp.write(uploaded_logo.read())
                logo_path = tmp.name

            class PDF(FPDF):
                def header(self):
                    if logo_path:
                        self.image(logo_path, 10, 8, 30)
                    self.set_font('Arial', 'B', 12)
                    self.cell(0, 10, "Grand Livre - Écritures comptables", border=False, ln=1, align='C')
                    self.ln(5)

                def footer(self):
                    self.set_y(-15)
                    self.set_font('Arial', 'I', 8)
                    self.cell(0, 10, f'Page {self.page_no()}', align='C')

            pdf = PDF(orientation='L', unit='mm', format='A4')
            pdf.add_page()
            pdf.set_font("Arial", size=9)

            colonnes = colonnes_presentes
            largeurs = [25, 20, 15, 25, 100, 25, 25]

            pdf.set_fill_color(220, 220, 220)
            pdf.set_font(style='B')
            for col, w in zip(colonnes, largeurs):
                pdf.cell(w, 8, col, border=1, align='C', fill=True)
            pdf.ln()

            pdf.set_font("Arial", size=8)
            for _, row in gl_df[colonnes].iterrows():
                for col, w in zip(colonnes, largeurs):
                    val = str(row[col])
                    align = 'R' if col in ["Débit", "Crédit"] else 'L'
                    pdf.cell(w, 7, val, border=1, align=align)
                pdf.ln()

            # Totaux
            pdf.set_font(style='B')
            pdf.cell(sum(largeurs[:-2]), 8, "TOTAL", border=1, align='R', fill=True)
            pdf.cell(largeurs[-2], 8, format_int(total_debit), border=1, align='R', fill=True)
            pdf.cell(largeurs[-1], 8, format_int(total_credit), border=1, align='R', fill=True)

            pdf_buffer = io.BytesIO()
            pdf.output(pdf_buffer)
            pdf_buffer.seek(0)

            st.download_button(
                label="📄 Télécharger le PDF avec logo",
                data=pdf_buffer,
                file_name="grand_livre_tableau.pdf",
                mime="application/pdf"
            )

# Balance
elif menu == "Balance":
    if not st.session_state.data_loaded:
        st.warning("📂 Veuillez d'abord importer un fichier Excel via le menu 'Import Fichier'.")
    else:
        st.subheader("📅 Balance à 8 colonnes")
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
