import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st

# Configurer l'application Streamlit
st.set_page_config(
    page_title="Analyse Acoustique Interactive",
    page_icon=":chart_with_upwards_trend:",
    layout="centered",
    initial_sidebar_state="expanded"
)

# Titre et configuration de la barre latérale
st.title("Outil d'Analyse Acoustique Interactive")
st.sidebar.title("Configuration")

# Liste des fréquences prédéfinies
frequencies = np.array([
    200, 250, 315, 400, 500, 630, 800, 1000,
    1250, 1600, 2000, 2500, 3150, 4000, 5000,
    6300, 8000, 10000
])

# Fonction pour charger et extraire des données
def load_and_extract_data(file):
    """
    Charge et extrait les données des séries dans la feuille 'Macro' d'un fichier Excel.
    Ignore les séries avec des valeurs non numériques.
    """
    try:
        df = pd.read_excel(file, sheet_name="Macro", header=None)
        extracted_data = []

        for col_start, name_cell, values_range in [
            (0, "A1", "C3:C20"), (4, "E1", "G3:G20"),
            (8, "I1", "K3:K20"), (12, "M1", "O3:O20"),
            (16, "Q1", "S3:S20"), (20, "U1", "W3:W20")
        ]:
            name = df.iloc[0, col_start]
            values = df.iloc[2:20, col_start + 2].values

            # Vérifie si toutes les valeurs sont numériques
            if np.issubdtype(values.dtype, np.number) and not np.isnan(values).all():
                extracted_data.append({"name": name, "values": values})

        return extracted_data

    except Exception as e:
        st.error(f"Erreur lors de la lecture du fichier : {e}")
        return []

# Téléchargement de fichiers
uploaded_files = st.sidebar.file_uploader(
    "Importez vos fichiers Excel (format .xlsx)", 
    type=["xlsx"], 
    accept_multiple_files=True
)

# Afficher les courbes pour chaque fichier et série
if uploaded_files:
    for file in uploaded_files:
        st.subheader(f"Données extraites de : {file.name}")

        # Extraire les données valides
        extracted_data = load_and_extract_data(file)

        # Si aucune série valide
        if not extracted_data:
            st.warning(f"Aucune série valide trouvée dans {file.name}.")
            continue

        # Génération des graphiques pour chaque série
        fig, ax = plt.subplots(figsize=(10, 6))
        for series in extracted_data:
            ax.plot(frequencies, series["values"], label=series["name"], marker="o")

        # Personnalisation du graphique
        ax.set_title(f"Courbes d'absorption pour {file.name}")
        ax.set_xscale("log")
        ax.set_xticks(frequencies)
        ax.get_xaxis().set_major_formatter(plt.ScalarFormatter())
        ax.set_xlabel("Fréquence (Hz)")
        ax.set_ylabel("Coefficient d'absorption")
        ax.legend()
        ax.grid(True, which="both", linestyle="--", linewidth=0.5)

        # Afficher le graphique dans Streamlit
        st.pyplot(fig)
else:
    st.info("Veuillez importer au moins un fichier Excel pour commencer l'analyse.")
