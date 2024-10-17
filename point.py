import streamlit as st
import pandas as pd
from io import BytesIO

# Fonction pour traiter le fichier
def process_attendance_data(uploaded_file):
    # Lire le fichier Excel
    df = pd.read_excel(uploaded_file)
    
    # Convertir la colonne 'Heure' en datetime
    df['Heure'] = pd.to_datetime(df['Heure'])
    
    # Créer une nouvelle colonne 'Date' basée sur 'Heure'
    df['Date'] = df['Heure'].dt.date
    
    # Trier par 'Nom' et 'Heure' pour garantir l'ordre correct
    df = df.sort_values(by=['Nom', 'Date', 'Heure'])
    
    # Extraire la première heure d'arrivée et la dernière heure de sortie par personne par jour
    df_filtered = df.groupby(['Nom', 'Date']).agg(
        Heure_Arrive=('Heure', 'first'),
        Heure_Sortie=('Heure', 'last')
    ).reset_index()
    
    # Formater les colonnes d'heure pour afficher uniquement les heures, minutes et secondes
    df_filtered['Heure d\'arrive et de sortie'] = df_filtered['Heure_Arrive'].dt.strftime('%H:%M:%S') + ' - ' + df_filtered['Heure_Sortie'].dt.strftime('%H:%M:%S')
    
    # Conserver uniquement les colonnes nécessaires
    df_final = df_filtered[['Date', 'Nom', 'Heure d\'arrive et de sortie']]
    
    return df_final

# Fonction pour convertir DataFrame en fichier Excel téléchargeable
def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Feuille1')
    writer.close()  # Fermeture du writer après écriture des données
    processed_data = output.getvalue()
    return processed_data

# Titre de l'application
st.title("App de traitement de pointage de présence")

# Zone pour uploader un fichier
uploaded_file = st.file_uploader("Choisir un fichier Excel", type=['xlsx'])

if uploaded_file is not None:
    # Traitement du fichier
    df_processed = process_attendance_data(uploaded_file)
    
    # Afficher le DataFrame traité
    st.write("Aperçu du fichier traité:")
    st.dataframe(df_processed)
    
    # Bouton pour télécharger le fichier traité
    st.download_button(
        label="Télécharger le fichier traité",
        data=to_excel(df_processed),
        file_name="pointage_traite.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
