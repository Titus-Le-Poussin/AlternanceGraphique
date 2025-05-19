import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.dates import DateFormatter

# Chemin vers le fichier Excel
file_path = r"C:\Users\Utilisateur\Documents\autre titi\Alternances.xlsx"

# Lecture du fichier Excel
try:
    df = pd.read_excel(file_path)

    # Nettoyer les noms des colonnes pour supprimer les espaces invisibles
    df.columns = df.columns.str.strip()

    # Vérification des colonnes nécessaires
    if "Date de candidature envoye?" not in df.columns or "Envois?" not in df.columns or "Poste?" not in df.columns:
        raise ValueError("Les colonnes 'Date de candidature envoye?', 'Envois?' ou 'Poste?' sont absentes du fichier Excel.")

    # Nettoyer les valeurs de la colonne 'Poste?'
    df["Poste?"] = df["Poste?"].str.strip().str.lower()

    # Nettoyage de la colonne pour supprimer le préfixe "le"
    df["Date de candidature envoye?"] = df["Date de candidature envoye?"].str.replace(r"^\s*le\s*", "", regex=True)

    # Conversion des dates en format datetime avec un format explicite
    df["Date de candidature envoye?"] = pd.to_datetime(
        df["Date de candidature envoye?"], format='%d/%m/%Y', errors='coerce'
    )

    # Suppression des lignes avec des dates invalides
    df = df.dropna(subset=["Date de candidature envoye?"])

    # Comptage des CV envoyés par jour
    daily_counts = df.groupby("Date de candidature envoye?").size()

    # Comptage des CV envoyés avec une lettre personnalisée
    personalized_counts = df[df["Envois?"].str.strip().str.lower() == "lettre personnalisé"].groupby("Date de candidature envoye?").size()

    # Comptage des candidatures spontanées
    spontaneous_counts = df[df["Poste?"].str.contains("spontan", na=False)].groupby("Date de candidature envoye?").size()

    # Générer une plage de dates complète
    start_date = pd.Timestamp("2025-04-01")
    end_date = daily_counts.index.max()
    all_dates = pd.date_range(start=start_date, end=end_date)

    # Réindexer pour inclure tous les jours, avec 0 pour les jours sans envois
    daily_counts = daily_counts.reindex(all_dates, fill_value=0)
    personalized_counts = personalized_counts.reindex(all_dates, fill_value=0)
    spontaneous_counts = spontaneous_counts.reindex(all_dates, fill_value=0)

    # Comptage des candidatures en entretien
    entretien_counts = df[df["Envois?"].str.strip().str.lower() == "candidature en entretien"].groupby("Date de candidature envoye?").size()
    entretien_counts = entretien_counts.reindex(all_dates, fill_value=0)

    # Création du graphique
    plt.figure(figsize=(10, 6))

    # Barres empilées : vert pour les lettres personnalisées, bleu pour les autres, rouge pour les candidatures en entretien
    plt.bar(daily_counts.index, personalized_counts, color='green', label='Lettre personnalisée')
    plt.bar(daily_counts.index, daily_counts - personalized_counts - entretien_counts, bottom=personalized_counts, color='blue', label='Autres')
    plt.bar(daily_counts.index, entretien_counts, bottom=personalized_counts + (daily_counts - personalized_counts - entretien_counts), color='red', label='Candidature en entretien')

    # Ajout des ronds rouges pour chaque candidature spontanée à la bonne hauteur
    for date in all_dates:
        # Sélectionner les candidatures du jour, en conservant l'ordre d'origine
        day_df = df[df["Date de candidature envoye?"] == date]
        # Réinitialiser l'index pour pouvoir itérer avec la position
        day_df = day_df.reset_index(drop=True)
        for idx, row in day_df.iterrows():
            if "spontan" in str(row["Poste?"]):
                # Le point est placé à la hauteur idx + 0.5 (au centre de la barre)
                plt.scatter(date, idx + 0.5, color='red', s=50, zorder=3)

    # Titre et labels
    plt.title("Nombre de CV envoyés par jour")
    plt.xlabel("Date")
    plt.ylabel("Nombre de CV envoyés")
    plt.grid(axis='y')

    # Formatage des dates sur l'axe des abscisses
    date_formatter = DateFormatter("%d/%m/%Y")
    plt.gca().xaxis.set_major_formatter(date_formatter)

    # Rotation des dates pour une meilleure lisibilité
    plt.xticks(rotation=45)

    # Ajouter tous les chiffres sur l'axe des ordonnées
    plt.yticks(range(0, int(daily_counts.max()) + 1))

    # Ajouter le total des CV envoyés sur le graphique
    total_cvs = daily_counts.sum()
    plt.text(0.01, 0.01, f"Total CV envoyés : {total_cvs}", transform=plt.gcf().transFigure, fontsize=10, color='black')

    # Ajouter une entrée pour les ronds rouges dans la légende
    plt.scatter([], [], color='red', s=50, label='Candidature spontanée')

    # Légende
    plt.legend()

    # Ajustement automatique des marges
    plt.tight_layout()

    # Affichage du graphique
    plt.show()

except FileNotFoundError:
    print(f"Le fichier '{file_path}' est introuvable.")
except ValueError as e:
    print(f"Erreur : {e}")
except Exception as e:
    print(f"Une erreur inattendue s'est produite : {e}")