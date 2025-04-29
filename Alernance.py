# ta grosse mere la pute

import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.dates import DateFormatter

# Chemin vers le fichier Excel
file_path = r"C:\Users\Utilisateur\Documents\autre titi\Alternances.xlsx"

# Lecture du fichier Excel
try:
    df = pd.read_excel(file_path)

    # Afficher les colonnes disponibles pour débogage
    print("Colonnes disponibles :", df.columns.tolist())

    # Nettoyer les noms des colonnes pour supprimer les espaces invisibles
    df.columns = df.columns.str.strip()

    # Vérification des colonnes nécessaires
    if "Date de candidature envoye?" not in df.columns or "Envois?" not in df.columns or "Poste?" not in df.columns:
        raise ValueError("Les colonnes 'Date de candidature envoye?', 'Envois?' ou 'Poste?' sont absentes du fichier Excel.")

    # Afficher les valeurs uniques dans la colonne 'Poste?' pour vérification
    print("Valeurs dans la colonne 'Poste?' :", df["Poste?"].unique())

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
    personalized_counts = df[df["Envois?"] == "Lettre personnalisé"].groupby("Date de candidature envoye?").size()

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

    # Création du graphique
    plt.figure(figsize=(10, 6))

    # Barres empilées : vert pour les lettres personnalisées, bleu pour les autres
    plt.bar(daily_counts.index, personalized_counts, color='green', label='Lettre personnalisée')
    plt.bar(daily_counts.index, daily_counts - personalized_counts, bottom=personalized_counts, color='blue', label='Autres')

    # Ajout des ronds rouges pour les candidatures spontanées
    for i, date in enumerate(all_dates):
        # Ajout des ronds rouges pour les candidatures spontanées avec lettre personnalisée
        bottom = personalized_counts[date]
        for _ in range(int(spontaneous_counts[date])):
            if date in personalized_counts and personalized_counts[date] > 0:
                plt.scatter(date, bottom + 0.5, color='red', s=50, zorder=3)  # Rond rouge sur vert
            else:
                plt.scatter(date, bottom + 0.5, color='red', s=50, zorder=3)  # Rond rouge sur bleu
            bottom += 1

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