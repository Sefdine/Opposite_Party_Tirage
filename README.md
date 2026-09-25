# Extraction Opposite Party

Petite application web Streamlit pour traiter plusieurs exports Excel ayant
le même format.

## Fonctionnement

1. Import de plusieurs fichiers `.xls` ou `.xlsx`.
2. Lecture avec `header=5`.
3. Conservation des transactions dont `Transaction Status = Completed`.
4. Nettoyage de `Opposite Party`.
5. Extraction de la première séquence de 10 chiffres.
6. Affichage des numéros commençant par `411` sous le titre **Numéro staff**.
7. Suppression optionnelle des numéros staff.
8. Visualisation du résultat.
9. Téléchargement CSV ou Excel.

## Installation

```bash
pip install -r requirements.txt
```

## Lancement

```bash
streamlit run app.py
```

L'application s'ouvrira ensuite dans le navigateur.
