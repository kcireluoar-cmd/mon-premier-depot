# Extracteur de dossier — Suivi DAO/AMI (version Windows)

Application de bureau qui lit un dossier (PDF, Word, texte), demande à Claude
d'en extraire les informations utiles, et **écrit directement une nouvelle
ligne** dans votre classeur `Suivi_Dossiers_DAO_AMI.xlsx`, dans la première
ligne vierge disponible, avec les bons formats (dates, montant, menus
déroulants).

## 1. Obtenir le fichier .exe (sans avoir besoin d'un PC Windows)

1. Créez un compte GitHub gratuit sur https://github.com si vous n'en avez pas.
2. Créez un nouveau dépôt (bouton **New repository**), par exemple nommé
   `extracteur-dao`. Laissez-le privé si vous préférez.
3. Déposez-y tous les fichiers de ce dossier (`app.py`, `requirements.txt`,
   `README.md`, et le dossier `.github/workflows/build-windows.yml`) —
   le plus simple est de faire glisser les fichiers dans l'interface web de
   GitHub (« Add file » → « Upload files »), en conservant bien le chemin
   `.github/workflows/build-windows.yml`.
4. Une fois les fichiers envoyés (« Commit »), allez dans l'onglet
   **Actions** du dépôt : une compilation démarre automatiquement (« Build
   Windows executable »). Elle prend 2 à 3 minutes.
5. Une fois terminée (coche verte), cliquez dessus, puis dans la section
   **Artifacts** en bas de page, téléchargez `ExtracteurDAO-windows.zip`.
6. Décompressez : vous obtenez `ExtracteurDAO.exe`. C'est votre application.

Vous pouvez répéter cette opération à chaque fois que je vous fournis une
nouvelle version de `app.py` : il suffit de remplacer le fichier dans votre
dépôt GitHub et de relancer l'onglet Actions (ou laissez faire, cela se
relance automatiquement à chaque envoi).

## 2. Installer / lancer sur Windows

`ExtracteurDAO.exe` est un exécutable autonome : pas d'installation, pas
besoin d'installer Python. Copiez-le où vous voulez (Bureau, dossier de
travail) et double-cliquez pour le lancer.

Windows peut afficher un avertissement « Éditeur inconnu / SmartScreen » la
première fois, car le fichier n'est pas signé numériquement (signer un
exécutable coûte un certificat payant, sans intérêt pour un usage
personnel). Cliquez sur « Informations complémentaires » puis « Exécuter
quand même ».

## 3. Première utilisation

1. Au premier lancement, ouvrez **Réglages → Clé API Anthropic…** et collez
   votre clé (obtenue sur https://console.anthropic.com/settings/keys).
   Elle est enregistrée localement dans
   `%APPDATA%\ExtracteurDAO\config.json` — jamais envoyée ailleurs qu'à
   l'API Anthropic au moment de l'extraction.
2. Ouvrez **Réglages → Classeur de suivi…** et choisissez votre fichier
   `Suivi_Dossiers_DAO_AMI.xlsx`. Ce chemin est mémorisé pour les prochaines
   fois.
3. Cliquez **Choisir un fichier**, sélectionnez le dossier à traiter
   (PDF, .docx ou .txt).
4. Cliquez **Extraire les informations (IA)** : les champs se remplissent.
5. Relisez, corrigez si besoin.
6. Cliquez **Ajouter cette ligne au classeur** : la ligne est écrite
   directement dans la feuille « Suivi DAO-AMI » de votre classeur, dans la
   première ligne vierge.

## Limites à connaître

- **Documents scannés (image)** : le texte n'est pas lisible automatiquement
  (pas d'OCR dans cette version). Remplissez la ligne à la main dans ce cas
  — le formulaire reste utilisable seul.
- **Capacité du tableau** : le classeur généré prévoit 12 lignes vierges
  prêtes à l'emploi. Une fois épuisées, l'outil continue d'ajouter des
  lignes juste en dessous (avec la même mise en forme), mais celles-ci ne
  seront pas comptées automatiquement par le Tableau de bord tant que vous
  n'avez pas réétendu le tableau Excel (sélectionnez le tableau, onglet
  **Création de tableau → Redimensionner le tableau**) et vérifié les
  formules. Revenez me voir si vous voulez une version qui gère cette
  extension automatiquement.
- **Rien n'est deviné** : un champ que l'IA ne trouve pas dans le document
  reste vide plutôt que d'être rempli au hasard.

## Mettre à jour le modèle utilisé

Le modèle Claude utilisé est `claude-sonnet-5` par défaut. Pour en changer,
ajoutez une clé `"model"` dans `%APPDATA%\ExtracteurDAO\config.json`,
par exemple :

```json
{
  "api_key": "sk-ant-...",
  "workbook_path": "C:\\Users\\...\\Suivi_Dossiers_DAO_AMI.xlsx",
  "model": "claude-opus-5-5"
}
```
