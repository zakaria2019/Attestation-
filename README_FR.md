# Application Windows — Gestion des attestations | Aéroport Ouarzazate

## Fonctionnalités
- Interface Windows indépendante de la feuille Excel : menu latéral, tableau de bord, formulaire RH, historique et paramètres.
- Lecture du personnel depuis la feuille `Personnel` du classeur existant.
- Structure attendue, d'après le VBA fourni :
  - Colonne A : Matricule
  - Colonne B : Nom
  - Colonne C : Prénom
  - Colonne D : Civilité
  - Colonne I : Emploi
- Génération automatique de l'adresse CC selon la convention existante : première lettre du prénom + "." + nom sans espaces + `@onda.ma`.
- Préparation d'un message Outlook en conservant la signature Outlook. Le message reste ouvert pour vérification avant envoi.
- Journal des demandes dans `Journal_Demandes` du même classeur.
- Choix du classeur RH dans Paramètres ; le chemin est mémorisé dans `config_attestations.json`.

## Prérequis
- Windows 10 ou 11.
- Microsoft Excel et Microsoft Outlook installés/configurés.
- Python 3.10 ou plus récent pour lancer le code source.
- Le classeur RH doit être fermé dans Excel pendant que l'application écrit dans le journal, afin d'éviter les conflits de fichier.

## Lancer le code source
Ouvrir l'invite de commandes dans ce dossier puis exécuter :

```bat
py -m pip install --upgrade pip
py -m pip install openpyxl pywin32
py app.py
```

Au premier lancement, choisissez votre classeur `.xlsx` ou `.xlsm`. Il doit contenir une feuille nommée exactement `Personnel`.

## Créer le fichier EXE
Dans l'invite de commandes ouverte dans ce dossier :

```bat
py -m pip install pyinstaller
pyinstaller --noconfirm --clean --windowed --name "Attestations Ouarzazate" --hidden-import win32timezone app.py
```

L'exécutable sera créé ici :

`dist\Attestations Ouarzazate\Attestations Ouarzazate.exe`

Pour un exécutable en fichier unique, essayez :

```bat
pyinstaller --noconfirm --clean --onefile --windowed --name "Attestations Ouarzazate" --hidden-import win32timezone app.py
```

Le fichier unique sera dans `dist\Attestations Ouarzazate.exe`.

## Première utilisation
1. Double-cliquer sur l'application.
2. Ouvrir Paramètres si le choix du classeur n'est pas proposé.
3. Sélectionner le classeur RH contenant `Personnel`.
4. Aller dans « Nouvelle demande », choisir le matricule et vérifier les informations.
5. Choisir une ou deux attestations puis cliquer sur « Préparer le courriel ».
6. Vérifier le message dans Outlook et cliquer sur Envoyer manuellement.
7. Consulter l'historique dans l'application.

## Notes importantes
- L'envoi direct est volontairement désactivé : l'application ouvre le message Outlook pour validation humaine.
- Le journal indique `Préparé` dès que le courriel a été ouvert avec succès, avant que l'utilisateur ne clique nécessairement sur Envoyer dans Outlook. Ce comportement reprend le principe de la macro VBA initiale.
- Pour les fichiers `.xlsm`, le code demande à openpyxl de conserver les macros. Faites une copie de sauvegarde avant les premiers essais.
- Ne laissez pas le classeur ouvert dans Excel pendant l'enregistrement de l'historique.
- Cette version est une base fonctionnelle à tester dans votre environnement ONDA. Les politiques de sécurité informatique peuvent bloquer l'automatisation Outlook ou les exécutables non signés ; sollicitez la DSI si nécessaire.
