# Obtenir l'EXE sans installer Python sur votre PC

Ce paquet contient le code source et un workflow qui construit l'application sur un environnement Windows distant (GitHub Actions). Votre PC n'a pas besoin d'avoir Python pour utiliser l'EXE final. Il faut toutefois un compte GitHub et un accès Internet pour déclencher la compilation et récupérer le fichier.

## Étapes
1. Connectez-vous à GitHub et créez un nouveau dépôt privé (recommandé pour un projet professionnel).
2. Téléversez tous les fichiers de ce dossier dans le dépôt, y compris le dossier caché `.github/workflows`.
3. Ouvrez l'onglet **Actions** du dépôt.
4. Sélectionnez **Build Windows EXE**, puis **Run workflow**. (Un premier push sur la branche `main` peut aussi lancer le workflow automatiquement.)
5. Quand l'exécution est terminée avec le statut vert, ouvrez l'exécution.
6. Dans **Artifacts**, téléchargez `Attestations_Ouarzazate_EXE`.
7. Décompressez l'archive téléchargée : elle contient `Attestations_Ouarzazate.exe`.

## Sur le PC ONDA
- Copiez l'EXE dans un dossier où vous avez le droit d'écriture, par exemple Documents.
- Au premier lancement, sélectionnez le classeur RH qui contient la feuille `Personnel`.
- Outlook et Excel doivent être installés/configurés.
- Ne gardez pas le classeur ouvert dans Excel lorsque l'application écrit dans le journal.
- L'EXE est non signé. Windows ou l'antivirus de l'entreprise peut le bloquer même sans demander les droits administrateur. Dans ce cas, demandez à la DSI de valider le fichier; ne contournez pas les contrôles de sécurité.

## Confidentialité
Le dépôt privé évite de publier le code publiquement, mais respectez les règles de l'ONDA et ne téléversez jamais le classeur RH ou des données personnelles dans GitHub. Seul le code de l'application doit être envoyé.
