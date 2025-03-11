# Présentation du Script

Ce script a été conçu pour orchestrer l'exécution de `user‑sync.exe` en mode test puis en mode production, tout en gérant les logs et en envoyant des alertes par email en cas de détection d'anomalies. L'approche adoptée permet d'automatiser et de sécuriser le processus de synchronisation des utilisateurs, tout en offrant une traçabilité complète grâce aux logs détaillés.

## Points Clés du Script

- **Initialisation et Configuration**  
  - Définition d'un répertoire dédié aux logs et création d'un fichier log unique, horodaté pour chaque exécution.
  - Paramétrage des seuils d'alerte, des délais d'attente, et des paramètres SMTP pour l'envoi d'emails (serveur, port, identifiants, etc.).

- **Fonctionnalité de Log Personnalisée**  
  - La fonction `Write-Log` ajoute systématiquement la date, l'heure et le contexte (`TEST` ou `PROD`) à chaque message loggué.
  - Cette catégorisation permet de distinguer facilement les messages issus de la phase de test de ceux en production.

- **Exécution en Mode Test**  
  - Le script lance `user‑sync.exe` en mode test (avec le flag `-t`) et capture toute la sortie, incluant les messages d’erreur.
  - La sortie est ensuite analysée pour extraire le nombre d'utilisateurs "Adobe-only" supprimés.  
  - Si la ligne attendue n'est pas trouvée ou en cas d'erreur, un message est loggué en mode `TEST`.

- **Décision et Alerte par Email**  
  - Si le nombre d'utilisateurs supprimés dépasse le seuil configuré, un email d'alerte est préparé et envoyé.
  - L'email inclut un sujet dynamique intégrant le nombre détecté, ainsi qu'un corps HTML formaté présentant :
    - La date/heure de lancement du script.
    - Le nombre d'utilisateurs détectés (affiché en rouge pour attirer l'attention).
    - Le seuil configuré.
    - La date/heure prévue pour la prochaine exécution.
  - Le log complet de l'exécution est joint à l'email.
  - Le script attend ensuite un délai défini avant de passer en mode production.

- **Exécution en Mode Production**  
  - Si aucune alerte n'est déclenchée ou après le délai d'attente, `user‑sync.exe` est exécuté en mode production.
  - La sortie est enregistrée dans le log avec le tag `PROD`, permettant de suivre distinctement l'exécution normale par rapport aux tests.


