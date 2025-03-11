<#
.SYNOPSIS
    Exécution de user‑sync.exe en mode test puis en mode PROD, avec envoi par email du log complet.

.DESCRIPTION
    - Le script exécute user‑sync.exe en mode test et analyse sa sortie pour extraire le nombre d’utilisateurs "Adobe-only" à supprimer.
    - Si ce nombre dépasse un seuil ($Threshold), un email d’alerte est envoyé.
         * Le sujet intègre le nombre d’utilisateurs détectés.
         * Le corps du mail est formaté en HTML avec un tableau présentant :
              - La date/heure de lancement du script
              - Le nombre d’utilisateurs détectés (affiché en rouge)
              - Le seuil configuré
              - La date/heure prévue pour la prochaine exécution (après un délai d’attente paramétrable)
         * Le log complet (fichier unique généré pour cette exécution) est joint.
    - Sinon, le script exécute immédiatement user‑sync.exe en mode normal.
    - Aucune sortie n’est renvoyée à la console (adapté aux tâches planifiées).

.NOTES
    Date          : 2025-03-06 
    Version       : 1.0
#>

#region Paramètres de configuration

# Répertoire de stockage des logs (à créer s'il n'existe pas)
$LogDirectory = "C:\UserSyncTool\Logs"
if (-not (Test-Path $LogDirectory)) {
    New-Item -Path $LogDirectory -ItemType Directory | Out-Null
}

# Date/heure d'exécution pour le nom du fichier log
$TimeStampFile = (Get-Date).ToString("yyyyMMdd_HHmmss")
$LogFile = Join-Path $LogDirectory "user-sync_log_$TimeStampFile.log"

# Créer un fichier log vide avec encodage UTF-8 (avec BOM)
"" | Out-File -FilePath $LogFile -Encoding utf8

# Paramètres de seuil et délai d'attente (modifiable)
$Threshold = 1            
$WaitTimeSeconds = 7200    

# Paramètres SMTP génériques (à adapter selon votre relais SMTP cloud)
$SmtpServer   = "127.0.0.1"    # Ex : smtp.protonmail.com ou autre relais cloud
$SmtpPort     = "587"                  # Port SMTP (587 pour STARTTLS, 465 pour SSL)
$SmtpUsername = "email@example.com"   # Adresse de l'expéditeur (doit correspondre au compte d'authentification)
$SmtpToken    = "smtptoken"  
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

# Conversion du jeton en SecureString et création des identifiants
$SecureToken = ConvertTo-SecureString $SmtpToken -AsPlainText -Force
$SmtpCredential = New-Object System.Management.Automation.PSCredential ($SmtpUsername, $SecureToken)

# Destinataires de l'email (liste pouvant contenir plusieurs adresses)
$EmailRecipients = @("destinataire1@example.com", "destinataire2@example.com")

#endregion Paramètres de configuration

#region Fonctions Utilitaires

# Fonction pour écrire dans le log avec horodatage et mode (TEST ou PROD)
function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [string]$Mode = "PROD"
    )
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $logEntry = "$timestamp - $Mode - $Message"
    # Écriture dans le fichier log avec encodage UTF-8
    Add-Content -Path $LogFile -Value $logEntry -Encoding utf8
}

#endregion Fonctions Utilitaires

#region Exécution en Mode Test

# Enregistrer la date/heure de lancement du script
$ScriptStartTime = Get-Date

# Exécuter user‑sync.exe en mode test et récupérer la sortie (y compris erreurs)
try {
    $TestOutput = & ".\user-sync.exe" -t 2>&1
}
catch {
    Write-Log "Erreur lors de l'execution en mode test : $($_.Exception.Message)" "TEST"
    exit 1
}

# Enregistrer la sortie du mode test dans le log (avec préfixe "TEST")
$TestOutput | ForEach-Object { Add-Content -Path $LogFile -Value ("TEST - " + $_) -Encoding utf8 }

# Joindre les lignes pour analyse
$TestOutputString = $TestOutput -join "`n"

# Extraire le nombre d'utilisateurs "Adobe-only" supprimés
if ($TestOutputString -match "Number of Adobe-only users removed:\s+(\d+)") {
    $NbUsersRemoved = [int]$matches[1]
}
else {
    Write-Log "La ligne 'Number of Adobe-only users removed:' n'a pas été trouvée dans la sortie." "TEST"
    $NbUsersRemoved = 0
}

Write-Log "Mode test : $NbUsersRemoved utilisateurs Adobe-only détectés pour suppression." "TEST"

#endregion Exécution en Mode Test

#region Décision et Envoi de l'Email

# Si le nombre détecté dépasse le seuil, envoyer un email d'alerte avec log en pièce jointe
if ($NbUsersRemoved -gt $Threshold) {
    # Calculer la date/heure prévue pour l'exécution normale
    $NextExecutionTime = $ScriptStartTime.AddSeconds($WaitTimeSeconds)
    
    # Construire le sujet de l'email en intégrant le nombre détecté
    $EmailSubject = "Alerte : user‑sync.exe - $NbUsersRemoved utilisateurs détectés"

    # Construire le corps de l'email en HTML avec mise en forme
    $EmailBody = @"
<html>
<head>
  <meta charset='UTF-8'>
  <style>
    body { font-family: Arial, sans-serif; margin: 20px; }
    h2 { color: #333; }
    table { border-collapse: collapse; width: 100%; }
    th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
    th { background-color: #f2f2f2; }
    .critical { color: red; font-weight: bold; }
  </style>
</head>
<body>
  <h2>Alerte : Exécution de user‑sync.exe</h2>
  <p>Le script a été lancé le : <strong>$($ScriptStartTime.ToString("yyyy-MM-dd HH:mm:ss"))</strong></p>
  <table>
    <tr>
      <th>Information</th>
      <th>Valeur</th>
    </tr>
    <tr>
      <td>Nombre d'utilisateurs Adobe-only détectés</td>
      <td class='critical'>$NbUsersRemoved</td>
    </tr>
    <tr>
      <td>Seuil configuré</td>
      <td>$Threshold</td>
    </tr>
    <tr>
      <td>Prochaine exécution prévue</td>
      <td>$($NextExecutionTime.ToString("yyyy-MM-dd HH:mm:ss"))</td>
    </tr>
  </table>
  <p>Veuillez consulter la pièce jointe pour le log complet de cette exécution.</p>
</body>
</html>
"@

    Write-Log "Nombre d'utilisateurs ($NbUsersRemoved) supérieur au seuil ($Threshold). Envoi d'un email d'alerte." "TEST"

    try {
      # Création de l'objet MailMessage
      $mailMessage = New-Object System.Net.Mail.MailMessage
      $mailMessage.From = $SmtpUsername
      foreach ($recipient in $EmailRecipients) {
          $mailMessage.To.Add($recipient)
      }
      $mailMessage.Subject = $EmailSubject
      $mailMessage.Body = $EmailBody
      $mailMessage.IsBodyHtml = $true
      
      # Définir l'encodage du sujet et du corps en UTF8
      $mailMessage.SubjectEncoding = [System.Text.Encoding]::UTF8
      $mailMessage.BodyEncoding = [System.Text.Encoding]::UTF8

      # Ajouter la pièce jointe
      $attachment = New-Object System.Net.Mail.Attachment($LogFile)
      $mailMessage.Attachments.Add($attachment)
      
      # Configuration du client SMTP
      $smtpClient = New-Object System.Net.Mail.SmtpClient($SmtpServer, [int]$SmtpPort)
      $smtpClient.EnableSsl = $true
      $smtpClient.Credentials = $SmtpCredential

      # Envoi de l'email
      $smtpClient.Send($mailMessage)
      
      Write-Log "E-mail d'alerte envoyé avec succès." "TEST"
  }
  catch {
      Write-Log "Échec de l'envoi de l'e-mail d'alerte : $($_.Exception.Message)" "TEST"
  }
    Write-Log "Attente de $WaitTimeSeconds secondes avant l'exécution en mode PROD." "TEST"
    Start-Sleep -Seconds $WaitTimeSeconds
}

#endregion Décision et Envoi de l'Email

#region Exécution en Mode PROD

Write-Log "Exécution de user‑sync.exe en mode PROD." "PROD"

try {
    $NormalOutput = & ".\user-sync.exe" 2>&1
}
catch {
    Write-Log "Erreur lors de l'exécution en mode PROD : $($_.Exception.Message)" "PROD"
    exit 1
}

$NormalOutput | ForEach-Object { Add-Content -Path $LogFile -Value ("PROD - " + $_) -Encoding utf8 }
Write-Log "Exécution de user‑sync.exe en mode PROD terminée." "PROD"

#endregion Exécution en Mode PROD

#region Fin du Script

Write-Log "Script terminé." "PROD"

#endregion Fin du Script
