<#
.SYNOPSIS
    Exécution de user‑sync.exe en mode test puis en mode PROD, avec envoi par email du log complet (configurable).

.DESCRIPTION
    - Le script exécute user‑sync.exe en mode test et analyse sa sortie pour extraire le nombre d’utilisateurs "Adobe-only" à supprimer.
    - Si ce nombre dépasse un seuil configurable ($Threshold), un email d’alerte est envoyé.
         * Le sujet intègre le nombre d’utilisateurs détectés, par exemple : 
           "Alerte: Adobe User Sync Tool s'apprête à supprimer x utilisateurs".
         * Le corps du mail est formaté en HTML et détaille :
              - La date/heure de lancement du script (avec indication "UTC Paris")
              - Le nombre d’utilisateurs détectés (affiché en rouge)
              - Le seuil configuré
              - La date/heure prévue pour la prochaine exécution (avec indication "UTC Paris")
         * Le log complet de l’exécution est joint (via une copie temporaire).
    - Si le seuil n'est pas dépassé, le script attend 1 heure (3600 secondes) avant d'exécuter user‑sync.exe en mode PROD afin de respecter les limitations de l'API Adobe.
    - Un booléen permet de désactiver le temps d'attente en mode test pour faciliter le débogage.
    - Aucune sortie n’est affichée à la console (adapté aux tâches planifiées).

.NOTES
    Date    : 2025-03-06
    Version : 1.4
#>

#region Paramètres de configuration

# Répertoire de stockage des logs (création si inexistant)
$LogDirectory = "C:\Users\Administrateur\Documents\AD to AC\logs"
if (-not (Test-Path $LogDirectory)) {
    New-Item -Path $LogDirectory -ItemType Directory | Out-Null
}

# Date/heure d'exécution pour le nom du fichier log
$TimeStampFile = (Get-Date).ToString("yyyyMMdd_HHmmss")
$LogFile = Join-Path $LogDirectory "user-sync_log_$TimeStampFile.log"

# Création d'un fichier log vide en UTF-8 (avec BOM)
"" | Out-File -FilePath $LogFile -Encoding utf8

# Paramètres généraux de seuil et délai d'attente
$Threshold = 1              # Seuil d'alerte (nombre d'utilisateurs "Adobe-only" supprimés)
$WaitTimeSeconds = 7200     # Délai en secondes avant exécution en mode PROD après alerte

# Booléen pour désactiver l'attente en mode test (pour faciliter les tests) $false => le script attend 1 heure afin d'éviter le blocage de l'API
# si $true => n'attend pas et lance directement le script en PROD
$DisableWaitForTest = $false

# Paramètres SMTP – à configurer selon votre environnement
$SmtpServer   = "127.0.0.1"    # Ex : smtp.protonmail.com ou autre relais cloud
$SmtpPort     = "1025"                  # Port SMTP (587 pour STARTTLS, 465 pour SSL)
$SmtpUsername = "username@example.com"   # Adresse de l'expéditeur (doit correspondre au compte d'authentification)
$SmtpToken    = "password"  
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

# Options pour l'envoi d'email
$EnableTls       = $true      # Mettre $false pour désactiver TLS
$EnableSmtpAuth  = $true      # Mettre $false pour désactiver l'authentification SMTP

# Préparation des identifiants SMTP si l'authentification est activée
if ($EnableSmtpAuth) {
    $SecureToken    = ConvertTo-SecureString $SmtpToken -AsPlainText -Force
    $SmtpCredential = New-Object System.Management.Automation.PSCredential ($SmtpUsername, $SecureToken)
}

# Liste des destinataires (plusieurs adresses possibles)
$EmailRecipients = @("destinataire1@example.com","destinataire2@example.com")

#endregion Paramètres de configuration

#region Fonctions Utilitaires

# Fonction pour écrire dans le log avec horodatage et tag de mode (TEST ou PROD)
function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [string]$Mode = "PROD"
    )
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $logEntry = "$timestamp - $Mode - $Message"
    Add-Content -Path $LogFile -Value $logEntry -Encoding utf8
}

#endregion Fonctions Utilitaires

#region Exécution en Mode Test

# Sauvegarder la date/heure de lancement
$ScriptStartTime = Get-Date

# Exécuter user‑sync.exe en mode test et récupérer toute la sortie (y compris erreurs)
try {
    $TestOutput = & ".\user-sync.exe" -t 2>&1
}
catch {
    Write-Log "Erreur lors de l'execution en mode test : $($_.Exception.Message)" "TEST"
    exit 1
}

# Enregistrer chaque ligne de la sortie avec le tag "TEST"
$TestOutput | ForEach-Object { Add-Content -Path $LogFile -Value ("TEST - " + $_) -Encoding utf8 }

# Assembler la sortie pour analyse
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

if ($NbUsersRemoved -gt $Threshold) {
    # Calculer la date/heure prévue pour l'exécution normale
    $NextExecutionTime = $ScriptStartTime.AddSeconds($WaitTimeSeconds)
    
    # Sujet de l'email intégrant le nombre d'utilisateurs détectés
    $EmailSubject = "Alerte: Adobe User Sync Tool s'apprête à supprimer $NbUsersRemoved utilisateurs"

    # Corps de l'email en HTML avec ajout de "UTC Paris" après les dates
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
  <p>Le script a été lancé le : <strong>$($ScriptStartTime.ToString("yyyy-MM-dd HH:mm:ss")) UTC Paris</strong></p>
  <table>
    <tr>
      <th>Informations</th>
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
      <td>$($NextExecutionTime.ToString("yyyy-MM-dd HH:mm:ss")) UTC Paris</td>
    </tr>
  </table>
  <p>Veuillez consulter la pièce jointe pour le log complet de cette exécution.</p>
</body>
</html>
"@

    Write-Log "Nombre d'utilisateurs ($NbUsersRemoved) supérieur au seuil ($Threshold). Envoi d'un email d'alerte." "TEST"

    try {
        # Création d'une copie temporaire du log pour l'attachement
        $TempLogFile = Join-Path $LogDirectory "temp_log_$TimeStampFile.log"
        Copy-Item -Path $LogFile -Destination $TempLogFile -Force

        # Création de l'objet MailMessage
        $mailMessage = New-Object System.Net.Mail.MailMessage
        $mailMessage.From = $SmtpUsername
        foreach ($recipient in $EmailRecipients) {
            $mailMessage.To.Add($recipient)
        }
        $mailMessage.Subject = $EmailSubject
        $mailMessage.Body = $EmailBody
        $mailMessage.IsBodyHtml = $true
        $mailMessage.SubjectEncoding = [System.Text.Encoding]::UTF8
        $mailMessage.BodyEncoding = [System.Text.Encoding]::UTF8

        # Ajouter le log (copie temporaire) en pièce jointe
        $attachment = New-Object System.Net.Mail.Attachment($TempLogFile)
        $mailMessage.Attachments.Add($attachment)
        
        # Configuration du client SMTP
        $smtpClient = New-Object System.Net.Mail.SmtpClient($SmtpServer, [int]$SmtpPort)
        $smtpClient.EnableSsl = $EnableTls
        if ($EnableSmtpAuth) {
            $smtpClient.Credentials = $SmtpCredential
        }
        
        # Envoi de l'email
        $smtpClient.Send($mailMessage)
        Write-Log "E-mail d'alerte envoyé avec succès." "TEST"
        
        # Libérer les ressources de l'attachement et du message
        $attachment.Dispose()
        $mailMessage.Dispose()
        
        # Suppression de la copie temporaire
        Remove-Item -Path $TempLogFile -Force
    }
    catch {
        Write-Log "Échec de l'envoi de l'e-mail d'alerte : $($_.Exception.Message)" "TEST"
    }
    
    Write-Log "Attente de $WaitTimeSeconds secondes avant l'exécution en mode PROD." "TEST"
    Start-Sleep -Seconds $WaitTimeSeconds
}
else {
    Write-Log "Nombre d'utilisateurs ($NbUsersRemoved) inférieur ou égal au seuil ($Threshold)." "TEST"
    if (-not $DisableWaitForTest) {
        Write-Log "Attente de 3600 secondes (limitation API Adobe) avant exécution en mode PROD." "TEST"
        Start-Sleep -Seconds 3600
    }
    else {
        Write-Log "Temps d'attente désactivé pour les tests. Passage immédiat en mode PROD." "TEST"
    }
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
