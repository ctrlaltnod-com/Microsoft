###########################################################################
# Script de détection - Alignement de la barre des tâches Windows 11
###########################################################################
# Auteur      : Emanuel DE ALMEIDA
# Site        : www.ctrlaltnod.com
# Description : Ce script détecte si la barre des tâches Windows 11 est 
#               alignée à gauche (valeur 0) ou au centre (valeur 1).
#               Il vérifie la clé de registre TaskbarAl dans HKCU.
#               
#               Code de sortie :
#               - Exit 0 : Barre des tâches alignée à gauche (conforme)
#               - Exit 1 : Barre des tâches non alignée à gauche (remédiation nécessaire)
#
# Version     : 1.0
# Date        : 17 octobre 2025
###########################################################################

# Définir le chemin du fichier de log
$logFile = Join-Path -Path $env:ProgramData -ChildPath "Microsoft\IntuneManagementExtension\Logs\Detect-Taskbar-Alignment.log"

# Fonction pour écrire dans le journal de logs
function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [ValidateSet("INFO", "WARN", "ERROR")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp][$Level] $Message"
    Add-Content -Path $logFile -Value $logEntry
}

# Démarrage du script
Write-Log -Message "=== Détection de l'alignement de la barre des tâches ===" -Level "INFO"
Write-Log -Message "Script de détection démarré - ctrlaltnod.com" -Level "INFO"

# Définir le chemin de registre et le nom de la valeur
$registryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
$registryValueName = "TaskbarAl"

try {
    # Récupérer la valeur actuelle de l'alignement de la barre des tâches
    $taskbarAlignment = Get-ItemPropertyValue -Path $registryPath -Name $registryValueName -ErrorAction SilentlyContinue
    
    if ($null -eq $taskbarAlignment) {
        Write-Log -Message "La clé de registre TaskbarAl n'existe pas. Valeur par défaut : Centre (1)" -Level "WARN"
        Write-Log -Message "Résultat : NON CONFORME - Remédiation nécessaire" -Level "WARN"
        exit 1
    }
    
    # Vérifier si la barre des tâches est alignée à gauche (valeur = 0)
    if ($taskbarAlignment -eq 0) {
        Write-Log -Message "Alignement détecté : GAUCHE (valeur = 0)" -Level "INFO"
        Write-Log -Message "Résultat : CONFORME - Aucune action nécessaire" -Level "INFO"
        Write-Log -Message "=== Script de détection terminé avec succès ===" -Level "INFO"
        exit 0
    }
    else {
        Write-Log -Message "Alignement détecté : CENTRE (valeur = $taskbarAlignment)" -Level "WARN"
        Write-Log -Message "Résultat : NON CONFORME - Remédiation nécessaire" -Level "WARN"
        Write-Log -Message "=== Script de détection terminé - Action requise ===" -Level "WARN"
        exit 1
    }
}
catch {
    Write-Log -Message "ERREUR lors de la lecture du registre : $_" -Level "ERROR"
    Write-Log -Message "=== Script de détection terminé avec erreur ===" -Level "ERROR"
    exit 1
}
