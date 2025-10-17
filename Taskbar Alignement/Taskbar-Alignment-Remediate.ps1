###########################################################################
# Script de remédiation - Alignement de la barre des tâches Windows 11
###########################################################################
# Site        : www.ctrlaltnod.com
# Description : Ce script configure automatiquement l'alignement de la barre 
#               des tâches Windows 11 à GAUCHE en modifiant la valeur de 
#               registre TaskbarAl à 0 dans HKCU.
#               
#               La modification prend effet après redémarrage de l'Explorateur
#               Windows ou déconnexion/reconnexion de l'utilisateur.
#
#               Code de sortie :
#               - Exit 0 : Configuration réussie
#               - Exit 1 : Échec de la configuration
#
# Version     : 1.0
# Date        : 17 octobre 2025
###########################################################################

# Définir le chemin du fichier de log
$logFile = Join-Path -Path $env:ProgramData -ChildPath "Microsoft\IntuneManagementExtension\Logs\Remediate-Taskbar-Alignment.log"

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
Write-Log -Message "=== Remédiation de l'alignement de la barre des tâches ===" -Level "INFO"
Write-Log -Message "Script de remédiation démarré - ctrlaltnod.com" -Level "INFO"

# Définir le chemin de registre et le nom de la valeur
$registryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
$registryValueName = "TaskbarAl"
$targetValue = 0  # 0 = Gauche, 1 = Centre

try {
    # Vérifier si le chemin de registre existe
    if (-not (Test-Path -Path $registryPath)) {
        Write-Log -Message "Le chemin de registre n'existe pas. Création en cours..." -Level "WARN"
        New-Item -Path $registryPath -Force | Out-Null
        Write-Log -Message "Chemin de registre créé avec succès" -Level "INFO"
    }
    
    # Définir l'alignement de la barre des tâches à GAUCHE (valeur = 0)
    Write-Log -Message "Tentative de modification de la valeur TaskbarAl à 0 (Gauche)" -Level "INFO"
    Set-ItemProperty -Path $registryPath -Name $registryValueName -Value $targetValue -Type DWord -ErrorAction Stop
    
    # Vérifier que la modification a bien été appliquée
    $verifyValue = Get-ItemPropertyValue -Path $registryPath -Name $registryValueName -ErrorAction Stop
    
    if ($verifyValue -eq $targetValue) {
        Write-Log -Message "✓ Configuration réussie : Barre des tâches alignée à GAUCHE (valeur = $verifyValue)" -Level "INFO"
        Write-Log -Message "NOTE : Le changement sera visible après redémarrage de l'Explorateur Windows ou reconnexion" -Level "INFO"
        Write-Log -Message "=== Script de remédiation terminé avec succès ===" -Level "INFO"
        
        # Optionnel : Redémarrer l'Explorateur Windows pour appliquer immédiatement
        # Décommentez les lignes suivantes si vous souhaitez un effet immédiat
        # Write-Log -Message "Redémarrage de l'Explorateur Windows pour application immédiate..." -Level "INFO"
        # Stop-Process -Name explorer -Force
        # Start-Sleep -Seconds 2
        # Write-Log -Message "Explorateur Windows redémarré" -Level "INFO"
        
        exit 0
    }
    else {
        Write-Log -Message "✗ ERREUR : La vérification a échoué. Valeur attendue: $targetValue, valeur obtenue: $verifyValue" -Level "ERROR"
        Write-Log -Message "=== Script de remédiation terminé avec erreur ===" -Level "ERROR"
        exit 1
    }
}
catch {
    Write-Log -Message "✗ ERREUR lors de la modification du registre : $_" -Level "ERROR"
    Write-Log -Message "Détails : $($_.Exception.Message)" -Level "ERROR"
    Write-Log -Message "=== Script de remédiation terminé avec erreur ===" -Level "ERROR"
    exit 1
}
