###########################################################################
# Script de détection - Espace disque faible avant mise à niveau Windows 11
###########################################################################
# Auteur      : Emanuel DE ALMEIDA
# Site        : www.ctrlaltnod.com
# Description : Ce script détecte si le lecteur C: dispose d'au moins 30 GB 
#               d'espace libre. Si l'espace est insuffisant, il calcule la 
#               taille combinée du dossier Téléchargements et de la Corbeille.
#               Si ces deux éléments totalisent plus de 500 MB, le script de 
#               remédiation sera déclenché pour aider l'utilisateur à libérer 
#               de l'espace.
#
#               Code de sortie :
#               - Exit 0 : Espace suffisant (>30GB) OU espace à récupérer <500MB
#               - Exit 1 : Espace insuffisant ET possibilité de libérer >500MB
#
# Version     : 1.0
# Date        : 17 octobre 2025
###########################################################################

# Définir le chemin du fichier de log
$logFile = Join-Path -Path $env:ProgramData -ChildPath "Microsoft\IntuneManagementExtension\Logs\Detect-LowDiskSpace.log"

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

# Configuration des seuils
$MinGBFree    = 30      # Espace minimum requis en GB
$MinReclaimMB = 500     # Espace minimum récupérable en MB pour déclencher la remédiation

# Démarrage du script
Write-Log -Message "=== Détection de l'espace disque disponible ===" -Level "INFO"
Write-Log -Message "Script de détection démarré - ctrlaltnod.com" -Level "INFO"
Write-Log -Message "Seuils configurés : Minimum $MinGBFree GB requis | $MinReclaimMB MB récupérables pour remédiation" -Level "INFO"

# Étape 1 : Vérifier l'espace libre sur le lecteur C:
Write-Log -Message "Étape 1/3 : Vérification de l'espace libre sur C:" -Level "INFO"

try {
    $drive = Get-PSDrive -Name C -ErrorAction Stop
    $freeBytes = $drive.Free
    $freeGB = [math]::Round($freeBytes / 1GB, 2)
    $totalGB = [math]::Round($drive.Used / 1GB + $freeGB, 2)
    $usedGB = [math]::Round($drive.Used / 1GB, 2)
    
    Write-Log -Message "Disque C: | Total: $totalGB GB | Utilisé: $usedGB GB | Libre: $freeGB GB" -Level "INFO"
    
    # Si l'espace libre est suffisant, pas besoin de continuer
    if ($freeBytes -ge ($MinGBFree * 1GB)) {
        Write-Log -Message "✓ Espace suffisant détecté ($freeGB GB libre > $MinGBFree GB requis)" -Level "INFO"
        Write-Log -Message "Résultat : CONFORME - Aucune action nécessaire" -Level "INFO"
        Write-Log -Message "=== Script de détection terminé avec succès ===" -Level "INFO"
        exit 0
    }
    
    Write-Log -Message "⚠ Espace insuffisant détecté ($freeGB GB libre < $MinGBFree GB requis)" -Level "WARN"
    Write-Log -Message "Analyse de l'espace récupérable en cours..." -Level "INFO"
    
} catch {
    Write-Log -Message "ERREUR lors de la lecture de l'espace disque C: - $_" -Level "ERROR"
    Write-Log -Message "=== Script de détection terminé avec erreur ===" -Level "ERROR"
    exit 1
}

# Étape 2 : Calculer la taille du dossier Téléchargements
Write-Log -Message "Étape 2/3 : Analyse du dossier Téléchargements" -Level "INFO"

$downloadsPath = Join-Path $env:USERPROFILE "Downloads"
$downloadBytes = 0

if (Test-Path -LiteralPath $downloadsPath) {
    try {
        $downloadItems = Get-ChildItem -LiteralPath $downloadsPath -Recurse -Force -File -ErrorAction SilentlyContinue
        $downloadBytes = ($downloadItems | Measure-Object -Sum Length).Sum
        
        if ($null -eq $downloadBytes) { $downloadBytes = 0 }
        
        $downloadMB = [math]::Round($downloadBytes / 1MB, 2)
        $fileCount = ($downloadItems | Measure-Object).Count
        
        Write-Log -Message "Dossier Téléchargements : $downloadMB MB ($fileCount fichiers)" -Level "INFO"
    } catch {
        Write-Log -Message "ERREUR lors de l'analyse du dossier Téléchargements : $_" -Level "WARN"
        $downloadBytes = 0
    }
} else {
    Write-Log -Message "Le dossier Téléchargements n'existe pas" -Level "INFO"
}

# Étape 3 : Calculer la taille de la Corbeille
Write-Log -Message "Étape 3/3 : Analyse de la Corbeille" -Level "INFO"

$RecycleBinBytes = 0

try {
    $shell = New-Object -ComObject Shell.Application
    $recycleBin = $shell.Namespace('shell:RecycleBinFolder')
    
    if ($recycleBin) {
        $itemCount = 0
        foreach ($item in $recycleBin.Items()) {
            try {
                $RecycleBinBytes += [int64]$item.Size
                $itemCount++
            } catch {
                # Ignorer les erreurs d'accès aux éléments individuels
            }
        }
        
        $recycleMB = [math]::Round($RecycleBinBytes / 1MB, 2)
        Write-Log -Message "Corbeille : $recycleMB MB ($itemCount éléments)" -Level "INFO"
    } else {
        Write-Log -Message "Impossible d'accéder à la Corbeille" -Level "WARN"
    }
} catch {
    Write-Log -Message "ERREUR lors de l'analyse de la Corbeille : $_" -Level "WARN"
    $RecycleBinBytes = 0
}

# Calcul de l'espace total récupérable
$combinedBytes = [int64]$downloadBytes + [int64]$RecycleBinBytes
$combinedMB = [math]::Round($combinedBytes / 1MB, 2)
$combinedGB = [math]::Round($combinedBytes / 1GB, 2)

Write-Log -Message "--- Résumé de l'analyse ---" -Level "INFO"
Write-Log -Message "Espace libre actuel : $freeGB GB" -Level "INFO"
Write-Log -Message "Espace récupérable total : $combinedMB MB ($combinedGB GB)" -Level "INFO"
Write-Log -Message "Téléchargements : $([math]::Round($downloadBytes / 1MB, 2)) MB" -Level "INFO"
Write-Log -Message "Corbeille : $([math]::Round($RecycleBinBytes / 1MB, 2)) MB" -Level "INFO"

# Décision finale : déclencher ou non la remédiation
if ($combinedBytes -ge ($MinReclaimMB * 1MB)) {
    Write-Log -Message "✓ Espace récupérable détecté : $combinedMB MB (seuil : $MinReclaimMB MB)" -Level "WARN"
    Write-Log -Message "Résultat : NON CONFORME - Remédiation nécessaire" -Level "WARN"
    Write-Log -Message "L'utilisateur sera invité à libérer de l'espace" -Level "WARN"
    Write-Log -Message "=== Script de détection terminé - Action requise ===" -Level "WARN"
    exit 1
} else {
    Write-Log -Message "⚠ Espace insuffisant MAIS espace récupérable trop faible : $combinedMB MB (seuil : $MinReclaimMB MB)" -Level "INFO"
    Write-Log -Message "Résultat : Remédiation automatique non recommandée" -Level "INFO"
    Write-Log -Message "L'intervention IT sera probablement nécessaire" -Level "INFO"
    Write-Log -Message "=== Script de détection terminé ===" -Level "INFO"
    exit 0
}
