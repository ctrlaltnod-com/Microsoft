#Requires -Version 5.1
<#
.SYNOPSIS
    Outil de Gestion Automapping Microsoft 365 - Version Française
    M365 Automapping Tool by ctrlaltnod.com

.DESCRIPTION
    Application PowerShell avec interface graphique pour gérer les permissions des boîtes aux lettres partagées
    et désactiver l'automapping dans Microsoft 365 Exchange Online.
    
    Fonctionnalités principales :
    - Interface WPF moderne et intuitive
    - Connexion sécurisée à Microsoft 365 Exchange Online
    - Suppression de toutes les permissions pour les utilisateurs sélectionnés
    - Désactivation de l'automapping tout en préservant l'accès complet
    - Suppression complète des permissions (Accès complet, Envoyer en tant que, Envoyer de la part de)
    - Ouverture automatique du navigateur pour l'authentification par périphérique
    - Gestion complète des erreurs et journalisation détaillée

.AUTHOR
    ctrlaltnod.com - Emanuel DE ALMEIDA

.VERSION
    1.0 Version 
.NOTES
    Prérequis :
    1. PowerShell 5.1 ou version ultérieure
    2. Module ExchangeOnlineManagement : Install-Module -Name ExchangeOnlineManagement -Force
    3. Permissions Administrateur Exchange ou Administrateur Global dans Microsoft 365
    
    Première utilisation :
    Unblock-File -Path ".\M365-Automapping-Tool.ps1"
    
    Développé par ctrlaltnod.com
#>

# Importation des assemblages requis pour WPF
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# Variables globales pour la synchronisation et l'état de l'application
$Global:SyncHash = [hashtable]::Synchronized(@{})  # Hashtable synchronisée pour les éléments UI
$Global:Connected = $false                          # État de la connexion Exchange Online
$Global:DeviceCode = $null                         # Code d'authentification de périphérique

# Définition XAML pour l'interface graphique - Entièrement en français
$Script:XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="M365 Automapping Tool by ctrlaltnod.com - Version Française" 
        Height="700" 
        Width="1000"
        WindowStartupLocation="CenterScreen"
        ResizeMode="CanMinimize">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- En-tête avec branding ctrlaltnod.com -->
        <Border Grid.Row="0" Background="#2E86AB" Padding="15">
            <StackPanel>
                <TextBlock Text="M365 Automapping Tool" 
                          FontSize="20" 
                          FontWeight="Bold" 
                          Foreground="White" 
                          HorizontalAlignment="Center"/>
                <TextBlock Text="by ctrlaltnod.com" 
                          FontSize="14" 
                          FontStyle="Italic"
                          Foreground="LightBlue" 
                          HorizontalAlignment="Center"
                          Margin="0,2,0,0"/>
            </StackPanel>
        </Border>
        
        <!-- Instructions pour le code d'authentification -->
        <GroupBox Grid.Row="1" Header="📱 Instructions Code d'Authentification" Margin="10" Background="#FFE4B5">
            <StackPanel>
                <TextBlock Text="Quand vous cliquez sur Connexion, regardez la FENÊTRE CONSOLE PowerShell pour le code d'authentification !" 
                          FontWeight="Bold" 
                          Foreground="Red"
                          TextWrapping="Wrap"
                          Margin="5"/>
                <TextBlock Text="Exemple : 'Pour vous connecter, utilisez un navigateur web pour ouvrir https://microsoft.com/devicelogin et entrez le code ABC123XYZ'" 
                          FontFamily="Courier New"
                          Background="LightGray"
                          Padding="5"
                          TextWrapping="Wrap"
                          Margin="5"/>
            </StackPanel>
        </GroupBox>
        
        <!-- Section de connexion Microsoft 365 -->
        <GroupBox Grid.Row="2" Header="🔐 Connexion Microsoft 365" Margin="10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <Label Grid.Column="0" Content="Email Administrateur :" VerticalAlignment="Center" FontWeight="Bold"/>
                <TextBox Grid.Column="1" Name="txtAdminEmail" Margin="5" Height="25" VerticalAlignment="Center"/>
                <Button Grid.Column="2" Name="btnConnect" Content="Se Connecter" Width="100" Height="30" Margin="5" Background="#4CAF50" Foreground="White"/>
                <Button Grid.Column="3" Name="btnDisconnect" Content="Se Déconnecter" Width="110" Height="30" Margin="5" Background="#f44336" Foreground="White" IsEnabled="False"/>
                <Button Grid.Column="4" Name="btnTestModule" Content="Tester Module" Width="100" Height="30" Margin="5" Background="#FF9800" Foreground="White"/>
            </Grid>
        </GroupBox>
        
        <!-- Statut de la connexion -->
        <Grid Grid.Row="3" Margin="10,0,10,10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Label Grid.Column="0" Content="Statut de Connexion :" FontWeight="Bold"/>
            <Label Grid.Column="1" Name="lblConnectionStatus" Content="Non Connecté" Foreground="Red" FontWeight="Bold"/>
        </Grid>
        
        <!-- Sélection utilisateur -->
        <GroupBox Grid.Row="4" Header="👤 Sélectionner l'Utilisateur" Margin="10,0,10,10" Name="grpUserSelection" IsEnabled="False">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <Label Grid.Column="0" Content="Choisir Utilisateur :" VerticalAlignment="Center" FontWeight="Bold"/>
                <ComboBox Grid.Column="1" Name="cmbUsers" Margin="5" Height="25" IsEditable="True"/>
                <Button Grid.Column="2" Name="btnRefreshUsers" Content="Actualiser Utilisateurs" Width="140" Height="30" Margin="5" Background="#2196F3" Foreground="White"/>
            </Grid>
        </GroupBox>
        
        <!-- Sélection boîte aux lettres partagée -->
        <GroupBox Grid.Row="5" Header="📮 Sélectionner la Boîte Partagée" Margin="10,0,10,10" Name="grpMailboxSelection" IsEnabled="False">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <Label Grid.Column="0" Content="Boîte Partagée :" VerticalAlignment="Center" FontWeight="Bold"/>
                <ComboBox Grid.Column="1" Name="cmbSharedMailboxes" Margin="5" Height="25" IsEditable="True"/>
                <Button Grid.Column="2" Name="btnRefreshMailboxes" Content="Actualiser Boîtes" Width="120" Height="30" Margin="5" Background="#2196F3" Foreground="White"/>
            </Grid>
        </GroupBox>
        
        <!-- Boutons d'actions -->
        <GroupBox Grid.Row="6" Header="⚡ Actions" Margin="10,0,10,10" Name="grpActions" IsEnabled="False">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                <Button Name="btnCheckPermissions" Content="Vérifier Permissions" Width="150" Height="40" Margin="5" Background="#9C27B0" Foreground="White"/>
                <Button Name="btnDisableAutomapping" Content="Désactiver Automapping" Width="160" Height="40" Margin="5" Background="#FF6B35" Foreground="White"/>
                <Button Name="btnRemoveAllPermissions" Content="Supprimer TOUTES Permissions" Width="180" Height="40" Margin="5" Background="#DC143C" Foreground="White"/>
            </StackPanel>
        </GroupBox>
        
        <!-- Journal d'activité -->
        <GroupBox Grid.Row="7" Header="📋 Journal d'Activité" Margin="10">
            <ScrollViewer VerticalScrollBarVisibility="Auto">
                <TextBox Name="txtLog" 
                         IsReadOnly="True" 
                         FontFamily="Courier New"
                         FontSize="10"
                         Background="#F5F5F5"
                         BorderThickness="0"
                         TextWrapping="Wrap"/>
            </ScrollViewer>
        </GroupBox>
        
        <!-- Barre de progression -->
        <ProgressBar Grid.Row="8" Name="progressBar" Height="20" Margin="10" Visibility="Collapsed"/>
        
        <!-- Barre de statut -->
        <StatusBar Grid.Row="9">
            <StatusBarItem>
                <TextBlock Name="txtStatus" Text="Prêt - Regardez la console pour le code d'authentification lors de la connexion" FontWeight="Bold"/>
            </StatusBarItem>
        </StatusBar>
    </Grid>
</Window>
"@

# Fonction pour écrire dans le journal d'activité avec horodatage
function Write-Log {
    param(
        [string]$Message,   # Message à enregistrer
        [string]$Level = "Info"  # Niveau : Info, Warning, Error
    )
    
    # Création de l'horodatage au format français
    $timestamp = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    try {
        # Mise à jour thread-safe du journal dans l'interface
        $Global:SyncHash.txtLog.Dispatcher.Invoke([action]{
            $Global:SyncHash.txtLog.AppendText("$logEntry`r`n")
            $Global:SyncHash.txtLog.ScrollToEnd()
        }, "Normal")
    } catch {
        # Fallback vers la console si l'interface n'est pas disponible
        Write-Host $logEntry -ForegroundColor $(if($Level -eq "Error"){"Red"}elseif($Level -eq "Warning"){"Yellow"}else{"White"})
    }
}

# Fonction pour mettre à jour la barre de statut
function Update-Status {
    param([string]$Status)  # Nouveau statut à afficher
    
    try {
        # Mise à jour thread-safe de la barre de statut
        $Global:SyncHash.txtStatus.Dispatcher.Invoke([action]{
            $Global:SyncHash.txtStatus.Text = $Status
        }, "Normal")
    } catch {
        # Fallback vers la console
        Write-Host "STATUT: $Status" -ForegroundColor Green
    }
}

# Fonction pour contrôler la barre de progression
function Set-ProgressBar {
    param(
        [bool]$Show,        # Afficher ou masquer la barre
        [int]$Value = 0     # Valeur de progression (0-100)
    )
    
    try {
        # Mise à jour thread-safe de la barre de progression
        $Global:SyncHash.progressBar.Dispatcher.Invoke([action]{
            if ($Show) {
                $Global:SyncHash.progressBar.Visibility = "Visible"
                $Global:SyncHash.progressBar.Value = $Value
            } else {
                $Global:SyncHash.progressBar.Visibility = "Collapsed"
            }
        }, "Normal")
    } catch {
        # Fallback vers la console
        Write-Host "PROGRESSION: $Value%" -ForegroundColor Yellow
    }
}

# Fonction pour tester la disponibilité du module Exchange Online
function Test-ExchangeModule {
    Write-Log "Test du module ExchangeOnlineManagement..." "Info"
    Update-Status "Test du module en cours..."
    
    try {
        # Vérification de l'installation du module
        $module = Get-Module -ListAvailable -Name ExchangeOnlineManagement
        if (-not $module) {
            $message = @"
Le module ExchangeOnlineManagement n'est PAS installé.

Commande d'installation :
Install-Module -Name ExchangeOnlineManagement -Force

Puis redémarrez cette application.
"@
            Write-Log "Module ExchangeOnlineManagement non trouvé" "Error"
            [System.Windows.MessageBox]::Show($message, "Module Manquant", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            return $false
        }
        
        # Tentative d'importation du module
        Import-Module ExchangeOnlineManagement -Force -ErrorAction Stop
        $moduleVersion = (Get-Module ExchangeOnlineManagement).Version.ToString()
        
        $message = "Module ExchangeOnlineManagement v$moduleVersion est prêt !"
        Write-Log $message "Info"
        Update-Status "Test du module réussi"
        
        [System.Windows.MessageBox]::Show($message, "Test Module Réussi", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return $true
        
    } catch {
        $message = "Échec du test du module : $($_.Exception.Message)"
        Write-Log $message "Error"
        Update-Status "Échec du test du module"
        [System.Windows.MessageBox]::Show($message, "Échec Test Module", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return $false
    }
}

# Fonction principale pour se connecter à Exchange Online
function Connect-ExchangeOnline365 {
    param([string]$UserPrincipalName)  # Email de l'administrateur
    
    # Validation de l'email administrateur
    if ([string]::IsNullOrWhiteSpace($UserPrincipalName)) {
        [System.Windows.MessageBox]::Show("Veuillez saisir une adresse email administrateur.", "Information Manquante", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return $false
    }
    
    Write-Log "Démarrage de la connexion à Exchange Online en tant que : $UserPrincipalName" "Info"
    Update-Status "Connexion à Exchange Online..."
    Set-ProgressBar -Show $true -Value 20
    
    # Affichage des instructions d'authentification
    $instructionMessage = @"
AUTHENTIFICATION PAR CODE D'APPAREIL

Instructions importantes :

1. Après avoir cliqué sur OK, le code d'appareil apparaîtra dans la FENÊTRE CONSOLE PowerShell (derrière cette interface)

2. Recherchez un message comme :
   "Pour vous connecter, utilisez un navigateur web pour ouvrir la page https://microsoft.com/devicelogin et entrez le code XXXXXXXX pour vous authentifier."

3. Copiez le code depuis la console
4. Ouvrez un navigateur vers https://microsoft.com/devicelogin
5. Entrez le code et terminez l'authentification

Le code d'appareil ressemblera à : ABC123XYZ

Cliquez sur OK pour commencer l'authentification...
"@
    
    [System.Windows.MessageBox]::Show($instructionMessage, "Instructions d'Authentification", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
    
    try {
        # Importation du module Exchange Online
        Import-Module ExchangeOnlineManagement -Force
        Set-ProgressBar -Show $true -Value 40
        
        Write-Log "Lancement de l'authentification par appareil..." "Info"
        Update-Status "Authentification par appareil en cours..."
        
        # Instructions dans la console
        Write-Host ""
        Write-Host "================================================" -ForegroundColor Yellow
        Write-Host "DÉBUT DE L'AUTHENTIFICATION PAR APPAREIL" -ForegroundColor Yellow
        Write-Host "Recherchez le code d'appareil ci-dessous..." -ForegroundColor Yellow
        Write-Host "================================================" -ForegroundColor Yellow
        Write-Host ""
        
        # Connexion avec authentification par appareil
        # Le code d'appareil apparaîtra automatiquement dans la console
        Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -Device -ShowBanner:$false
        
        # Si nous arrivons ici, l'authentification a réussi
        Set-ProgressBar -Show $true -Value 90
        
        # Test de la connexion
        $testResult = Get-EXOMailbox -ResultSize 1 -ErrorAction Stop
        
        # Mise à jour de l'interface pour le succès
        $Global:SyncHash.lblConnectionStatus.Content = "Connecté en tant que $UserPrincipalName"
        $Global:SyncHash.lblConnectionStatus.Foreground = "Green"
        $Global:SyncHash.btnConnect.IsEnabled = $false
        $Global:SyncHash.btnDisconnect.IsEnabled = $true
        $Global:SyncHash.grpUserSelection.IsEnabled = $true
        $Global:SyncHash.grpMailboxSelection.IsEnabled = $true
        $Global:SyncHash.grpActions.IsEnabled = $true
        
        Set-ProgressBar -Show $true -Value 100
        Start-Sleep -Seconds 1
        Set-ProgressBar -Show $false
        
        # Mise à jour des variables globales
        $Global:Connected = $true
        Write-Log "Connexion réussie à Exchange Online !" "Info"
        Update-Status "Connecté à Exchange Online"
        
        # Message de succès dans la console
        Write-Host ""
        Write-Host "================================================" -ForegroundColor Green
        Write-Host "AUTHENTIFICATION RÉUSSIE !" -ForegroundColor Green
        Write-Host "Connecté à Exchange Online" -ForegroundColor Green
        Write-Host "================================================" -ForegroundColor Green
        Write-Host ""
        
        [System.Windows.MessageBox]::Show("Connexion réussie à Exchange Online !", "Connexion Réussie", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return $true
        
    } catch {
        # Gestion des erreurs de connexion
        $errorMessage = $_.Exception.Message
        Write-Log "Échec de la connexion : $errorMessage" "Error"
        Set-ProgressBar -Show $false
        Update-Status "Échec de la connexion"
        
        # Message d'erreur dans la console
        Write-Host ""
        Write-Host "================================================" -ForegroundColor Red
        Write-Host "ÉCHEC DE L'AUTHENTIFICATION" -ForegroundColor Red
        Write-Host "Erreur : $errorMessage" -ForegroundColor Red
        Write-Host "================================================" -ForegroundColor Red
        Write-Host ""
        
        [System.Windows.MessageBox]::Show("Échec de la connexion : $errorMessage", "Erreur de Connexion", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return $false
    }
}

# Fonction pour se déconnecter d'Exchange Online
function Disconnect-ExchangeOnline365 {
    try {
        # Déconnexion si connecté
        if ($Global:Connected) {
            Disconnect-ExchangeOnline -Confirm:$false
            Write-Log "Déconnecté d'Exchange Online" "Info"
        }
        
        # Réinitialisation de l'interface utilisateur
        $Global:SyncHash.lblConnectionStatus.Content = "Non Connecté"
        $Global:SyncHash.lblConnectionStatus.Foreground = "Red"
        $Global:SyncHash.btnConnect.IsEnabled = $true
        $Global:SyncHash.btnDisconnect.IsEnabled = $false
        $Global:SyncHash.grpUserSelection.IsEnabled = $false
        $Global:SyncHash.grpMailboxSelection.IsEnabled = $false
        $Global:SyncHash.grpActions.IsEnabled = $false
        $Global:SyncHash.cmbUsers.Items.Clear()
        $Global:SyncHash.cmbSharedMailboxes.Items.Clear()
        
        # Réinitialisation des variables globales
        $Global:Connected = $false
        Update-Status "Déconnecté"
        
    } catch {
        Write-Log "Erreur lors de la déconnexion : $($_.Exception.Message)" "Error"
    }
}

# Fonction pour actualiser la liste des utilisateurs
function Refresh-Users {
    # Vérification de la connexion
    if (-not $Global:Connected) {
        [System.Windows.MessageBox]::Show("Veuillez vous connecter à Exchange Online d'abord.", "Non Connecté", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    
    Write-Log "Chargement des boîtes aux lettres utilisateur..." "Info"
    Update-Status "Chargement des utilisateurs..."
    Set-ProgressBar -Show $true -Value 30
    
    try {
        # Récupération des boîtes aux lettres utilisateur (limité à 50 pour les performances)
        $users = Get-EXOMailbox -RecipientTypeDetails UserMailbox -ResultSize 50 | 
                 Select-Object DisplayName, PrimarySmtpAddress | 
                 Sort-Object DisplayName
        
        Set-ProgressBar -Show $true -Value 80
        
        # Remplissage de la liste déroulante des utilisateurs
        $Global:SyncHash.cmbUsers.Items.Clear()
        foreach ($user in $users) {
            $displayText = "$($user.DisplayName) ($($user.PrimarySmtpAddress))"
            $Global:SyncHash.cmbUsers.Items.Add($displayText)
        }
        
        Write-Log "Chargé $($users.Count) boîtes aux lettres utilisateur" "Info"
        Set-ProgressBar -Show $false
        Update-Status "Utilisateurs chargés avec succès"
        
    } catch {
        Write-Log "Erreur lors du chargement des utilisateurs : $($_.Exception.Message)" "Error"
        Set-ProgressBar -Show $false
        Update-Status "Erreur lors du chargement des utilisateurs"
    }
}

# Fonction pour actualiser la liste des boîtes aux lettres partagées
function Refresh-SharedMailboxes {
    # Vérification de la connexion
    if (-not $Global:Connected) {
        [System.Windows.MessageBox]::Show("Veuillez vous connecter à Exchange Online d'abord.", "Non Connecté", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    
    Write-Log "Chargement des boîtes aux lettres partagées..." "Info"
    Update-Status "Chargement des boîtes partagées..."
    Set-ProgressBar -Show $true -Value 30
    
    try {
        # Récupération des boîtes aux lettres partagées (limité à 50 pour les performances)
        $sharedMailboxes = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize 50 | 
                          Select-Object DisplayName, PrimarySmtpAddress | 
                          Sort-Object DisplayName
        
        Set-ProgressBar -Show $true -Value 80
        
        # Remplissage de la liste déroulante des boîtes partagées
        $Global:SyncHash.cmbSharedMailboxes.Items.Clear()
        foreach ($mailbox in $sharedMailboxes) {
            $displayText = "$($mailbox.DisplayName) ($($mailbox.PrimarySmtpAddress))"
            $Global:SyncHash.cmbSharedMailboxes.Items.Add($displayText)
        }
        
        Write-Log "Chargé $($sharedMailboxes.Count) boîtes aux lettres partagées" "Info"
        Set-ProgressBar -Show $false
        Update-Status "Boîtes partagées chargées avec succès"
        
    } catch {
        Write-Log "Erreur lors du chargement des boîtes partagées : $($_.Exception.Message)" "Error"
        Set-ProgressBar -Show $false
        Update-Status "Erreur lors du chargement des boîtes partagées"
    }
}

# Fonction utilitaire pour extraire l'adresse email d'une sélection
function Get-EmailFromSelection {
    param([string]$Selection)  # Texte sélectionné au format "Nom (email@domain.com)"
    
    # Expression régulière pour extraire l'email entre parenthèses
    if ($Selection -match '\(([^)]+)\)$') {
        return $matches[1]
    }
    return $Selection
}

# Fonction pour vérifier les permissions actuelles
function Check-Permissions {
    # Récupération des sélections utilisateur
    $selectedUser = $Global:SyncHash.cmbUsers.SelectedItem
    $selectedMailbox = $Global:SyncHash.cmbSharedMailboxes.SelectedItem
    
    # Validation des sélections
    if (-not $selectedUser -or -not $selectedMailbox) {
        [System.Windows.MessageBox]::Show("Veuillez sélectionner à la fois un utilisateur et une boîte aux lettres partagée.", "Sélection Requise", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    
    # Extraction des adresses email
    $userEmail = Get-EmailFromSelection -Selection $selectedUser
    $mailboxEmail = Get-EmailFromSelection -Selection $selectedMailbox
    
    Write-Log "Vérification des permissions pour $userEmail sur $mailboxEmail" "Info"
    Update-Status "Vérification des permissions..."
    Set-ProgressBar -Show $true -Value 50
    
    try {
        # Vérification des permissions de boîte aux lettres (Accès complet)
        $permissions = Get-EXOMailboxPermission -Identity $mailboxEmail -User $userEmail -ErrorAction SilentlyContinue
        
        if ($permissions) {
            Write-Log "Permissions de boîte aux lettres actuelles trouvées :" "Info"
            foreach ($perm in $permissions) {
                Write-Log "  - Droits d'accès : $($perm.AccessRights -join ', ')" "Info"
                Write-Log "  - Hérité : $($perm.IsInherited)" "Info"
                Write-Log "  - Refusé : $($perm.Deny)" "Info"
            }
        } else {
            Write-Log "Aucune permission explicite de boîte aux lettres trouvée" "Info"
        }
        
        # Vérification des permissions "Envoyer en tant que"
        $recipientPerms = Get-EXORecipientPermission -Identity $mailboxEmail -Trustee $userEmail -ErrorAction SilentlyContinue
        if ($recipientPerms) {
            Write-Log "Permissions 'Envoyer en tant que' trouvées :" "Info"
            foreach ($perm in $recipientPerms) {
                Write-Log "  - Droits 'Envoyer en tant que' : $($perm.AccessRights -join ', ')" "Info"
            }
        } else {
            Write-Log "Aucune permission 'Envoyer en tant que' trouvée" "Info"
        }
        
        # Vérification des permissions "Envoyer de la part de"
        $mailbox = Get-EXOMailbox -Identity $mailboxEmail -Properties GrantSendOnBehalfTo -ErrorAction SilentlyContinue
        if ($mailbox.GrantSendOnBehalfTo -and $mailbox.GrantSendOnBehalfTo -contains $userEmail) {
            Write-Log "Permissions 'Envoyer de la part de' trouvées pour cet utilisateur" "Info"
        } else {
            Write-Log "Aucune permission 'Envoyer de la part de' trouvée" "Info"
        }
        
        Set-ProgressBar -Show $false
        Update-Status "Vérification des permissions terminée"
        
        [System.Windows.MessageBox]::Show("Vérification des permissions terminée. Consultez le journal d'activité pour les détails.", "Vérification Terminée", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        
    } catch {
        Write-Log "Erreur lors de la vérification des permissions : $($_.Exception.Message)" "Error"
        Set-ProgressBar -Show $false
        Update-Status "Erreur lors de la vérification des permissions"
        [System.Windows.MessageBox]::Show("Erreur lors de la vérification des permissions : $($_.Exception.Message)", "Erreur", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
}

# Fonction principale pour désactiver l'automapping (fonctionnalité principale)
function Disable-Automapping {
    # Récupération des sélections utilisateur
    $selectedUser = $Global:SyncHash.cmbUsers.SelectedItem
    $selectedMailbox = $Global:SyncHash.cmbSharedMailboxes.SelectedItem
    
    # Validation des sélections
    if (-not $selectedUser -or -not $selectedMailbox) {
        [System.Windows.MessageBox]::Show("Veuillez sélectionner à la fois un utilisateur et une boîte aux lettres partagée.", "Sélection Requise", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    
    # Extraction des adresses email
    $userEmail = Get-EmailFromSelection -Selection $selectedUser
    $mailboxEmail = Get-EmailFromSelection -Selection $selectedMailbox
    
    # Dialogue de confirmation avec explication détaillée
    $confirmMessage = @"
Ceci va désactiver l'automapping pour :

Utilisateur : $userEmail
Boîte Partagée : $mailboxEmail

Processus :
1. Supprimer les permissions d'Accès Complet existantes
2. Re-ajouter les permissions d'Accès Complet avec AutoMapping désactivé

L'utilisateur conservera l'accès à la boîte aux lettres mais celle-ci n'apparaîtra plus automatiquement dans Outlook.

Voulez-vous continuer ?
"@
    
    $result = [System.Windows.MessageBox]::Show($confirmMessage, "Confirmer Désactivation Automapping", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
    
    if ($result -ne [System.Windows.MessageBoxResult]::Yes) {
        Write-Log "Opération de désactivation automapping annulée par l'utilisateur" "Info"
        return
    }
    
    Write-Log "Désactivation de l'automapping pour $userEmail sur $mailboxEmail" "Info"
    Update-Status "Désactivation de l'automapping..."
    Set-ProgressBar -Show $true -Value 25
    
    try {
        # Étape 1 : Suppression des permissions d'accès complet existantes
        Write-Log "Étape 1 : Suppression des permissions d'Accès Complet existantes..." "Info"
        Remove-MailboxPermission -Identity $mailboxEmail -User $userEmail -AccessRights FullAccess -Confirm:$false -ErrorAction SilentlyContinue
        Write-Log "Permissions d'Accès Complet existantes supprimées (s'il y en avait)" "Info"
        
        Set-ProgressBar -Show $true -Value 75
        
        # Étape 2 : Re-ajout avec automapping désactivé
        Write-Log "Étape 2 : Re-ajout des permissions d'Accès Complet avec AutoMapping désactivé..." "Info"
        Add-MailboxPermission -Identity $mailboxEmail -User $userEmail -AccessRights FullAccess -AutoMapping:$false -Confirm:$false
        Write-Log "Permissions d'Accès Complet ajoutées avec AutoMapping désactivé" "Info"
        
        Set-ProgressBar -Show $true -Value 100
        Start-Sleep -Seconds 1
        Set-ProgressBar -Show $false
        
        Write-Log "Automapping désactivé avec succès !" "Info"
        Update-Status "Automapping désactivé avec succès"
        
        # Message de succès détaillé
        $successMessage = @"
Automapping désactivé avec succès !

Utilisateur : $userEmail
Boîte Partagée : $mailboxEmail

Notes importantes :
- L'utilisateur a toujours l'Accès Complet à la boîte aux lettres
- La boîte aux lettres n'apparaîtra plus automatiquement dans Outlook
- L'utilisateur doit redémarrer Outlook pour voir les changements
- L'utilisateur peut ajouter manuellement la boîte aux lettres si nécessaire

Opération terminée avec succès.
"@
        
        [System.Windows.MessageBox]::Show($successMessage, "Automapping Désactivé", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        
    } catch {
        Write-Log "Erreur lors de la désactivation de l'automapping : $($_.Exception.Message)" "Error"
        Set-ProgressBar -Show $false
        Update-Status "Erreur lors de la désactivation de l'automapping"
        [System.Windows.MessageBox]::Show("Erreur lors de la désactivation de l'automapping : $($_.Exception.Message)", "Erreur", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
}

# Fonction pour supprimer toutes les permissions
function Remove-AllPermissions {
    # Récupération des sélections utilisateur
    $selectedUser = $Global:SyncHash.cmbUsers.SelectedItem
    $selectedMailbox = $Global:SyncHash.cmbSharedMailboxes.SelectedItem
    
    # Validation des sélections
    if (-not $selectedUser -or -not $selectedMailbox) {
        [System.Windows.MessageBox]::Show("Veuillez sélectionner à la fois un utilisateur et une boîte aux lettres partagée.", "Sélection Requise", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    
    # Extraction des adresses email
    $userEmail = Get-EmailFromSelection -Selection $selectedUser
    $mailboxEmail = Get-EmailFromSelection -Selection $selectedMailbox
    
    # Dialogue de confirmation avec avertissement fort
    $confirmMessage = @"
ATTENTION : Ceci va supprimer TOUTES les permissions pour :

Utilisateur : $userEmail
Boîte Partagée : $mailboxEmail

Cela inclut :
✗ Permissions d'Accès Complet
✗ Permissions 'Envoyer en tant que'
✗ Permissions 'Envoyer de la part de'

Après cette opération :
- L'utilisateur perdra TOUT accès à la boîte aux lettres
- L'utilisateur ne pourra plus ouvrir ou accéder à la boîte aux lettres
- L'utilisateur ne pourra plus envoyer d'emails en tant que ou de la part de cette boîte

Cette action ne peut pas être facilement annulée.

Êtes-vous absolument sûr de vouloir continuer ?
"@
    
    $result = [System.Windows.MessageBox]::Show($confirmMessage, "Confirmer Suppression TOUTES Permissions", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
    
    if ($result -ne [System.Windows.MessageBoxResult]::Yes) {
        Write-Log "Opération de suppression de toutes les permissions annulée par l'utilisateur" "Info"
        return
    }
    
    Write-Log "Suppression de TOUTES les permissions pour $userEmail sur $mailboxEmail" "Info"
    Update-Status "Suppression de toutes les permissions..."
    Set-ProgressBar -Show $true -Value 33
    
    try {
        # Suppression des permissions d'Accès Complet
        Write-Log "Suppression des permissions d'Accès Complet..." "Info"
        Remove-MailboxPermission -Identity $mailboxEmail -User $userEmail -AccessRights FullAccess -Confirm:$false -ErrorAction SilentlyContinue
        Write-Log "Permissions d'Accès Complet supprimées" "Info"
        
        Set-ProgressBar -Show $true -Value 66
        
        # Suppression des permissions "Envoyer en tant que"
        Write-Log "Suppression des permissions 'Envoyer en tant que'..." "Info"
        Remove-RecipientPermission -Identity $mailboxEmail -Trustee $userEmail -AccessRights SendAs -Confirm:$false -ErrorAction SilentlyContinue
        Write-Log "Permissions 'Envoyer en tant que' supprimées" "Info"
        
        Set-ProgressBar -Show $true -Value 85
        
        # Suppression des permissions "Envoyer de la part de"
        Write-Log "Suppression des permissions 'Envoyer de la part de'..." "Info"
        try {
            $mailbox = Get-EXOMailbox -Identity $mailboxEmail -Properties GrantSendOnBehalfTo
            if ($mailbox.GrantSendOnBehalfTo -and $mailbox.GrantSendOnBehalfTo -contains $userEmail) {
                $newGrantList = $mailbox.GrantSendOnBehalfTo | Where-Object { $_ -ne $userEmail }
                Set-Mailbox -Identity $mailboxEmail -GrantSendOnBehalfTo $newGrantList -ErrorAction SilentlyContinue
                Write-Log "Permissions 'Envoyer de la part de' supprimées" "Info"
            } else {
                Write-Log "Aucune permission 'Envoyer de la part de' trouvée à supprimer" "Info"
            }
        } catch {
            Write-Log "Suppression des permissions 'Envoyer de la part de' : $($_.Exception.Message)" "Warning"
        }
        
        Set-ProgressBar -Show $true -Value 100
        Start-Sleep -Seconds 1
        Set-ProgressBar -Show $false
        
        Write-Log "TOUTES les permissions supprimées avec succès !" "Info"
        Update-Status "Toutes les permissions supprimées"
        
        # Message de succès
        $successMessage = @"
Toutes les permissions ont été supprimées avec succès !

Utilisateur : $userEmail
Boîte Partagée : $mailboxEmail

Permissions supprimées :
✓ Permissions d'Accès Complet
✓ Permissions 'Envoyer en tant que'
✓ Permissions 'Envoyer de la part de'

L'utilisateur n'a maintenant AUCUN accès à cette boîte aux lettres partagée.

Opération terminée avec succès.
"@
        
        [System.Windows.MessageBox]::Show($successMessage, "Toutes les Permissions Supprimées", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        
    } catch {
        Write-Log "Erreur lors de la suppression des permissions : $($_.Exception.Message)" "Error"
        Set-ProgressBar -Show $false
        Update-Status "Erreur lors de la suppression des permissions"
        [System.Windows.MessageBox]::Show("Erreur lors de la suppression des permissions : $($_.Exception.Message)", "Erreur", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
}

# Fonction principale pour créer et afficher l'interface graphique
function Show-AutomappingManagerGUI {
    try {
        # Analyse du XAML et création de la fenêtre
        [xml]$xamlXml = $Script:XAML
        $reader = New-Object System.Xml.XmlNodeReader $xamlXml
        $Global:SyncHash.Window = [Windows.Markup.XamlReader]::Load($reader)
        
        # Récupération de tous les éléments UI nommés
        $xamlXml.SelectNodes("//*[@Name]") | ForEach-Object {
            $Global:SyncHash.($_.Name) = $Global:SyncHash.Window.FindName($_.Name)
        }
        
        Write-Log "Application démarrée - M365 Automapping Tool by ctrlaltnod.com" "Info"
        Update-Status "Prêt - Regardez la console pour le code d'authentification lors de la connexion"
        
        # Liaison des gestionnaires d'événements avec syntaxe appropriée
        
        # Bouton Test Module
        $Global:SyncHash.btnTestModule.Add_Click({
            Test-ExchangeModule
        })
        
        # Bouton Se Connecter
        $Global:SyncHash.btnConnect.Add_Click({
            $adminEmail = $Global:SyncHash.txtAdminEmail.Text.Trim()
            if ([string]::IsNullOrWhiteSpace($adminEmail)) {
                [System.Windows.MessageBox]::Show("Veuillez saisir une adresse email administrateur.", "Information Manquante", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }
            
            # Désactivation du bouton pendant la tentative de connexion
            $Global:SyncHash.btnConnect.IsEnabled = $false
            try {
                Connect-ExchangeOnline365 -UserPrincipalName $adminEmail
            } finally {
                # Réactivation du bouton si la connexion a échoué
                if (-not $Global:Connected) {
                    $Global:SyncHash.btnConnect.IsEnabled = $true
                }
            }
        })
        
        # Bouton Se Déconnecter
        $Global:SyncHash.btnDisconnect.Add_Click({
            Disconnect-ExchangeOnline365
        })
        
        # Bouton Actualiser Utilisateurs
        $Global:SyncHash.btnRefreshUsers.Add_Click({
            Refresh-Users
        })
        
        # Bouton Actualiser Boîtes aux Lettres
        $Global:SyncHash.btnRefreshMailboxes.Add_Click({
            Refresh-SharedMailboxes
        })
        
        # Bouton Vérifier Permissions
        $Global:SyncHash.btnCheckPermissions.Add_Click({
            Check-Permissions
        })
        
        # Bouton Désactiver Automapping (fonctionnalité principale)
        $Global:SyncHash.btnDisableAutomapping.Add_Click({
            Disable-Automapping
        })
        
        # Bouton Supprimer Toutes les Permissions
        $Global:SyncHash.btnRemoveAllPermissions.Add_Click({
            Remove-AllPermissions
        })
        
        # Événement de fermeture de fenêtre
        $Global:SyncHash.Window.Add_Closing({
            Write-Log "Fermeture de l'application" "Info"
            if ($Global:Connected) {
                Write-Log "Nettoyage de la connexion Exchange Online..." "Info"
                try {
                    Disconnect-ExchangeOnline365
                } catch {
                    Write-Log "Erreur lors du nettoyage : $($_.Exception.Message)" "Warning"
                }
            }
        })
        
        # Affichage des informations système
        $psVersion = $PSVersionTable.PSVersion.ToString()
        $osVersion = [System.Environment]::OSVersion.VersionString
        Write-Log "=== M365 Automapping Tool by ctrlaltnod.com ===" "Info"
        Write-Log "Informations Système :" "Info"
        Write-Log "Version PowerShell : $psVersion" "Info"
        Write-Log "Système d'Exploitation : $osVersion" "Info"
        Write-Log "Application prête à l'utilisation" "Info"
        
        # Affichage de la fenêtre
        $null = $Global:SyncHash.Window.ShowDialog()
        
    } catch {
        # Gestion des erreurs critiques de création de l'interface
        $errorMessage = "Erreur critique lors de la création de l'interface : $($_.Exception.Message)"
        Write-Host $errorMessage -ForegroundColor Red
        Write-Host "Trace de la pile : $($_.ScriptStackTrace)" -ForegroundColor Red
        [System.Windows.MessageBox]::Show($errorMessage, "Erreur Critique Interface", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
}

# Fonction pour vérifier les prérequis système
function Test-Prerequisites {
    Write-Host "Vérification des prérequis système..." -ForegroundColor Yellow
    $allGood = $true
    
    # Vérification de la version PowerShell
    if ($PSVersionTable.PSVersion.Major -lt 5) {
        Write-Host "❌ ERREUR : PowerShell 5.1 ou version ultérieure requis (trouvé $($PSVersionTable.PSVersion))" -ForegroundColor Red
        $allGood = $false
    } else {
        Write-Host "✅ Version PowerShell : $($PSVersionTable.PSVersion)" -ForegroundColor Green
    }
    
    # Vérification si exécuté en tant qu'administrateur (recommandé mais pas obligatoire)
    try {
        $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
        $isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        
        if ($isAdmin) {
            Write-Host "✅ Exécuté en tant qu'Administrateur" -ForegroundColor Green
        } else {
            Write-Host "⚠️  Avertissement : Non exécuté en tant qu'Administrateur (certaines fonctionnalités peuvent ne pas fonctionner de manière optimale)" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "⚠️  Impossible de déterminer le statut administrateur" -ForegroundColor Yellow
    }
    
    # Vérification de la politique d'exécution
    $executionPolicy = Get-ExecutionPolicy
    if ($executionPolicy -eq "Restricted") {
        Write-Host "❌ ERREUR : La Politique d'Exécution est Restreinte. Exécutez : Set-ExecutionPolicy RemoteSigned -Scope CurrentUser" -ForegroundColor Red
        $allGood = $false
    } else {
        Write-Host "✅ Politique d'Exécution : $executionPolicy" -ForegroundColor Green
    }
    
    Write-Host ""
    if ($allGood) {
        Write-Host "✅ Tous les prérequis sont satisfaits !" -ForegroundColor Green
    } else {
        Write-Host "❌ Veuillez résoudre les problèmes ci-dessus avant de continuer." -ForegroundColor Red
    }
    
    return $allGood
}

# Fonction principale d'exécution
function Main {
    # Affichage de la bannière avec branding ctrlaltnod.com
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "   M365 Automapping Tool by ctrlaltnod.com" -ForegroundColor Cyan
    Write-Host "   Solution complète pour gérer l'automapping des boîtes partagées" -ForegroundColor Cyan
    Write-Host "   Version Française avec commentaires détaillés" -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""
    
    # Vérification si le fichier script est bloqué (avertissement de sécurité)
    $scriptPath = $MyInvocation.MyCommand.Path
    if ($scriptPath) {
        try {
            $zones = Get-Content -Path "$scriptPath" -Stream Zone.Identifier -ErrorAction SilentlyContinue
            if ($zones) {
                Write-Host "🔒 AVIS DE SÉCURITÉ : Ce script est bloqué par la sécurité Windows." -ForegroundColor Yellow
                Write-Host "   Pour supprimer l'avertissement de sécurité, exécutez cette commande :" -ForegroundColor Yellow
                Write-Host "   Unblock-File -Path '$scriptPath'" -ForegroundColor White
                Write-Host ""
                $response = Read-Host "Continuer quand même ? (o/n)"
                if ($response -ne 'o' -and $response -ne 'O') {
                    Write-Host "Fermeture..." -ForegroundColor Yellow
                    return
                }
                Write-Host ""
            }
        } catch {
            # La vérification Zone.Identifier a échoué, continuer quand même
        }
    }
    
    # Vérification des prérequis
    if (-not (Test-Prerequisites)) {
        Write-Host ""
        Write-Host "Veuillez résoudre les problèmes de prérequis et réessayer." -ForegroundColor Red
        Read-Host "Appuyez sur Entrée pour quitter"
        return
    }
    
    Write-Host ""
    Write-Host "IMPORTANT : Le code d'authentification apparaîtra dans cette fenêtre console !" -ForegroundColor Yellow
    Write-Host "Recherchez un message comme : 'Pour vous connecter, utilisez un navigateur web...'" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Démarrage de l'interface graphique..." -ForegroundColor Green
    Write-Host ""
    
    # Lancement de l'interface graphique
    try {
        Show-AutomappingManagerGUI
    } catch {
        Write-Host "Échec du démarrage de l'application : $($_.Exception.Message)" -ForegroundColor Red
        Read-Host "Appuyez sur Entrée pour quitter"
    }
}

# Exécution de la fonction principale
Main