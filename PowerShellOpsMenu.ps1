# ==============================================
# SCRIPT HEADER
# ==============================================
# Script PowerShell pour Réparation et Configuration
# Auteur : Maxime Simard
# Date : 26-09-2024
# Description : Script De Configuration Divers
# ==============================================

# ==============================================
# Configuration globale et variables
# ==============================================
# Déclare ici toutes les variables globales

$NOM_DU_SCRIPT = "PowerShell Ops Menu"
$VERSION_LOGICIEL = "0.62"

$LIEN_DOWNLOAD_EVERYTHING = "https://www.voidtools.com/Everything-1.4.1.1009.x64-Setup.exe"
$LIEN_DOWNLOAD_SERTI = "https://rdt.serti.com/Upload/SDSSuite_Complete_Install.exe"
$LIEN_DOWNLOAD_WINGET = "https://aka.ms/getwinget"

# ==============================================
# Configuration variables de Styles Globales
# ==============================================
# Déclare ici toutes les variables de Styles

#Black - DarkBlue - DarkGreen - DarkCyan - DarkRed - DarkMagenta - DarkYellow - Gray - DarkGray - Blue - Green - Cyan - Red - Magenta - Yellow - White
$BACKGROUND_COLOR = "Black"      # Fond sombre mais pas aussi intense que le noir
$DEFAULT_TEXT_COLOR = "White"       # Texte par défaut en blanc pour un bon contraste avec le fond

$TITLE_COLOR = "Blue"           # Couleur pour les titres, sombre mais toujours visible
$CATEGORY_SEPARATOR_COLOR = "Green"  # Couleur plus douce pour les séparateurs de catégorie
$CREDIT_COLOR = "Cyan"              # Couleur pour les crédits, claire et visible
$LEAVE_COLOR = "DarkRed"         # Texte pour quitter en jaune sombre, visible sans être trop vif
$EXIT_COLOR = "DarkRed"             # Texte de sortie en rouge sombre, moins agressif qu'un rouge vif
$REBOOT_ADMIN_COLOR = "Green"       # Texte pour relancer en admin en vert, énergique mais pas trop lumineux

$ANIMATION_COLOR = "white"
# ==============================================
# Functions : Toutes les fonctions sont regroupées ici
# ==============================================
#renomme la fenetre en cours d'execution
[System.Console]::Title = $NOM_DU_SCRIPT
$host.UI.RawUI.BackgroundColor = $BACKGROUND_COLOR  # Exemple de couleur pour l'arrière-plan
$host.UI.RawUI.ForegroundColor = $DEFAULT_TEXT_COLOR  # Changer la couleur du texte
Clear-Host  # Applique immédiatement la couleur en rafraîchissant l'écran
# ---------------------------------------------------------------------- Fonction 0 : Animation -----------------------------------------------------------------------------------------
#region Function_Animation
function AnimationTML {
Write-Host "";
Write-Host "";
Write-Host "    :::::::::       ::::::::::         :::   :::       ::::::::  ";
Write-Host "    :+:    :+:      :+:               :+:+: :+:+:     :+:    :+: ";
Write-Host "   +:+    +:+      +:+              +:+ +:+:+ +:+    +:+    +:+  ";
Write-Host "  +#+    +:+      +#++:++#         +#+  +:+  +#+    +#+    +:+   ";
Write-Host " +#+    +#+      +#+              +#+       +#+    +#+    +#+    ";
Write-Host "#+#    #+#      #+#              #+#       #+#    #+#    #+#     ";
Write-Host "#########       ##########       ###       ###     ########      ";
    Start-Sleep -Milliseconds 650
    #[console]::beep(1000,100)
    #[console]::beep(600,300)
    #[console]::beep(1000,500)
}
#endregion Function_Animation
# ---------------------------------------------------------------------- Fonction 1 : Menu -----------------------------------------------------------------------------------------
#region Function_Menu
function MenuPrincipale() {
    Clear-Host
    AdminModeWrite
    Write-Host ""
    Write-Host "---------------------------------------------" -ForegroundColor $TITLE_COLOR
    Write-Host "            $($NOM_DU_SCRIPT)"-ForegroundColor $TITLE_COLOR
    Write-Host "---------------------------------------------" -ForegroundColor $TITLE_COLOR
    Write-Host "-----MENUS PRINCIPAUX------------------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR 
    Write-Host "1.  Exchange Partner Connect 3.0"
    Write-Host "2.  Exchange Admin Connect 2.0"
    Write-Host "3.  Datto RMM                        (New)(Sabotage pour Démo)"
    Write-Host "4.  Site Web Diagnostique"
    Write-Host "5.  Registre"
    Write-Host "6.  Services                         $(AdminModeRequis)"
    Write-Host "7.  Configuration"
    Write-Host "8.  Reparation                       $(AdminModeRequis)"
    Write-Host "9.  Installation Logiciel"
    Write-Host "-----MENUS SUPPLEMENTAIRES-------------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR 
    Write-Host "10. Outils Divers                    $(AdminModeRequis)"
    Write-Host "11. Generateur de Mot de passe"
    Write-Host "12. Chronometre"
    Write-Host "-----MENUS INFORMATION-----------------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR 
    Write-Host "15. Information System"
    Write-Host "16. Information de Base Compagnie"
    Write-Host "-----OPTIONS---------------------------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR
    Write-Host "01. Relancer en mode Admin"-ForegroundColor $REBOOT_ADMIN_COLOR   
    Write-Host "02. Credit"-ForegroundColor $CREDIT_COLOR
    Write-Host "0.  Quitter                           V.$($VERSION_LOGICIEL)"-ForegroundColor $EXIT_COLOR
    Write-Host ""
    return Read-Host "Selection"
}
function MenuPartageTrois {
    #Menu
    Clear-Host
    AdminModeWrite              
    Write-Host ""
    Write-Host "---------------------------------------------" -ForegroundColor $TITLE_COLOR
    Write-Host "          MENU DE PARTAGE 3.0 - New "-ForegroundColor $TITLE_COLOR
    Write-Host "---------------------------------------------" -ForegroundColor $TITLE_COLOR
    Write-Host "-----GESTION BOITES ET CALENDRIERS PARTAGES--"-ForegroundColor $CATEGORY_SEPARATOR_COLOR  
    Write-Host "1. Ajouter Boite Partager"
    Write-Host "2. Enlever Boite Partager"
    Write-Host "3. Ajouter Calendrier Partager"
    Write-Host "4. Enlever Calendrier Partager"
    Write-Host "5. Changer Time zone Courriel + FR"
    Write-Host "-----CALENDRIERS PARTAGES KSA ONLY-----------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR  
    Write-Host "6. Ajouter Calendrier Partager (Compagnie A)"
    Write-Host "7. Enlever Calendrier Partager (Compagnie A)"
    Write-Host "-----ACTIVATION ET ARCHIVAGE-----------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR  
    Write-Host "20.Forcer Activation Archivage"
    Write-Host "-----COMMANDES MANUELLES---------------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR  
    Write-Host "98. Deconnexion"
    Write-Host "99. Commande Manuel"
    Write-Host "100. Installer Module exchange"
    Write-Host "---------------------------------------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR  
    Write-Host "0. Retour"-ForegroundColor $LEAVE_COLOR
    Write-Host ""
    return Read-Host "Selection"                    
}
function MenuPartageDeux($AdminConnecter) {
    #Menu
    Clear-Host
    AdminModeWrite
    if ($AdminConnecter -eq "") { Write-Warning "Aucun Admin Connecte!" }
    if ($AdminConnecter -ne "") { Write-Warning "Connecte : $AdminConnecter" }

    Write-Host ""
    Write-Host "---------------------------------------------" -ForegroundColor $TITLE_COLOR
    Write-Host "            MENU DE PARTAGE 2.0" -ForegroundColor $TITLE_COLOR
    Write-Host "---------------------------------------------" -ForegroundColor $TITLE_COLOR
    Write-Host "-----CONNEXION ADMIN DE TENANT---------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR 
    Write-Host "1. Connecter un Admin"
    Write-Host "2. Deconnexion"
    Write-Host "-----GESTION BOITES ET CALENDRIERS PARTAGES--"-ForegroundColor $CATEGORY_SEPARATOR_COLOR  
    Write-Host "3. Ajouter Boite Partagee"
    Write-Host "4. Enlever Boite Partagee"
    Write-Host "5. Ajouter Calendrier Partage"
    Write-Host "6. Enlever Calendrier Partage"
    Write-Host "7. Changer Time zone Courriel + FR"
    Write-Host "-----CALENDRIERS PARTAGES KSA ONLY-----------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR  
    Write-Host "8. Ajouter Calendrier Partager (Compagnie A)"
    Write-Host "9. Enlever Calendrier Partager (Compagnie A)"
    Write-Host "-----ACTIVATION ET ARCHIVAGE-----------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR  
    Write-Host "20. Forcer Activation Archivage"
    Write-Host "-----COMMANDES MANUELLES---------------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR  
    Write-Host "99. Commande Manuelle"
    Write-Host "100. Installer Module Exchange"
    Write-Host "---------------------------------------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR  
    Write-Host "0. Retour" -ForegroundColor $LEAVE_COLOR
    Write-Host ""
    return Read-Host "Selection"                   
}
function MenuSiteWeb() {
    Clear-Host
    AdminModeWrite
    Write-Host ""
    Write-Host "---------------------------------------------" -ForegroundColor $TITLE_COLOR
    Write-Host "        Liste Site Web / Diagnostique"-ForegroundColor $TITLE_COLOR
    Write-Host "---------------------------------------------" -ForegroundColor $TITLE_COLOR
    Write-Host "-----TESTS DIVERS----------------------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR
    Write-Host "1. Stress CPU-Test Matthew-X83"
    Write-Host "2. Stress GPU-Test Matthew-X83"
    Write-Host "3. Stress Test All OCCT            (Download)"
    Write-Host "4. WebCam Test Matthew-X83"
    Write-Host "5. Microphone Test Matthew-X83"
    Write-Host "6. Speaker Test Matthew-X83"
    Write-Host "7. Browser Test Matthew-X83"
    Write-Host "-----BENCHMARKS------------------------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR
    Write-Host "10. Benchmark CPU Matthew-X83"
    Write-Host "11. Benchmark GPU Matthew-X83"
    Write-Host "-----OUTILS NETTOYAGE ET DIAGNOSTIC----------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR
    Write-Host "90. Win Memory Cleaner             (Download)"
    Write-Host "-----TESTS EN LIGNE ET SCANS-----------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR
    Write-Host "100. Verifier Domain, Serveur, DNS, Mail"
    Write-Host "101. Verifier Site web Down"
    Write-Host "102. Virus Scan File and URL"
    Write-Host "-----BIOS------------------------------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR
    Write-Host "110. Lenovo BIOS Simulator"                                                                  
    Write-Host "---------------------------------------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR               
    Write-Host "0. Retour"-ForegroundColor $LEAVE_COLOR
    Write-Host ""
    return Read-Host "Selection"
}
function MenuRegistre() {
    Clear-Host
    AdminModeWrite
    Write-Host ""
    Write-Host "---------------------------------------------" -ForegroundColor $TITLE_COLOR
    Write-Host "           Editeur de Registre"-ForegroundColor $TITLE_COLOR
    Write-Host "---------------------------------------------" -ForegroundColor $TITLE_COLOR
    Write-Host "-----OUVRIR L'EDITEUR------------------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR
    Write-Host "1. Ouvrire l'editeur du Registre"
    Write-Host "-----NUMLOCK---------------------------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR              
    Write-Host "2. Verrouiller NumLock"
    Write-Host "3. Deverrouiller NumLock"
    Write-Host "-----BING DANS LA RECHERCHE------------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR  
    Write-Host "4. Supprimer Bing de la recherche"
    Write-Host "5. Ajouter Bing a la recherche"
    #Write-Host "6. n/a"
    #Write-Host "---------------------------------------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR             
    #Write-Host "50. "
    #Write-Host "51. "                                                                                                                      
    Write-Host "---------------------------------------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR               
    Write-Host "0. Retour"-ForegroundColor $LEAVE_COLOR
    Write-Host ""
    return Read-Host "Selection"
                    
}
function MenuServices() {
    $srvSpooler = Get-Service -Name "Spooler" | Select-Object Status
    $srvRemoteAccess = Get-Service -Name "RemoteAccess" | Select-Object Status
    $srvIKEEXT = Get-Service -Name "IKEEXT" | Select-Object Status
    $srvDigTrack = Get-Service -Name "DiagTrack" | Select-Object Status
    $srvTelecopie = Get-Service -Name "Telecopie" | Select-Object Status
    $srvSCardSvr = Get-Service -Name "SCardSvr" | Select-Object Status
    $srvXblGameSave = Get-Service -Name "XblGameSave" | Select-Object Status
    $srvXblAuthManager = Get-Service -Name "XblAuthManager" | Select-Object Status

    Clear-Host
    AdminModeWrite
    Write-Host ""
    Write-Host "---------------------------------------------" -ForegroundColor $TITLE_COLOR
    Write-Host "              liste Services         $(AdminModeRequis)"-ForegroundColor $TITLE_COLOR
    Write-Host "---------------------------------------------" -ForegroundColor $TITLE_COLOR
    Write-Host "-----OPTIONS DE BASE DES SERVICES------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR
    Write-Host "1. Liste des Services"
    Write-Host "2. Liste des Services Running"
    Write-Host "3. Liste des Services Stopped"
    Write-Host "-----SERVICES SPECIFIQUES--------------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR
    Write-Host "10. $($srvSpooler) Spouleur deimpression " 
    Write-Host "11. $($srvRemoteAccess) Routage et acces Distant (VPN)"
    Write-Host "12. $($srvIKEEXT) Modules de generation de cles IKEd"
    Write-Host "-----SERVICES ADDITIONNELS-------------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR
    Write-Host "50. $($srvDigTrack) Telemetrie" #DiagTrack
    Write-Host "51. $($srvTelecopie) null Telecopie (Logiciel fax)" #Telecopie
    Write-Host "52. $($srvSCardSvr) Carte a puce (Gere l'acces aux cartes puce lues par l'ordinateur)" #SCardSvr
    Write-Host "53. $($srvXblGameSave) Xbox" #XblGameSave
    Write-Host "54. $($srvXblAuthManager) Xbox Authentification" #XblAuthManager
    Write-Host "55. n/a"
    Write-Host "56. n/a"
    Write-Host "---------------------------------------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR
    Write-Host "0. Retour"-ForegroundColor $LEAVE_COLOR
    Write-Host ""
    return Read-Host "Selection"
}
function MenuConfiguration() {
    Clear-Host
    AdminModeWrite
    Write-Host ""
    Write-Host "---------------------------------------------" -ForegroundColor $TITLE_COLOR
    Write-Host "           MENU Configuration"-ForegroundColor $TITLE_COLOR
    Write-Host "---------------------------------------------" -ForegroundColor $TITLE_COLOR
    Write-Host "-----OPTIONS UTILISATEUR---------------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR
    Write-Host "1. Cree User IT                  (mdp par Defaut)"
    Write-Host "2. Set Mise en veille a jamais"
    Write-Host "3. Set Clavier en FR"
    Write-Host "4. Set Clavier en EN"
    Write-Host "-----OPTIONS SYSTEME-------------------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR               
    Write-Host "50. Installation Framework 3.5"
    Write-Host "51. Mise a jour Driver"                                                                                                                      
    Write-Host "---------------------------------------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR                
    Write-Host "0. Retour"-ForegroundColor $LEAVE_COLOR
    Write-Host ""
    return Read-Host "Selection"
                    
}
function MenuReparation() {
    Clear-Host
    AdminModeWrite
    Write-Host ""
    Write-Host "---------------------------------------------" -ForegroundColor $TITLE_COLOR
    Write-Host "              MENU REPARATION        $(AdminModeRequis)" -ForegroundColor $TITLE_COLOR
    Write-Host "---------------------------------------------" -ForegroundColor $TITLE_COLOR
    Write-Host "-----GESTION DES MISES A JOUR WINDOWS--------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR
    Write-Host "1. Desinstaller Mise a Jour Windows Update"
    Write-Host "2. SFC                                     (reparer les fichiers systemes)"
    Write-Host "3. BCD et demarrage de Windows"
    Write-Host "4. Point de restauration"
    Write-Host "-----GESTION DU DISQUE-----------------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR
    Write-Host "10. ChkDsk Classic"
    Write-Host "11. ChkDsk Corrompu"
    Write-Host "12. ChkDsk Offline"
    Write-Host "13. ChkDsk Online"
    Write-Host "-----REPARATION ET MAINTENANCE SYSTEME-------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR
    Write-Host "20. DISM                                   (Recherche Corruption)"
    Write-Host "21. DISM                                   (Restaure les Fichiers corrompus)"
    Write-Host "22. Repair-WindowsImage                    (Repare L'image de Windows)"
    Write-Host "-----GESTION DES PARTITIONS------------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR
    Write-Host "100. Resize RE Partition                   (W.Update KB5034441)"
    Write-Host "-----WINRE ET RESTAURATION-------------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR
    Write-Host "120. Recreer WinRE Partition"
    Write-Host "---------------------------------------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR
    Write-Host "0. Retour"-ForegroundColor $LEAVE_COLOR
    Write-Host ""
    return Read-Host "Selection"
}
function MenuWinGet() {
    Clear-Host
    AdminModeWrite
    Write-Host ""
    Write-Host "---------------------------------------------" -ForegroundColor $TITLE_COLOR
    Write-Host "        Installation Logiciel (WinGet)"-ForegroundColor $TITLE_COLOR
    Write-Host "---------------------------------------------" -ForegroundColor $TITLE_COLOR
    Write-Host "-----RECHERCHE ET INSTALLATION---------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR
    Write-Host "1. Search"
    Write-Host "-----APPLICATIONS COURANTES------------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR
    Write-Host "2. Chrome"#Google.Chrome
    Write-Host "3. Adobe Acrobat Reader DC (64-bit)" #Adobe.Acrobat.Reader.64-bit
    Write-Host "4. Java 8"#Oracle.JavaRuntimeEnvironment
    Write-Host "-----MICROSOFT-------------------------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR
    Write-Host "10. Microsoft OneDrive"#Microsoft.OneDrive
    Write-Host "11. Microsoft Outlook New" #9NRX63209R7B
    Write-Host "12. Microsoft Teams classic"#Microsoft.Teams.Classic
    Write-Host "13. Suite Microsoft 365"# Function Installer Microsoft
    Write-Host "-----OUTILS SPECIFIQUES----------------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR
    Write-Host "20. Lenovo Vantage" #9WZDNCRFJ4MV
    Write-Host "21. SDS (Serti)"
    Write-Host "-----GESTION DU DISQUE-----------------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR
    Write-Host "30. CrystalDiskInfo" #CrystalDewWorld.CrystalDiskInfo 
    Write-Host "31. Defraggler"#Piriform.Defraggler
    Write-Host "32. DiskGenius" #Eassos.DiskGenius               
    Write-Host "33. GDU (Gers de WinDirStat)" #dundee.gdu
    Write-Host "34. MiniTool Partition Wizard Free"#MiniTool.PartitionWizard.Free
    Write-Host "35. OCCT Personal (Stress test All)" #OCBase.OCCT.Personal 
    Write-Host "36. WinDirStat" #WinDirStat.WinDirStat
    Write-Host "-----BUREAU A DISTANCE-----------------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR
    Write-Host "50. AnyDesk" #AnyDeskSoftwareGmbH.AnyDesk
    Write-Host "51. GoToAssist Agent Desktop Console" #GoTo.GoToAssistAgentDesktopConsole
    #Write-Host "52. " #
    #Write-Host "53. " #
    Write-Host "-----PROGRAMMATION---------------------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR
    Write-Host "70. GitHub Desktop" #GitHub.GitHubDesktop
    Write-Host "71. PowerShell 7:Lastest" #Microsoft.PowerShell
    Write-Host "72. Visual Studio Code" #Microsoft.VisualStudioCode
    Write-Host "73. Visual Studio Community 2022" #Microsoft.VisualStudio.2022.Community
    Write-Host "-----GESTIONNAIRE ET EDITEUR DE TEXT---------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR
    Write-Host "90. Bitwarden (Gestionnaire de mot de passe)" #Bitwarden.Bitwarden
    Write-Host "91. LibreOffice " #TheDocumentFoundation.LibreOffice
    Write-Host "92. Obsidian" #Obsidian.Obsidian
    Write-Host "-----INSTALLATION GROUPEE--------------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR
    Write-Host "100. Essentiel (Chrome-Vantage-Acrobat)"
    Write-Host "101. Programmation (GitHub-Powershell 7-Studio Code-Studio Community)"                
    Write-Host "102. Partition Manager (DiskGenus-CrystalDiskInfo)"
    Write-Host "103. Machine Test (GDU-OCCT)"                
    Write-Host "---------------------------------------------"-ForegroundColor $CATEGORY_SEPARATOR_COLOR                                                  
    Write-Host "0. Retour"-ForegroundColor $LEAVE_COLOR
    Write-Host ""
    return Read-Host "Selection"
}
function MenuInfoSystem() {
    Clear-Host
    AdminModeWrite
    Write-Host ""
    Write-Host "---------------------------------------------" -ForegroundColor $TITLE_COLOR
    Write-Host "           Information System"-ForegroundColor $TITLE_COLOR
    Write-Host "---------------------------------------------" -ForegroundColor $TITLE_COLOR
    Write-Host "-----INFORMATIONS RESEAU---------------------" -ForegroundColor $CATEGORY_SEPARATOR_COLOR
    Write-Host "1. Configuration Reseau"
    Write-Host "-----INFORMATIONS SYSTEME DETAILLEES---------" -ForegroundColor $CATEGORY_SEPARATOR_COLOR
    Write-Host "2. Information System                      (Model, S/N, etc)"
    Write-Host "3. Information System                      (Moins Detaille)"
    Write-Host "4. Slot de Ram + info"
    Write-Host "-----RAPPORT A GENERER-----------------------" -ForegroundColor $CATEGORY_SEPARATOR_COLOR
    Write-Host "10. Rapport de la batterie"
    Write-Host "11. Rapport des peripheriques              (.txt > html)"
    Write-Host "12. Rapport de diagnostic reseau complet   $(AdminModeRequis)"
    Write-Host "13. Rapport de consommation electrique     $(AdminModeRequis)"                                                                    
    Write-Host "14. Rapport de fiabilite des peripherique"                                                                    
    Write-Host "15. Rapport de diagnostic systeme"                                                                    
    Write-Host "16. rapport SMART des disques"                                                                                                                                 
    Write-Host "---------------------------------------------" -ForegroundColor $CATEGORY_SEPARATOR_COLOR              
    Write-Host "0. Retour"-ForegroundColor $LEAVE_COLOR
    Write-Host ""
    return Read-Host "Selection"
}
function MenuOutilsDivers() {
    Clear-Host
    AdminModeWrite
    Write-Host ""
    Write-Host "---------------------------------------------" -ForegroundColor $TITLE_COLOR
    Write-Host "             Outils Divers"-ForegroundColor $TITLE_COLOR
    Write-Host "---------------------------------------------" -ForegroundColor $TITLE_COLOR
    Write-Host "-----EXTRACTEURS PHOTOS VIDEOS---------------" -ForegroundColor $CATEGORY_SEPARATOR_COLOR 
    Write-Host "1. Extracteurs de Photos                   $(AdminModeRequis)"
    Write-Host "2. Extracteurs de Videos                   $(AdminModeRequis)"
    Write-Host "-----OBSERVATEUR ET EXPLORATEUR--------------" -ForegroundColor $CATEGORY_SEPARATOR_COLOR 
    Write-Host "3. Explorateur de Fichier                  $(AdminModeRequis)" 
    Write-Host "4. Observateur d'evenement"
    Write-Host "-----OUTILS DE RECHERCHE---------------------" -ForegroundColor $CATEGORY_SEPARATOR_COLOR 
    Write-Host "5. Recherche en Powershell                 $(AdminModeRequis)"
    Write-Host "6. Recherche Avec Everything.exe           (auto install)"
    #Write-Host "-----A COMPLETER-----------------------------" -ForegroundColor $CATEGORY_SEPARATOR_COLOR                
    #Write-Host "20."
    #Write-Host "21."                                                   
    Write-Host "---------------------------------------------" -ForegroundColor $CATEGORY_SEPARATOR_COLOR               
    Write-Host "0. Retour"-ForegroundColor $LEAVE_COLOR
    Write-Host ""
    return Read-Host "Selection"
}
function MenuPasswordGenerators() {
    Clear-Host
    AdminModeWrite
    Write-Host ""
    Write-Host "---------------------------------------------" -ForegroundColor $TITLE_COLOR
    Write-Host "         Generateur de Mot de passe"-ForegroundColor $TITLE_COLOR
    Write-Host "---------------------------------------------" -ForegroundColor $TITLE_COLOR
    Write-Host "-----GENERER MOT DE PASSE--------------------" -ForegroundColor $CATEGORY_SEPARATOR_COLOR  
    Write-Host "1. De Base Format Exchange         (ABC12345)"
    Write-Host "2. Full Random longueur ajustable  (*?856Yt%)"
    Write-Host "3. Mots Random ajustable           (De?-Et!-)"              
    Write-Host "---------------------------------------------" -ForegroundColor $CATEGORY_SEPARATOR_COLOR               
    Write-Host "0. Retour"-ForegroundColor $LEAVE_COLOR
    Write-Host ""
    $selection = Read-Host "Selection"
    return $selection
    
}

#endregion Function_Menu
# ------------------------------------------------------------------- Fonction Graphique : Secondaire -----------------------------------------------------------------------------------------
#region Function_Graphique
#Selecteur de dossier, retourne un path
function Select-Folder {
    Add-Type -AssemblyName System.Windows.Forms

    # Crée un objet FolderBrowserDialog
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog

    # Affiche la fenêtre et vérifie si l'utilisateur a sélectionné un dossier
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.SelectedPath
    }
    else {
        return $null
    }
}
#Selecteur de fichier, retourne un path
function Select-File {
    Add-Type -AssemblyName System.Windows.Forms

    # Crée un objet OpenFileDialog
    $dialog = New-Object System.Windows.Forms.OpenFileDialog

    # Affiche la fenêtre et vérifie si l'utilisateur a sélectionné un fichier
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.FileName
    }
    else {
        return $null
    }
}
#endregion Function_Graphique
# ---------------------------------------------------------------------- Fonction 2 : Secondaire -----------------------------------------------------------------------------------------
#region Function_Secondaire
function AdminModeWrite() {
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-Host "Admin Mode : Disable"-ForegroundColor Red
        Write-Host "*Certaine Fonctionaliter seront restrinte"-ForegroundColor gray
    }
    else {
        Write-Host "Admin Mode : Enable" -ForegroundColor green
    }
}

function AdminModeRequis() {
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        return "(*Admin)"
    }
    else {
        return ""
    }
}

function StartStopService {
    param (
        [string]$ServiceName
    )
    if ((Get-Service -Name $ServiceName).Status -eq "Running") {
        Stop-Service -Name $ServiceName
        return
    }
    if ((Get-Service -Name $ServiceName).Status -eq "Stopped") {
        Start-Service -Name $ServiceName
        return
    }
}
function DefaultError404 { Write-Host "Erreur 404 : Aucune correspondance trouvee"; pause; }


#endregion Function_Secondaire
# ---------------------------------------------------------------------- Fonction 3 : principale -------------------------------------------------------------------------------------------
#region Function_principale
function OptionQuiterLeLogiciel() {
    Clear-Host
    Write-Host "Logiciel en cours de fermeture ..."
    Write-Host "     ________  _____ " -ForegroundColor cyan;
    Write-Host "    /_  __/  |/  / / " -ForegroundColor cyan;
    Write-Host "     / / / /|_/ / /  " -ForegroundColor cyan;
    Write-Host "    / / / /  / / /___" -ForegroundColor cyan;
    Write-Host "   /_/ /_/  /_/_____/ informatique." -ForegroundColor cyan;
    Write-Host "                     " -ForegroundColor cyan;
    Write-Host "Version $($VERSION_LOGICIEL)"
    Start-Sleep -Milliseconds 500
    $pid
    Clear-Host
    Stop-Process -ID $pid
}
function OptionPartageTrois {
    Import-Module ExchangeOnlineManagement
    $leaveMenuPartage = $true
    While ($leaveMenuPartage) {

        $ChoixMENU = MenuPartageTrois
        Clear-Host

        #Ajout Acces Courriel
        switch ($ChoixMENU) {
            
            #Ajouter Boite Partager
            "1" { PartagerCourriel }
            "2" { DePartagerCourriel }
            "3" { PartagerCalendrier }
            "4" { DePartagerCalendrier }
            "5" { TimeZoneCourrielFR }
            "6" { PartagerCalendrierKSA }
            "7" { DePartagerCalendrierKSA }
            "20" { ForcerActivationArchivage }
            #Deconnexion
            "98" { DeconnexionAdmin }
            "99" { ExchangeManuel }
            "100" { InstallerModuleExchange }
            "0" { $leaveMenuPartage = $false }
            default { DefaultError404 }
        }
    }
}
function OptionPartageDeux {
    Import-Module ExchangeOnlineManagement
    $AdminConnecter = ""
    $leaveMenuPartage = $true

    while ($leaveMenuPartage) {
        $ChoixMENU = MenuPartageDeux $AdminConnecter
        Clear-Host

        switch ($ChoixMENU) {
            "1" { $AdminConnecter = ConnectAdminDeux }
            "2" { $AdminConnecter = DisconnectAdminDeux }
            "3" { AjouterBoitePartageDeux $AdminConnecter }
            "4" { EnleverBoitePartageDeux $AdminConnecter }
            "5" { AjouterCalendrierPartageDeux $AdminConnecter }
            "6" { EnleverCalendrierPartageDeux $AdminConnecter }
            "7" { ChangeMailboxLanguageDeux $AdminConnecter }
            "20" { ActiverArchivageDeux $AdminConnecter }
            "100" { InstallerModuleExchangeDeux }
            "0" { $leaveMenuPartage = $false }
            Default { Write-Host "Aucune correspondance trouvee" }
        }
    }
}
function OptionDattoRMM {
    $securePassword = ConvertTo-SecureString -String 'public' -AsPlainText -Force

    $params = @{
        Credential  = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList ('public-client', $securePassword)
        Uri = "https://google.com"
        Method = "POST"
        ContentType = "application/x-zzz-form-urlencoded"
        Body = "grant_type=password&username=123456789&password=123456789"
    }

    $token = (Invoke-WebRequest -UseBasicParsing @params | ConvertFrom-Json).access_token

    $params = @{
        Credential  = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList ('public-client', $securePassword)
        Uri = "https://google.com"
        Method = "GET"
        ContentType = "application/json"
        Headers = @{
            Authorization = "Bearer $token"
        }
    }

    $sites = Invoke-WebRequest -UseBasicParsing @params | ConvertFrom-Json

    $site = $sites.sites | Select-Object name, uid | Sort-Object name | Out-GridView -Title "Datto RMM" -PassThru | Select-Object -ExpandProperty uid

    # clear l'installation au cas où qu'elle est corrompue
    if (((Get-Service -Name "CagService" -ErrorAction SilentlyContinue).Status -eq "Stopped")) {
    if (Test-Path -Path "C:\Program Files (x86)\CentraStage") {
        Get-Process | ? {$_.Path -like "*CentraStage*"} | Stop-Process -Force
        Remove-Item -Path "C:\Program Files (x86)\CentraStage" -Force -Recurse -ErrorAction Continue
    }
    }

    # enable tls 1.2
    # https://www.reddit.com/r/PowerShell/comments/ozr6ye/psa_enabling_tls12_and_you/
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
    (New-Object System.Net.WebClient).DownloadFile("https://google.com/$site", "$env:TEMP/AgentInstall.exe");start-process "$env:TEMP/AgentInstall.exe"
}
function OptionSiteWeb {
    $leaveMenuSiteWeb = $true
    While ($leaveMenuSiteWeb) {

        $ChoixMenuSiteWeb = MenuSiteWeb ############################ Menu Site Web #####################################
        Clear-Host

        switch ($ChoixMenuSiteWeb) {
            "1" { Start-Process "https://www.matthew-x83.com/online/cpu-test.php" } 
            "2" { Start-Process "https://www.matthew-x83.com/online/gpu-test.php" } 
            "3" { Start-Process "https://www.ocbase.com/" }
            "4" { Start-Process "https://www.matthew-x83.com/online/webcam-test.php" }
            "5" { Start-Process "https://www.matthew-x83.com/online/microphone-test.php" }
            "6" { Start-Process "https://www.matthew-x83.com/online/speaker-test.php" }
            "7" { Start-Process "https://www.matthew-x83.com/online/browser-test.php" }
            "10" { Start-Process "https://www.matthew-x83.com/online/cpu-benchmark.php" }
            "11" { Start-Process "https://www.matthew-x83.com/online/gpu-benchmark-test.php" }
            "90" { Start-Process "https://github.com/IgorMundstein/WinMemoryCleaner?tab=readme-ov-file" }
            "100" { Start-Process "https://mxtoolbox.com/" }
            "101" { Start-Process "https://www.isitdownrightnow.com/" }
            "110" { Start-Process "https://download.lenovo.com/bsco/index.html#/" }                                                                            
            "0" { $leaveMenuSiteWeb = $false } 
            Default { DefaultError404 }                  
        }
    }
}
function OptionRegistre {
    $leaveMenuRegistre = $true
    While ($leaveMenuRegistre) {

        $ChoixMenuRegistre = MenuRegistre ############################ MENU Registre #####################################
        Clear-Host

        switch ($ChoixMenuRegistre) {
            "1" { regedit; } 
            "2" { Set-ItemProperty -Path "HKCU:\Control Panel\Keyboard" -Name "InitialKeyboardIndicators" -Value 2; pause; }
            "3" { Set-ItemProperty -Path "HKCU:\Control Panel\Keyboard" -Name "InitialKeyboardIndicators" -Value 0; pause; }
            "4" { Remove-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Search" -Name "BingSearchEnabled"; pause; }
            "5" { Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Search" -Name "BingSearchEnabled" -Value 1; pause; }                                                                            
            "0" { $leaveMenuRegistre = $false } 
            Default { DefaultError404 }                  
        }
    }
}
function OptionServices {
    $leaveMenuServices = $true
    While ($leaveMenuServices) {

        $ChoixMenuServices = MenuServices ############################ MENU Services #####################################
        Clear-Host

        switch ($ChoixMenuServices) {
            "1" { Clear-Host; Get-Service | Select-Object Status, Name, DisplayName; pause; }
            "2" { Clear-Host; Get-Service | Where-Object { $_.Status -ne 'Stopped' } | Select-Object Status, Name, DisplayName; pause; }
            "3" { Clear-Host; Get-Service | Where-Object { $_.Status -eq 'Stopped' } | Select-Object Status, Name, DisplayName; pause; }
            "10" { Clear-Host; StartStopService("Spooler"); }
            "11" { Clear-Host; StartStopService("RemoteAccess"); }
            "12" { Clear-Host; StartStopService("IKEEXT"); }
            "50" { Clear-Host; StartStopService("DiagTrack"); }
            "51" { Clear-Host; StartStopService("Telecopie"); }
            "52" { Clear-Host; StartStopService("SCardSvr"); }
            "53" { Clear-Host; StartStopService("XblGameSave"); }
            "54" { Clear-Host; StartStopService("XblAuthManager"); }

            #"10"{stop-Service -Name Spooler; start-Service -Name Spooler; pause; }
                                                                                                                                                                                                   
            "0" { $leaveMenuServices = $false }
            Default { DefaultError404 }                  
        }
    }
}
function OptionConfiguration {
    $leaveMenuConfiguration = $true
    While ($leaveMenuConfiguration) {

        $ChoixMenuConfiguration = MenuConfiguration ############################ MENU CONFIGURATION #####################################
        Clear-Host

        switch ($ChoixMenuConfiguration) {
            "1" { Clear-Host; New-LocalUser -Name "IT" -Password "Dem0!$"; Add-LocalGroupMember -Group "Administrators" -Member "IT"; Add-LocalGroupMember -Group "Administrateurs" -Member "IT"; Write-Host "L'utilisateur ITUser a ete cre avec succes."; pause; }
            "2" { powercfg /Change -monitor-timeout-ac 0; powercfg /Change -monitor-timeout-dc 0; powercfg /Change -standby-timeout-ac 0; powercfg /Change -standby-timeout-dc 0; pause; }
            "3" { Set-WinUserLanguageList -LanguageList fr-CA -Force; }
            "4" { Set-WinUserLanguageList -LanguageList en-US -Force; } 
            "50" { Start-Process "cmd.exe" -ArgumentList '/c DISM.exe /online /Enable-Feature /FeatureName:NetFx3' -Verb RunAs; }
            "51" { Start-Process "powershell.exe" -ArgumentList "Install-Module 'LSUClient' Force; Import-Module 'LSUClient'; Write-Host 'Installing driver updates. The computer will reboot.'; Write-Host 'Loading update list...'; $updates = Get-LSUpdate; $updates | Select-Object -ExpandProperty Title; $updates | Install-LSUpdate -Verbose" -Verb RunAs; } 
                                                                                                                                                                                                   
            "0" { $leaveMenuConfiguration = $false } 
            Default { DefaultError404 }                  
        }
    }
}
function OptionReparation {
    $leaveMenuReparation = $true
    while ($leaveMenuReparation) {
        $ChoixMenuReparation = MenuReparation ############################ MENU REPARATION #######################################
        Clear-Host
               
        switch ($ChoixMenuReparation) {
            "1" {
                $installedUpdates = Get-HotFix | Select-Object HotfixID, Description, InstalledOn | Sort-Object InstalledOn; $installedUpdates | Format-Table -AutoSize;
                Write-host "Entree le Numero de KB a Desinstaller EX:(KB: 5035853)"; $NumKB = read-Host "KB "; $Restart = Read-Host "Redemarer apres la Desintallation ? [O/N]" ;
                if ($Restart -eq "O" -or $Restart -eq "0") { wusa.exe /uninstall /kb:$($NumKB) /log ; }else { wusa.exe /uninstall /kb:$($NumKB) /norestart /log; } pause; 
            }
            "2" { sfc /scannow; pause; }                  
            "3" {
                Clear-Host; $volume = Get-Partition -DiskNumber 0 -PartitionNumber 1; Set-Partition -InputObject $volume -NewDriveLetter Z; bcdboot C:\Windows /s Z: /f ALL; $restart = Read-Host "Redemarrer l'ordinateur maintenant? (O/N)";
                if ($restart -eq "O") { Restart-Computer; }
            }
            "4" { get-computerrestorepoint; $NumRestaurePoint = read-host "Entrer le numero de sequence"; restore-computer -restaurePoint $NumRestaurePoint; get-computerrestorepoint -LastStatus; write-host "L'ordinateur devrait redemarer . . ."; pause; }
            "10" { Get-PSDrive -PSProvider FileSystem; Write-Host ""; $lettreDiskSelec = Read-Host "Entrer la lettre du lecteur a Verifier"; chkdsk $lettreDiskSelec /F /V /perf; pause; }
            "11" { Get-PSDrive -PSProvider FileSystem; Write-Host ""; $lettreDiskSelec = Read-Host "Entrer la lettre du lecteur a Verifier"; chkdsk $lettreDiskSelec /R /Scan /perf; pause; }
            "12" { Get-PSDrive -PSProvider FileSystem; Write-Host ""; $lettreDiskSelec = Read-Host "Entrer la lettre du lecteur a Verifier"; chkdsk $lettreDiskSelec /scan /forceofflinefix /perf; pause; }
            "13" { Get-PSDrive -PSProvider FileSystem; Write-Host ""; $lettreDiskSelec = Read-Host "Entrer la lettre du lecteur a Verifier"; chkdsk $lettreDiskSelec /scan /F /V /perf; pause; }
            "20" { dism /online /cleanup-image /checkhealth; pause; }
            "21" { dism /online /cleanup-image /restorehealth; pause; }
            "22" { Repair-WindowsImage -Online -RestoreHealth; pause; }
            "100" { powershell.exe -NoProfile -ExecutionPolicy Bypass -File "Resize_script.ps1" -SkipConfirmation $true -BackupFolder "c:\\winre_backup"; Pause; }
            "120" {
                Get-Disk | Where-Object { $_.Number -eq 0 } | Get-Partition | Where-Object { $_.PartitionNumber -eq 4 } | Set-Partition -NewDriveLetter R
                New-Item -Path R:\Recovery -ItemType Directory
                New-Item -Path R:\Recovery\WindowsRE -ItemType Directory
                Copy-Item -Path C:\mount\Windows\System32\Recovery\winre.wim -Destination R:\Recovery\WindowsRE\
                Set-Partition -DriveLetter R -GptType DE94BBA4-06D1-4D40-A16A-BFD50179D6AC
                reagentc /disable; reagentc /setreimage /path R:\Recovery\WindowsRE\; reagentc /enable; reagentc /info; pause;
            }
                               
            "0" { $leaveMenuReparation = $false }
            Default { DefaultError404 }
        }
    }
}
function OptionWinGet {
    $leaveMenuWinget = $true
    #Verifie si winget est installer
    $VerifierWinGetInstaller = Get-AppPackage *Microsoft.DesktopAppInstaller* | Select-Object Name

    if ($VerifierWinGetInstaller -eq "") {
        Write-host "Installation de WinGet . . ."
        #$progressPreference = 'silentlyContinue'
        Write-Information "Telechargement de WinGet et de ses dependances..."
        Invoke-WebRequest -Uri $LIEN_DOWNLOAD_WINGET -OutFile "$HOME\Downloads\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"
        # Installation du package telecharge
        Add-AppxPackage -Path "$HOME\Downloads\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"
    }
    While ($leaveMenuWinget) {
        $ChoixMenuWinget = MenuWinGet ################################### MENU WINGET #########################################
        Clear-Host

        switch ($ChoixMenuWinget) {
            #Search
            "1" {
                write-host "q. Quitter" -ForegroundColor Red; $logicielRechercher = Read-Host "Logiciel a rechercher"; Clear-Host;
                if ($logicielRechercher -ne "q") {
                    winget search $logicielRechercher; write-host "  "; write-host "q. Quitter" -ForegroundColor Red; $IDlogicielAInstaller = Read-Host "ID logiciel a installer";                                
                    if ($IDlogicielAInstaller -ne "q") { winget install $IDlogicielAInstaller; pause; }
                }
            }
            #Application courantes
            "2" { Winget install Google.Chrome }
            "3" { Winget install Adobe.Acrobat.Reader.64-bit }
            "4" { Winget install Oracle.JavaRuntimeEnvironment }
            #Microsoft
            "10" { Winget install Microsoft.OneDrive }
            "11" { Winget install 9NRX63209R7B }
            "12" { Winget install Microsoft.Teams.Classic }
            "13" { Installer_Microsoft }
            #Outils specifique
            "20" { Winget install 9WZDNCRFJ4MV }
            "21" { Set-Location /d %temp%; curl.exe --output sds.exe --url $LIEN_DOWNLOAD_SERTI; powershell.exe "start .\sds.exe"; exit; }
            #gestion du disk
            "30" { Winget install CrystalDewWorld.CrystalDiskInfo }
            "31" { Winget install Piriform.Defraggler }
            "32" { Winget install Eassos.DiskGenius }
            "33" { Winget install dundee.gdu }
            "34" { Winget install MiniTool.PartitionWizard.Free }
            "35" { Winget install OCBase.OCCT.Personal }
            "36" { Winget install WinDirStat.WinDirStat }
            #Bureau a distance
            "50" { Winget install AnyDeskSoftwareGmbH.AnyDesk }
            "51" { Winget install GoTo.GoToAssistAgentDesktopConsole }
            #Programmation
            "70" { Winget install GitHub.GitHubDesktop }
            "71" { Winget install Microsoft.PowerShell }
            "72" { Winget install Microsoft.VisualStudioCode }
            "73" { Winget install Microsoft.VisualStudio.2022.Community }
            #Gestionnaire et editeur de text
            "90" { Winget install Bitwarden.Bitwarden }
            "91" { Winget install TheDocumentFoundation.LibreOffice }
            "92" { Winget install Obsidian.Obsidian }
            #"" { Winget install }

            #Lot d'application
            "100" {
                #Essentiel
                winget install 9WZDNCRFJ4MV --silent --accept-source-agreements --accept-package-agreements --force --source msstore
                winget install Google.Chrome --silent --accept-source-agreements --accept-package-agreements --force --source winget
                winget install Adobe.Acrobat.Reader.64-bit --silent --accept-source-agreements --accept-package-agreements --force --source winget
            }
            "101" {
                #Programmation
                Winget install GitHub.GitHubDesktop --silent --accept-source-agreements --accept-package-agreements --force --source winget
                Winget install Microsoft.PowerShell --silent --accept-source-agreements --accept-package-agreements --force --source winget
                Winget install Microsoft.VisualStudioCode --silent --accept-source-agreements --accept-package-agreements --force --source winget
                Winget install Microsoft.VisualStudio.2022.Community --silent --accept-source-agreements --accept-package-agreements --force --source winget
            }
            "102" {
                #Partition Manager
                Winget install CrystalDewWorld.CrystalDiskInfo --silent --accept-source-agreements --accept-package-agreements --force --source winget
                Winget install Eassos.DiskGenius --silent --accept-source-agreements --accept-package-agreements --force --source winget
            }
            "103" {
                #Machine Test
                Winget install dundee.gdu --silent --accept-source-agreements --accept-package-agreements --force --source winget
                Winget install OCBase.OCCT.Personal --silent --accept-source-agreements --accept-package-agreements --force --source winget
            }
            "0" { $leaveMenuWinget = $false } 
            Default { DefaultError404 }                  
        }
    } 
}
function OptionInfoSystem {
    $leaveMenuInfoSystem = $true
    While ($leaveMenuInfoSystem) {
        $ChoixMenuInfoSystem = MenuInfoSystem ################################### MENU Info System #########################################
        Clear-Host

        switch ($ChoixMenuInfoSystem) {
            #Search
            "1" { ipconfig /all; pause; }
            "2" { msinfo32; }
            "3" { Write-Host "System Info Loading ..."; systeminfo; pause; }
            "4" { Clear-Host; Write-Host "Info sur la Ram"; Get-WmiObject -Class Win32_PhysicalMemoryArray | Select-Object -Property MemoryDevices; Get-WmiObject -Class Win32_PhysicalMemory | Format-Table -Property DeviceLocator, Capacity, Manufacturer, PartNumber, Speed; pause; }
            "10" { CheckBatteryStatusHtml }
            "11" { CheckDeviceStatusHtml }
            "12" { GenerateNetworkReport }
            "13" { GenerateEnergyReport }
            "14" { OpenReliabilityReport }
            "15" { GenerateSystemDiagnosticsReport }
            "16" { GenerateSMARTReport }

            "0" { $leaveMenuInfoSystem = $false } 
            Default { DefaultError404 }                  
        }
    } 
}
function OptionOutilsDivers {
    $leaveMenuOutilsDivers = $true
    While ($leaveMenuOutilsDivers) {
        $ChoixMenuOutilsDivers = MenuOutilsDivers
        Clear-Host

        switch ($ChoixMenuOutilsDivers) {
            #Extracteur de photo
            "1" { ExtracteurDePhotos }
            #extracteur de Video
            "2" { ExtracteurDeVideos }
            #Explorateur Powershell
            "3" { ExplorateurPowershellCustom }
            #Observateur d'evenement
            "4" { ObservateurEvenement }
            "5" { RecherchePowershell }
            "6" { RechercheEverything }
            "0" { $leaveMenuOutilsDivers = $false } 
            Default { DefaultError404 }                  
        }
    } 
}
function OptionPasswordGenerator {

    $numberIsCheck = $false

    $leaveMenuPasswordGenerators = $true
    While ($leaveMenuPasswordGenerators) {
        $ChoixMenuPasswordGenerators = MenuPasswordGenerators
        Clear-Host

        switch ($ChoixMenuPasswordGenerators) {
            "1" { GeneratePasswordBase }
            "2" { GeneratePasswordInterComplex }
            "3" { GeneratePasswordMots }
            "0" { $leaveMenuPasswordGenerators = $false } 
            Default { DefaultError404 }                  
        }
    } 
}
function OptionCrono {
    $isRunning = $false
    $startTime = $null
    $elapsedTime = 0

    Write-Host "Appuyez sur 'S' pour demarrer ou arreter le chronometre."
    Write-Host "Appuyez sur 'Q' pour quitter."

    while ($true) {
        
        
        $key = [System.Console]::ReadKey($true).Key
        
        if ($key -eq 'S') {
            if ($isRunning) {
                # Arrêter le chronomètre
                $elapsedTime += (Get-Date) - $startTime
                Write-Host "Chronometre arrete. Temps ecoule : $([math]::Round($elapsedTime.TotalSeconds, 2)) secondes."
                $isRunning = $false
            }
            else {
                # Démarrer le chronomètre
                $startTime = Get-Date
                Write-Host "Chronometre demarre."
                $isRunning = $true
            }
        }
        elseif ($key -eq 'Q') {
            # Quitter
            Write-Host "Sortie du chronometre."
            break
        }
        Start-Sleep -Milliseconds 500  # Attendre un peu avant la prochaine lecture de touche
    }
}
function OptionInfoCompagnie {
    # Vérifie si Out-GridView est disponible
    if (Get-Command Out-GridView -ErrorAction SilentlyContinue) {

        # Création de la liste des compagnies avec informations
        $compagnies = @(
            [PSCustomObject]@{ Nom = "Groupe Demo";                     Num = "(418) 123-1234"; Tenant = "Groupe Demo";     Lieu = "QC"; Datto = "Oui"; Antivirus = "Defender" }

            #[PSCustomObject]@{ Nom = "Template";                    Num = "(418) 000-0000"; Tenant = ""; Lieu = ""; Datto = ""; Antivirus = "" }
        )

        # Tri des compagnies par Nom avant l'affichage
        $compagnies = $compagnies | Sort-Object -Property Nom

        # Affichage des compagnies dans Out-GridView
        $selectionCompagnie = $compagnies | Out-GridView -Title "List des Compagnies" -PassThru

        # Vérifie si une compagnie a été sélectionnée
        if ($selectionCompagnie) {
            # Affiche les informations de la compagnie sélectionnée
            Write-Host "Informations de la compagnie selectionnee :"
            Write-Host "Nom : $($selectionCompagnie.Nom)"
            Write-Host "Numero : $($selectionCompagnie.Num)"
            Write-Host "# Tenant : $($selectionCompagnie.Tenant)"
            Write-Host "Lieu : $($selectionCompagnie.Lieu)"
            Write-Host "Datto : $($selectionCompagnie.Datto)"
            Write-Host "Antivirus : $($selectionCompagnie.Antivirus)"
            pause
        } else {
            Write-Host "Aucune compagnie selectionnee."
        }

    } else {
        Write-Host "La commande 'Out-GridView' n'est pas disponible sur ce système."
    }
}
function OptionCredit {
    Clear-Host
    Write-Host ""
    Write-Host "---------------------------------------------" -ForegroundColor $TITLE_COLOR
    Write-Host "                   Credit"-ForegroundColor $TITLE_COLOR
    Write-Host "---------------------------------------------" -ForegroundColor $TITLE_COLOR
    Write-Host "---------------------------------------------" -ForegroundColor $CATEGORY_SEPARATOR_COLOR
    Write-Host "Developpe par Maxime S."
    Write-Host "Animations realisees par Maxime S."
    Write-Host "Logo concu a l'aide d'une intelligence artificielle."
    Write-Host "---------------------------------------------" -ForegroundColor $CATEGORY_SEPARATOR_COLOR
    Write-Host "Date de creation : 27 Juin 2024"
    Write-Host "Version actuelle du logiciel : $($VERSION_LOGICIEL)"
    Write-Host "c 2025 Maxime S. Tous droits reserves."
    Write-Host "---------------------------------------------" -ForegroundColor $CATEGORY_SEPARATOR_COLOR
    Write-Host "Pour toute question ou retour, veuillez contacter : demo@demo.com"

    Write-Host "_|_|_|     _|_|_|_|   _|      _|     _|_|" -ForegroundColor DarkRed
    Write-Host "_|    _|   _|         _|_|  _|_|   _|    _|" -ForegroundColor DarkRed
    Write-Host "_|    _|   _|_|_|     _|  _|  _|   _|    _|" -ForegroundColor DarkRed
    Write-Host "_|    _|   _|         _|      _|   _|    _|" -ForegroundColor DarkRed
    Write-Host "_|_|_|     _|_|_|_|   _|      _|     _|_|" -ForegroundColor DarkRed
    Write-Host "                     " -ForegroundColor DarkRed

    pause
    Clear-Host
}
function OptionAdminMode {
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit;
    }
}
#endregion Function_principale
#----------------------------------------------------------------------- Fonction WinGet  ----------------------------------------------------------------------------------------------------
#region Function_Menu_WinGet
function Installer_Microsoft {
        # Chemins et URLs
    $tempDir = "$env:TEMP\repository-main"
    $zipUrl = "https://github.com/MaxSim2001/Microsoft236_Unattended_Installer/archive/refs/heads/main.zip"
    $zipPath = "$env:TEMP\repository-main.zip"

    # Télécharger le fichier ZIP depuis GitHub
    Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath

    # Vérifier si le fichier ZIP est téléchargé et tenter de l'extraire
    if (Test-Path -Path $zipPath) {
        try {
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            Expand-Archive -Path $zipPath -DestinationPath $tempDir -Force
            Write-Output "Extraction réussie."
        } catch {
            Write-Output "Erreur : Le fichier ZIP semble corrompu ou endommage."
            exit
        }
    } else {
        Write-Output "Erreur : Telechargement du fichier ZIP a echoue."
        exit
    }

    # Vérification de la présence de odt.exe
    $odtPath = Join-Path -Path $tempDir -ChildPath "Microsoft236_Unattended_Installer-main\odt.exe"
    if (Test-Path -Path $odtPath) {
        Write-Output "Fichier odt.exe trouvé."
    } else {
        Write-Output "Erreur : 'odt.exe' introuvable."
        exit
    }

    # Rechercher et exécuter le fichier install.cmd dans le bon répertoire
    $installScriptPath = Join-Path -Path $tempDir -ChildPath "Microsoft236_Unattended_Installer-main\install.cmd"
    if (Test-Path -Path $installScriptPath) {
        $originalPath = Get-Location
        Set-Location -Path (Split-Path -Path $installScriptPath)
        & $installScriptPath
        Set-Location -Path $originalPath
    } else {
        Write-Output "Erreur : Le script install.cmd est introuvable dans $tempDir"
        exit
    }

    # Nettoyage après exécution
    Remove-Item -Path $zipPath -Force
    Remove-Item -Path $tempDir -Recurse -Force
}
#endregion Function_Menu_WinGet
#----------------------------------------------------------------------- Fonction PARTAGE 3.0  ----------------------------------------------------------------------------------------------------
#region Function_Menu_Partage
# Menu Partage - Récupere domaine d'un courriel
function Get-DomainFromEmail {
    param (
        [string]$EmailAddress
    )

    try {
        # Vérifier si l'adresse e-mail contient bien un '@'
        if ($EmailAddress -match "^.+@(.+)$") {
            $Domain = $matches[1]
            return $Domain
        }
        else {
            throw "Format d'adresse e-mail invalide. L'adresse doit contenir '@'."
        }
    }
    catch {
        Write-Host "Erreur : $($_.Exception.Message)"
    }
}
# Menu Partage - CMD Exchange Manuel
function ExchangeManuel {
    Clear-Host
    $leaveWhileLibreExchange = $true
    While ($leaveWhileLibreExchange) {
        write-host ""
        write-host "Entrer vos commande Exchange Librement!"
        Write-Host "0. Retour"-ForegroundColor $LEAVE_COLOR
        write-host ""
        $commandeUserPS = Read-Host "PS Microsoft365>"

        if ($commandeUserPS -eq "0") { $leaveWhileLibreExchange = $false; }
        else { Invoke-Expression $commandeUserPS; }
                            
    }
}
# Menu Partage - Installer module exchange lastest
function InstallerModuleExchange {
    Set-ExecutionPolicy unrestricted; Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber; pause;
}

function DeconnexionAdmin {
    try {
        Disconnect-ExchangeOnline -Confirm:$false
    }
    catch {
        Write-Host "Erreur lors de la déconnexion d'Exchange Online : $($_.Exception.Message)"
    }
}

function PartagerCourriel {
    try {
        Clear-Host
        Write-Host "Adresse de la boite partagee"
        $SharedMailbox = Read-Host "Courriel"
        Write-Host "Adresse de l'utilisateur a qui vous donnez l'acces"
        $RecipientEmail = Read-Host "Courriel"

        if ($SharedMailbox -ne "" -and $RecipientEmail -ne "") {
            Clear-Host
            $DomainEmail = Get-DomainFromEmail($SharedMailbox)
            Connect-ExchangeOnline -DelegatedOrganization $DomainEmail

            Add-MailboxPermission -Identity $SharedMailbox -User $RecipientEmail -AccessRights FullAccess -InheritanceType All -Automapping $false
            Add-RecipientPermission -Identity $SharedMailbox -Trustee $RecipientEmail -AccessRights SendAs
            Write-Host "Acces accorde a $RecipientEmail pour la boite aux lettres partagee $SharedMailbox."
            pause
        }
        else {
            Write-Host "Adresse e-mail ou boite partagee vide."
        }
    }
    catch {
        Write-Host "Erreur lors du partage de la boîte : $($_.Exception.Message)"
    }
}

function DePartagerCourriel {
    try {
        Clear-Host
        Write-Host "Adresse de la boite partagee"
        $SharedMailbox = Read-Host "Courriel"
        Write-Host "Adresse de l'utilisateur à qui vous donnez l'acces"
        $RecipientEmail = Read-Host "Courriel"

        if ($SharedMailbox -ne "" -and $RecipientEmail -ne "") {
            $DomainEmail = Get-DomainFromEmail($SharedMailbox)
            Connect-ExchangeOnline -DelegatedOrganization $DomainEmail

            Remove-RecipientPermission $SharedMailbox -AccessRights SendAs -Trustee $RecipientEmail
            Remove-MailboxPermission -Identity $SharedMailbox -User $RecipientEmail -AccessRights FullAccess -InheritanceType All
            Write-Host "Acces supprime a $RecipientEmail pour la boîte aux lettres partagee $SharedMailbox."
            pause
        }
        else {
            Write-Host "Adresse e-mail ou boîte partagee vide."
        }
    }
    catch {
        Write-Host "Erreur lors de la suppression du partage de la boite : $($_.Exception.Message)"
    }
}

function PartagerCalendrier {
    try {
        Clear-Host
        Write-Host "Adresse du courriel a partager le Calendrier"
        $SharedMailbox = Read-Host "Courriel"
        Write-Host "Adresse de l'utilisateur a qui vous donnez l'acces"
        $RecipientEmail = Read-Host "Courriel"

        if ($SharedMailbox -ne "" -and $RecipientEmail -ne "") {
            Clear-Host
            $DomainEmail = Get-DomainFromEmail($SharedMailbox)
            Connect-ExchangeOnline -DelegatedOrganization $DomainEmail

            Add-MailboxPermission -Identity $SharedMailbox -User $RecipientEmail -AccessRights FullAccess -InheritanceType All -Automapping $false
            Write-Host "Acces accorde a $RecipientEmail pour le Calendrier de $SharedMailbox."
            pause
        }
        else {
            Write-Host "Adresse e-mail ou boite partagee vide."
        }
    }
    catch {
        Write-Host "Erreur lors du partage du calendrier : $($_.Exception.Message)"
    }
}

function DePartagerCalendrier {
    try {
        Clear-Host
        Write-Host "Adresse du courriel a supprimer le partage du Calendrier"
        $SharedMailbox = Read-Host "Courriel"
        Write-Host "Adresse de l'utilisateur a qui vous enlevez l'acces"
        $RecipientEmail = Read-Host "Courriel"

        if ($SharedMailbox -ne "" -and $RecipientEmail -ne "") {
            $DomainEmail = Get-DomainFromEmail($SharedMailbox)
            Connect-ExchangeOnline -DelegatedOrganization $DomainEmail

            Remove-MailboxPermission -Identity $SharedMailbox -User $RecipientEmail -AccessRights FullAccess -InheritanceType All
            Write-Host "Acces supprime a $RecipientEmail pour le calendrier de $SharedMailbox."
            pause
        }
        else {
            Write-Host "Adresse e-mail ou boîte partagee vide."
        }
    }
    catch {
        Write-Host "Erreur lors de la suppression du partage du calendrier : $($_.Exception.Message)"
    }
}

function TimeZoneCourrielFR {
    try {
        Clear-Host
        Write-Host "Adresse du courriel a changer la Time Zone en FR"
        $RecipientEmail = Read-Host "Courriel"

        if ($RecipientEmail -ne "") {
            $DomainEmail = Get-DomainFromEmail($RecipientEmail)
            Connect-ExchangeOnline -DelegatedOrganization $DomainEmail

            Write-host "AVANT"
            Get-MailboxRegionalConfiguration $RecipientEmail
            Write-Host ""
            Set-MailboxRegionalConfiguration $RecipientEmail -Language "fr-FR" -DateFormat "dd/MM/yyyy" -TimeFormat "HH:mm" -TimeZone "Eastern Standard Time"
            Write-host "APRES MODIFICATION"
            Get-MailboxRegionalConfiguration $RecipientEmail
            Write-Host "Time Zone modifiée pour $RecipientEmail"
            pause
        }
        else {
            Write-Host "Adresse e-mail vide."
        }
    }
    catch {
        Write-Host "Erreur lors de la modification de la Time Zone : $($_.Exception.Message)"
    }
}

function PartagerCalendrierCompagnie {
    try {
        Clear-Host
        Write-Host "Donner les droits a tout les Calendrier Salles de Compagnie"
        $RecipientEmail = Read-Host "Courriel"

        $DomainEmail = Get-DomainFromEmail($RecipientEmail)
        Connect-ExchangeOnline -DelegatedOrganization $DomainEmail

        if ($RecipientEmail -ne "") {
            $users = @($RecipientEmail)

            $mailboxesFirst = Get-Mailbox -Filter "Name -like '*salle*'"
            $mailboxes = $mailboxesFirst | Where-Object { $_.Name -like '*demo-*' }
            Clear-Host;
            $mailboxes | % {$currentMailbox = $_; Write-Warning $currentMailbox }
            $confirmer = Read-Host "Confirmer les Calendriers a ajouter O/N "

            if ($confirmer -eq "o" -or $confirmer -eq "O") {
                $mailboxes | % {
                    $currentMailbox = $_
                    $users | % {
                        $currentUser = $_
                        Add-MailboxPermission -Identity $currentMailbox -User $currentUser -AccessRights FullAccess -InheritanceType All -AutoMapping $false
                    }
                }
                Pause
            }
        }
    }
    catch {
        Write-Host "Erreur lors du partage des calendriers Compagnie : $($_.Exception.Message)"
    }
}

function DePartagerCalendrierCompagnie {
    try {
        Clear-Host
        Write-Host "Enlever les droits a tout les Calendrier Salles de Compagnie"
        $RecipientEmail = Read-Host "Courriel"

        $DomainEmail = Get-DomainFromEmail($RecipientEmail)
        Connect-ExchangeOnline -DelegatedOrganization $DomainEmail

        if ($RecipientEmail -ne "") {
            $users = @($RecipientEmail)

            $mailboxesFirst = Get-Mailbox -Filter "Name -like '*salle*'"
            $mailboxes = $mailboxesFirst | Where-Object { $_.Name -like '*demo-*' }

            $mailboxes | % { Clear-Host; $currentMailbox = $_; Write-Warning $currentMailbox }
            $confirmer = Read-Host "Confirmer les Calendriers a enlever O/N "

            if ($confirmer -eq "o" -or $confirmer -eq "O") {
                $mailboxes | % {
                    $currentMailbox = $_
                    $users | % {
                        $currentUser = $_
                        Remove-MailboxPermission -Identity $currentMailbox -User $currentUser -AccessRights FullAccess -InheritanceType All
                    }
                }
            }
        }
    }
    catch {
        Write-Host "Erreur lors de la suppression des partages des calendriers Compagnie : $($_.Exception.Message)"
    }
}

function ForcerActivationArchivage {
    try {
        Clear-Host
        Write-Host "Forcer l'activation de l'archivage pour l'usager"
        $RecipientEmail = Read-Host "Courriel"
        $DomainEmail = Get-DomainFromEmail($RecipientEmail)
        Connect-ExchangeOnline -DelegatedOrganization $DomainEmail

        if ($RecipientEmail -ne "") {
            Clear-Host
            Get-Mailbox $RecipientEmail | Format-List RetentionPolicy
            Write-Host ""
            $ArchiveConfirm = Read-Host "Valider l'application de l'archivage (o/n)"
            if ($ArchiveConfirm -eq "o" -or $ArchiveConfirm -eq "O") {
                Start-ManagedFolderAssistant $RecipientEmail
                pause
            }
        }
    }
    catch {
        Write-Host "Erreur lors de l'activation de l'archivage : $($_.Exception.Message)"
    }
}


#endregion Function_Menu_Partage
#----------------------------------------------------------------------- Fonction PARTAGE 2.0  ----------------------------------------------------------------------------------------------------
#region Function_Menu_Partage_deux
function ConnectAdminDeux() {
    Clear-Host
    $AdminConnecter = Read-Host "Connecter Admin"
    Connect-ExchangeOnline -UserPrincipalName $AdminConnecter
    return $AdminConnecter
}

function DisconnectAdminDeux() {
    return ""
}

function AjouterBoitePartageDeux($AdminConnecter) {
    if ($AdminConnecter -eq "") { Write-Warning "Aucun Admin connecte!" ; return }
    $CourrielAPartager = Read-Host "Courriel a partager"
    $CourrielDuUser = Read-Host "Donner les droits a"

    if ($CourrielAPartager -and $CourrielDuUser) {
        Add-MailboxPermission -Identity $CourrielAPartager -User $CourrielDuUser -AccessRights FullAccess -InheritanceType All -Automapping $false
        Add-RecipientPermission $CourrielAPartager -AccessRights SendAs -Trustee $CourrielDuUser
        Write-Host "Boite partagee ajoutee avec succes."
    } else {
        Write-Host "Erreur: Variable Vide"
    }
}

function EnleverBoitePartageDeux($AdminConnecter) {
    if ($AdminConnecter -eq "") { Write-Warning "Aucun Admin connecte!" ; return }
    $CourrielADepartager = Read-Host "Courriel a departager"
    $CourrielDuUserAEnlever = Read-Host "ENLEVER les droits a"

    if ($CourrielADepartager -and $CourrielDuUserAEnlever) {
        Remove-RecipientPermission $CourrielADepartager -AccessRights SendAs -Trustee $CourrielDuUserAEnlever
        Remove-MailboxPermission -Identity $CourrielADepartager -User $CourrielDuUserAEnlever -AccessRights FullAccess -InheritanceType All
        Write-Host "Boite partagee enlevee avec succes."
    } else {
        Write-Host "Erreur: Variable Vide"
    }
}

function AjouterCalendrierPartageDeux($AdminConnecter) {
    if ($AdminConnecter -eq "") { Write-Warning "Aucun Admin connecte!" ; return }
    $CalendrierAPartager = Read-Host "Calendrier a partager"
    $CourrielDuUser = Read-Host "Donner les droits a"

    if ($CalendrierAPartager -and $CourrielDuUser) {
        Add-MailboxPermission -Identity $CalendrierAPartager -User $CourrielDuUser -AccessRights FullAccess -InheritanceType All -Automapping $false
        Write-Host "Acces au calendrier ajoute avec succes."
    } else {
        Write-Host "Erreur: Variable Vide"
    }
}

function EnleverCalendrierPartageDeux($AdminConnecter) {
    if ($AdminConnecter -eq "") { Write-Warning "Aucun Admin connecte!" ; return }
    $CalendrierADepartager = Read-Host "Calendrier a departager"
    $CourrielDuUser = Read-Host "ENLEVER les droits a"

    if ($CalendrierADepartager -and $CourrielDuUser) {
        Remove-MailboxPermission -Identity $CalendrierADepartager -User $CourrielDuUser -AccessRights FullAccess -InheritanceType All
        Write-Host "Acces au calendrier enleve avec succes."
    } else {
        Write-Host "Erreur: Variable Vide"
    }
}

function ChangeMailboxLanguageDeux($AdminConnecter) {
    if ($AdminConnecter -eq "") { Write-Warning "Aucun Admin connecte!" ; return }
    $CourrielAChangerLangue = Read-Host "Courriel a changer la langue"

    if ($CourrielAChangerLangue) {
        Set-MailboxRegionalConfiguration $CourrielAChangerLangue -Language "fr-FR" -DateFormat "dd/MM/yyyy" -TimeFormat "HH:mm" -TimeZone "Eastern Standard Time"
        Write-Host "Parametres regionaux modifies pour le compte $CourrielAChangerLangue."
    } else {
        Write-Host "Erreur: Variable Vide"
    }
}

function ActiverArchivageDeux($AdminConnecter) {
    if ($AdminConnecter -eq "") { Write-Warning "Aucun Admin connecte!" ; return }
    $ActiverArchivageUser = Read-Host "Forcer l'archivage de"

    if ($ActiverArchivageUser) {
        Start-ManagedFolderAssistant $ActiverArchivageUser
        Write-Host "Archivage force pour le compte $ActiverArchivageUser."
    } else {
        Write-Host "Erreur: Variable Vide"
    }
}

function InstallerModuleExchangeDeux() {
    Set-ExecutionPolicy unrestricted -Force
    Install-Module -Name ExchangeOnlineManagement -RequiredVersion 3.1.0
    Write-Host "Module Exchange installe avec succes."
}
#endregion Function_Menu_Partage_deux
#----------------------------------------------------------------------- Fonction Outis Divers  ---------------------------------------------------------------------------------------------------
#region Function_Outis_Divers

function ObservateurEvenement { Start-Process eventvwr.msc }
# Fonction pour créer un nouveau dossier #Explorateur Powershell
function New-dir {
    param (
        [string]$directoryName
    )
    try {
        New-Item -Path . -Name $directoryName -ItemType Directory -ErrorAction Stop
        Write-Host "Le dossier $directoryName a été créé avec succès."
    }
    catch {
        Write-Host "Impossible de créer le dossier $directoryName. Raison : $($_.Exception.Message)" -ForegroundColor DarkRed
    }
}
# Fonction pour obtenir la taille d'un dossier #Explorateur Powershell
function Get-FolderSize {
    param (
        [string]$folderPath
    )
    $folderSize = 0
    try {
        $items = Get-ChildItem -Path $folderPath -Recurse -Force -ErrorAction SilentlyContinue
        foreach ($item in $items) {
            try {
                $folderSize += $item.Length
            }
            catch {
                Write-Host "Erreur lors de la lecture de l'élément $($item.FullName): $($_.Exception.Message)" -ForegroundColor DarkRed
            }
        }
    }
    catch {
        Write-Host "Impossible d'accéder à $folderPath. Raison : $($_.Exception.Message)" -ForegroundColor DarkRed
    }
    return [math]::Round($folderSize / 1GB, 2)
}
# Fonction pour obtenir la taille d'un fichier #Explorateur Powershell
function Get-FilesSize {
    param (
        [string]$filePath
    )
    $fileSize = 0
    try {
        $item = Get-Item -Path $filePath -Force -ErrorAction SilentlyContinue
        try {
            $fileSize = $item.Length
        }
        catch {
            Write-Host "Erreur lors de la lecture de l'élément $($filePath): $($_.Exception.Message)" -ForegroundColor DarkRed
        }
    }
    catch {
        Write-Host "Impossible d'accéder à $filePath. Raison : $($_.Exception.Message)" -ForegroundColor DarkRed
    }
    return [math]::Round($fileSize / 1GB, 2)
}
# Function supprimer contenu d'un fichier #Explorateur Powershell
function Clear-FileContent {
    param (
        [string]$filePath
    )

    try {
        if (Test-Path -Path $filePath -PathType Leaf) {
            # Supprime le contenu du fichier
            Clear-Content -Path $filePath -Force
            Write-Host "Le contenu du fichier $filePath a été supprimé avec succès." -ForegroundColor Green
        }
        else {
            Write-Host "Le fichier $filePath n'existe pas ou n'est pas un fichier." -ForegroundColor Red
        }
    }
    catch {
        Write-Host "Impossible de vider le fichier $filePath. Raison : $($_.Exception.Message)" -ForegroundColor DarkRed
    }
}
# Fonction pour l'explorateur PowerShell personnalisé
function ExplorateurPowershellCustom {
    Clear-Host
    $leaveExplorateur = $true
    Set-Location "C:\"  # Définir le chemin initial
    While ($leaveExplorateur) {
        # Affiche le chemin actuel
        $cheminActuel = Get-Location
        Write-Host $cheminActuel
        Write-Host "-------------------------------------------------------" -ForegroundColor Yellow
        Write-Host ""

        # Récupère les dossiers et fichiers dans le chemin actuel
        $rootPath = Get-Location
        try {
            $folders = Get-ChildItem -Path $rootPath -Directory -ErrorAction SilentlyContinue
            $files = Get-ChildItem -Path $rootPath -File -ErrorAction SilentlyContinue
        }
        catch {
            Write-Host "Erreur d'accès au chemin : $($_.Exception.Message)" -ForegroundColor DarkRed
            continue
        }

        # Boucle pour afficher les dossiers avec leur taille
        foreach ($folder in $folders) {
            try {
                $folderSize = Get-FolderSize -folderPath $folder.FullName
                Write-Host "$($folder.Name.PadRight(40)) $($folderSize.ToString('F2').PadRight(5)) GB" -ForegroundColor Yellow
            }
            catch {
                Write-Host "Impossible de lire le dossier (Vérifiez les droits)" -ForegroundColor DarkRed
            }
        }

        # Boucle pour afficher les fichiers avec leur taille
        foreach ($file in $files) {
            try {
                $fileSize = Get-FilesSize -filePath $file.FullName
                Write-Host "$($file.Name.PadRight(40)) $($fileSize.ToString('F2').PadRight(5)) GB"
            }
            catch {
                Write-Host "Impossible de lire le fichier (Vérifiez les droits)" -ForegroundColor DarkRed
            }
        }

        Write-Host ""
        Write-Host "-------------------------------------------------------" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Supprimer         : Remove-Item (Nom)" -ForegroundColor DarkRed
        Write-Host "Supprimer Contenu : Clear-FileContent"
        Write-Host "Creer Fichier     : New-Item (Nom.txt)"
        Write-Host "Creer Dossier     : New-Item -ItemType Directory (Nom)"
        Write-Host "Entrer            : cd (Nom)"
        Write-Host "Sortir            : cd .."
        Write-Host "Quitter           : q" -ForegroundColor Red
        Write-Host ""
        
        # Récupère la commande utilisateur
        $Commande = Read-Host "Commande"
        if ($Commande -eq "q" -or $Commande -eq "Q") {
            $leaveExplorateur = $false
        }
        else {
            Clear-Host
            try {
                # Exécute la commande saisie par l'utilisateur
                Invoke-Expression $Commande
            }
            catch {
                Write-Host "Erreur lors de l'exécution de la commande : $($_.Exception.Message)" -ForegroundColor DarkRed
            }
        }

        Write-Host "-------------------------------------------------------" -ForegroundColor Yellow
        Write-Host ""
    }
}
function ExtracteurDeVideos {
    Clear-Host
    Write-Host "Activer Admin Mode pour faire des recherches Avancees"
    AdminModeWrite
    Write-Host ""
    Get-PSDrive -PSProvider FileSystem
    Write-Host ""

    Write-Host "Veuillez selectionner le dossier source pour la recherche de videos"
    $DossierSource = Select-Folder
    Write-Host "Dossier source selectionne : $DossierSource"
    Write-Host ""

    Write-Host "Veuillez selectionner le dossier de destination pour enregistrer les videos trouvees"
    $DossierDestination = Select-Folder
    Write-Host "Dossier de destination sélectionne : $DossierDestination"
    Write-Host ""

    # Demande des extensions de fichiers à rechercher
    $extensions = Read-Host "Entrez les extensions de video a rechercher (separees par des virgules, par exemple : mp4,avi,mov,wmv,mkv,flv,mpeg,3gp) "
    $extensionsArray = $extensions -split ',\s*' | ForEach-Object { "*.$_" }

    # Confirmation pour exécuter l'opération
    $confirmationExecution = Read-Host "Souhaitez-vous executer l'operation ? (O/N)"
    if ($DossierSource -and $DossierDestination -and ($confirmationExecution -eq "O" -or $confirmationExecution -eq "o")) {

        # Création du dossier de destination pour les vidéos si non existant
        $cheminExtraction = Join-Path -Path $DossierDestination -ChildPath "VideosExtraction"
        if (-not (Test-Path $cheminExtraction)) {
            New-Item -Path $cheminExtraction -ItemType Directory | Out-Null
        }

        # Extraction des vidéos et copie vers le dossier de destination
        Write-Host "Recherche et copie des videos en cours..."
        foreach ($extension in $extensionsArray) {
            Get-ChildItem -Path $DossierSource -Recurse -Include $extension | ForEach-Object {
                $cheminVideoDestination = Join-Path -Path $cheminExtraction -ChildPath "$($_.BaseName)$($_.Extension)"
                Write-Host "Copie de : $($_.FullName) vers $cheminVideoDestination"
                Copy-Item $_.FullName -Destination $cheminVideoDestination -ErrorAction SilentlyContinue
            }
        }

        Write-Host "`nOperation terminee ! Toutes les videos ont été copiees dans le dossier : $cheminExtraction"
        pause
    } else {
        Write-Host "Operation annulee par l'utilisateur."
        pause
    }
}


function ExtracteurDePhotos {
    Clear-Host
    Write-Host "Activer Admin Mode pour faire des recherches Avancees"
    AdminModeWrite
    Write-Host ""

    Get-PSDrive -PSProvider FileSystem
    Write-Host ""

    # Sélection du dossier source
    Write-Host "Selectionnez le dossier source pour la recherche d'images"
    $DossierSource = Select-Folder
    Write-Host "Dossier source selectionne : $DossierSource"
    Write-Host ""
    
    # Sélection du dossier de destination
    Write-Host "Selectionnez le dossier de destination pour enregistrer les images trouvees"
    $DossierDestination = Select-Folder
    Write-Host "Dossier de destination selectionne : $DossierDestination"
    Write-Host ""
    
    # Sélection des extensions de fichier à rechercher
    $extensions = Read-Host "Entrez les extensions a rechercher (ex : bmp,gif,heic,heif,jpg,jpeg,png,tif,tiff) (autres : cr2,nef,arw) "
    $extensionsArray = $extensions -split ",\s*"

    # Confirmation pour exécuter l'opération
    $confirmationExecution = Read-Host "Souhaitez-vous executer l'operation ? (O/N)"
    if ($DossierSource -and $DossierDestination -and ($confirmationExecution -eq "O" -or $confirmationExecution -eq "o")) {

        # Création du dossier de destination pour les images si non existant
        $cheminExtraction = Join-Path -Path $DossierDestination -ChildPath "ImagesExtraction"
        if (-not (Test-Path $cheminExtraction)) {
            New-Item -Path $cheminExtraction -ItemType Directory | Out-Null
        }

        # Extraction des images et copie vers le dossier de destination
        Write-Host "Recherche et copie des images en cours..."
        foreach ($extension in $extensionsArray) {
            Get-ChildItem -Path $DossierSource -Recurse -Include "*.$extension" | ForEach-Object {
                $cheminImageDestination = Join-Path -Path $cheminExtraction -ChildPath "$($_.BaseName)$($_.Extension)"
                Write-Host "Copie de : $($_.FullName) vers $cheminImageDestination"
                Copy-Item $_.FullName -Destination $cheminImageDestination -ErrorAction SilentlyContinue
            }
        }

        Write-Host "`nOperation terminee ! Toutes les images ont ete copiees dans le dossier : $cheminExtraction"
        pause
    } else {
        Write-Host "Operation annulee par l'utilisateur."
        pause
    }
}


function RecherchePowershell {
    AdminModeWrite; Write-Host ""
    Write-Host "Selectionner emplacement a rechercher"
    $emplacementDeRecherche = Select-Folder
    $recherche = Read-host "Rechercher Item dans $($emplacementDeRecherche)"
    Clear-Host
    Get-ChildItem -Path "$($emplacementDeRecherche)" -Recurse -File -Filter "*$($recherche)*" | 
    Select-Object @{Name='Nom du fichier'; Expression={$_.Name}},
                  #@{Name='Droits'; Expression={(Get-Acl $_.FullName).AccessToString}},
                  @{Name='Localisation'; Expression={$_.FullName}} |
    Format-Table -AutoSize

    pause           
}
# Recherche via everything, le DOWNLOAD si pas installer et execute
function RechercheEverything {

    # Définir l'URL de téléchargement pour Everything
    $downloadUrl = $LIEN_DOWNLOAD_EVERYTHING
    $installerPath = "$env:TEMP\Everything-Setup.exe"
    $installPath = "C:\Program Files\Everything\Everything.exe"

    # Vérifier si Everything est déjà installé
    if (Test-Path -Path $installPath) {
        Write-Output "Everything est déjà installé. Lancement de l'application..."
        Start-Process -FilePath $installPath
    } else {
        Write-Output "Everything n'est pas installé. Téléchargement en cours..."
        
        # Télécharger le fichier d'installation
        Invoke-WebRequest -Uri $downloadUrl -OutFile $installerPath

        # Vérifier si le fichier est téléchargé
        if (Test-Path -Path $installerPath) {
            Write-Output "Téléchargement terminé. Installation de Everything..."
            
            # Exécuter l'installateur
            Start-Process -FilePath $installerPath -ArgumentList "/S" -Wait

            # Supprimer le fichier d'installation après l'installation
            Remove-Item -Path $installerPath -Force

            # Vérifier si l'installation a réussi
            if (Test-Path -Path $installPath) {
                Write-Output "Installation réussie. Lancement de Everything..."
                Start-Process -FilePath $installPath
            } else {
                Write-Output "Erreur : L'installation de Everything a échoué."
            }
        } else {
            Write-Output "Erreur : Téléchargement de Everything échoué."
        }
    }
}

#endregion Function_Outis_Divers
#----------------------------------------------------------------------- Fonction Information System  ---------------------------------------------------------------------------------------------------
#region Function_Information_System
function CheckBatteryStatusHtml {
    # Emplacement du fichier de rapport
    $reportPath = "$env:TEMP\battery-report.html"

    # Générer le rapport d'état de la batterie
    powercfg /batteryreport /output $reportPath

    # Vérifier si le rapport a été généré
    if (Test-Path $reportPath) {
        Write-Host "Le rapport de batterie a ete genere avec succes. Ouverture dans le navigateur..."
        # Ouvrir le fichier HTML dans le navigateur par défaut
        Start-Process $reportPath
    } else {
        Write-Host "Erreur : le rapport de batterie n'a pas pu etre genere."
    }
}

function CheckDeviceStatusHtml {
    # Emplacement du fichier de rapport
    $reportPath = "$env:TEMP\device-report.html"

    # Générer le rapport de configuration système au format texte
    msinfo32 /report "$env:TEMP\device-report.txt"

    # Conversion du rapport texte en HTML
    $content = Get-Content "$env:TEMP\device-report.txt"
    $htmlContent = "<html><body><pre>" + ($content -join "`n") + "</pre></body></html>"
    Set-Content -Path $reportPath -Value $htmlContent

    # Vérifier si le rapport a été généré
    if (Test-Path $reportPath) {
        Write-Host "Le rapport de peripheriques a ete genere avec succes. Ouverture dans le navigateur..."
        # Ouvrir le fichier HTML dans le navigateur par défaut
        Start-Process $reportPath
    } else {
        Write-Host "Erreur : le rapport de peripheriques n'a pas pu etre genere."
    }
    pause
}
function GenerateNetworkReport {
    # Commande pour générer le rapport réseau
    netsh wlan show wlanreport
    $reportPath = "C:\ProgramData\Microsoft\Windows\WlanReport\wlan-report-latest.html"

    # Vérification et ouverture
    if (Test-Path $reportPath) {
        Start-Process $reportPath
        Write-Host "Rapport reseau genere et ouvert avec succes."
    } else {
        Write-Host "Erreur : Le rapport reseau n'a pas pu être genere."
    }
}
function GenerateEnergyReport {
    # Génération du rapport
    powercfg /energy /output "energy-report.html"
    $reportPath = ".\energy-report.html"

    # Vérification et ouverture du rapport
    if (Test-Path $reportPath) {
        Start-Process $reportPath
        Write-Host "Rapport d'energie genere et ouvert avec succes."
    } else {
        Write-Host "Erreur : le rapport d'energie n'a pas pu etre genere."
    }
}
function OpenReliabilityReport {
    # Ouvre le moniteur de fiabilité
    Start-Process "perfmon.exe" -ArgumentList "/rel"
    Write-Host "Le Moniteur de fiabilite s'est ouvert. Consulte le rapport pour plus de details."
}
function GenerateSystemDiagnosticsReport {
    # Génération et ouverture du rapport de diagnostic système
    Start-Process "perfmon.exe" -ArgumentList "/report"
    Write-Host "Rapport de diagnostic systeme lance dans l'Observateur d'evenements."
}
function GenerateSMARTReport {
    # Obtient les disques physiques
    $disks = Get-PhysicalDisk

    # Crée un rapport sur chaque disque
    $report = @()
    foreach ($disk in $disks) {
        $diskInfo = New-Object PSObject -Property @{
            "Nom du Disque" = $disk.DeviceID
            "État de Fonctionnement" = $disk.OperationalStatus -join ", "
            "Type" = $disk.MediaType
            "Capacité (GB)" = [math]::Round($disk.Size / 1GB, 2)
            "Échec Prédit" = $disk.FailurePredicted
        }
        $report += $diskInfo
    }

    # Affiche le rapport
    $report | Format-Table -AutoSize
    pause
}

#endregion Function_Information_System
#----------------------------------------------------------------------- Fonction Password Generators  ---------------------------------------------------------------------------------------------------
#region Function_PasswordGenerators

function GeneratePasswordBase {
     # Générer la première lettre en majuscule (A-Z)
    $firstLetter = [char](65..90 | Get-Random)

    # Générer deux lettres minuscules aléatoires (a-z)
    $lowercaseLetters = -join ((97..122) | Get-Random -Count 2 | ForEach-Object { [char]$_ })

    # Générer cinq chiffres aléatoires (0-9)
    $numbers = -join ((48..57) | Get-Random -Count 5 | ForEach-Object { [char]$_ })

    # Combiner les parties pour former le mot de passe
    $password = "$($firstLetter)$($lowercaseLetters)$($numbers)"

    write-host ""
    write-host "Mot de passe Generer : "$password
    write-host ""
    pause
}

function GeneratePasswordInterComplex {

    write-host ""
    
    # Demander la longueur du mot de passe et vérifier que l'entrée est un nombre valide
    $passwordLength = Read-Host "Selectionner la longueur du mot de passe"
    
    # Vérifier si l'entrée est un nombre et si elle est supérieure à 0
    while (-not ($passwordLength -match '^\d+$') -or $passwordLength -lt 1) {
        write-host "Veuillez entrer un nombre valide supérieur à 0."
        $passwordLength = Read-Host "Selectionner la longueur du mot de passe"
    }
    
    write-host ""

    $password = ""
    $specialChars = "!@#$%^&*()_-+=<>?{}[]|"
    $allSpecialChars = $specialChars.ToCharArray()

    try {
        for ($i = 0; $i -lt $passwordLength; $i++) {
            # Choisir un type de caractère aléatoire : majuscule, minuscule, chiffre ou spécial
            $randomNumberChoice = Get-Random -Minimum 0 -Maximum 4
            
            if ($randomNumberChoice -eq 0) {
                # Majuscule aléatoire
                $firstLetter = [char](65..90 | Get-Random)
                $password += $firstLetter
            }
            elseif ($randomNumberChoice -eq 1) {
                # Minuscule aléatoire
                $lowercaseLetter = [char](97..122 | Get-Random)
                $password += $lowercaseLetter
            }
            elseif ($randomNumberChoice -eq 2) {
                # Chiffre aléatoire
                $number = [char](48..57 | Get-Random)
                $password += $number
            }
            elseif ($randomNumberChoice -eq 3) {
                # Caractère spécial aléatoire
                $randomChar = $allSpecialChars | Get-Random
                $password += $randomChar
            }
        }
    }
    catch {
        Write-Host "Une erreur s'est produite lors de la generation du mot de passe : $_"
        pause
        return
    }
   
    # Affichage du mot de passe généré
    write-host ""
    Write-Host "Mot de passe genere : $password"
    write-host ""
    Pause
}

function GeneratePasswordMots {
    Write-Host ""
    
    # Demander la longueur du mot de passe (en nombre de mots) et vérifier que l'entrée est un nombre valide
    $passwordLength = Read-Host "Selectionnez le nombre de mots"
    
    # Vérifier si l'entrée est un nombre et si elle est supérieure à 0
    while (-not ($passwordLength -match '^\d+$') -or $passwordLength -lt 1) {
        Write-Host "Veuillez entrer un nombre valide superieur à 0."
        $passwordLength = Read-Host "Selectionnez le nombre de mots"
    }
    
    Write-Host ""
    
    # Initialiser les variables
    $password = ""
    $specialChars = "!@#$%^&*()_-+=<>?{}[]|"
    $allSpecialChars = $specialChars.ToCharArray()

    # Liste de mots pour la génération
    $ListeDeMots = @(
        "Aventure", "Mystere", "Enigme", "Decouverte", "Creation", "Inspiration", "Energie", "Passion", 
        "Courage", "Esprit", "Vision", "Force", "Etoile", "Ocean", "Voyage", "Espace", "Nature", "Tempete", 
        "Harmonie", "Liberte", "Connexion", "Horizon", "Mouvement", "Symbiose", "Art", "Lumiere", "Destin", 
        "Philosophie", "Realite", "Evasion", "Symbole", "Culture", "Technologie", "Element", "Univers", 
        "Exploration", "Perspective", "Transition", "Voyant", "Structure", "Environnement", "Evolution", 
        "Invention", "Concept", "Expansion", "Influence", "Creativite", "Innovation", "Serenite", 
        "Enigmatique", "Enthousiasme", "Essence", "Progres", "Infiniment", "Mecanisme", "Inattendu", 
        "Renaissance", "Flux", "Transforme", "Merveille", "Ideal", "Sens", "Aube", "Vitalite", "Futur", 
        "Nouvelle", "Echange", "Ambition", "Impact", "Reflexion", "Dynamisme", "Recul", "Confiance", 
        "Patience", "Amelioration", "Experience", "Esprit", "Transparence", "Ordre", "Raison", 
        "Reverie", "Intuition", "Envol", "Mystique", "Anticipation", "Charisme", "Energie", 
        "Equilibre", "Perspective", "Monde", "Idee", "Emancipation", "Originalite", "Recit", 
        "Emprise", "Origine", "Essence", "Savoir", "Point", "Decalage", "Mystique", "Unite",
        "arbre", "soleil", "montagne", "riviere", "etoile", "foret", "lac", "nuage", 
        "tempete", "pluie", "feu", "vent", "glace", "sable", "mer", "ocean", "oiseau",
        "poisson", "ile", "plante", "rocher", "desert", "galaxie", "comete", "planete",
        "univers", "galet", "cristal", "mystere", "orage", "flamme", "terre", "eclair",
        "lune", "champ", "ciel", "falaise", "cascade", "prairie", "canyon", "volcan",
        "glacier", "vague", "ruisseau", "tornade", "mont", "grotte", "reflet", "horizon", "de", "et", "les", "des", "mes", "ses", "moi", "toi"
    )

    # Construction du mot de passe en choisissant des mots aléatoires et des caractères spéciaux
    for ($i = 1; $i -le $passwordLength; $i++) {
        # Ajouter un mot aléatoire
        $motAleatoire = Get-Random -InputObject $ListeDeMots
        $password += $motAleatoire
        
        # Ajouter un caractère spécial aléatoire, sauf pour le dernier mot
        if ($i -lt $passwordLength) {
            $specialChar = Get-Random -InputObject $allSpecialChars
            $password += $specialChar + "-"
        }
    }

    # Afficher le mot de passe généré
    Write-Output "Mot de passe genere : $password"
    pause
}

#endregion Function_PasswordGenerators
# ==============================================
# Main Script : Le flux principal du script
# ==============================================
#region Main_Script
AnimationTML
While ($true) {    
    #MenuPrincipale
    $ChoixMenuPrincipale = MenuPrincipale
    Clear-Host

    switch ($ChoixMenuPrincipale) {
        #Quitter
        "0" { OptionQuiterLeLogiciel }

        #Menu Partager 3.0 Auto Admin
        "1" { OptionPartageTrois }
        #Menu Partager 2.0 Admin
        "2" { OptionPartageDeux }
        #Ninja One
        "3" { OptionDattoRMM }
        #Site Web Test
        "4" { OptionSiteWeb }
        #Menu Registre
        "5" { OptionRegistre }
        #Menu Services
        "6" { OptionServices }
        #Menu Configuration
        "7" { OptionConfiguration }
        #Menu Reparation
        "8" { OptionReparation } 
        #Menu WinGet
        "9" { OptionWinGet }

        
        #Outils Divers
        "10" { OptionOutilsDivers }
        #Password Generator
        "11" { OptionPasswordGenerator }
        #Chronometre
        "12" { OptionCrono }

        #Menu Information System
        "15" { OptionInfoSystem }
        #Information de base compagnie
        "16" { OptionInfoCompagnie }

        #Administrateur
        "01" { OptionAdminMode }
        #Credit
        "02" { OptionCredit } 
        #Default 
        Default { DefaultError404 }
    }
}

#endregion Main_Script