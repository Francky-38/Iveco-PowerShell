<#
.SYNOPSIS
    Script d'initialisation du projet PowerShell Iveco

.DESCRIPTION
    Configure l'environnement et charge les dependances necessaires

.EXAMPLE
    .\setup.ps1
#>

param(
    [switch]$Force
)

Write-Host ""
Write-Host "====== Setup Projet PowerShell Iveco ======" -ForegroundColor Green
Write-Host ""

# Verification de la version PowerShell
$PSVersionActual = $PSVersionTable.PSVersion

if ($PSVersionActual.Major -eq 5 -and $PSVersionActual.Minor -ge 1) {
    Write-Host "OK - PowerShell v$($PSVersionActual.Major).$($PSVersionActual.Minor) detecte" -ForegroundColor Green
} elseif ($PSVersionActual.Major -lt 5) {
    Write-Host "Attention: PowerShell 5.1 ou superieur requis" -ForegroundColor Red
    Write-Host "Version actuelle: $($PSVersionActual.Major).$($PSVersionActual.Minor)" -ForegroundColor Red
    exit 1
}

Write-Host "OK - PowerShell v$PSVersionActual detecte" -ForegroundColor Green
Write-Host ""

# Verifier la structure des dossiers
$Folders = @("Functions", "Scripts", "Tests", "Configs")

foreach ($Folder in $Folders) {
    if (Test-Path -Path $Folder) {
        Write-Host "OK - Dossier '$Folder' trouve" -ForegroundColor Green
    } else {
        Write-Host "Création du dossier '$Folder'" -ForegroundColor Yellow
        New-Item -ItemType Directory -Name $Folder -Force | Out-Null
    }
}

Write-Host ""

# Charger la configuration
Write-Host "Chargement de la configuration..." -ForegroundColor Yellow
$ConfigPath = Join-Path -Path $PSScriptRoot -ChildPath "Configs\config.ps1"
if (Test-Path -Path $ConfigPath) {
    . $ConfigPath
    Write-Host "OK - Configuration chargee" -ForegroundColor Green
    Write-Host "  - Projet: $($Global:ProjectName)" -ForegroundColor Cyan
    Write-Host "  - Version: $($Global:ProjectVersion)" -ForegroundColor Cyan
    Write-Host "  - Environnement: $($Config.Environment)" -ForegroundColor Cyan
} else {
    Write-Host "Attention: Fichier de configuration non trouve" -ForegroundColor Red
}

Write-Host ""
Write-Host "OK - Projet pret a etre utilise!" -ForegroundColor Green
Write-Host ""
Write-Host "Prochaines etapes:" -ForegroundColor Yellow
Write-Host "  1. Ajouter vos fonctions dans ./Functions/" -ForegroundColor Yellow
Write-Host "  2. Ajouter vos scripts dans ./Scripts/" -ForegroundColor Yellow
Write-Host "  3. Charger les fonctions avec : . ./Functions/*.ps1" -ForegroundColor Yellow
Write-Host ""
