#Requires -Version 5.1
<#
.SYNOPSIS
Script de compilation de présentation PowerPoint

.DESCRIPTION
Ce script compile une nouvelle présentation PowerPoint en assemblant
des slides issus de plusieurs fichiers sources selon une table descriptive.

.AUTHOR
Franck Ginhoux
#>

# ==================== INITIALISATION ====================
# Charger les fichiers de configuration et de fonctions

# Determiner le repertoire racine du projet
if ($PSScriptRoot) {
    $ScriptRoot = Split-Path -Path $PSScriptRoot -Parent
} else {
    # Si PSScriptRoot n'est pas defini (ISE), essayer d'utiliser le repertoire courant
    # ou demander a l'utilisateur
    $CurrentDir = Get-Location
    
    # Verifier si on est deja dans le bon repertoire
    if (Test-Path "Configs\config.ps1") {
        $ScriptRoot = $CurrentDir
    } else {
        # Chercher le dossier PowerShell en remontant les repertoires
        $TestPath = $CurrentDir
        $Found = $false
        for ($i = 0; $i -lt 5; $i++) {
            if (Test-Path (Join-Path $TestPath "Configs\config.ps1")) {
                $ScriptRoot = $TestPath
                $Found = $true
                break
            }
            $TestPath = Split-Path -Parent $TestPath
        }
        
        if (-not $Found) {
            Write-Error "Impossible de trouver le repertoire du projet. Assurez-vous de lancer le script depuis le repertoire D:\W\Iveco\PowerShell"
            exit 1
        }
    }
}

# Changer le repertoire de travail vers la racine du projet
Set-Location $ScriptRoot

$ConfigPath = Join-Path $ScriptRoot "Configs\config.ps1"
$HelperPath = Join-Path $ScriptRoot "Functions\Helper.ps1"

Write-Host "Repertoire racine: $ScriptRoot" -ForegroundColor Gray
Write-Host "Config: $ConfigPath" -ForegroundColor Gray
Write-Host "Helper: $HelperPath" -ForegroundColor Gray

if (-not (Test-Path $ConfigPath)) {
    Write-Error "Fichier de configuration non trouve : $ConfigPath"
    exit 1
}

if (-not (Test-Path $HelperPath)) {
    Write-Error "Fichier de fonctions non trouve : $HelperPath"
    exit 1
}

# Sourcer les fichiers
. $ConfigPath
. $HelperPath

# ==================== CONFIGURATION ====================

# Paramètres de la présentation à compiler
$ModelePath = "D:\W\Iveco\PowerShell\Modele.pptx"
$OutputFileName = "Test_compil_01.pptx"
$OutputPath = "D:\W\Iveco"

# Table descriptive des pages à compiler
# Chaque élément contient :
# - Source : Chemin du fichier PPTX source
# - PageSource : Numéro de la page à copier (1-indexed)
$Pages = @(
    @{
        Source = "D:\W\Iveco\serveur\Affaires 2024000150SA\01-Dossiers ligne EL-EG\LIGNE EG0\Poste1\SOP P1 1.pptx"
        PageSource = 2
    },
    @{
        Source = "D:\W\Iveco\serveur\Affaires 2024000150SA\01-Dossiers ligne EL-EG\LIGNE EG0\Poste1\SOP P1 2.pptx"
        PageSource = 3
    },
    @{
        Source = "D:\W\Iveco\serveur\Affaires 2024000150SA\01-Dossiers ligne EL-EG\LIGNE EG0\Poste1\SOP P1 1.pptx"
        PageSource = 1
    }
)

# ==================== EXECUTION ====================

Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "Compilation de Presentation PowerPoint" -ForegroundColor Cyan
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host ""

# Appeler la fonction de compilation
New-CompiledPresentation -ModelePath $ModelePath `
                         -OutputFileName $OutputFileName `
                         -OutputPath $OutputPath `
                         -Pages $Pages

Write-Host ""
Write-Host "Compilation terminee !" -ForegroundColor Green
