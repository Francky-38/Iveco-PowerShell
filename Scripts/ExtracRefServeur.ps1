<#
.SYNOPSIS
    Script d'extraction de references depuis des archives PPTX

.DESCRIPTION
    Parcourt l'arborescence et traite tous les fichiers PPTX pour extraire 
    les references (5-10 chiffres avec optionnel T/R/S en debut)
    Genere un unique fichier XML avec toutes les references trouvees

.EXAMPLE
    .\Scripts\ExtracRefServeur.ps1
    
    Pour traiter un dossier specifique:
    .\Scripts\ExtracRefServeur.ps1 -RootPath "D:\W\Iveco\serveur"
    
    Avec un fichier de sortie personnalise:
    .\Scripts\ExtracRefServeur.ps1 -RootPath "D:\W\Iveco\serveur" -OutputFile "D:\resultats.xml"
#>

param(
    [string]$RootPath = "",
    [string]$OutputFile = ""
)

# Charger la configuration
$ConfigPath = Join-Path -Path $PSScriptRoot -ChildPath "..\Configs\config.ps1"
. $ConfigPath

# Utiliser les valeurs de config si les paramètres ne sont pas fournis
if ([string]::IsNullOrEmpty($RootPath)) {
    $RootPath = $Config.ExtractionRootPath
}
if ([string]::IsNullOrEmpty($OutputFile)) {
    $OutputFile = $Config.ExtractXmlData
}

# Charger les fonctions
$ScriptPath = Split-Path -Path $MyInvocation.MyCommand.Definition
$FunctionsPath = Join-Path -Path $ScriptPath -ChildPath "..\Functions\Helper.ps1"
. $FunctionsPath

Write-Host ""
Write-Host "====== Extraction de References depuis PPTX - Projet Iveco ======" -ForegroundColor Cyan
Write-Host ""

Get-WelcomeMessage -Message "Extraction de references depuis l'arborescence des affaires"
Get-SystemInfo

# Lancer l'extraction sur l'arborescence
Export-PptxReferencesFromTree -RootPath $RootPath -OutputFile $OutputFile

Write-Host ""
Write-Host "OK - Script termine avec succes!" -ForegroundColor Green
Write-Host ""
