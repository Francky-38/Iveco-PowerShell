<#
.SYNOPSIS
    Interface graphique de recherche de texte a la volee

.DESCRIPTION
    Affiche une interface WinForms pour rechercher un texte libre directement
    dans les fichiers PPTX de l'arborescence (memes fichiers et meme configuration
    que ExtracRefServeur.ps1). Aucune base n'est generee : les fichiers sont
    parcourus a la volee et les resultats apparaissent au fur et a mesure.

.PARAMETER RootPath
    Chemin racine de l'arborescence a scruter.
    Par defaut: valeur de config ExtractionRootPath

.EXAMPLE
    .\Scripts\SearchGui-LiveText.ps1

    .\Scripts\SearchGui-LiveText.ps1 -RootPath "D:\W\Iveco\serveur"
#>

param(
    [string]$RootPath = "",
    [array]$SubPathStructures = @()
)

# Charger la configuration
$ConfigPath = Join-Path -Path $PSScriptRoot -ChildPath "..\Configs\config.ps1"
. $ConfigPath

# Utiliser les valeurs de config si les parametres ne sont pas fournis
if ([string]::IsNullOrEmpty($RootPath)) {
    $RootPath = $Config.ExtractionRootPath
}
if ($SubPathStructures.Count -eq 0) {
    $SubPathStructures = $Config.SubPathStructures
}

# Charger les fonctions
$FunctionsPath = Join-Path -Path $PSScriptRoot -ChildPath "..\Functions\Helper.ps1"
. $FunctionsPath

# Lancer l'interface graphique de recherche a la volee
Show-LiveTextSearchGui -RootPath $RootPath -SubPathStructures $SubPathStructures
