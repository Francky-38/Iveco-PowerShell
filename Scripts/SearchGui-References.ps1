<#
.SYNOPSIS
    Interface graphique de recherche de references

.DESCRIPTION
    Affiche une interface WinForms pour rechercher des références
    dans le fichier ZIP genere par le script d'extraction

.PARAMETER DataPath
    Chemin du fichier ZIP (base sans extension ou .zip)
    Par defaut: valeur de config ExtractXmlDataPath

.EXAMPLE
    .\Scripts\SearchGui-References.ps1
    
    .\Scripts\SearchGui-References.ps1 -DataPath "D:\RefServeur"
#>

param(
    [string]$DataPath = ""
)

# Charger la configuration
$ConfigPath = Join-Path -Path $PSScriptRoot -ChildPath "..\Configs\config.ps1"
. $ConfigPath

# Utiliser la valeur de config si le paramètre n'est pas fourni
if ([string]::IsNullOrEmpty($DataPath)) {
    $DataPath = $Config.ExtractXmlDataPath
}

# Charger les fonctions
$FunctionsPath = Join-Path -Path $PSScriptRoot -ChildPath "..\Functions\Helper.ps1"
. $FunctionsPath

# Lancer l'interface graphique avec le chemin des données
Show-SearchGui -DataPath $DataPath
