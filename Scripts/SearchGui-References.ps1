<#
.SYNOPSIS
    Interface graphique de recherche de references

.DESCRIPTION
    Affiche une interface WinForms pour rechercher des références
    dans le fichier XML genere par le script d'extraction

.PARAMETER XmlPath
    Chemin du fichier XML d'extraction
    Par defaut: D:\W\Iveco\RefServeur.xml

.EXAMPLE
    .\Scripts\SearchGui-References.ps1
    
    .\Scripts\SearchGui-References.ps1 -XmlPath "D:\resultats.xml"
#>

param(
    [string]$XmlPath = ""
)

# Charger la configuration
$ConfigPath = Join-Path -Path $PSScriptRoot -ChildPath "..\Configs\config.ps1"
. $ConfigPath

# Utiliser la valeur de config si le paramètre n'est pas fourni
if ([string]::IsNullOrEmpty($XmlPath)) {
    $XmlPath = $Config.ExtractXmlData
}

# Charger les fonctions
$FunctionsPath = Join-Path -Path $PSScriptRoot -ChildPath "..\Functions\Helper.ps1"
. $FunctionsPath

# Lancer l'interface graphique
Show-SearchGui -XmlPath $XmlPath
