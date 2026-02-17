# Configuration du projet Iveco

# Variables globales
$Global:ProjectName = "MY-Project-PowerShell"
$Global:ProjectVersion = "9.0.0"
$Global:ProjectAuthor = "Franck Ginhoux"

# Paramètres
$Config = @{
    Environment = "Development"  # Development, Testing, Production
    LogLevel = "Info"            # Debug, Info, Warning, Error
    LogFile = ".\Logs\project.log"
}

# Paramètres d'extraction et recherche des references
$Config.ExtractionRootPath = "D:\W\Iveco\serveur"
# Chemin relatif du fichier de données (base sans extension) - a la racine du projet
$Config.ExtractXmlData = "RefServeur"
$ProjectRoot = Split-Path -Parent $PSScriptRoot
$Config.ExtractXmlDataPath = Join-Path -Path $ProjectRoot -ChildPath $Config.ExtractXmlData
# Regex pour extraire les references (ex: [THRS]? suivi de 5 a 10 chiffres)
$Config.ReferenceRegexPattern = '[THRS]?\d{5,10}'

# Sous-chemins à scruter dans chaque dossier d'affaire
# Laissez vide pour scruter directement le dossier d'affaire
# Exemple avec plusieurs dossiers: 
<#  
$Config.SubPathStructures = @(
    "01-Dossiers ligne EL-EG\CATAPHORESE"
    "01-Dossiers ligne EL-EG\LIGNE B1"
    "01-Dossiers ligne EL-EG\LIGNE B2"
    "01-Dossiers ligne EL-EG\LIGNE B90"
    "01-Dossiers ligne EL-EG\LIGNE EG0"
    "01-Dossiers ligne EL-EG\LIGNE EL0"
    "01-Dossiers ligne EL-EG\LIGNE L40"
    "01-Dossiers ligne EL-EG\LIGNE UA0"
    "01-Dossiers ligne EL-EG\M1"
    "01-Dossiers ligne EL-EG\M2"
    )
#>
$Config.SubPathStructures = @(
    "01-Dossiers ligne EL-EG\LIGNE EG0"
    )

# Paramètres de l'interface graphique
$Config.BaseName = "Locale" #"Globale" #"locale"
$Config.FormBackColor = "Honeydew" #"WhiteSmoke"  #"Honeydew"

# Export des paramètres
$Config
