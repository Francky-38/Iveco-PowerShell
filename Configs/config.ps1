# Configuration du projet Iveco

# Variables globales
$Global:ProjectName = "MY-Project-PowerShell"
$Global:ProjectVersion = "8.2.0"
$Global:ProjectAuthor = "Franck Ginhoux"

# Paramètres
$Config = @{
    Environment = "Development"  # Development, Testing, Production
    LogLevel = "Info"            # Debug, Info, Warning, Error
    LogFile = ".\Logs\project.log"
}

# Paramètres d'extraction et recherche des references
$Config.ExtractionRootPath = "D:\W\Iveco\serveur"
$Config.ExtractXmlData = "D:\W\Iveco\RefServeur"

# Sous-chemins à scruter dans chaque dossier d'affaire
# Laissez vide pour scruter directement le dossier d'affaire
# Exemple avec plusieurs dossiers: @("01-Dossiers ligne EL-EG\LIGNE EG0", "02-Autre structure")
$Config.SubPathStructures = @(
    "01-Dossiers ligne EL-EG\LIGNE EG0"
    "01-Dossiers ligne EL-EG\LIGNE EL0"
    )

# Paramètres de l'interface graphique
$Config.BaseName = "Locale" #"Globale" #"locale"
$Config.FormBackColor = "Honeydew" #"WhiteSmoke"  #"Honeydew"

# Export des paramètres
$Config
