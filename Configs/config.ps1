# Configuration du projet Iveco

# Variables globales
$Global:ProjectName = "MY-Project-PowerShell"
$Global:ProjectVersion = "7.0.2"
$Global:ProjectAuthor = "Franck Ginhoux"

# Paramètres
$Config = @{
    Environment = "Development"  # Development, Testing, Production
    LogLevel = "Info"            # Debug, Info, Warning, Error
    LogFile = ".\Logs\project.log"
}

# Paramètres d'extraction et recherche des references
$Config.ExtractionRootPath = "D:\W\Iveco\serveur"
$Config.ExtractXmlData = "D:\W\Iveco\RefServeur.xml"
$Config.SubPathStructure = "01-Dossiers ligne EL-EG\LIGNE EG0"

# Paramètres de l'interface graphique
$Config.BaseName = "Locale" #"Globale" #"locale"
$Config.FormBackColor = "Honeydew" #"PaleGreen" #"WhiteSmoke"

# Export des paramètres
$Config
