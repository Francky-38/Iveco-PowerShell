# Configuration du projet Iveco

# Variables globales
$Global:ProjectName = "Iveco-PowerShell"
$Global:ProjectVersion = "1.0.0"
$Global:ProjectAuthor = "Iveco Team"

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
