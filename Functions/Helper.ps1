<#
.SYNOPSIS
    Exemple de fonction PowerShell réutilisable

.DESCRIPTION
    Cette fonction démontre la structure recommandée pour les fonctions du projet

.PARAMETER Message
    Le message à afficher

.EXAMPLE
    Get-WelcomeMessage -Message "Bienvenue!"
#>

<#
.SYNOPSIS
    Récupère la version du projet depuis les tags Git

.DESCRIPTION
    Utilise git describe pour obtenir le tag le plus récent du projet

.EXAMPLE
    Get-ProjectVersion
#>

function Get-ProjectVersion {
    try {
        # Aller au répertoire racine du projet
        Push-Location (Split-Path -Path $PSScriptRoot -Parent)
        
        # Récupérer le dernier tag Git
        $latestTag = git describe --tags --abbrev=0 2>&1
        
        Pop-Location
        
        if ($latestTag -and $LASTEXITCODE -eq 0) {
            # Nettoyer le 'v' du début si présent
            return $latestTag -replace '^v', ''
        }
    }
    catch {
        Write-Host "Erreur Get-ProjectVersion: $_" -ForegroundColor Red
    }
    return $Global:ProjectVersion
}

function Get-WelcomeMessage {
    param(
        [string]$Message = "Bienvenue !"
    )

    Write-Host $Message -ForegroundColor Cyan
}

<#
.SYNOPSIS
    Affiche les informations du système

.DESCRIPTION
    Récupère et affiche les informations système

.EXAMPLE
    Get-SystemInfo
#>

function Get-SystemInfo {
    Write-Host "`n--- Informations Système ---" -ForegroundColor Yellow
    Write-Host "Nom d'hôte: $(hostname)" -ForegroundColor White
    Write-Host "PowerShell: $($PSVersionTable.PSVersion)" -ForegroundColor White
    Write-Host "OS: $([System.Environment]::OSVersion.VersionString)" -ForegroundColor White
    Write-Host "Utilisateur: $($env:USERNAME)" -ForegroundColor White
    Write-Host "Répertoire: $(Get-Location)" -ForegroundColor White
}

<#
.SYNOPSIS
    Extrait les références de tous les fichiers PPTX dans une arborescence

.DESCRIPTION
    Parcourt récursivement l'arborescence et cherche tous les fichiers PPTX
    dans les chemins: [Affaires]\01-Dossiers ligne EL-EG\LIGNE EG0\[PostesTravail]\*.pptx
    Traite chaque fichier et génère un unique fichier XML de sortie

.PARAMETER RootPath
    Chemin racine de départ (ex: D:\W\Iveco\serveur)

.PARAMETER OutputFile
    Chemin complet du fichier XML de sortie (optionnel)

.EXAMPLE
    Export-PptxReferencesFromTree -RootPath "D:\W\Iveco\serveur"
#>

function Export-PptxReferencesFromTree {
    param(
        [string]$RootPath = "D:\W\Iveco\serveur",
        [string]$OutputFile = "",
        [array]$SubPathStructures = @("01-Dossiers ligne EL-EG\LIGNE EG0")
    )

    # Enregistrer l'heure de départ
    $StartTime = Get-Date

    # Vérifier que le chemin existe
    if (-not (Test-Path -Path $RootPath)) {
        Write-Host "Erreur: Le chemin '$RootPath' n'existe pas" -ForegroundColor Red
        return
    }

    # Normaliser SubPathStructures (peut être null ou vide)
    if ($null -eq $SubPathStructures -or $SubPathStructures.Count -eq 0) {
        $SubPathStructures = @("")
    }

    # Déterminer le fichier de sortie (format CLIXML uniquement)
    if ([string]::IsNullOrEmpty($OutputFile)) {
        $OutputFile = Join-Path -Path $RootPath -ChildPath "References_Extraites.clixml"
    } else {
        # Assurer l'extension .clixml
        $OutputFile = [System.IO.Path]::ChangeExtension($OutputFile, "clixml")
    }

    $ZipOutputPath = [System.IO.Path]::ChangeExtension($OutputFile, "zip")
    Write-Host "`nRecherche des fichiers PPTX dans l'arborescence..." -ForegroundColor Yellow
    Write-Host "Chemin racine: $RootPath" -ForegroundColor Cyan
    Write-Host "Sous-chemins a scruter: $($SubPathStructures -join ', ')" -ForegroundColor Cyan
    Write-Host "Sortie (ZIP): $ZipOutputPath" -ForegroundColor Cyan

    # Créer le fichier XML de sortie principal
    $XmlOutput = New-Object System.Xml.XmlDocument
    $Root = $XmlOutput.CreateElement("References")
    $XmlOutput.AppendChild($Root) | Out-Null

    $TotalReferences = 0
    $TotalFiles = 0

    try {
        Write-Host "Recherche des dossiers Affaires (SA, SB)..." -ForegroundColor Yellow
        
        # Étape 1: Chercher les dossiers Affaires (terminant par SA ou SB)
        $AffairesFolders = Get-ChildItem -Path $RootPath -Directory -ErrorAction SilentlyContinue | 
                          Where-Object { $_.Name -match '(SA|SB)$' }
        
        if ($AffairesFolders.Count -eq 0) {
            Write-Host "Attention: Aucun dossier Affaire (SA/SB) trouve" -ForegroundColor Yellow
            return
        }
        
        Write-Host "  - $($AffairesFolders.Count) dossier(s) Affaire trouve(s)" -ForegroundColor Cyan
        Write-Host ""
        
        # Étape 2: Pour chaque Affaire, chercher dans chaque sous-chemin
        foreach ($AffaireFolder in $AffairesFolders) {
            Write-Host "Exploration Affaire: $($AffaireFolder.Name)" -ForegroundColor Cyan
            
            # Itérer sur chaque sous-chemin à scruter
            foreach ($SubPath in $SubPathStructures) {
                # Construire le chemin vers les dossiers postes
                if ([string]::IsNullOrEmpty($SubPath)) {
                    # Si SubPath est vide, chercher directement dans le dossier d'affaires
                    $PostesPath = $AffaireFolder.FullName
                } else {
                    # Sinon, utiliser le chemin spécifié
                    $PostesPath = Join-Path -Path $AffaireFolder.FullName -ChildPath $SubPath
                }
                
                if (Test-Path -Path $PostesPath) {
                    # Étape 3: Chercher les dossiers de postes (1er niveau, pas récursif)
                    $PostesFolders = Get-ChildItem -Path $PostesPath -Directory -ErrorAction SilentlyContinue
                    
                    foreach ($PosteFolder in $PostesFolders) {
                        Write-Host "  Poste: $($PosteFolder.Name) [Structure: $SubPath]" -ForegroundColor Gray
                        
                        # Étape 4: Chercher les fichiers .pptx dans ce dossier (pas récursif)
                        $FilesInPoste = Get-ChildItem -Path $PosteFolder.FullName -Filter "*.pptx" -ErrorAction SilentlyContinue
                        
                        if ($FilesInPoste.Count -gt 0) {
                            Write-Host "      - $($FilesInPoste.Count) fichier(s) PPTX" -ForegroundColor Gray
                        
                        # Traiter chaque fichier PPTX directement
                        foreach ($PptxFile in $FilesInPoste) {
                            $TotalFiles++
                            Write-Host "      [$TotalFiles] Traitement: $($PptxFile.Name)" -ForegroundColor Gray

                            try {
                                # Récupérer le propriétaire du fichier une seule fois
                                $Owner = ""
                                try {
                                    $Acl = Get-Acl -Path $PptxFile.FullName
                                    $Owner = $Acl.Owner
                                    # Extraire uniquement le nom d'utilisateur (après le \)
                                    if ($Owner -match '\\(.+)$') {
                                        $Owner = $Matches[1]
                                    }
                                }
                                catch {
                                    $Owner = "Inconnu"
                                }

                                # Ouvrir l'archive PPTX sans l'extraire
                                Add-Type -AssemblyName System.IO.Compression.FileSystem
                                $ZipArchive = [System.IO.Compression.ZipFile]::OpenRead($PptxFile.FullName)
                                
                                try {
                                    # Récupérer les fichiers XML des slides directement depuis l'archive
                                    $SlidesEntries = $ZipArchive.Entries | Where-Object { $_.FullName -like "ppt/slides/slide*.xml" }
                                    
                                    # Collecter tous les slides du fichier PPTX
                                    $AllSlidesData = @()
                                    
                                    foreach ($SlideEntry in $SlidesEntries) {
                                        try {
                                            # Lire le contenu XML directement depuis le stream sans passer par le disque
                                            $Stream = $SlideEntry.Open()
                                            $XmlContent = [System.Xml.XmlDocument]::new()
                                            $XmlContent.Load($Stream)
                                            $Stream.Close()
                                            
                                            # Récupérer tous les nœuds <a:r> (run) avec gestion du namespace
                                            $NamespaceManager = New-Object System.Xml.XmlNamespaceManager($XmlContent.NameTable)
                                            $NamespaceManager.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
                                            $RunNodes = $XmlContent.SelectNodes("//a:r", $NamespaceManager)
                                            
                                            # Collecter toutes les références de ce slide avec leur type
                                            $SlideReferences = @()
                                            foreach ($RunNode in $RunNodes) {
                                                # Récupérer le texte du nœud <a:t>
                                                $TextNode = $RunNode.SelectSingleNode("a:t", $NamespaceManager)
                                                if (-not $TextNode) { continue }
                                                
                                                $Text = $TextNode.InnerText.Trim()
                                                
                                                # Chercher toutes les références dans le texte (pattern depuis config)
                                                $References = [regex]::Matches($Text, $Config.ReferenceRegexPattern)
                                                
                                                foreach ($RefMatch in $References) {
                                                    $Reference = $RefMatch.Value
                                                    
                                                    # Déterminer le type de référence (S ou H)
                                                    $RefType = "H"  # Défaut
                                                    $RunPrNode = $RunNode.SelectSingleNode("a:rPr", $NamespaceManager)
                                                    if ($RunPrNode) {
                                                        $SolidFillNode = $RunPrNode.SelectSingleNode("a:solidFill", $NamespaceManager)
                                                        if (-not $SolidFillNode) {
                                                            $RefType = "S"
                                                        } else {
                                                            $SrgbClrNode = $SolidFillNode.SelectSingleNode("a:srgbClr[@val='0000FF']", $NamespaceManager)
                                                            if ($SrgbClrNode) {
                                                                $RefType = "S"
                                                            }
                                                        }
                                                    }
                                                    
                                                    $SlideReferences += @{
                                                        Value = $Reference
                                                        Type = $RefType
                                                    }
                                                }
                                            }
                                            
                                            # Ajouter les données du slide à la collection si des références ont été trouvées
                                            if ($SlideReferences.Count -gt 0) {
                                                $AllSlidesData += @{
                                                    Name = $SlideEntry.Name
                                                    References = $SlideReferences
                                                }
                                                $TotalReferences += $SlideReferences.Count
                                            }
                                        }
                                        catch {
                                            Write-Host "      Erreur lors du traitement de $($SlideEntry.Name): $_" -ForegroundColor Red
                                        }
                                    }
                                    
                                    # Si des références ont été trouvées dans ce fichier PPTX, créer une seule entrée
                                    if ($AllSlidesData.Count -gt 0) {
                                        $AffaireNom = $AffaireFolder.Name
                                        $PosteNom = $PosteFolder.Name
                                        
                                        $Entry = $XmlOutput.CreateElement("Entree")
                                        
                                        $PathElem = $XmlOutput.CreateElement("Path")
                                        $PathElem.InnerText = $PptxFile.FullName
                                        $Entry.AppendChild($PathElem) | Out-Null
                                        
                                        $AffaireElem = $XmlOutput.CreateElement("Affaire")
                                        $AffaireElem.InnerText = $AffaireNom
                                        $Entry.AppendChild($AffaireElem) | Out-Null
                                        
                                        $PosteElem = $XmlOutput.CreateElement("Poste")
                                        $PosteElem.InnerText = $PosteNom
                                        $Entry.AppendChild($PosteElem) | Out-Null
                                        
                                        $NameElem = $XmlOutput.CreateElement("SOP")
                                        $NameElem.InnerText = $PptxFile.Name
                                        $Entry.AppendChild($NameElem) | Out-Null
                                        
                                        $DateModElem = $XmlOutput.CreateElement("DateModification")
                                        $DateModElem.InnerText = $PptxFile.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
                                        $Entry.AppendChild($DateModElem) | Out-Null
                                        
                                        $OwnerElem = $XmlOutput.CreateElement("Auteur")
                                        $OwnerElem.InnerText = $Owner
                                        $Entry.AppendChild($OwnerElem) | Out-Null
                                        
                                        # Créer la balise <Pages> contenant tous les slides de ce fichier
                                        $PagesElem = $XmlOutput.CreateElement("Pages")
                                        foreach ($SlideData in $AllSlidesData) {
                                            $PageElem = $XmlOutput.CreateElement("Page")
                                            
                                            $PageNameElem = $XmlOutput.CreateElement("Name")
                                            $PageNameElem.InnerText = $SlideData.Name
                                            $PageElem.AppendChild($PageNameElem) | Out-Null
                                            
                                            # Créer la balise <References> pour ce slide
                                            $ReferencesElem = $XmlOutput.CreateElement("References")
                                            foreach ($Ref in $SlideData.References) {
                                                $RefElem = $XmlOutput.CreateElement("Reference")
                                                $RefElem.InnerText = $Ref.Value
                                                $RefElem.SetAttribute("Type", $Ref.Type)
                                                $ReferencesElem.AppendChild($RefElem) | Out-Null
                                            }
                                            $PageElem.AppendChild($ReferencesElem) | Out-Null
                                            
                                            $PagesElem.AppendChild($PageElem) | Out-Null
                                        }
                                        $Entry.AppendChild($PagesElem) | Out-Null
                                        
                                        $Root.AppendChild($Entry) | Out-Null
                                    }
                                }
                                finally {
                                    # Fermer l'archive
                                    if ($ZipArchive) { $ZipArchive.Dispose() }
                                }
                            }
                            catch {
                                Write-Host "    Erreur lors du traitement de $($PptxFile.Name): $_" -ForegroundColor Red
                            }
                        }
                    }
                }
            } else {
                    $PathDisplay = if ([string]::IsNullOrEmpty($SubPath)) { "le dossier d'affaires" } else { "'$SubPath'" }
                    Write-Host "    Attention: Chemin $PathDisplay non trouve" -ForegroundColor Yellow
                }
            }
        }
        
        if ($TotalFiles -eq 0) {
            Write-Host "Attention: Aucun fichier PPTX trouve" -ForegroundColor Yellow
            return
        }
        
        Write-Host ""
        
        # Convertir les données XML en PowerShell Objects
        Write-Host "Conversion en format CLIXML PowerShell..." -ForegroundColor Yellow
        $AllEntries = $XmlOutput.SelectNodes("//Entree") | ForEach-Object {
            $Entry = $_
            $Pages = @()
            $PagesNode = $Entry.SelectSingleNode("Pages")
            if ($PagesNode) {
                $Pages = $PagesNode.SelectNodes("Page") | ForEach-Object {
                    $PageNode = $_
                    $References = @()
                    $ReferencesNode = $PageNode.SelectSingleNode("References")
                    if ($ReferencesNode) {
                        $References = $ReferencesNode.SelectNodes("Reference") | ForEach-Object {
                            @{
                                Value = $_.InnerText
                                Type = $_.GetAttribute("Type")
                            }
                        }
                    }
                    [PSCustomObject]@{
                        Name = $PageNode.SelectSingleNode("Name").InnerText
                        References = $References
                    }
                }
            }
            
            [PSCustomObject]@{
                Path = $Entry.SelectSingleNode("Path").InnerText
                Affaire = $Entry.SelectSingleNode("Affaire").InnerText
                Poste = $Entry.SelectSingleNode("Poste").InnerText
                SOP = $Entry.SelectSingleNode("SOP").InnerText
                DateModification = $Entry.SelectSingleNode("DateModification").InnerText
                Auteur = $Entry.SelectSingleNode("Auteur").InnerText
                Pages = $Pages
            }
        }

        # Créer les index en mémoire et sauvegarder une base pré-indexée
        # (le temps de création des index est absorbé ici, en tâche de fond)
        Write-Host "Creation des index et sauvegarde base pre-indexee..." -ForegroundColor Yellow
        $SearchIndex = Build-SearchIndexes -AllEntries $AllEntries
        $SearchIndex.IndexVersion = 1  # Marqueur pour New-SearchIndex (format pre-indexe)
        $SearchIndex | Export-Clixml -Path $OutputFile -Force
        Write-Host "  - Fichier CLIXML sauvegarde: $OutputFile" -ForegroundColor Cyan
        $FileSizeClixml = (Get-Item $OutputFile).Length / 1MB
        Write-Host "  - Taille CLIXML: $([math]::Round($FileSizeClixml, 2)) MB" -ForegroundColor Cyan

        # Compresser le .clixml en ZIP et supprimer le fichier non compressé
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $ZipPath = [System.IO.Path]::ChangeExtension($OutputFile, "zip")
        $ClixmlName = [System.IO.Path]::GetFileName($OutputFile)
        if (Test-Path $ZipPath) { Remove-Item $ZipPath -Force }
        $ZipArchive = [System.IO.Compression.ZipFile]::Open($ZipPath, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($ZipArchive, $OutputFile, $ClixmlName, [System.IO.Compression.CompressionLevel]::Optimal) | Out-Null
        }
        finally {
            $ZipArchive.Dispose()
        }
        Remove-Item $OutputFile -Force
        $OutputPath = $ZipPath
        $FileSizeZip = (Get-Item $ZipPath).Length / 1MB
        Write-Host "  - Fichier ZIP cree: $ZipPath" -ForegroundColor Cyan
        Write-Host "  - Taille ZIP: $([math]::Round($FileSizeZip, 2)) MB" -ForegroundColor Cyan
        
        Write-Host ""
        
        # Calculer le temps écoulé
        $EndTime = Get-Date
        $Duration = $EndTime - $StartTime
        
        Write-Host "OK - Traitement termine!" -ForegroundColor Green
        Write-Host "  - Fichiers PPTX traites: $TotalFiles" -ForegroundColor Cyan
        Write-Host "  - Nombre total de references extraites: $TotalReferences" -ForegroundColor Cyan
        Write-Host "  - Fichier de sortie: $OutputPath" -ForegroundColor Cyan
        Write-Host "  - Duree de traitement: $($Duration.Hours)h $($Duration.Minutes)m $($Duration.Seconds)s" -ForegroundColor Cyan
    }
    catch {
        Write-Host "Erreur lors du traitement: $_" -ForegroundColor Red
    }
    finally {
        # Aucun nettoyage nécessaire (lecture directe sans extraction)
    }
}

<#
.SYNOPSIS
    Construit les index de recherche en mémoire à partir des entrées brutes (usage interne)

.DESCRIPTION
    Crée ReferenceIndex, AffaireIndex et PagesFlat pour la recherche rapide.
#>
function Build-SearchIndexes {
    param([Parameter(Mandatory=$true)] $AllEntries)

    $ReferenceIndex = @{}
    $AffaireIndex = @{}
    $PagesFlat = @()

    foreach ($Entry in $AllEntries) {
        $Affaire = $Entry.Affaire
        if (-not $AffaireIndex.ContainsKey($Affaire)) {
            $AffaireIndex[$Affaire] = @()
        }

        foreach ($Page in $Entry.Pages) {
            # Extraire les valeurs des références pour stocker uniquement les strings
            $ReferenceValues = @($Page.References | ForEach-Object { $_.Value })
            
            $PageKey = @{
                Affaire = $Entry.Affaire
                Poste = $Entry.Poste
                SOP = $Entry.SOP
                Path = $Entry.Path
                DateModification = $Entry.DateModification
                Auteur = $Entry.Auteur
                PageName = $Page.Name
                References = $ReferenceValues  # Stocker les valeurs uniquement pour la recherche
                ReferenceDetails = $Page.References  # Garder les détails (Value+Type) pour les affichages
                PageNumber = if ($Page.Name -match 'slide(\d+)') { [int]$Matches[1] } else { 0 }
            }

            $PagesFlat += $PageKey
            $AffaireIndex[$Affaire] += $PageKey

            foreach ($RefDetail in $Page.References) {
                $RefValue = $RefDetail.Value
                if (-not $ReferenceIndex.ContainsKey($RefValue)) {
                    $ReferenceIndex[$RefValue] = @()
                }
                $ReferenceIndex[$RefValue] += $PageKey
            }
        }
    }

    return @{
        AllEntries = $AllEntries
        ReferenceIndex = $ReferenceIndex
        AffaireIndex = $AffaireIndex
        PagesFlat = $PagesFlat
    }
}

<#
.SYNOPSIS
    Charge les données pré-indexées depuis un fichier ZIP

.DESCRIPTION
    Extrait le CLIXML du ZIP et charge les index en mémoire (format généré par Export-PptxReferencesFromTree)

.PARAMETER DataPath
    Chemin du fichier ZIP (avec ou sans extension .zip)

.EXAMPLE
    $SearchIndex = New-SearchIndex -DataPath "D:\W\Iveco\RefServeur"
#>

function New-SearchIndex {
    param(
        [string]$DataPath
    )

    $StartLoadTime = Get-Date
    Write-Host "Chargement des donnees ZIP..." -ForegroundColor Yellow

    $ZipPath = if ([System.IO.Path]::HasExtension($DataPath)) { $DataPath } else { "$DataPath.zip" }
    if (-not (Test-Path $ZipPath)) {
        Write-Host "Erreur: Le fichier '$ZipPath' n'existe pas" -ForegroundColor Red
        return $null
    }

    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $ZipArchive = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
        try {
            $ClixmlEntry = $ZipArchive.Entries | Where-Object { $_.Name -like "*.clixml" } | Select-Object -First 1
            if (-not $ClixmlEntry) {
                Write-Host "Erreur: Aucun fichier .clixml dans l'archive ZIP" -ForegroundColor Red
                return $null
            }
            $TempFile = [System.IO.Path]::GetTempFileName()
            $TempClixml = [System.IO.Path]::ChangeExtension($TempFile, "clixml")
            Remove-Item $TempFile -Force -ErrorAction SilentlyContinue
            $EntryStream = $ClixmlEntry.Open()
            try {
                $FileStream = [System.IO.File]::Create($TempClixml)
                try {
                    $EntryStream.CopyTo($FileStream)
                }
                finally {
                    $FileStream.Close()
                }
            }
            finally {
                $EntryStream.Close()
            }
            try {
                $SearchIndex = Import-Clixml -Path $TempClixml
            }
            finally {
                Remove-Item $TempClixml -Force -ErrorAction SilentlyContinue
            }
        }
        finally {
            $ZipArchive.Dispose()
        }
    }
    catch {
        Write-Host "Erreur lors du chargement: $_" -ForegroundColor Red
        return $null
    }

    $LoadTime = (Get-Date) - $StartLoadTime
    Write-Host "  Donnees chargees: $($SearchIndex.PagesFlat.Count) pages, $($SearchIndex.ReferenceIndex.Count) references uniques" -ForegroundColor Cyan
    Write-Host "  Temps de chargement: $($LoadTime.TotalSeconds) secondes" -ForegroundColor Cyan

    return $SearchIndex
}

<#
.SYNOPSIS
    Retourne l'ensemble des references de type 'H' pour un marche (Affaire)

.PARAMETER SearchIndex
    Index de recherche charge par New-SearchIndex

.PARAMETER Affaire
    Nom de l'affaire/marche
#>
function Get-MarketHReferences {
    param(
        [Parameter(Mandatory=$true)] $SearchIndex,
        [Parameter(Mandatory=$true)] [string]$Affaire
    )
    $Result = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    if (-not $SearchIndex.AffaireIndex -or -not $SearchIndex.AffaireIndex.ContainsKey($Affaire)) {
        return @($Result)
    }
    foreach ($Page in $SearchIndex.AffaireIndex[$Affaire]) {
        $Details = $Page.ReferenceDetails
        if (-not $Details) { $Details = @() }
        foreach ($RefDetail in $Details) {
            $Val = if ($RefDetail.Value) { $RefDetail.Value } elseif ($RefDetail["Value"]) { $RefDetail["Value"] } else { $null }
            $Typ = if ($RefDetail.Type) { $RefDetail.Type } elseif ($RefDetail["Type"]) { $RefDetail["Type"] } else { $null }
            if ($Val -and $Typ -eq "H") {
                [void]$Result.Add($Val)
            }
        }
    }
    return @($Result)
}

<#
.SYNOPSIS
    Interface interactive de recherche de références

.DESCRIPTION
    Affiche un menu interactif pour rechercher des références dans le fichier XML

.PARAMETER DataPath
    Chemin du fichier ZIP (base sans extension ou chemin complet .zip)

.EXAMPLE
    Show-SearchGui -DataPath "D:\W\Iveco\RefServeur"
#>

function Show-SearchGui {
    param(
        [string]$DataPath
    )

    # Charger la configuration
    $ConfigPath = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath "Configs\config.ps1"
    if (Test-Path -Path $ConfigPath) {
        . $ConfigPath
    }

    # Charger les assemblies
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    # Déterminer le chemin des données (base sans extension)
    if ([string]::IsNullOrEmpty($DataPath)) {
        $DataPath = $Config.ExtractXmlDataPath
    }

    $ZipPath = if ([System.IO.Path]::HasExtension($DataPath)) { $DataPath } else { "$DataPath.zip" }
    if (-not (Test-Path $ZipPath)) {
        [System.Windows.Forms.MessageBox]::Show("Erreur: Le fichier ZIP '$ZipPath' n'existe pas", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # Charger l'index en mémoire
    $SearchIndex = New-SearchIndex -DataPath $DataPath
    $ReferenceIndex = $SearchIndex.ReferenceIndex
    $AffaireIndex = $SearchIndex.AffaireIndex
    $PagesFlat = $SearchIndex.PagesFlat

    # Créer la fenêtre principale
    $Version = Get-ProjectVersion
    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = "Recherche SOP avec r" + [char]233 + "f" + [char]233 + "rence - V$Version"
    $Form.Size = New-Object System.Drawing.Size(1220, 550)
    $Form.StartPosition = "CenterScreen"
    $Form.BackColor = [System.Drawing.Color]::($Config.FormBackColor)
    
    # Ajouter l'icône à la Form
    $IconPath = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath "nono bleu.ico"
    if (Test-Path -Path $IconPath) {
        $Form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($IconPath)
    }

    # Panel supérieur (recherche)
    $SearchPanel = New-Object System.Windows.Forms.Panel
    $SearchPanel.Location = New-Object System.Drawing.Point(10, 10)
    $SearchPanel.Size = New-Object System.Drawing.Size(1180, 60)
    $SearchPanel.BackColor = [System.Drawing.Color]::White
    $SearchPanel.BorderStyle = "Fixed3D"
    
    # Ajouter une info-bulle sur le SearchPanel avec le nom du développeur
    $ToolTip = New-Object System.Windows.Forms.ToolTip
    $ToolTip.SetToolTip($SearchPanel, "D" + [char]233 + "velopp" + [char]233 + " par : Franck Ginhoux")

    # Label pour Références
    $Label = New-Object System.Windows.Forms.Label
    $Label.Text = "R" + [char]233 + "f" + [char]233 + "rence(s):"
    $Label.Location = New-Object System.Drawing.Point(10, 15)
    $Label.Size = New-Object System.Drawing.Size(100, 20)
    $Label.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)

    # TextBox pour Références
    $TextBox = New-Object System.Windows.Forms.TextBox
    $TextBox.Location = New-Object System.Drawing.Point(110, 15)
    $TextBox.Size = New-Object System.Drawing.Size(440, 25)
    $TextBox.Font = New-Object System.Drawing.Font("Arial", 10)

    # Label pour Crit. Marché
    $LabelMarket = New-Object System.Windows.Forms.Label
    $LabelMarket.Text = "Crit. March" + [char]233 + ":"
    $LabelMarket.Location = New-Object System.Drawing.Point(570, 15)
    $LabelMarket.Size = New-Object System.Drawing.Size(100, 20)
    $LabelMarket.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)

    # Ajouter une info-bulle sur le LabelMarket pour aider à la saisie
    $ToolTip = New-Object System.Windows.Forms.ToolTip
    $ToolTip.SetToolTip($LabelMarket, "_ pour exclure le crit" + [char]232 + "re")

    # TextBox pour Crit. Marché
    $TextBoxMarket = New-Object System.Windows.Forms.TextBox
    $TextBoxMarket.Location = New-Object System.Drawing.Point(670, 15)
    $TextBoxMarket.Size = New-Object System.Drawing.Size(130, 25)
    $TextBoxMarket.Font = New-Object System.Drawing.Font("Arial", 10)

    # Bouton Rechercher
    $SearchButton = New-Object System.Windows.Forms.Button
    $SearchButton.Text = "Rechercher"
    $SearchButton.Location = New-Object System.Drawing.Point(810, 10)
    $SearchButton.Size = New-Object System.Drawing.Size(100, 30)
    $SearchButton.BackColor = [System.Drawing.Color]::LightBlue
    $SearchButton.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $SearchButton.FlatStyle = "Flat"

    # Bouton Réinitialiser
    $ClearButton = New-Object System.Windows.Forms.Button
    $ClearButton.Text = "R" + [char]233 + "initialiser"
    $ClearButton.Location = New-Object System.Drawing.Point(920, 10)
    $ClearButton.Size = New-Object System.Drawing.Size(100, 30)
    $ClearButton.BackColor = [System.Drawing.Color]::LightGray
    $ClearButton.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $ClearButton.FlatStyle = "Flat"

    # Label info base de données
    $BaseInfoLabel = New-Object System.Windows.Forms.Label
    $BaseInfoLabel.Text = "Base : $($Config.BaseName)"
    $BaseInfoLabel.Location = New-Object System.Drawing.Point(1030, 10)
    $BaseInfoLabel.Size = New-Object System.Drawing.Size(140, 30)
    $BaseInfoLabel.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Italic)
    $BaseInfoLabel.TextAlign = "MiddleRight"
    $BaseInfoLabel.ForeColor = [System.Drawing.Color]::DarkGray

    $SearchPanel.Controls.Add($Label)
    $SearchPanel.Controls.Add($TextBox)
    $SearchPanel.Controls.Add($LabelMarket)
    $SearchPanel.Controls.Add($TextBoxMarket)
    $SearchPanel.Controls.Add($SearchButton)
    $SearchPanel.Controls.Add($ClearButton)
    $SearchPanel.Controls.Add($BaseInfoLabel)

    # DataGridView pour les résultats
    $DataGridView = New-Object System.Windows.Forms.DataGridView
    $DataGridView.Location = New-Object System.Drawing.Point(10, 80)
    $DataGridView.Size = New-Object System.Drawing.Size(1180, 410)
    $DataGridView.AllowUserToAddRows = $false
    $DataGridView.AllowUserToDeleteRows = $false
    $DataGridView.ReadOnly = $true
    $DataGridView.SelectionMode = "FullRowSelect"
    $DataGridView.BackgroundColor = [System.Drawing.Color]::White
    $DataGridView.GridColor = [System.Drawing.Color]::LightGray
    $DataGridView.Font = New-Object System.Drawing.Font("Arial", 9)

    # Ajouter les colonnes
    $DataGridView.ColumnCount = 8
    $DataGridView.Columns[0].Name = "March" + [char]233
    $DataGridView.Columns[0].Width = 300
    $DataGridView.Columns[1].Name = "Poste"
    $DataGridView.Columns[1].Width = 255
    $DataGridView.Columns[2].Name = "SOP"
    $DataGridView.Columns[2].Width = 240
    $DataGridView.Columns[3].Name = "Page"
    $DataGridView.Columns[3].Width = 50
    $DataGridView.Columns[5].Name = "Date"
    $DataGridView.Columns[5].Width = 130
    $DataGridView.Columns[4].Name = "Auteur"
    $DataGridView.Columns[4].Width = 80
    $DataGridView.Columns[6].Name = "Nb Ref"
    $DataGridView.Columns[6].Width = 80
    $DataGridView.Columns[7].Name = "PathPptx"
    $DataGridView.Columns[7].Visible = $false  # Masquer cette colonne

    # En-tête
    $DataGridView.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::DarkBlue
    $DataGridView.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
    $DataGridView.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)

    # Événement du bouton Rechercher
    $SearchButton.Add_Click({
        $SearchInput = $TextBox.Text.Trim()
        
        if ([string]::IsNullOrEmpty($SearchInput)) {
            [System.Windows.Forms.MessageBox]::Show("Veuillez entrer une ou plusieurs references a rechercher (separees par ;)", "Attention", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        # Diviser les références par ';' et nettoyer chaque terme
        $SearchTerms = @($SearchInput -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        
        if ($SearchTerms.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Veuillez entrer une ou plusieurs references valides", "Attention", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        # Récupérer le critère marché
        $MarketCriteria = $TextBoxMarket.Text.Trim()
        $IsNegativeMarketFilter = $false
        $MarketFilterValue = ""
        
        if (-not [string]::IsNullOrEmpty($MarketCriteria)) {
            # Vérifier si c'est un filtre négatif (commence par _)
            if ($MarketCriteria.StartsWith("_")) {
                $IsNegativeMarketFilter = $true
                $MarketFilterValue = $MarketCriteria.Substring(1)
            } else {
                $IsNegativeMarketFilter = $false
                $MarketFilterValue = $MarketCriteria
            }
        }

        $Form.Text = "Recherche En cours..."
        $DataGridView.Rows.Clear()
        $FoundCount = 0
        
        $SearchStartTime = Get-Date

        # Optimisation avec index: si une seule référence, utiliser l'index direct
        if ($SearchTerms.Count -eq 1) {
            # Recherche simple rapide par index
            $Term = $SearchTerms[0]
            $MatchedPages = if ($ReferenceIndex -ne $null) { $ReferenceIndex[$Term] } else { $null }
            
            if ($MatchedPages) {
                foreach ($Page in $MatchedPages) {
                    # Appliquer le filtre marché
                    $AffaireMatchesMarketFilter = $true
                    if (-not [string]::IsNullOrEmpty($MarketFilterValue)) {
                        if ($IsNegativeMarketFilter) {
                            $AffaireMatchesMarketFilter = -not ($Page.Affaire -like "*$MarketFilterValue*")
                        } else {
                            $AffaireMatchesMarketFilter = $Page.Affaire -like "*$MarketFilterValue*"
                        }
                    }
                    
                    if ($AffaireMatchesMarketFilter) {
                        $NbRefInPage = $Page.References.Count
                        $Percentage = [math]::Round((1 / $NbRefInPage) * 100)
                        $RefCount = "$Percentage% (1/$NbRefInPage)"
                        
                        $DataGridView.Rows.Add($Page.Affaire, $Page.Poste, $Page.SOP, $Page.PageNumber, $Page.Auteur, $Page.DateModification, $RefCount, $Page.Path)
                        $FoundCount++
                    }
                }
            }
        } else {
            # Recherche multi-référence : chercher les pages contenant TOUTES les références
            foreach ($Page in $PagesFlat) {
                # Appliquer le filtre marché
                $AffaireMatchesMarketFilter = $true
                if (-not [string]::IsNullOrEmpty($MarketFilterValue)) {
                    if ($IsNegativeMarketFilter) {
                        $AffaireMatchesMarketFilter = -not ($Page.Affaire -like "*$MarketFilterValue*")
                    } else {
                        $AffaireMatchesMarketFilter = $Page.Affaire -like "*$MarketFilterValue*"
                    }
                }
                
                if (-not $AffaireMatchesMarketFilter) {
                    continue
                }
                
                # Vérifier que TOUS les termes sont présents
                $AllTermsFound = $true
                $ReferencesText = $Page.References -join " "
                foreach ($Term in $SearchTerms) {
                    if ($ReferencesText -notlike "*$Term*") {
                        $AllTermsFound = $false
                        break
                    }
                }
                
                if ($AllTermsFound) {
                    $NbRefInPage = $Page.References.Count
                    $NbRefSearched = $SearchTerms.Count
                    $Percentage = [math]::Round(($NbRefSearched / $NbRefInPage) * 100)
                    $RefCount = "$Percentage% ($NbRefSearched/$NbRefInPage)"
                    
                    $DataGridView.Rows.Add($Page.Affaire, $Page.Poste, $Page.SOP, $Page.PageNumber, $Page.Auteur, $Page.DateModification, $RefCount, $Page.Path)
                    $FoundCount++
                }
            }
        }

        # Trier par DateModification (décroissant - plus récentes en haut)
        if ($DataGridView.Rows.Count -gt 0) {
            $DataGridView.Sort($DataGridView.Columns["Date"], [System.ComponentModel.ListSortDirection]::Descending)
        }
        
        $SearchDuration = (Get-Date) - $SearchStartTime
        $Form.Text = "Recherche SOP avec r" + [char]233 + "f" + [char]233 + "rence [$FoundCount resultat(s) en $($SearchDuration.TotalMilliseconds)ms] V$Version"
    
    })

    # Événement du bouton Réinitialiser
    $ClearButton.Add_Click({
        $TextBox.Text = ""
        $TextBoxMarket.Text = ""
        $DataGridView.Rows.Clear()
        $Form.Text = "Recherche SOP avec r" + [char]233 + "f" + [char]233 + "rence - V$Version"
    })

    # Événement Enter dans la textbox
    $TextBox.Add_KeyDown({
        if ($_.KeyCode -eq "Return") {
            $SearchButton.PerformClick()
        }
    })

    # Événement double-clic sur une cellule de la colonne SOP pour ouvrir le fichier
    $DataGridView.Add_CellDoubleClick({
        $RowIndex = $_.RowIndex
        
        # Colonne 0 = Marché (Affaire)
        if ($_.ColumnIndex -eq 0) {
            $Affaire = $DataGridView.Rows[$RowIndex].Cells[0].Value
            
            # Construire le chemin du dossier Affaire
            $AffairePath = Join-Path -Path $Config.ExtractionRootPath -ChildPath $Affaire
            
            if (Test-Path $AffairePath) {
                try {
                    Invoke-Item $AffairePath
                }
                catch {
                    [System.Windows.Forms.MessageBox]::Show("Erreur lors de l'ouverture du dossier Marche: $_", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                }
            } else {
                [System.Windows.Forms.MessageBox]::Show("Dossier Marche non trouve: $AffairePath", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
        
        # Colonne 1 = Poste
        if ($_.ColumnIndex -eq 1) {
            $Affaire = $DataGridView.Rows[$RowIndex].Cells[0].Value
            $Poste = $DataGridView.Rows[$RowIndex].Cells[1].Value
            
            # Construire le chemin du dossier Poste
            # Essayer d'abord avec le premier sous-chemin de la config
            $PostePath = $null
            foreach ($SubPath in $Config.SubPathStructures) {
                $PossiblePath = Join-Path -Path $Config.ExtractionRootPath -ChildPath $Affaire | Join-Path -ChildPath $SubPath | Join-Path -ChildPath $Poste
                if (Test-Path $PossiblePath) {
                    $PostePath = $PossiblePath
                    break
                }
            }
            
            # Si pas trouvé avec SubPathStructures, essayer directement dans l'Affaire
            if (-not $PostePath) {
                $PostePath = Join-Path -Path $Config.ExtractionRootPath -ChildPath $Affaire | Join-Path -ChildPath $Poste
            }
            
            if (Test-Path $PostePath) {
                try {
                    Invoke-Item $PostePath
                }
                catch {
                    [System.Windows.Forms.MessageBox]::Show("Erreur lors de l'ouverture du dossier Poste: $_", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                }
            } else {
                [System.Windows.Forms.MessageBox]::Show("Dossier Poste non trouve: $PostePath", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
        
        # Colonne 2 = SOP
        if ($_.ColumnIndex -eq 2) {
            $FilePath = $DataGridView.Rows[$RowIndex].Cells[7].Value
            
            if (Test-Path $FilePath) {
                try {
                    Invoke-Item $FilePath
                }
                catch {
                    [System.Windows.Forms.MessageBox]::Show("Erreur lors de l'ouverture: $_", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                }
            } else {
                [System.Windows.Forms.MessageBox]::Show("Fichier non trouve: $FilePath", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
        
        # Colonne 3 = Page
        if ($_.ColumnIndex -eq 3) {
            $FilePath = $DataGridView.Rows[$RowIndex].Cells[7].Value
            $index = $DataGridView.Rows[$RowIndex].Cells[3].Value
            
            if (Test-Path $FilePath) {
                try {
                    $pp = New-Object -ComObject PowerPoint.Application
                    $pp.Visible = -1
                    $presentation = $pp.Presentations.Open($FilePath)
                    $pp.ActiveWindow.View.GotoSlide($index)
                }
                catch {
                    [System.Windows.Forms.MessageBox]::Show("Erreur lors de l'ouverture: $_", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                }
                finally {
                    # Libérer les objets COM
                    if ($presentation) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null }
                    if ($pp) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp) | Out-Null }
                    [GC]::Collect()
                    [GC]::WaitForPendingFinalizers()
                }
            } else {
                [System.Windows.Forms.MessageBox]::Show("Fichier non trouve: $FilePath", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
        
    })

    <# Événement MouseMove pour afficher le chemin du fichier en survolant la colonne SOP
    $DataGridView.Add_MouseMove({
        $HitTest = $DataGridView.HitTest($_.X, $_.Y)
        
        if ($HitTest.ColumnIndex -eq 2 -and $HitTest.RowIndex -ge 0) {
            $RowIndex = $HitTest.RowIndex
            $PageNum = $DataGridView.Rows[$HitTest.RowIndex].Cells[3].Value
            $FilePath = $DataGridView.Rows[$HitTest.RowIndex].Cells[7].Value
            $Form.Text = "Chemin: " + $FilePath + " | Page: " + $PageNum
        } else {
            $SearchCount = $DataGridView.Rows.Count
            $Form.Text = "Recherche de References - Projet Iveco [$SearchCount resultat(s)]"
        }
    })
    #>
    
    # Ajouter les contrôles à la forme
    $Form.Controls.Add($SearchPanel)
    $Form.Controls.Add($DataGridView)

    # Afficher la fenêtre
    $Form.ShowDialog() | Out-Null
}


function Export-PptxTextContent {
    param(
        [string]$PptxPath,
        [string]$OutputPath = ""
    )

    # Vérifier que le fichier existe
    if (-not (Test-Path -Path $PptxPath)) {
        Write-Host "Erreur: Le fichier '$PptxPath' n'existe pas" -ForegroundColor Red
        return
    }

    # Déterminer le dossier de sortie
    if ([string]::IsNullOrEmpty($OutputPath)) {
        $OutputPath = Split-Path -Path $PptxPath
    }

    $ArchiveName = [System.IO.Path]::GetFileNameWithoutExtension($PptxPath)
    $OutputFile = Join-Path -Path $OutputPath -ChildPath "$($ArchiveName)_contenu.xml"
    
    Write-Host "`nTraitement du fichier PPTX..." -ForegroundColor Yellow
    Write-Host "Source: $PptxPath" -ForegroundColor Cyan
    Write-Host "Sortie: $OutputFile" -ForegroundColor Cyan

    # Créer un dossier temporaire
    $TempDir = Join-Path -Path $env:TEMP -ChildPath "PPTX_Extract_$(Get-Random)"
    New-Item -ItemType Directory -Path $TempDir -Force | Out-Null
    
    try {
        # Extraire l'archive PPTX (c'est un ZIP)
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::ExtractToDirectory($PptxPath, $TempDir)
        
        # Créer le fichier XML de sortie
        $XmlOutput = New-Object System.Xml.XmlDocument
        $Root = $XmlOutput.CreateElement("Contenu")
        $XmlOutput.AppendChild($Root) | Out-Null
        
        # Parcourir les fichiers XML dans ppt\slides
        $SlidesPath = Join-Path -Path $TempDir -ChildPath "ppt\slides"
        
        if (Test-Path -Path $SlidesPath) {
            $XmlFiles = Get-ChildItem -Path $SlidesPath -Filter "*.xml" -ErrorAction SilentlyContinue
            
            $TextCount = 0
            foreach ($XmlFile in $XmlFiles) {
                Write-Host "  - Traitement de $($XmlFile.Name)..." -ForegroundColor Gray
                
                try {
                    $XmlContent = [System.Xml.XmlDocument]::new()
                    $XmlContent.Load($XmlFile.FullName)
                    
                    # Récupérer tous les nœuds <a:t> avec gestion du namespace
                    $NamespaceManager = New-Object System.Xml.XmlNamespaceManager($XmlContent.NameTable)
                    $NamespaceManager.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
                    $TextNodes = $XmlContent.SelectNodes("//a:t", $NamespaceManager)
                    
                    foreach ($TextNode in $TextNodes) {
                        $Text = $TextNode.InnerText.Trim()
                        
                        # Vérifier si le texte est une référence valide
                        # Format: [TRS]?\d{5,10} (optionnellement T, R ou S suivi de 5 à 10 chiffres)
                        if ($Text -match '^[TRS]?\d{5,10}$') {
                            $Entry = $XmlOutput.CreateElement("Entree")
                            
                            $PathElem = $XmlOutput.CreateElement("CheminArchive")
                            $PathElem.InnerText = [System.IO.Path]::GetDirectoryName($PptxPath)
                            $Entry.AppendChild($PathElem) | Out-Null
                            
                            $NameElem = $XmlOutput.CreateElement("NomArchive")
                            $NameElem.InnerText = [System.IO.Path]::GetFileName($PptxPath)
                            $Entry.AppendChild($NameElem) | Out-Null
                            
                            $FileElem = $XmlOutput.CreateElement("NomFichierXml")
                            $FileElem.InnerText = $XmlFile.Name
                            $Entry.AppendChild($FileElem) | Out-Null
                            
                            $TextElem = $XmlOutput.CreateElement("Reference")
                            $TextElem.InnerText = $Text
                            $Entry.AppendChild($TextElem) | Out-Null
                            
                            $Root.AppendChild($Entry) | Out-Null
                            $TextCount++
                        }
                    }
                }
                catch {
                    Write-Host "    Erreur lors du traitement de $($XmlFile.Name): $_" -ForegroundColor Red
                }
            }
            
            # Sauvegarder le fichier XML
            $XmlOutput.Save($OutputFile)
            Write-Host "`nOK - Traitement termine!" -ForegroundColor Green
            Write-Host "  - Nombre de textes extraits: $TextCount" -ForegroundColor Cyan
            Write-Host "  - Fichier de sortie: $OutputFile" -ForegroundColor Cyan
        }
        else {
            Write-Host "Erreur: Le dossier 'ppt\slides' n'a pas ete trouve dans l'archive" -ForegroundColor Red
        }
    }
    catch {
        Write-Host "Erreur lors du traitement: $_" -ForegroundColor Red
    }
    finally {
        # Nettoyer le dossier temporaire
        Remove-Item -Path $TempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

<#
.SYNOPSIS
    Ouvre un fichier PowerPoint à la page spécifiée via VBScript

.DESCRIPTION
    Crée un VBScript temporaire pour ouvrir PowerPoint à la slide spécifique
    Cette approche est plus fiable que la COM PowerPoint directe

.PARAMETER FilePath
    Chemin complet du fichier PPTX

.PARAMETER SlideNumber
    Numéro de la slide à afficher

.EXAMPLE
    Open-PowerPointAtSlide -FilePath "D:\file.pptx" -SlideNumber 5
#>

function Open-PowerPointAtSlide {
    param(
        [string]$FilePath,
        [int]$SlideNumber
    )
    
    # Vérifier que le fichier existe
    if (-not (Test-Path $FilePath)) {
        [System.Windows.Forms.MessageBox]::Show("Fichier non trouve: $FilePath", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # Créer un script VBScript temporaire
    $VBScriptPath = Join-Path $env:TEMP "PowerPoint_Goto_$([System.Guid]::NewGuid()).vbs"
    
    $VBScript = @"
Dim objPPT, objPresentation, slideNum

Set objPPT = CreateObject("PowerPoint.Application")
objPPT.Visible = True

' Ouvrir la présentation
Set objPresentation = objPPT.Presentations.Open("$FilePath", , , msoTrue)

' Attendre le chargement complet
WScript.Sleep(2000)

' Aller à la slide
slideNum = CInt($SlideNumber)

' DEBUG: Afficher la valeur pour vérifier
' MsgBox "Slide Number: " & slideNum & " Total Slides: " & objPresentation.Slides.Count

If slideNum <= objPresentation.Slides.Count And slideNum > 0 Then
    On Error Resume Next
    
    ' Configurer le diaporama pour démarrer à cette slide
    With objPresentation.SlideShowSettings
        .StartingSlide = CInt(slideNum)
        .EndingSlide = CInt(slideNum)
        .ShowType = 3  ' Normal view (pas plein écran)
        .Run  ' Lancer le diaporama
    End With
    
    On Error GoTo 0
Else
    MsgBox "Slide " & slideNum & " invalide. Total: " & objPresentation.Slides.Count
End If

' Garder PowerPoint ouvert
Set objPresentation = Nothing
Set objPPT = Nothing
"@

    # Écrire le VBScript
    Set-Content -Path $VBScriptPath -Value $VBScript -Encoding ASCII
    
    try {
        # Exécuter le VBScript en mode non-interactif
        & cmd /c "cscript.exe //nologo `"$VBScriptPath`""
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Erreur lors de l'ouverture: $_", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
    finally {
        # Nettoyer le fichier temporaire
        Start-Sleep -Milliseconds 1000
        Remove-Item -Path $VBScriptPath -Force -ErrorAction SilentlyContinue
    }
}

function New-CompiledPresentation {
    <#
    .SYNOPSIS
    Cree une presentation PowerPoint compilee en copiant des slides depuis plusieurs sources vers un modele.

    .DESCRIPTION
    Cette fonction copie le modele vierge vers le chemin de sortie, puis y ajoute des slides
    selectionnees a partir de fichiers PPTX source selon une table descriptive.

    .PARAMETER ModelePath
    Chemin du modele PPTX vierge (contient uniquement les masques)

    .PARAMETER OutputFileName
    Nom du fichier de sortie PPTX

    .PARAMETER OutputPath
    Chemin du dossier de sortie

    .PARAMETER Pages
    Tableau contenant les pages a copier. Chaque element doit avoir :
    - Source : Chemin du fichier PPTX source
    - PageSource : Numero de la page a copier (1-indexed)

    .EXAMPLE
    $Pages = @(
        @{ Source = "C:\Source1.pptx"; PageSource = 2 },
        @{ Source = "C:\Source2.pptx"; PageSource = 3 }
    )
    New-CompiledPresentation -ModelePath "C:\Model.pptx" `
                             -OutputFileName "Compiled.pptx" `
                             -OutputPath "C:\Output" `
                             -Pages $Pages
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$ModelePath,

        [Parameter(Mandatory = $true)]
        [string]$OutputFileName,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath,

        [Parameter(Mandatory = $true)]
        [array]$Pages
    )

    # Validation des parametres
    if (-not (Test-Path $ModelePath)) {
        Write-Error "Le modele n'existe pas."
        return
    }

    if (-not (Test-Path $OutputPath)) {
        Write-Error "Le dossier de sortie n'existe pas."
        return
    }

    # Chemin complet du fichier de sortie
    $OutputFilePath = Join-Path $OutputPath $OutputFileName

    try {
        # Copier le modele vers le chemin de sortie
        Copy-Item -Path $ModelePath -Destination $OutputFilePath -Force
        Write-Host "Modele copie vers : $OutputFilePath" -ForegroundColor Green

        # Charger PowerPoint COM
        $PPTApp = New-Object -ComObject PowerPoint.Application
        $PPTApp.Visible = 1  # PowerPoint ne permet pas de masquer la fenetre

        # Ouvrir la presentation de compilation
        $FullOutputPath = [System.IO.Path]::GetFullPath($OutputFilePath)
        $CompilationPres = $PPTApp.Presentations.Open($FullOutputPath)

        # Supprimer la premiere page vierge du modele
        if ($CompilationPres.Slides.Count -gt 0) {
            $CompilationPres.Slides(1).Delete()
        }

        Write-Host "Compilation en cours..." -ForegroundColor Cyan

        foreach ($Page in $Pages) {
            $SourceFile = $Page.Source
            $SourcePageNum = $Page.PageSource

            # Verifier l'existence du fichier source
            if (-not (Test-Path $SourceFile)) {
                Write-Warning "Fichier source non trouve. Ignore."
                continue
            }

            try {
                # Ouvrir le fichier source
                $FullSourcePath = [System.IO.Path]::GetFullPath($SourceFile)
                $SourcePres = $PPTApp.Presentations.Open($FullSourcePath)

                # Verifier que le numero de page existe
                if ($SourcePageNum -gt $SourcePres.Slides.Count -or $SourcePageNum -lt 1) {
                    Write-Warning "La page n'existe pas. Ignore."
                    $SourcePres.Close()
                    continue
                }

                # Copier la slide
                $SourceSlide = $SourcePres.Slides($SourcePageNum)
                
                # Copier vers le presse-papiers et coller dans la compilation
                $SourceSlide.Copy()
                $CompilationPres.Slides.Paste() | Out-Null
                
                $FileName = [System.IO.Path]::GetFileName($SourceFile)
                Write-Host "Page $SourcePageNum de $FileName copiee" -ForegroundColor Green

                # Fermer le fichier source
                $SourcePres.Close()
            }
            catch {
                $ErrorMsg = $_
                Write-Error "Erreur lors du traitement : $ErrorMsg"
            }
        }

        # Sauvegarder la presentation de compilation
        $CompilationPres.Save()
        Write-Host "Presentation compilee sauvegardee" -ForegroundColor Green

        # Fermer et nettoyer
        $CompilationPres.Close()
        $PPTApp.Quit()
    }
    catch {
        $ErrorMsg = $_
        Write-Error "Erreur lors de la compilation : $ErrorMsg"
    }
    finally {
        # Liberer les ressources COM
        if ($null -ne $CompilationPres) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($CompilationPres) | Out-Null
        }
        if ($null -ne $PPTApp) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($PPTApp) | Out-Null
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

<#
.SYNOPSIS
    Interface graphique de recherche de texte a la volee dans les PPTX

.DESCRIPTION
    Affiche une interface WinForms qui recherche un texte saisi par l'utilisateur
    directement dans le contenu des fichiers PPTX de l'arborescence (memes fichiers
    et meme configuration que ExtracRefServeur.ps1). Aucune base n'est generee :
    les fichiers sont parcourus a la volee et les resultats sont ajoutes a la liste
    au fur et a mesure de leur decouverte.

.EXAMPLE
    Show-LiveTextSearchGui
#>
function Show-LiveTextSearchGui {
    param(
        [string]$RootPath = "",
        [array]$SubPathStructures = @()
    )

    # Charger la configuration
    $ConfigPath = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath "Configs\config.ps1"
    if (Test-Path -Path $ConfigPath) {
        . $ConfigPath
    }

    # Charger les assemblies
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    # Utiliser les valeurs de config si les parametres ne sont pas fournis
    if ([string]::IsNullOrEmpty($RootPath)) {
        $RootPath = $Config.ExtractionRootPath
    }
    if ($null -eq $SubPathStructures -or $SubPathStructures.Count -eq 0) {
        $SubPathStructures = $Config.SubPathStructures
    }
    if ($null -eq $SubPathStructures -or $SubPathStructures.Count -eq 0) {
        $SubPathStructures = @("")
    }

    if (-not (Test-Path -Path $RootPath)) {
        [System.Windows.Forms.MessageBox]::Show("Erreur: Le chemin racine '$RootPath' n'existe pas", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # Etat partage entre les gestionnaires d'evenements
    $State = @{ Cancel = $false; Running = $false }

    # Creer la fenetre principale
    $Version = Get-ProjectVersion
    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = "Recherche de texte a la vol" + [char]233 + "e - V$Version"
    $Form.Size = New-Object System.Drawing.Size(1220, 600)
    $Form.StartPosition = "CenterScreen"
    $Form.BackColor = [System.Drawing.Color]::($Config.FormBackColor)

    # Ajouter l'icone a la Form
    $IconPath = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath "nono bleu.ico"
    if (Test-Path -Path $IconPath) {
        $Form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($IconPath)
    }

    # Panel superieur (recherche)
    $SearchPanel = New-Object System.Windows.Forms.Panel
    $SearchPanel.Location = New-Object System.Drawing.Point(10, 10)
    $SearchPanel.Size = New-Object System.Drawing.Size(1180, 60)
    $SearchPanel.BackColor = [System.Drawing.Color]::White
    $SearchPanel.BorderStyle = "Fixed3D"

    $ToolTip = New-Object System.Windows.Forms.ToolTip
    $ToolTip.SetToolTip($SearchPanel, "D" + [char]233 + "velopp" + [char]233 + " par : Franck Ginhoux")

    # Label pour le texte recherche
    $Label = New-Object System.Windows.Forms.Label
    $Label.Text = "Texte :"
    $Label.Location = New-Object System.Drawing.Point(10, 15)
    $Label.Size = New-Object System.Drawing.Size(60, 20)
    $Label.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)

    # TextBox pour le texte recherche
    $TextBox = New-Object System.Windows.Forms.TextBox
    $TextBox.Location = New-Object System.Drawing.Point(75, 15)
    $TextBox.Size = New-Object System.Drawing.Size(420, 25)
    $TextBox.Font = New-Object System.Drawing.Font("Arial", 10)

    # Label pour Crit. Marche
    $LabelMarket = New-Object System.Windows.Forms.Label
    $LabelMarket.Text = "Crit. March" + [char]233 + ":"
    $LabelMarket.Location = New-Object System.Drawing.Point(510, 15)
    $LabelMarket.Size = New-Object System.Drawing.Size(95, 20)
    $LabelMarket.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)

    $ToolTipMarket = New-Object System.Windows.Forms.ToolTip
    $ToolTipMarket.SetToolTip($LabelMarket, "_ pour exclure le crit" + [char]232 + "re")

    # TextBox pour Crit. Marche
    $TextBoxMarket = New-Object System.Windows.Forms.TextBox
    $TextBoxMarket.Location = New-Object System.Drawing.Point(605, 15)
    $TextBoxMarket.Size = New-Object System.Drawing.Size(120, 25)
    $TextBoxMarket.Font = New-Object System.Drawing.Font("Arial", 10)

    # Bouton Rechercher
    $SearchButton = New-Object System.Windows.Forms.Button
    $SearchButton.Text = "Rechercher"
    $SearchButton.Location = New-Object System.Drawing.Point(740, 10)
    $SearchButton.Size = New-Object System.Drawing.Size(100, 30)
    $SearchButton.BackColor = [System.Drawing.Color]::LightBlue
    $SearchButton.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $SearchButton.FlatStyle = "Flat"

    # Bouton Arreter
    $StopButton = New-Object System.Windows.Forms.Button
    $StopButton.Text = "Arr" + [char]234 + "ter"
    $StopButton.Location = New-Object System.Drawing.Point(850, 10)
    $StopButton.Size = New-Object System.Drawing.Size(90, 30)
    $StopButton.BackColor = [System.Drawing.Color]::LightSalmon
    $StopButton.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $StopButton.FlatStyle = "Flat"
    $StopButton.Enabled = $false

    # Bouton Reinitialiser
    $ClearButton = New-Object System.Windows.Forms.Button
    $ClearButton.Text = "R" + [char]233 + "initialiser"
    $ClearButton.Location = New-Object System.Drawing.Point(950, 10)
    $ClearButton.Size = New-Object System.Drawing.Size(100, 30)
    $ClearButton.BackColor = [System.Drawing.Color]::LightGray
    $ClearButton.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $ClearButton.FlatStyle = "Flat"

    # Label info base
    $BaseInfoLabel = New-Object System.Windows.Forms.Label
    $BaseInfoLabel.Text = "Serveur (live)"
    $BaseInfoLabel.Location = New-Object System.Drawing.Point(1060, 10)
    $BaseInfoLabel.Size = New-Object System.Drawing.Size(110, 30)
    $BaseInfoLabel.Font = New-Object System.Drawing.Font("Arial", 11, [System.Drawing.FontStyle]::Italic)
    $BaseInfoLabel.TextAlign = "MiddleRight"
    $BaseInfoLabel.ForeColor = [System.Drawing.Color]::DarkGray

    $SearchPanel.Controls.Add($Label)
    $SearchPanel.Controls.Add($TextBox)
    $SearchPanel.Controls.Add($LabelMarket)
    $SearchPanel.Controls.Add($TextBoxMarket)
    $SearchPanel.Controls.Add($SearchButton)
    $SearchPanel.Controls.Add($StopButton)
    $SearchPanel.Controls.Add($ClearButton)
    $SearchPanel.Controls.Add($BaseInfoLabel)

    # DataGridView pour les resultats
    $DataGridView = New-Object System.Windows.Forms.DataGridView
    $DataGridView.Location = New-Object System.Drawing.Point(10, 80)
    $DataGridView.Size = New-Object System.Drawing.Size(1180, 430)
    $DataGridView.AllowUserToAddRows = $false
    $DataGridView.AllowUserToDeleteRows = $false
    $DataGridView.ReadOnly = $true
    $DataGridView.SelectionMode = "FullRowSelect"
    $DataGridView.BackgroundColor = [System.Drawing.Color]::White
    $DataGridView.GridColor = [System.Drawing.Color]::LightGray
    $DataGridView.Font = New-Object System.Drawing.Font("Arial", 9)

    $DataGridView.ColumnCount = 7
    $DataGridView.Columns[0].Name = "March" + [char]233
    $DataGridView.Columns[0].Width = 250
    $DataGridView.Columns[1].Name = "Poste"
    $DataGridView.Columns[1].Width = 200
    $DataGridView.Columns[2].Name = "SOP"
    $DataGridView.Columns[2].Width = 220
    $DataGridView.Columns[3].Name = "Page"
    $DataGridView.Columns[3].Width = 50
    $DataGridView.Columns[4].Name = "Date"
    $DataGridView.Columns[4].Width = 130
    $DataGridView.Columns[5].Name = "Extrait"
    $DataGridView.Columns[5].Width = 320
    $DataGridView.Columns[6].Name = "PathPptx"
    $DataGridView.Columns[6].Visible = $false

    $DataGridView.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::DarkBlue
    $DataGridView.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
    $DataGridView.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)

    # Label de statut (progression)
    $StatusLabel = New-Object System.Windows.Forms.Label
    $StatusLabel.Location = New-Object System.Drawing.Point(10, 520)
    $StatusLabel.Size = New-Object System.Drawing.Size(1180, 30)
    $StatusLabel.Font = New-Object System.Drawing.Font("Arial", 9)
    $StatusLabel.TextAlign = "MiddleLeft"
    $StatusLabel.Text = "Pr" + [char]234 + "t."

    # Fonction interne : extraire le texte de chaque slide d'un PPTX
    $GetSlidesText = {
        param([string]$PptxPath)

        $Result = @()
        $ZipArchive = $null
        try {
            $ZipArchive = [System.IO.Compression.ZipFile]::OpenRead($PptxPath)
            $SlidesEntries = $ZipArchive.Entries | Where-Object { $_.FullName -like "ppt/slides/slide*.xml" }
            foreach ($SlideEntry in $SlidesEntries) {
                try {
                    $Stream = $SlideEntry.Open()
                    $XmlContent = [System.Xml.XmlDocument]::new()
                    $XmlContent.Load($Stream)
                    $Stream.Close()

                    $NamespaceManager = New-Object System.Xml.XmlNamespaceManager($XmlContent.NameTable)
                    $NamespaceManager.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
                    $TextNodes = $XmlContent.SelectNodes("//a:t", $NamespaceManager)

                    $Builder = New-Object System.Text.StringBuilder
                    foreach ($TextNode in $TextNodes) {
                        [void]$Builder.Append($TextNode.InnerText)
                        [void]$Builder.Append(" ")
                    }

                    $PageNumber = if ($SlideEntry.Name -match 'slide(\d+)') { [int]$Matches[1] } else { 0 }
                    $Result += [PSCustomObject]@{
                        PageNumber = $PageNumber
                        Text = $Builder.ToString()
                    }
                }
                catch { }
            }
        }
        catch { }
        finally {
            if ($ZipArchive) { $ZipArchive.Dispose() }
        }
        return $Result
    }

    # Evenement du bouton Rechercher
    $SearchButton.Add_Click({
        if ($State.Running) { return }

        $SearchInput = $TextBox.Text.Trim()
        if ([string]::IsNullOrEmpty($SearchInput)) {
            [System.Windows.Forms.MessageBox]::Show("Veuillez entrer un texte a rechercher", "Attention", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        # Critere marche (filtre sur le nom de l'affaire)
        $MarketCriteria = $TextBoxMarket.Text.Trim()
        $IsNegativeMarketFilter = $false
        $MarketFilterValue = ""
        if (-not [string]::IsNullOrEmpty($MarketCriteria)) {
            if ($MarketCriteria.StartsWith("_")) {
                $IsNegativeMarketFilter = $true
                $MarketFilterValue = $MarketCriteria.Substring(1)
            } else {
                $MarketFilterValue = $MarketCriteria
            }
        }

        # Preparer l'UI pour la recherche
        $State.Cancel = $false
        $State.Running = $true
        $SearchButton.Enabled = $false
        $StopButton.Enabled = $true
        $ClearButton.Enabled = $false
        $DataGridView.Rows.Clear()
        $FoundCount = 0
        $FilesScanned = 0
        $SearchStartTime = Get-Date

        try {
            # Etape 1 : dossiers Affaires (terminant par SA ou SB)
            $AffairesFolders = @(Get-ChildItem -Path $RootPath -Directory -ErrorAction SilentlyContinue |
                               Where-Object { $_.Name -match '(SA|SB)$' })

            if ($AffairesFolders.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Aucun dossier Affaire (SA/SB) trouve dans :`n$RootPath",
                    "Attention",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                ) | Out-Null
            }

            foreach ($AffaireFolder in $AffairesFolders) {
                if ($State.Cancel) { break }

                # Filtre marche sur le nom de l'affaire
                if (-not [string]::IsNullOrEmpty($MarketFilterValue)) {
                    if ($IsNegativeMarketFilter) {
                        if ($AffaireFolder.Name -like "*$MarketFilterValue*") { continue }
                    } else {
                        if (-not ($AffaireFolder.Name -like "*$MarketFilterValue*")) { continue }
                    }
                }

                foreach ($SubPath in $SubPathStructures) {
                    if ($State.Cancel) { break }

                    $PostesPath = if ([string]::IsNullOrEmpty($SubPath)) {
                        $AffaireFolder.FullName
                    } else {
                        Join-Path -Path $AffaireFolder.FullName -ChildPath $SubPath
                    }
                    if (-not (Test-Path -Path $PostesPath)) { continue }

                    $PostesFolders = Get-ChildItem -Path $PostesPath -Directory -ErrorAction SilentlyContinue
                    foreach ($PosteFolder in $PostesFolders) {
                        if ($State.Cancel) { break }

                        $FilesInPoste = Get-ChildItem -Path $PosteFolder.FullName -Filter "*.pptx" -ErrorAction SilentlyContinue
                        foreach ($PptxFile in $FilesInPoste) {
                            if ($State.Cancel) { break }

                            $FilesScanned++
                            $StatusLabel.Text = "Analyse : $($PptxFile.Name)  |  Fichiers scrut" + [char]233 + "s : $FilesScanned  |  R" + [char]233 + "sultats : $FoundCount"

                            $DateMod = $PptxFile.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")

                            $Slides = & $GetSlidesText -PptxPath $PptxFile.FullName
                            foreach ($Slide in $Slides) {
                                if ($State.Cancel) { break }

                                $idx = $Slide.Text.IndexOf($SearchInput, [System.StringComparison]::OrdinalIgnoreCase)
                                if ($idx -ge 0) {
                                    # Construire un extrait autour de la correspondance
                                    $Start = [Math]::Max(0, $idx - 30)
                                    $Len = [Math]::Min($Slide.Text.Length - $Start, $SearchInput.Length + 60)
                                    $Snippet = $Slide.Text.Substring($Start, $Len).Trim() -replace '\s+', ' '
                                    if ($Start -gt 0) { $Snippet = "..." + $Snippet }

                                    $DataGridView.Rows.Add(
                                        $AffaireFolder.Name,
                                        $PosteFolder.Name,
                                        $PptxFile.Name,
                                        $Slide.PageNumber,
                                        $DateMod,
                                        $Snippet,
                                        $PptxFile.FullName
                                    ) | Out-Null
                                    $FoundCount++
                                }
                            }

                            # Rafraichir l'interface au fur et a mesure (resultats a la volee)
                            [System.Windows.Forms.Application]::DoEvents()
                        }
                    }
                }
            }
        }
        finally {
            # Trier par date decroissante (plus recentes en haut)
            if ($DataGridView.Rows.Count -gt 0) {
                $DataGridView.Sort($DataGridView.Columns["Date"], [System.ComponentModel.ListSortDirection]::Descending)
            }

            $State.Running = $false
            $SearchButton.Enabled = $true
            $StopButton.Enabled = $false
            $ClearButton.Enabled = $true

            $SearchDuration = (Get-Date) - $SearchStartTime
            $EndMsg = if ($State.Cancel) { "Recherche interrompue" } else { "Recherche termin" + [char]233 + "e" }
            $StatusLabel.Text = "$EndMsg  |  Fichiers scrut" + [char]233 + "s : $FilesScanned  |  R" + [char]233 + "sultats : $FoundCount  |  Dur" + [char]233 + "e : $([math]::Round($SearchDuration.TotalSeconds, 1))s"
            $Form.Text = "Recherche de texte a la vol" + [char]233 + "e [$FoundCount resultat(s)] - V$Version"
        }
    })

    # Evenement du bouton Arreter
    $StopButton.Add_Click({
        $State.Cancel = $true
    })

    # Evenement du bouton Reinitialiser
    $ClearButton.Add_Click({
        if ($State.Running) { return }
        $TextBox.Text = ""
        $TextBoxMarket.Text = ""
        $DataGridView.Rows.Clear()
        $StatusLabel.Text = "Pr" + [char]234 + "t."
    })

    # Evenement Enter dans la textbox
    $TextBox.Add_KeyDown({
        if ($_.KeyCode -eq "Return") {
            $SearchButton.PerformClick()
        }
    })

    # Double-clic : ouvrir dossier marche / poste / fichier / slide
    $DataGridView.Add_CellDoubleClick({
        $RowIndex = $_.RowIndex
        if ($RowIndex -lt 0) { return }

        # Colonne 0 = Marche (Affaire)
        if ($_.ColumnIndex -eq 0) {
            $Affaire = $DataGridView.Rows[$RowIndex].Cells[0].Value
            $AffairePath = Join-Path -Path $Config.ExtractionRootPath -ChildPath $Affaire
            if (Test-Path $AffairePath) {
                try { Invoke-Item $AffairePath }
                catch { [System.Windows.Forms.MessageBox]::Show("Erreur lors de l'ouverture du dossier Marche: $_", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) }
            } else {
                [System.Windows.Forms.MessageBox]::Show("Dossier Marche non trouve: $AffairePath", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }

        # Colonne 1 = Poste
        if ($_.ColumnIndex -eq 1) {
            $Affaire = $DataGridView.Rows[$RowIndex].Cells[0].Value
            $Poste = $DataGridView.Rows[$RowIndex].Cells[1].Value
            $PostePath = $null
            foreach ($SubPath in $Config.SubPathStructures) {
                $PossiblePath = Join-Path -Path $Config.ExtractionRootPath -ChildPath $Affaire | Join-Path -ChildPath $SubPath | Join-Path -ChildPath $Poste
                if (Test-Path $PossiblePath) { $PostePath = $PossiblePath; break }
            }
            if (-not $PostePath) {
                $PostePath = Join-Path -Path $Config.ExtractionRootPath -ChildPath $Affaire | Join-Path -ChildPath $Poste
            }
            if (Test-Path $PostePath) {
                try { Invoke-Item $PostePath }
                catch { [System.Windows.Forms.MessageBox]::Show("Erreur lors de l'ouverture du dossier Poste: $_", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) }
            } else {
                [System.Windows.Forms.MessageBox]::Show("Dossier Poste non trouve: $PostePath", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }

        # Colonne 2 = SOP (ouvrir le fichier)
        if ($_.ColumnIndex -eq 2) {
            $FilePath = $DataGridView.Rows[$RowIndex].Cells[6].Value
            if (Test-Path $FilePath) {
                try { Invoke-Item $FilePath }
                catch { [System.Windows.Forms.MessageBox]::Show("Erreur lors de l'ouverture: $_", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) }
            } else {
                [System.Windows.Forms.MessageBox]::Show("Fichier non trouve: $FilePath", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }

        # Colonne 3 = Page (ouvrir PowerPoint a la slide)
        if ($_.ColumnIndex -eq 3) {
            $FilePath = $DataGridView.Rows[$RowIndex].Cells[6].Value
            $index = $DataGridView.Rows[$RowIndex].Cells[3].Value
            if (Test-Path $FilePath) {
                try {
                    $pp = New-Object -ComObject PowerPoint.Application
                    $pp.Visible = -1
                    $presentation = $pp.Presentations.Open($FilePath)
                    $pp.ActiveWindow.View.GotoSlide($index)
                }
                catch { [System.Windows.Forms.MessageBox]::Show("Erreur lors de l'ouverture: $_", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) }
                finally {
                    if ($presentation) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null }
                    if ($pp) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp) | Out-Null }
                    [GC]::Collect()
                    [GC]::WaitForPendingFinalizers()
                }
            } else {
                [System.Windows.Forms.MessageBox]::Show("Fichier non trouve: $FilePath", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    })

    # Arreter proprement la recherche si la fenetre se ferme
    $Form.Add_FormClosing({
        $State.Cancel = $true
    })

    # Ajouter les controles a la forme
    $Form.Controls.Add($SearchPanel)
    $Form.Controls.Add($DataGridView)
    $Form.Controls.Add($StatusLabel)

    # Afficher la fenetre
    $Form.ShowDialog() | Out-Null
}