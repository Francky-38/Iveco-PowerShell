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

function Get-WelcomeMessage {
    param(
        [string]$Message = "Bienvenue dans le projet Iveco!"
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
        [string]$OutputFile = ""
    )

    # Vérifier que le chemin existe
    if (-not (Test-Path -Path $RootPath)) {
        Write-Host "Erreur: Le chemin '$RootPath' n'existe pas" -ForegroundColor Red
        return
    }

    # Déterminer le fichier de sortie
    if ([string]::IsNullOrEmpty($OutputFile)) {
        $OutputFile = Join-Path -Path $RootPath -ChildPath "References_Extraites.xml"
    }

    Write-Host "`nRecherche des fichiers PPTX dans l'arborescence..." -ForegroundColor Yellow
    Write-Host "Chemin racine: $RootPath" -ForegroundColor Cyan
    Write-Host "Sortie: $OutputFile" -ForegroundColor Cyan

    # Créer le fichier XML de sortie principal
    $XmlOutput = New-Object System.Xml.XmlDocument
    $Root = $XmlOutput.CreateElement("References")
    $XmlOutput.AppendChild($Root) | Out-Null

    $TotalReferences = 0
    $TotalFiles = 0
    $TempDir = Join-Path -Path $env:TEMP -ChildPath "PPTX_Extract_$(Get-Random)"

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
        
        # Étape 2: Pour chaque Affaire, chercher dans le chemin spécifique
        foreach ($AffaireFolder in $AffairesFolders) {
            Write-Host "Exploration Affaire: $($AffaireFolder.Name)" -ForegroundColor Cyan
            
            # Construire le chemin vers les dossiers postes
            $PostesPath = Join-Path -Path $AffaireFolder.FullName -ChildPath "01-Dossiers ligne EL-EG\LIGNE EG0"
            
            if (Test-Path -Path $PostesPath) {
                # Étape 3: Chercher les dossiers de postes (1er niveau, pas récursif)
                $PostesFolders = Get-ChildItem -Path $PostesPath -Directory -ErrorAction SilentlyContinue
                
                foreach ($PosteFolder in $PostesFolders) {
                    Write-Host "  Poste: $($PosteFolder.Name)" -ForegroundColor Gray
                    
                    # Étape 4: Chercher les fichiers .pptx dans ce dossier (pas récursif)
                    $FilesInPoste = Get-ChildItem -Path $PosteFolder.FullName -Filter "*.pptx" -ErrorAction SilentlyContinue
                    
                    if ($FilesInPoste.Count -gt 0) {
                        Write-Host "    - $($FilesInPoste.Count) fichier(s) PPTX" -ForegroundColor Gray
                        
                        # Traiter chaque fichier PPTX directement
                        foreach ($PptxFile in $FilesInPoste) {
                            $TotalFiles++
                            Write-Host "      [$TotalFiles] Traitement: $($PptxFile.Name)" -ForegroundColor Gray

                            try {
                                # Créer un dossier temporaire unique pour chaque fichier
                                $FileTempDir = Join-Path -Path $env:TEMP -ChildPath "PPTX_Extract_$(Get-Random)"
                                New-Item -ItemType Directory -Path $FileTempDir -Force | Out-Null

                                # Extraire l'archive PPTX
                                Add-Type -AssemblyName System.IO.Compression.FileSystem
                                [System.IO.Compression.ZipFile]::ExtractToDirectory($PptxFile.FullName, $FileTempDir)

                                # Parcourir les fichiers XML dans ppt\slides
                                $SlidesPath = Join-Path -Path $FileTempDir -ChildPath "ppt\slides"

                                if (Test-Path -Path $SlidesPath) {
                                    $XmlFiles = Get-ChildItem -Path $SlidesPath -Filter "*.xml" -ErrorAction SilentlyContinue

                                    foreach ($XmlFile in $XmlFiles) {
                                        try {
                                            $XmlContent = [System.Xml.XmlDocument]::new()
                                            $XmlContent.Load($XmlFile.FullName)

                                            # Récupérer tous les nœuds <a:t> avec gestion du namespace
                                            $NamespaceManager = New-Object System.Xml.XmlNamespaceManager($XmlContent.NameTable)
                                            $NamespaceManager.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
                                            $TextNodes = $XmlContent.SelectNodes("//a:t", $NamespaceManager)

                                            foreach ($TextNode in $TextNodes) {
                                                $Text = $TextNode.InnerText.Trim()

                                                # Chercher toutes les références dans le texte
                                                # Format: [TRS]?\d{5,10} (peut être au milieu d'une chaîne)
                                                $References = [regex]::Matches($Text, '[TRS]?\d{5,10}')
                                                
                                                foreach ($RefMatch in $References) {
                                                    $Reference = $RefMatch.Value
                                                    
                                                    # Les noms d'Affaire et Poste viennent des boucles actuelles
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

                                                    $FileElem = $XmlOutput.CreateElement("Page")
                                                    $FileElem.InnerText = $XmlFile.Name
                                                    $Entry.AppendChild($FileElem) | Out-Null

                                                    $RefElem = $XmlOutput.CreateElement("Reference")
                                                    $RefElem.InnerText = $Reference
                                                    $Entry.AppendChild($RefElem) | Out-Null

                                                    $Root.AppendChild($Entry) | Out-Null
                                                    $TotalReferences++
                                                }
                                            }
                                        }
                                        catch {
                                            Write-Host "      Erreur lors du traitement de $($XmlFile.Name): $_" -ForegroundColor Red
                                        }
                                    }
                                }

                                # Nettoyer le dossier temporaire pour ce fichier
                                Remove-Item -Path $FileTempDir -Recurse -Force -ErrorAction SilentlyContinue
                            }
                            catch {
                                Write-Host "    Erreur lors du traitement de $($PptxFile.Name): $_" -ForegroundColor Red
                            }
                        }
                    }
                }
            } else {
                Write-Host "  Attention: Chemin '01-Dossiers ligne EL-EG\LIGNE EG0' non trouve" -ForegroundColor Yellow
            }
        }
        
        if ($TotalFiles -eq 0) {
            Write-Host "Attention: Aucun fichier PPTX trouve" -ForegroundColor Yellow
            return
        }
        
        Write-Host ""

        # Sauvegarder le fichier XML
        $XmlOutput.Save($OutputFile)
        
        Write-Host ""
        Write-Host "OK - Traitement termine!" -ForegroundColor Green
        Write-Host "  - Fichiers PPTX traites: $TotalFiles" -ForegroundColor Cyan
        Write-Host "  - Nombre total de references extraites: $TotalReferences" -ForegroundColor Cyan
        Write-Host "  - Fichier de sortie: $OutputFile" -ForegroundColor Cyan
    }
    catch {
        Write-Host "Erreur lors du traitement: $_" -ForegroundColor Red
    }
    finally {
        # Nettoyer le dossier temporaire principal
        Remove-Item -Path $TempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

<#
.SYNOPSIS
    Interface interactive de recherche de références

.DESCRIPTION
    Affiche un menu interactif pour rechercher des références dans le fichier XML

.PARAMETER XmlPath
    Chemin du fichier XML d'extraction

.EXAMPLE
    Show-SearchMenu -XmlPath "D:\W\Iveco\RefServeur.xml"
#>

function Show-SearchGui {
    param(
        [string]$XmlPath
    )

    # Vérifier que le fichier XML existe
    if (-not (Test-Path -Path $XmlPath)) {
        [System.Windows.Forms.MessageBox]::Show("Erreur: Le fichier XML '$XmlPath' n'existe pas", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # Charger les assemblies
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    # Charger le fichier XML
    $XmlContent = [System.Xml.XmlDocument]::new()
    $XmlContent.Load($XmlPath)
    $AllEntries = $XmlContent.SelectNodes("//Entree")
    
    # Dictionnaire pour stocker les chemins des fichiers PPTX
    $PptxPaths = @{}

    # Créer la fenêtre principale
    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = "Recherche de References - Projet Iveco"
    $Form.Size = New-Object System.Drawing.Size(1120, 550)
    $Form.StartPosition = "CenterScreen"
    $Form.BackColor = [System.Drawing.Color]::WhiteSmoke

    # Panel supérieur (recherche)
    $SearchPanel = New-Object System.Windows.Forms.Panel
    $SearchPanel.Location = New-Object System.Drawing.Point(10, 10)
    $SearchPanel.Size = New-Object System.Drawing.Size(1080, 60)
    $SearchPanel.BackColor = [System.Drawing.Color]::White
    $SearchPanel.BorderStyle = "Fixed3D"

    # Label
    $Label = New-Object System.Windows.Forms.Label
    $Label.Text = "R" + [char]233 + "f" + [char]233 + "rence cherch" + [char]233 + "e:"
    $Label.Location = New-Object System.Drawing.Point(10, 10)
    $Label.Size = New-Object System.Drawing.Size(150, 20)
    $Label.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)

    # TextBox
    $TextBox = New-Object System.Windows.Forms.TextBox
    $TextBox.Location = New-Object System.Drawing.Point(160, 10)
    $TextBox.Size = New-Object System.Drawing.Size(200, 25)
    $TextBox.Font = New-Object System.Drawing.Font("Arial", 10)

    # Bouton Rechercher
    $SearchButton = New-Object System.Windows.Forms.Button
    $SearchButton.Text = "Rechercher"
    $SearchButton.Location = New-Object System.Drawing.Point(370, 10)
    $SearchButton.Size = New-Object System.Drawing.Size(100, 30)
    $SearchButton.BackColor = [System.Drawing.Color]::LightBlue
    $SearchButton.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $SearchButton.FlatStyle = "Flat"

    # Bouton Réinitialiser
    $ClearButton = New-Object System.Windows.Forms.Button
    $ClearButton.Text = "R" + [char]233 + "initialiser"
    $ClearButton.Location = New-Object System.Drawing.Point(480, 10)
    $ClearButton.Size = New-Object System.Drawing.Size(100, 30)
    $ClearButton.BackColor = [System.Drawing.Color]::LightGray
    $ClearButton.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $ClearButton.FlatStyle = "Flat"

    $SearchPanel.Controls.Add($Label)
    $SearchPanel.Controls.Add($TextBox)
    $SearchPanel.Controls.Add($SearchButton)
    $SearchPanel.Controls.Add($ClearButton)

    # DataGridView pour les résultats
    $DataGridView = New-Object System.Windows.Forms.DataGridView
    $DataGridView.Location = New-Object System.Drawing.Point(10, 80)
    $DataGridView.Size = New-Object System.Drawing.Size(1080, 410)
    $DataGridView.AllowUserToAddRows = $false
    $DataGridView.AllowUserToDeleteRows = $false
    $DataGridView.ReadOnly = $true
    $DataGridView.SelectionMode = "FullRowSelect"
    $DataGridView.BackgroundColor = [System.Drawing.Color]::White
    $DataGridView.GridColor = [System.Drawing.Color]::LightGray
    $DataGridView.Font = New-Object System.Drawing.Font("Arial", 9)

    # Ajouter les colonnes
    $DataGridView.ColumnCount = 4
    $DataGridView.Columns[0].Name = "March" + [char]233
    $DataGridView.Columns[0].Width = 450
    $DataGridView.Columns[1].Name = "Poste"
    $DataGridView.Columns[1].Width = 330
    $DataGridView.Columns[2].Name = "SOP"
    $DataGridView.Columns[2].Width = 200
    $DataGridView.Columns[3].Name = "Page"
    $DataGridView.Columns[3].Width = 50

    # En-tête
    $DataGridView.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::DarkBlue
    $DataGridView.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
    $DataGridView.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)

    # Événement du bouton Rechercher
    $SearchButton.Add_Click({
        $SearchTerm = $TextBox.Text.Trim()
        
        if ([string]::IsNullOrEmpty($SearchTerm)) {
            [System.Windows.Forms.MessageBox]::Show("Veuillez entrer une reference a rechercher", "Attention", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $DataGridView.Rows.Clear()
        $FoundCount = 0

        foreach ($Entry in $AllEntries) {
            $Reference = $Entry.SelectSingleNode("Reference").InnerText
            
            if ($Reference -like "*$SearchTerm*") {
                $Affaire = $Entry.SelectSingleNode("Affaire").InnerText
                $Poste = $Entry.SelectSingleNode("Poste").InnerText
                $Archive = $Entry.SelectSingleNode("SOP").InnerText
                $Fichier = $Entry.SelectSingleNode("Page").InnerText
                
                # Extraire le numéro de page de "slide1.xml" → 1
                $PageNumber = ""
                if ($Fichier -match 'slide(\d+)') {
                    $PageNumber = $Matches[1]
                }

                $DataGridView.Rows.Add($Affaire, $Poste, $Archive, $PageNumber)
                $Path = $Entry.SelectSingleNode("Path").InnerText
                $PptxPaths[$FoundCount] = $Path
                $FoundCount++
            }
        }

        $Form.Text = "Recherche de References - Projet Iveco [$FoundCount resultat(s)]"
    })

    # Événement du bouton Réinitialiser
    $ClearButton.Add_Click({
        $TextBox.Text = ""
        $DataGridView.Rows.Clear()
        $Form.Text = "Recherche de References - Projet Iveco"
    })

    # Événement Enter dans la textbox
    $TextBox.Add_KeyDown({
        if ($_.KeyCode -eq "Return") {
            $SearchButton.PerformClick()
        }
    })

    # Événement double-clic sur une cellule de la colonne SOP pour ouvrir le fichier
    $DataGridView.Add_CellDoubleClick({
        if ($_.ColumnIndex -eq 2) {  # Colonne 2 = SOP
            $RowIndex = $_.RowIndex
            $SopName = $DataGridView.Rows[$RowIndex].Cells[2].Value
            
            if ($PptxPaths.ContainsKey($RowIndex)) {
                $FilePath = $PptxPaths[$RowIndex]
                
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
            } else {
                [System.Windows.Forms.MessageBox]::Show("Chemin non trouve pour: $SopName", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
        if ($_.ColumnIndex -eq 3) {  # Colonne 3 = Page
            $RowIndex = $_.RowIndex
            $SopName = $DataGridView.Rows[$RowIndex].Cells[2].Value
            $index=$DataGridView.Rows[$RowIndex].Cells[3].Value
            if ($PptxPaths.ContainsKey($RowIndex)) {
                $FilePath = $PptxPaths[$RowIndex]
                
                if (Test-Path $FilePath) {
                    try {
                        $pp = New-Object -ComObject PowerPoint.Application
                        $pp.Visible = -1
                        $presentation = $pp.Presentations.Open($FilePath)
                        $pp.ActiveWindow.View.GotoSlide($index)

                        
                        #Invoke-Item $FilePath
                    }
                    catch {
                        [System.Windows.Forms.MessageBox]::Show("Erreur lors de l'ouverture: $_", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    }
                } else {
                    [System.Windows.Forms.MessageBox]::Show("Fichier non trouve: $FilePath", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                }
            } else {
                [System.Windows.Forms.MessageBox]::Show("Chemin non trouve pour: $SopName", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
        
    })

    # Événement MouseMove pour afficher le chemin du fichier en survolant la colonne SOP
    $DataGridView.Add_MouseMove({
        $HitTest = $DataGridView.HitTest($_.X, $_.Y)
        
        if ($HitTest.ColumnIndex -eq 2 -and $HitTest.RowIndex -ge 0) {
            $RowIndex = $HitTest.RowIndex
            $PageNum = $DataGridView.Rows[$HitTest.RowIndex].Cells[3].Value
            if ($PptxPaths.ContainsKey($RowIndex)) {
                $Form.Text = "Chemin: " + $PptxPaths[$RowIndex] + " | Page: " + $PageNum
            }
        } else {
            $SearchCount = $DataGridView.Rows.Count
            $Form.Text = "Recherche de References - Projet Iveco [$SearchCount resultat(s)]"
        }
    })

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