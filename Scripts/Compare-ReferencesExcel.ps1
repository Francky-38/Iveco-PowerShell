<#
.SYNOPSIS
    Compare les references type H du fichier Excel avec les references H des marches de la base

.DESCRIPTION
    Lit les refs H depuis un fichier Excel (col 14 = ref, col 3 = type, 1ere ligne = titre).
    Compare avec les refs H de chaque marche (Affaire) de la base.
    Resultat : tableau recapitulatif avec pour chaque marche :
    - Nom de l'affaire
    - (refs base trouvees dans Excel) / (total refs base du marche)
    - (refs Excel trouvees dans base) / (total refs Excel)

.EXAMPLE
    .\Scripts\Compare-ReferencesExcel.ps1
#>

param()

# Charger la configuration
$ConfigPath = Join-Path -Path $PSScriptRoot -ChildPath "..\Configs\config.ps1"
. $ConfigPath

# Charger les fonctions
$FunctionsPath = Join-Path -Path $PSScriptRoot -ChildPath "..\Functions\Helper.ps1"
. $FunctionsPath

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Lancer l'interface
$DataPath = $Config.ExtractXmlDataPath
$ZipPath = if ([System.IO.Path]::HasExtension($DataPath)) { $DataPath } else { "$DataPath.zip" }

if (-not (Test-Path $ZipPath)) {
    [System.Windows.Forms.MessageBox]::Show("Le fichier de donnees '$ZipPath' n'existe pas. Lancez d'abord l'extraction.", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit 1
}

$SearchIndex = New-SearchIndex -DataPath $DataPath
if (-not $SearchIndex) {
    [System.Windows.Forms.MessageBox]::Show("Impossible de charger l'index de recherche.", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit 1
}

function Mark-NewReferencesInExcel {
    param(
        [string]$ExcelPath,
        $Index,
        $Config,
        $LabelResult
    )
    if (-not $ExcelPath -or -not (Test-Path $ExcelPath)) {
        [System.Windows.Forms.MessageBox]::Show("Chemin Excel invalide ou fichier inexistant.", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    try {
        $Excel = New-Object -ComObject Excel.Application
        $Excel.Visible = $true
        $Excel.DisplayAlerts = $false
        $Workbook = $Excel.Workbooks.Open($ExcelPath, 0, $false)
        $Sheet = $Workbook.Sheets.Item(1)

        # Collecter toutes les refs H de la base (tous les marchés)
        $AllBaseRefs = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($Affaire in @($Index.AffaireIndex.Keys)) {
            $MarketRefs = Get-MarketHReferences -SearchIndex $Index -Affaire $Affaire
            foreach ($ref in $MarketRefs) {
                [void]$AllBaseRefs.Add($ref)
            }
        }

        # Parcourir les refs du fichier Excel et colorer
        $UsedRange = $Sheet.UsedRange
        $MaxRow = $UsedRange.Rows.Count
        $NewCount = 0

        for ($r = 2; $r -le $MaxRow; $r++) {
            $TypeVal = $Sheet.Cells.Item($r, $Config.ExcelTypeColumn).Text
            if ($TypeVal -eq "H") {
                $RefVal = $Sheet.Cells.Item($r, $Config.ExcelReferenceColumn).Text.Trim()
                $TypeCell = $Sheet.Cells.Item($r, $Config.ExcelTypeColumn)
                
                if ($AllBaseRefs.Contains($RefVal)) {
                    # Trouvée dans la base : enlever la couleur
                    $TypeCell.Interior.ColorIndex = -4142  # xlNone
                } else {
                    # Nouvelle : colorer en rouge
                    $TypeCell.Interior.Color = 255
                    $TypeCell.Interior.Pattern = 1  # xlSolid
                    $NewCount++
                }
            }
        }

        $Workbook.Save()
        $LabelResult.Text = "Ref. nouvelles : $NewCount"
        [System.Windows.Forms.MessageBox]::Show("Traitement termine. $NewCount reference(s) nouvelle(s) identifiee(s) et coloriee(s) en rouge.", "Info", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Erreur : $_", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
    finally {
        if ($Workbook) { $Workbook.Close($true) }
        if ($Excel) { $Excel.Quit() }
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Sheet) 2>$null | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Workbook) 2>$null | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) 2>$null | Out-Null
        [GC]::Collect()
    }
}

function Read-ExcelHReferences {
    param(
        [string]$ExcelPath,
        [int]$TypeColumn = 3,
        [int]$ReferenceColumn = 14
    )
    $Refs = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    try {
        $Excel = New-Object -ComObject Excel.Application
        $Excel.Visible = $false
        $Excel.DisplayAlerts = $false
        $Workbook = $Excel.Workbooks.Open($ExcelPath, 0, $true)
        $Sheet = $Workbook.Sheets.Item(1)
        $UsedRange = $Sheet.UsedRange
        $MaxRow = $UsedRange.Rows.Count
        for ($r = 2; $r -le $MaxRow; $r++) {
            $TypeVal = $Sheet.Cells.Item($r, $TypeColumn).Text
            if ($TypeVal -eq "H") {
                $RefVal = $Sheet.Cells.Item($r, $ReferenceColumn).Text
                if ($RefVal) {
                    [void]$Refs.Add($RefVal.Trim())
                }
            }
        }
        $Workbook.Close($false)
        $Excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Sheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
        [GC]::Collect()
    }
    catch {
        throw $_
    }
    return @($Refs)
}

function Start-Comparison {
    param(
        [string]$ExcelPath,
        [System.Windows.Forms.DataGridView]$Grid,
        $Index,
        [string[]]$SelectedMarkets,
        $Config
    )
    if (-not $ExcelPath -or -not (Test-Path $ExcelPath)) {
        [System.Windows.Forms.MessageBox]::Show("Chemin Excel invalide ou fichier inexistant.", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    $Grid.Rows.Clear()
    try {
        $ExcelRefs = Read-ExcelHReferences -ExcelPath $ExcelPath -TypeColumn $Config.ExcelTypeColumn -ReferenceColumn $Config.ExcelReferenceColumn
        $ExcelTotal = $ExcelRefs.Count
        if ($ExcelTotal -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Aucune reference de type H trouvee dans le fichier Excel.", "Info", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        $ExcelSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($r in $ExcelRefs) { [void]$ExcelSet.Add($r) }

        $Affaires = @($SelectedMarkets)
        foreach ($Affaire in $Affaires) {
            $MarketRefs = Get-MarketHReferences -SearchIndex $Index -Affaire $Affaire
            $MarketTotal = $MarketRefs.Count
            if ($MarketTotal -eq 0) {
                $Grid.Rows.Add($Affaire, "0% (0/0)", "0% (0/$ExcelTotal)") | Out-Null
                continue
            }
            $MarketSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            foreach ($r in $MarketRefs) { [void]$MarketSet.Add($r) }

            $MarketFoundInExcel = 0
            foreach ($m in $MarketRefs) {
                if ($ExcelSet.Contains($m)) { $MarketFoundInExcel++ }
            }
            $ExcelFoundInMarket = 0
            foreach ($e in $ExcelRefs) {
                if ($MarketSet.Contains($e)) { $ExcelFoundInMarket++ }
            }

            # Col2 : (refs du marché trouvées dans Excel) / (total refs H dans le marché)
            $PercentageMarketInExcel = if ($MarketTotal -gt 0) { [math]::Round(($MarketFoundInExcel / $MarketTotal) * 100) } else { 0 }
            # Col3 : (refs Excel trouvées dans le marché) / (total refs H dans Excel)
            $PercentageExcelInMarket = if ($ExcelTotal -gt 0) { [math]::Round(($ExcelFoundInMarket / $ExcelTotal) * 100) } else { 0 }
            $Col2 = "$PercentageMarketInExcel% ($MarketFoundInExcel/$MarketTotal)"
            $Col3 = "$PercentageExcelInMarket% ($ExcelFoundInMarket/$ExcelTotal)"
            $Grid.Rows.Add($Affaire, $Col2, $Col3) | Out-Null
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Erreur : $_", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# Interface
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Comparaison references H - Excel vs Base"
$Form.Size = New-Object System.Drawing.Size(700, 520)
$Form.StartPosition = "CenterScreen"

# Ajouter l'icône à la Form
$IconPath = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath "nono bleu.ico"
if (Test-Path -Path $IconPath) {
    $Form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($IconPath)
}

$LabelMarkets = New-Object System.Windows.Forms.Label
$LabelMarkets.Text = "Selectionner les march" + [char]233 + "s :"
$LabelMarkets.Location = New-Object System.Drawing.Point(10, 15)
$LabelMarkets.Size = New-Object System.Drawing.Size(200, 20)
$Form.Controls.Add($LabelMarkets)

$ButtonSelectGX = New-Object System.Windows.Forms.Button
$ButtonSelectGX.Text = "GX"
$ButtonSelectGX.Location = New-Object System.Drawing.Point(10, 38)
$ButtonSelectGX.Size = New-Object System.Drawing.Size(50, 25)
$ButtonSelectGX.Add_Click({
    $GridMarkets.Rows.Clear()
    foreach ($Affaire in $script:AllAffaires | Where-Object { $_ -like "*GX*" }) {
        $Numero = ""
        if ($Affaire -match '(20\d{8})') {
            $Numero = $Matches[1]
        }
        [void]$GridMarkets.Rows.Add($Affaire, $Numero)
    }
})
$Form.Controls.Add($ButtonSelectGX)

$ButtonSelectOther = New-Object System.Windows.Forms.Button
$ButtonSelectOther.Text = "Autres"
$ButtonSelectOther.Location = New-Object System.Drawing.Point(65, 38)
$ButtonSelectOther.Size = New-Object System.Drawing.Size(50, 25)
$ButtonSelectOther.Add_Click({
    $GridMarkets.Rows.Clear()
    foreach ($Affaire in $script:AllAffaires | Where-Object { $_ -notlike "*GX*" }) {
        $Numero = ""
        if ($Affaire -match '(20\d{8})') {
            $Numero = $Matches[1]
        }
        [void]$GridMarkets.Rows.Add($Affaire, $Numero)
    }
})
$Form.Controls.Add($ButtonSelectOther)

$ButtonSelectAll = New-Object System.Windows.Forms.Button
$ButtonSelectAll.Text = "Tous"
$ButtonSelectAll.Location = New-Object System.Drawing.Point(120, 38)
$ButtonSelectAll.Size = New-Object System.Drawing.Size(50, 25)
$ButtonSelectAll.Add_Click({
    $GridMarkets.Rows.Clear()
    foreach ($Affaire in $script:AllAffaires) {
        $Numero = ""
        if ($Affaire -match '(20\d{8})') {
            $Numero = $Matches[1]
        }
        [void]$GridMarkets.Rows.Add($Affaire, $Numero)
    }
})
$Form.Controls.Add($ButtonSelectAll)

$ButtonRemove = New-Object System.Windows.Forms.Button
$ButtonRemove.Text = "Sup. sel."
$ButtonRemove.Location = New-Object System.Drawing.Point(340, 38)
$ButtonRemove.Size = New-Object System.Drawing.Size(70, 25)
$ButtonRemove.Add_Click({
    # Supprimer en ordre inverse pour éviter les problèmes d'index
    for ($i = $GridMarkets.Rows.Count - 1; $i -ge 0; $i--) {
        if ($GridMarkets.Rows[$i].Selected) {
            $GridMarkets.Rows.RemoveAt($i)
        }
    }
})
$Form.Controls.Add($ButtonRemove)

$GridMarkets = New-Object System.Windows.Forms.DataGridView
$GridMarkets.Location = New-Object System.Drawing.Point(10, 68)
$GridMarkets.Size = New-Object System.Drawing.Size(400, 180)
$GridMarkets.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$GridMarkets.MultiSelect = $true
$GridMarkets.AllowUserToAddRows = $false
$GridMarkets.AllowUserToDeleteRows = $false
$GridMarkets.AllowUserToResizeRows = $false
$GridMarkets.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
$GridMarkets.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::None

# Créer les colonnes
$ColName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$ColName.Name = "Marche"
$ColName.HeaderText = "March" + [char]233
$ColName.Width = 287
$GridMarkets.Columns.Add($ColName) | Out-Null

$ColNum = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$ColNum.Name = "Numero"
$ColNum.HeaderText = "Num" + [char]233 + "ro"
$ColNum.Width = 70
$GridMarkets.Columns.Add($ColNum) | Out-Null

# Remplir avec tous les marchés de la base
$script:AllAffaires = @($SearchIndex.AffaireIndex.Keys | Sort-Object)
foreach ($Affaire in $script:AllAffaires) {
    # Extraire le numéro (10 chiffres commençant par 20)
    $Numero = ""
    if ($Affaire -match '(20\d{8})') {
        $Numero = $Matches[1]
    }
    [void]$GridMarkets.Rows.Add($Affaire, $Numero)
}
# Trier par Numero en ordre décroissant
$GridMarkets.Sort($GridMarkets.Columns["Numero"], [System.ComponentModel.ListSortDirection]::Descending)
$Form.Controls.Add($GridMarkets)

$LabelPath = New-Object System.Windows.Forms.Label
$LabelPath.Text = "Selectionner l'export Excel :"
$LabelPath.Location = New-Object System.Drawing.Point(430, 15)
$LabelPath.Size = New-Object System.Drawing.Size(200, 20)
$Form.Controls.Add($LabelPath)

$TextBoxPath = New-Object System.Windows.Forms.TextBox
$TextBoxPath.Location = New-Object System.Drawing.Point(430, 40)
$TextBoxPath.Size = New-Object System.Drawing.Size(200, 25)
$Form.Controls.Add($TextBoxPath)

$ButtonBrowse = New-Object System.Windows.Forms.Button
$ButtonBrowse.Text = "..."
$ButtonBrowse.Location = New-Object System.Drawing.Point(635, 38)
$ButtonBrowse.Size = New-Object System.Drawing.Size(35, 25)
$ButtonBrowse.Add_Click({
    $Dlg = New-Object System.Windows.Forms.OpenFileDialog
    $Dlg.Filter = "Fichiers Excel (*.xlsx;*.xls)|*.xlsx;*.xls"
    $Dlg.Title = "Selectionner le fichier Excel"
    if ($Dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:TextBoxPath.Text = $Dlg.FileName
    }
})
$Form.Controls.Add($ButtonBrowse)

$ButtonRun = New-Object System.Windows.Forms.Button
$ButtonRun.Text = "Lancer"
$ButtonRun.Location = New-Object System.Drawing.Point(430, 70)
$ButtonRun.Size = New-Object System.Drawing.Size(80, 30)
$ButtonRun.BackColor = [System.Drawing.Color]::LightGreen
$ButtonRun.Add_Click({
    $SelectedMarkets = @()
    foreach ($Row in $GridMarkets.SelectedRows) {
        $SelectedMarkets += $Row.Cells["Marche"].Value
    }
    if ($SelectedMarkets.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Veuillez selectionner au moins un march" + [char]233 + ".", "Avertissement", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    Start-Comparison -ExcelPath $TextBoxPath.Text.Trim() -Grid $DataGrid -Index $SearchIndex -SelectedMarkets $SelectedMarkets -Config $Config
})
$Form.Controls.Add($ButtonRun)

$ButtonInit = New-Object System.Windows.Forms.Button
$ButtonInit.Text = "Initialiser"
$ButtonInit.Location = New-Object System.Drawing.Point(515, 70)
$ButtonInit.Size = New-Object System.Drawing.Size(80, 30)
$ButtonInit.Add_Click({
    $TextBoxPath.Text = ""
    $DataGrid.Rows.Clear()
    $LabelNewCount.Text = "Ref. nouvelles : -"
})
$Form.Controls.Add($ButtonInit)

$ButtonNovelty = New-Object System.Windows.Forms.Button
$ButtonNovelty.Text = "Taux news"
$ButtonNovelty.Location = New-Object System.Drawing.Point(600, 70)
$ButtonNovelty.Size = New-Object System.Drawing.Size(70, 30)
$ButtonNovelty.BackColor = [System.Drawing.Color]::LightCyan
$ButtonNovelty.Add_Click({
    Mark-NewReferencesInExcel -ExcelPath $TextBoxPath.Text.Trim() -Index $SearchIndex -Config $Config -LabelResult $LabelNewCount
})
$Form.Controls.Add($ButtonNovelty)

$LabelNewCount = New-Object System.Windows.Forms.Label
$LabelNewCount.Text = "Ref. nouvelles : -"
$LabelNewCount.Location = New-Object System.Drawing.Point(430, 110)
$LabelNewCount.Size = New-Object System.Drawing.Size(300, 20)
$LabelNewCount.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$LabelNewCount.ForeColor = [System.Drawing.Color]::Red
$Form.Controls.Add($LabelNewCount)

$DataGrid = New-Object System.Windows.Forms.DataGridView
$DataGrid.Location = New-Object System.Drawing.Point(10, 255)
$DataGrid.Size = New-Object System.Drawing.Size(660, 215)
$DataGrid.ReadOnly = $true
$DataGrid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$DataGrid.ColumnCount = 3
$DataGrid.Columns[0].Name = "March" + [char]233
$DataGrid.Columns[1].Name = "Excel / March" + [char]233
$DataGrid.Columns[2].Name = "March" + [char]233 + " / Excel"
$Form.Controls.Add($DataGrid)
$Form.Add_Shown({ $Form.Activate() })
[void]$Form.ShowDialog()
