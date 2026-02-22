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
        $Index
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

        $Affaires = @($Index.AffaireIndex.Keys | Select-Object -First 20)
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
$Form.Size = New-Object System.Drawing.Size(700, 450)
$Form.StartPosition = "CenterScreen"

$LabelPath = New-Object System.Windows.Forms.Label
$LabelPath.Text = "Fichier Excel (chemin complet) :"
$LabelPath.Location = New-Object System.Drawing.Point(10, 15)
$LabelPath.Size = New-Object System.Drawing.Size(200, 20)
$Form.Controls.Add($LabelPath)

$TextBoxPath = New-Object System.Windows.Forms.TextBox
$TextBoxPath.Location = New-Object System.Drawing.Point(10, 40)
$TextBoxPath.Size = New-Object System.Drawing.Size(540, 25)
$Form.Controls.Add($TextBoxPath)

$ButtonBrowse = New-Object System.Windows.Forms.Button
$ButtonBrowse.Text = "..."
$ButtonBrowse.Location = New-Object System.Drawing.Point(555, 38)
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
$ButtonRun.Location = New-Object System.Drawing.Point(10, 75)
$ButtonRun.Size = New-Object System.Drawing.Size(80, 30)
$ButtonRun.BackColor = [System.Drawing.Color]::LightGreen
$ButtonRun.Add_Click({
    Start-Comparison -ExcelPath $TextBoxPath.Text.Trim() -Grid $DataGrid -Index $SearchIndex
})
$Form.Controls.Add($ButtonRun)

$ButtonInit = New-Object System.Windows.Forms.Button
$ButtonInit.Text = "Initialiser"
$ButtonInit.Location = New-Object System.Drawing.Point(95, 75)
$ButtonInit.Size = New-Object System.Drawing.Size(80, 30)
$ButtonInit.Add_Click({
    $TextBoxPath.Text = ""
    $DataGrid.Rows.Clear()
    $LabelNewCount.Text = "Ref. nouvelles : -"
})
$Form.Controls.Add($ButtonInit)

$ButtonNovelty = New-Object System.Windows.Forms.Button
$ButtonNovelty.Text = "Taux news"
$ButtonNovelty.Location = New-Object System.Drawing.Point(180, 75)
$ButtonNovelty.Size = New-Object System.Drawing.Size(120, 30)
$ButtonNovelty.BackColor = [System.Drawing.Color]::LightCyan
$ButtonNovelty.Add_Click({
    Mark-NewReferencesInExcel -ExcelPath $TextBoxPath.Text.Trim() -Index $SearchIndex -Config $Config -LabelResult $LabelNewCount
})
$Form.Controls.Add($ButtonNovelty)

$LabelNewCount = New-Object System.Windows.Forms.Label
$LabelNewCount.Text = "Ref. nouvelles : -"
$LabelNewCount.Location = New-Object System.Drawing.Point(310, 80)
$LabelNewCount.Size = New-Object System.Drawing.Size(300, 20)
$LabelNewCount.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$LabelNewCount.ForeColor = [System.Drawing.Color]::Red
$Form.Controls.Add($LabelNewCount)

$DataGrid = New-Object System.Windows.Forms.DataGridView
$DataGrid.Location = New-Object System.Drawing.Point(10, 115)
$DataGrid.Size = New-Object System.Drawing.Size(660, 280)
$DataGrid.ReadOnly = $true
$DataGrid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$DataGrid.ColumnCount = 3
$DataGrid.Columns[0].Name = "March" + [char]233
$DataGrid.Columns[1].Name = "Excel / March" + [char]233
$DataGrid.Columns[2].Name = "March" + [char]233 + " / Excel"
$Form.Controls.Add($DataGrid)
$Form.Add_Shown({ $Form.Activate() })
[void]$Form.ShowDialog()
