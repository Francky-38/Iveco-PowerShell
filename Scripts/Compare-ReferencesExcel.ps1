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

function Read-ExcelHReferences {
    param([string]$ExcelPath)
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
            $TypeVal = $Sheet.Cells.Item($r, 3).Text
            if ($TypeVal -eq "H") {
                $RefVal = $Sheet.Cells.Item($r, 14).Text
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
        $ExcelRefs = Read-ExcelHReferences -ExcelPath $ExcelPath
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
                $Grid.Rows.Add($Affaire, "0/0", "0/$ExcelTotal") | Out-Null
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

            $Col2 = "$MarketFoundInExcel/$MarketTotal"
            $Col3 = "$ExcelFoundInMarket/$ExcelTotal"
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
})
$Form.Controls.Add($ButtonInit)

$DataGrid = New-Object System.Windows.Forms.DataGridView
$DataGrid.Location = New-Object System.Drawing.Point(10, 115)
$DataGrid.Size = New-Object System.Drawing.Size(660, 280)
$DataGrid.ReadOnly = $true
$DataGrid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$DataGrid.ColumnCount = 3
$DataGrid.Columns[0].Name = "Affaire"
$DataGrid.Columns[1].Name = "Base trouvee dans Excel"
$DataGrid.Columns[2].Name = "Excel trouve dans Base"
$Form.Controls.Add($DataGrid)
$Form.Add_Shown({ $Form.Activate() })
[void]$Form.ShowDialog()
