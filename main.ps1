# Set-StrictMode -Version 3.0 # For development purposes only

if (-not (Get-command -module psexcel)) {
    Install-module PSExcel -Scope CurrentUser
    Get-command -module psexcel
    import-module psexcel
}

$host.PrivateData.ProgressBackgroundColor = $host.UI.RawUI.BackgroundColor
$host.privatedata.ProgressForegroundColor = "green";

$currentFolder = (Get-Location | Select-Object -ExpandProperty Path)

$file = $args[0]
$excel = new-object -com Excel.Application -Property @{ Visible = $false }

try {
    # Test-Path $file
    $workbook = $excel.Workbooks.Open($file)
}
catch {
    Write-Output "Invalid path: '$file'"
    exit
}

$myCustomObject = New-Object -TypeName psobject

$startRow = 1
$startColumn = 1
$tableHeaderRow = 1

# Loop though all the available worksheets
foreach ($sheet in $workbook.WorkSheets) {
    $rows = @()
    $columnCount = ($sheet.UsedRange.Columns).count
    $rowCount = ($sheet.UsedRange.Rows).count

    $rowEnd = $rowCount
    $columnEnd = $columnCount

    Write-Output "Worksheet $($sheet.Name)"

    # Which row is reperesents the table head
    $tableHeaderRowPrompt = Read-Host -Prompt "Table header row: $($tableHeaderRow)? [Overwrite]"

    if (-not [string]::IsNullOrWhiteSpace($tableHeaderRowPrompt)) {
        $tableHeaderRow = [int32]$tableHeaderRowPrompt
    }

    # Which row to start from
    $startRowPrompt = Read-Host -Prompt "Start row: $($startRow)? [Overwrite]"

    if (-not [string]::IsNullOrWhiteSpace($startRowPrompt)) {
        $startRow = [int32]$startRowPrompt
    }

    # Which column to start from
    $startColumnPrompt = Read-Host -Prompt "Start column: $($startColumn)? [Overwrite]"

    if (-not [string]::IsNullOrWhiteSpace($startColumnPrompt)) {
        $startColumn = [int32]$startColumnPrompt
    }

    # Set max row
    $rowEndPrompt = Read-Host -Prompt "$rowCount rows? [Overwrite]"

    if (-not [string]::IsNullOrWhiteSpace($rowEndPrompt)) {
        $rowEnd = [int32]$rowEndPrompt
    }

    #Set max column
    $columnEndPrompt = Read-Host -Prompt "$columnCount columns? [Overwrite]"

    if (-not [string]::IsNullOrWhiteSpace($columnEndPrompt)) {
        $columnEnd = [int32]$columnEndPrompt
    }

    Write-Output "Table head row $tableHeaderRow - Dimensions: Rows $($startRow)-$($rowEnd), Columns $($startColumn)-$($columnEnd)"

    $rowIterationCount = $rowEnd - $startRow + 1

    # Loop through rows and columns
    for ($row = $startRow; $row -le $rowEnd; $row++) {
        Write-Progress -Activity "Reading rows ..." -Status "Row $($row - $startRow + 1) of $rowIterationCount" -PercentComplete $($row / $rowEnd * 100)
        # $columns = @()
        [PSCustomObject]$columns = New-Object -TypeName psobject
        for ($column = $startColumn; $column -le $columnEnd; $column++) {
            $columnHead = $sheet.Cells.Item($tableHeaderRow, $column).Text
            $value = $sheet.Cells.Item($row, $column).Text

            if ($column -eq $startColumn -or -not ($columns.psobject.Properties.name -contains $columnHead)) {
                # Check if key already exists or object is empty
                $columns | Add-Member -MemberType NoteProperty -Name $columnHead -Value $value # Add all values in row to $columns
            }
        }
        $rows += , $columns # Add array of values to $rows
    }
    $myCustomObject | Add-Member -MemberType NoteProperty -Name $sheet.Name -Value $rows # Create new Key for each Worksheet with corresponding rows

    Write-Output "--------"
}

Write-Output $myCustomObject

# Save result to JSON file
[PSCustomObject]$myCustomObject | ConvertTo-Json -Depth 3 -Compress | Set-Content -Path "$($currentFolder)\psobject.json" -Encoding UTF8
