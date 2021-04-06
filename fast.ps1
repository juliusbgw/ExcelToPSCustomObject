#fast.ps1
param ([
    Parameter(Mandatory)]
    [string]$file,
    [string]$include,
    [Boolean]$auto = $True
)

# Set-StrictMode -Version 3.0 # For development purposes only
Function ExcelToCsv () {
    param (
        $file,
        $csvFolder
    )
    $myDir = (Get-Location | Select-Object -ExpandProperty Path)
    $excel = new-object -com Excel.Application -Property @{ Visible = $false }
    $excel.DisplayAlerts = $false
    $wb = $Excel.Workbooks.Open($file)
    $xlCSV = 6
    try {
        foreach ($ws in $wb.Worksheets) {
            $x = $ws
            $x.SaveAs("$myDir\$csvFolder\$($ws.name).csv", $xlCSV)
        }
    }
    catch {
        Write-Error $_
    }
    $Excel.Workbooks.Close()
    $Excel.Quit()
}

function FindTable {
    param (
        $pathToTable
    )

    # Load CSV or text file
    $file = Get-Content -Path $pathToTable

    # Initialize some variables for later
    $longestTableHeadCount = 0
    $row = 0
    $tableHeadRow = 0
    $startColumn = 0
    $endColumn = 10
    $concatenatedLine = ""
    $rows = @()
    $defaultDelimiter = ","

    # Read file line by line
    foreach ($line in $file) {
        $lineSegment = $line.Split($defaultDelimiter)

        # Check for empty lines -> continue with next line
        if (($lineSegment -join "") -eq "") {
            continue
        }

        $currentLine = $line

        # Check if the first character of the last item in the line is equal to '"' -> Indicates line break
        # -> continue with reading the remaining items in this line
        # Example: x,y,"z

        if ($lineSegment[-1][0] -eq '"' ) {
            $concatenatedLine += $line
            continue
        }

        # Check if the last character of the first item in the line is equal to '"' -> Indicates end of total line
        # -> continue with reading the remaining items in this line
        # Example: x",y,z

        elseif ($lineSegment[0][-1] -eq '"') {
            $concatenatedLine += $line
            $currentLine = $concatenatedLine
            $concatenatedLine = ""
        }

        $cleanColumns = $currentLine.Replace("`n", " ").Replace("`r", " ") # Remove special characters
        $completeColumns = $cleanColumns.Split($defaultDelimiter)
        $longestTableHead = @()

        # Remove all empty items and get the longest series of items in current row
        [regex]::split($cleanColumns, '[,]{2,}') | ForEach-Object {
            $headers = $_.Split($defaultDelimiter)
            if ( $headers.Length -gt $longestTableHead.Length) {
                $longestTableHead = $headers
            }
        }

        $rows += , $completeColumns

        # Get table head
        if ($longestTableHead.Length -gt $longestTableHeadCount) {
            $longestTableHeadCount = $longestTableHead.Length
            $startColumn = $completeColumns.IndexOf($longestTableHead[0])
            $endColumn = $completeColumns.IndexOf($longestTableHead[-1])
            $tableHeadRow = $row
        }
        $row++
    }

    $table = [PSCustomObject]@{
        rows         = $rows
        tableHeadRow = $tableHeadRow;
        startColumn  = $startColumn;
        endColumn    = $endColumn;
        lastRow      = $row;
    }

    return $table
}

function Run {
    param (
        [Parameter(Mandatory)]
        [string]$file,
        [string]$sheetsToInclude,
        [Boolean]$auto = $True
    )

    if (-not (Get-command -module psexcel)) {
        Install-module PSExcel -Scope CurrentUser
        Get-command -module psexcel
        import-module psexcel
    }

    $host.PrivateData.ProgressBackgroundColor = $host.UI.RawUI.BackgroundColor
    $host.privatedata.ProgressForegroundColor = "green";

    $csvFolderName = ".csv-files" # Temporary storage
    $currentFolder = (Get-Location | Select-Object -ExpandProperty Path)

    $pathToCSVs = $currentFolder + "\$csvFolderName"

    if (-not(Test-Path $pathToCSVs)) {
        mkdir $csvFolderName | Out-Null
    }

    ExcelToCsv $file $csvFolderName # Export excel sheets as CSVs

    $BigPSCustomObject = New-Object -TypeName psobject # Custom object to store the sheets

    # Excel does not encode as UTF8 by default
    Get-ChildItem "$currentFolder\$csvFolderName" | ForEach-Object {
        $content = Get-Content -Path "$pathToCSVs\$($_.Name)"
        $content | Set-Content -Path "$pathToCSVs\$($_.Name)" -Encoding UTF8
    }

    # Loop through all of the CSV files created that represent the sheets in the Excel file
    Get-ChildItem "$currentFolder\$csvFolderName" | ForEach-Object {
        $rowsWithHeader = @()
        $sheetName = $_.Name.Replace(".csv", "")

        if ($sheetsToInclude -and -not($sheetsToInclude -match $sheetName)) {
            return
        }

        Write-Progress -Activity "Reading sheets..." -Status "WS '$($sheetName)'"

        $pathToSheet = "$pathToCSVs\$sheetName.csv"

        $table = FindTable $pathToSheet # Make predictions about the table dimensions

        $rows, $tableHeadRow, $startColumn, $endColumn, $lastRow = `
            $table.rows, $table.tableHeadRow, `
            $table.startColumn, $table.endColumn, $table.lastRow

        $startRow = $tableHeadRow + 1

        if ($auto -eq $False) {
            $tableHeadRowPrompt = Read-Host -Prompt "Table head row [Default: $($tableHeadRow)]?"

            if (-not [string]::IsNullOrWhiteSpace($tableHeadRowPrompt)) {
                $tableHeadRow = [int32]$tableHeadRowPrompt
            }

            $startRowPrompt = Read-Host -Prompt "Start row [Default: $($startRow)]?"

            if (-not [string]::IsNullOrWhiteSpace($startRowPrompt)) {
                $startRow = [int32]$startRowPrompt
            }
        }

        # And loop ...
        for ($row = $startRow; $row -lt $lastRow; $row++) {
            [PSCustomObject]$columnsWithHeader = New-Object -TypeName psobject

            for ($column = $startColumn; $column -lt $endColumn; $column++) {
                $columnHead = $rows[$tableHeadRow][$column]
                $cellValue = $rows[$row][$column]

                if ($columnsWithHeader.psobject.Properties.name -contains $columnHead) {
                    # Check if key already exists or object is empty
                    continue
                }

                $columnsWithHeader | Add-Member -MemberType NoteProperty -Name $columnHead -Value $cellValue
            }

            $rowsWithHeader += , $columnsWithHeader
        }
        $BigPSCustomObject | Add-Member -MemberType NoteProperty -Name $sheetName -Value `
            $rowsWithHeader # Create new Key for each Worksheet with corresponding rows
    }

    $jsonFileName = "$($file.Split("\")[-1].Replace('.xlsx', '')).json"
    Write-Host "`nSaving results to $($jsonFileName)..." -ForegroundColor Green


    [PSCustomObject]$BigPSCustomObject `
    | ConvertTo-Json -Depth 3 -Compress `
    | Set-Content -Path "$($currentFolder)\$jsonFileName" -Encoding UTF8

    Write-Host "Done!" -ForegroundColor Green
    Remove-Item -R $csvFolderName
}

# Run
Run -file $file -sheetsToInclude $include -auto $auto