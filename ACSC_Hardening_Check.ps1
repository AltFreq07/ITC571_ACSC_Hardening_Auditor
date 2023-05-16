param (
    [string] $title = "",
    [string] $output = ""
)



function GetDateString() {
    return Get-Date -Format "ddMMyy_HHmm"
}

# Check if script has admin privileges.
function CheckAdmin() {
    return $null -eq ${env:=::}
}

# Function to load CSV file
function LoadCSV {
    param(
        [string]$path
    )
    return Import-Csv $path
}

function AppendToCSV() {
    param(
        [string]$str
    )
    Add-Content -Path $global:OutputPath -Value $str
}

function CreateCSV() {
    Out-File -FilePath $global:OutputPath
    AppendToCSV "Section,Query,Expected Value,Actual Value,Bool,Reversed,Unit"
}

function CheckRule {
    param(
        [object]$rule
    )
    Write-Output "Checking rule:" $rule
    if ($rule."Registry Location") {
        # Write-Output "Registry Value exists"
        # Gather Current Registry Value
        $currentValue = (Get-ItemProperty -ErrorAction SilentlyContinue -Path $rule."Registry Location").$($rule."Registry Key")
        if ($currentValue -is [Boolean] -Or $currentValue -is [int]) {
            $currentValue = $currentValue.ToString()
        }
        Write-Output "Registry Value: $(IF($currentValue) {$currentValue} ELSE {Write-Output null})"
    }
    elseif ($rule.Command) {
        # Check for admin and run command\
        if (([boolean]$rule.Admin -And $global:Admin) -Or ![boolean]$rule.Admin) {
            try {
                $currentValue = Invoke-Expression -Command $rule.Command
                if ($currentValue -is [Boolean] -Or $currentValue -is [int]) {
                    $currentValue = $currentValue.ToString()
                }
            }
            catch {
                $currentValue = ""
            }
            Write-Output "Command Output: $(IF($currentValue) {$currentValue} ELSE {Write-Output null})"
        }
        else {
            # Admin credentials required but not provided
            AppendToCSV "$($rule.Section),$(IF($rule."Group Policy") {$rule."Group Policy"} Else {$rule.Command}),$($rule."Expected Value"),$(IF($($currentValue)){$currentValue} Else {Write-Output "Admin credentials required"}),$($rule.Bool),$($rule.Reversed),$($rule.Unit)"
            continue
        }
        
    }
    else {
        # Missing registry value or command
        Write-Output "Registry Key or Command does not exist"
        break
    }

    # Check for expected value
    if (!$rule."Expected Value") {
        Write-Output "Expected Value does not exist"
        break
    }

    if ($rule."Expected Value" -eq "-1" ) {
        if (!$currentValue) {
            $rule."Expected Value" = "!Exist"
            $currentValue = "!Exist"
        }
    }
    elseif ($rule."Expected Value" -eq "-2") {
        $rule."Expected Value" = "Per User Settings Required"
        $currentValue = "Per User Settings Required"
    }

    # Check if registry/command value matches expected value
    if ($currentValue -eq $rule."Expected Value") {
        # Write-Output "Value matches expected value"
        $global:GoodCount += 1
    }
    else {
        # Write-Output "Value does not match expected value"
        $global:BadCount += 1
    }
    AppendToCSV "$($rule.Section),$(IF($rule."Group Policy") {$rule."Group Policy"} Else {$rule.Command}),`"$($rule."Expected Value")`",$(IF($($currentValue)){$currentValue} Else {Write-Output "Not Configured"}),$($rule.Bool),$($rule.Reversed),$($rule.Unit)"
}

function ProcessRules {
    param(
        [object]$rules,
        [switch]$admin
    )
    foreach ($rule in $rules) {
        CheckRule $rule $admin
    }
}

function createFile() {
    param(
        [string]$Path,
        [string]$Value
    )
    New-Item -Path $Path -Value $Value
}

function createFileFromBase64() {
    param(
        [string]$b64File,
        [string]$location
    )
    Write-Output "Creating file from base64"
    Write-Output "$location"
    certutil -decode $b64File $location
}

function deleteFile() {
    param(
        [string]$location
    )
    Remove-Item -Path $location
}

function combineResultsAndReport() {
    ### Set input and output path
    $inputCSV = $global:OutputPath
    $inputXLSX = "$((Get-Item .).FullName)\$($global:Title).xlsm"

    ### Create a new Excel Workbook with one empty sheet
    $excel = New-Object -ComObject excel.application 
    $workbook = $excel.Workbooks.Open($inputXLSX)
    $report = $workbook.Worksheets.Item(1)
    # "$($global:Title)"
    $report.Range("A5").Cells(1, 1).Value = "$($global:Title)"
    $report.Range("A6").Cells(1, 1).Value = "$(Get-Date)"
    $report.Range("A8").Cells(1, 1).Value = "The device complied with $($global:GoodCount) out of $($global:GoodCount+$global:BadCount) rules"
    $worksheet = $workbook.worksheets.add()
    $worksheet.name = "Raw_Results"
    $worksheet.visible = $false

    ### Build the QueryTables.Add command
    ### QueryTables does the same as when clicking "Data » From Text" in Excel
    $TxtConnector = ("TEXT;" + $inputCSV)
    $Connector = $worksheet.QueryTables.add($TxtConnector, $worksheet.Range("A1"))
    $query = $worksheet.QueryTables.item($Connector.name)

    ### Set the delimiter (, or ;) according to your regional settings
    $query.TextFileOtherDelimiter = ','

    ### Set the format to delimited and text for every column
    ### A trick to create an array of 2s is used with the preceding comma
    $query.TextFileParseType = 1
    $query.TextFileColumnDataTypes = , 2 * $worksheet.Cells.Columns.Count
    $query.AdjustColumnWidth = 1

    ### Execute & delete the import query
    $query.Refresh()
    $query.Delete()

    $Range = $worksheet.UsedRange.Cells

    $Table = $worksheet.ListObjects.Add(
        [Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange,
        $worksheet.Range($worksheet.Cells(1, 1), $worksheet.Cells($Range.rows.count, $Range.columns.count)), 
        "Raw_Results",
        [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes
    )

    $Table.name = "Raw_Results"

    ### Save & close the Workbook as XLSX. Change the output extension for Excel 2003
    $Workbook.Save()
    $excel.Quit()
}


$global:OutputPath = "$((Get-Item .).FullName)\$(GetDateString)_results.csv"


if ($title -eq "") {
    $global:Title = "$(GetDateString)_Report"
}
else {
    $global:Title = $title
}


# ----------------------------------------------------------------------------------------------------------------------
#
# CSV Structure:
# Section,Group Policy,Registry Location,Registry Key,Command,Expected Value,Bool,Reversed,Unit,Admin
#
# ----------------------------------------------------------------------------------------------------------------------

# Global variable for rule count
$global:GoodCount = 0
$global:BadCount = 0


$global:Admin = $true
if ($global:Admin) {
    Write-Output "Admin privileges are enabled"
}
else {
    Write-Host -ForegroundColor Red @"
===============================================================
WARNING: Admin privileges are not enabled
===============================================================
"@
}

Start-Sleep -Seconds 2

CreateCSV

$tempCSV = "$pwd\rules.csv"

# Export report template to "$pwd\$($global:Title).xlsm"
# mv report_template to $pwd\$($global:Title).xlsm
Move-Item -Path "$pwd\report_template.xlsm" -Destination "$pwd\$($global:Title).xlsm"

$rules = LoadCSV($tempCSV)

deleteFile $tempCSV

Write-Output "$($rules.length) rules loaded."

ProcessRules $rules

if ($ExecutionContext.SessionState.LanguageMode -eq "FullLanguage") {
    combineResultsAndReport
    deleteFile $global:OutputPath
    Write-Host -ForegroundColor Green "--- Scan Completed, results are found at $pwd\$($global:Title).xlsm ---"
    Read-Host -Prompt "Press Enter to continue"
}
else {
    Write-Output ""
    Write-Output ""
    Write-Host -ForegroundColor Red "--- Actions Required, Unable to automate in current language mode ---"
    Write-Output "1. Open the generated report.xlsm"
    Write-Output "2. Click Enable Content to enable macros"
    Write-Output "3. Click the Data tab"
    Write-Output "4. Click New Query > From File > From CSV"
    Write-Output "5. Import the results.csv"
    Write-Output "6. Click Close & Save"
    Write-Output "7. Rename the generated sheet to Raw_Results"
    Read-Host -Prompt "Press Enter to continue"
}


Write-Output "Scan completed"
