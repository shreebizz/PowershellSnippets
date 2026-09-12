<#
.SYNOPSIS
    Searches a folder recursively for Visual Studio Solution (.sln) files, parses all C# / .NET project files (.csproj), 
    extracts project-to-project references, and generates a report in Excel / CSV format.

.DESCRIPTION
    This script performs the following steps:
    1. Recursively scans the target RootPath for .sln files.
    2. Parses each .sln file to extract associated project files (.csproj, .vbproj, etc.).
    3. Parses each project file XML for <ProjectReference> elements.
    4. Aggregates solution name, project name, referenced project name, and relative/absolute paths.
    5. Exports the results to an Excel (.xlsx) file or CSV (.csv) file.

.PARAMETER RootPath
    The root folder path to search recursively for .sln files. Defaults to the current folder.

.PARAMETER OutputPath
    The output file path for the Excel/CSV report. Defaults to '.\Solution_Project_References.xlsx'.

.PARAMETER ExportFormat
    Desired output format: 'Auto', 'Excel', or 'CSV'. Defaults to 'Auto'.
    - 'Auto': Tries ImportExcel module, then Excel COM Automation, then CSV.
    - 'Excel': Forces .xlsx export using available Excel engine.
    - 'CSV': Exports as CSV file (compatible with Excel).

.PARAMETER IncludeStandaloneProjects
    If specified, also includes .csproj files found in subfolders that are not referenced by any .sln file.

.PARAMETER IncludePackageReferences
    If specified, also includes NuGet <PackageReference> entries in the output report.

.EXAMPLE
    .\Get-ProjectReferenceReport.ps1 -RootPath "C:\Projects\MyRepo" -OutputPath "C:\Reports\ProjectDependencies.xlsx"

.EXAMPLE
    .\Get-ProjectReferenceReport.ps1 -RootPath "." -IncludePackageReferences
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false, Position = 0, ValueFromPipeline = $true)]
    [string]$RootPath = ".",

    [Parameter(Mandatory = $false, Position = 1)]
    [string]$OutputPath = ".\Solution_Project_References.xlsx",

    [Parameter(Mandatory = $false)]
    [ValidateSet("Auto", "Excel", "CSV")]
    [string]$ExportFormat = "Auto",

    [Parameter(Mandatory = $false)]
    [switch]$IncludeStandaloneProjects,

    [Parameter(Mandatory = $false)]
    [switch]$IncludePackageReferences
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Resolve full path of RootPath
if (-not (Test-Path -Path $RootPath)) {
    Write-Error "Root path does not exist: '$RootPath'"
    return
}
$ResolvedRootPath = (Get-Item -Path $RootPath).FullName
Write-Host "=================================================================" -ForegroundColor Cyan
Write-Host " Scanning Root Directory: $ResolvedRootPath" -ForegroundColor Cyan
Write-Host "=================================================================" -ForegroundColor Cyan

# 1. Find all .sln files
$slnFiles = @(Get-ChildItem -Path $ResolvedRootPath -Filter "*.sln" -Recurse -File -ErrorAction SilentlyContinue)

Write-Host "Found $($slnFiles.Count) solution (.sln) file(s)." -ForegroundColor Green

$reportRows = [System.Collections.Generic.List[PSCustomObject]]::new()
$processedProjectPaths = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

# Helper function to extract ProjectReferences from a .csproj XML file
function Get-ProjectReferencesFromFile {
    param (
        [string]$ProjectPath,
        [string]$SolutionName,
        [string]$SolutionPath
    )

    $projectName = [System.IO.Path]::GetFileNameWithoutExtension($ProjectPath)
    $projectDir = [System.IO.Path]::GetDirectoryName($ProjectPath)

    if (-not (Test-Path -Path $ProjectPath)) {
        Write-Warning "Project file not found on disk: '$ProjectPath'"
        $reportRows.Add([PSCustomObject]@{
            "Solution Name"               = $SolutionName
            "Solution Path"               = $SolutionPath
            "Project Name"                = $projectName
            "Project Path"                = $ProjectPath
            "Referenced Project Name"     = "(File Not Found)"
            "Referenced Project Path"     = ""
            "Reference Type"              = "Error"
        })
        return
    }

    try {
        [xml]$xml = Get-Content -Path $ProjectPath -Raw -ErrorAction Stop
    }
    catch {
        Write-Warning "Failed to parse XML for project '$ProjectPath': $_"
        $reportRows.Add([PSCustomObject]@{
            "Solution Name"               = $SolutionName
            "Solution Path"               = $SolutionPath
            "Project Name"                = $projectName
            "Project Path"                = $ProjectPath
            "Referenced Project Name"     = "(Invalid XML)"
            "Referenced Project Path"     = ""
            "Reference Type"              = "Error"
        })
        return
    }

    $refCount = 0

    # Process <ProjectReference> nodes
    $projRefNodes = $xml.SelectNodes("//*[local-name()='ProjectReference']")
    if ($projRefNodes) {
        foreach ($node in $projRefNodes) {
            $incAttr = $node.GetAttribute("Include")
            if ([string]::IsNullOrWhiteSpace($incAttr)) { continue }

            # Replace MSBuild properties if any basic ones exist
            $cleanIncAttr = $incAttr -replace '\$\(SolutionDir\)', "$([System.IO.Path]::GetDirectoryName($SolutionPath))\"
            
            # Resolve full target path
            $refFullPath = ""
            try {
                $refFullPath = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($projectDir, $cleanIncAttr))
            }
            catch {
                $refFullPath = $cleanIncAttr
            }

            # Referenced project name
            $refProjName = ""
            $nameNode = $node.SelectSingleNode("*[local-name()='Name']")
            if ($nameNode -and -not [string]::IsNullOrWhiteSpace($nameNode.InnerText)) {
                $refProjName = $nameNode.InnerText.Trim()
            }
            else {
                $refProjName = [System.IO.Path]::GetFileNameWithoutExtension($cleanIncAttr)
            }

            $reportRows.Add([PSCustomObject]@{
                "Solution Name"           = $SolutionName
                "Solution Path"           = $SolutionPath
                "Project Name"            = $projectName
                "Project Path"            = $ProjectPath
                "Referenced Project Name" = $refProjName
                "Referenced Project Path" = $refFullPath
                "Reference Type"          = "Project Reference"
            })
            $refCount++
        }
    }

    # Process <PackageReference> nodes if enabled
    if ($IncludePackageReferences) {
        $pkgRefNodes = $xml.SelectNodes("//*[local-name()='PackageReference']")
        if ($pkgRefNodes) {
            foreach ($node in $pkgRefNodes) {
                $pkgName = $node.GetAttribute("Include")
                if ([string]::IsNullOrWhiteSpace($pkgName)) {
                    $pkgNameNode = $node.SelectSingleNode("*[local-name()='Include']")
                    if ($pkgNameNode) { $pkgName = $pkgNameNode.InnerText }
                }
                $pkgVersion = $node.GetAttribute("Version")
                if ([string]::IsNullOrWhiteSpace($pkgVersion)) {
                    $verNode = $node.SelectSingleNode("*[local-name()='Version']")
                    if ($verNode) { $pkgVersion = $verNode.InnerText }
                }

                if (-not [string]::IsNullOrWhiteSpace($pkgName)) {
                    $reportRows.Add([PSCustomObject]@{
                        "Solution Name"           = $SolutionName
                        "Solution Path"           = $SolutionPath
                        "Project Name"            = $projectName
                        "Project Path"            = $ProjectPath
                        "Referenced Project Name" = "$pkgName (v$pkgVersion)"
                        "Referenced Project Path" = "NuGet Package"
                        "Reference Type"          = "Package Reference"
                    })
                    $refCount++
                }
            }
        }
    }

    # If no references were found, add a summary row so project is still visible
    if ($refCount -eq 0) {
        $reportRows.Add([PSCustomObject]@{
            "Solution Name"           = $SolutionName
            "Solution Path"           = $SolutionPath
            "Project Name"            = $projectName
            "Project Path"            = $ProjectPath
            "Referenced Project Name" = "(None)"
            "Referenced Project Path" = ""
            "Reference Type"          = "No References"
        })
    }
}

# 2. Iterate each .sln file and extract projects
foreach ($sln in $slnFiles) {
    $slnName = $sln.BaseName
    $slnPath = $sln.FullName
    $slnDir = $sln.DirectoryName

    Write-Host "Analyzing Solution: $slnName ($($sln.FullName))" -ForegroundColor Yellow

    $slnContent = Get-Content -Path $slnPath -ErrorAction SilentlyContinue
    
    # Regex pattern to match Project entries in .sln file:
    # Project("{...}") = "ProjectName", "RelativePath.csproj", "{...}"
    $projectRegex = 'Project\("\{[^"]+\}"\)\s*=\s*"([^"]+)"\s*,\s*"([^"]+)"\s*,\s*"\{[^"]+\}"'

    $slnProjectCount = 0
    foreach ($line in $slnContent) {
        if ($line -match $projectRegex) {
            $projNameInSln = $Matches[1]
            $projRelPath = $Matches[2]

            # Filter out solution folders (which don't end in project file extensions like .csproj, .vbproj, .fsproj, .vcxproj)
            $ext = [System.IO.Path]::GetExtension($projRelPath).ToLower()
            if ($ext -in @('.csproj', '.vbproj', '.fsproj', '.vcxproj', '.dbproj', '.proj')) {
                $projFullPath = ""
                try {
                    $projFullPath = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($slnDir, $projRelPath))
                }
                catch {
                    $projFullPath = $projRelPath
                }

                $null = $processedProjectPaths.Add($projFullPath)
                $slnProjectCount++

                Write-Host "  -> Found Project: $projNameInSln" -ForegroundColor Gray
                Get-ProjectReferencesFromFile -ProjectPath $projFullPath -SolutionName $slnName -SolutionPath $slnPath
            }
        }
    }

    if ($slnProjectCount -eq 0) {
        Write-Host "  (No project files matched inside solution)" -ForegroundColor DarkGray
    }
}

# 3. Optionally process standalone projects not in any .sln
if ($IncludeStandaloneProjects) {
    Write-Host "`nScanning for standalone project files not listed in any solution..." -ForegroundColor Yellow
    $allProjFiles = @(Get-ChildItem -Path "$ResolvedRootPath\*" -Include "*.csproj", "*.vbproj", "*.fsproj" -Recurse -File -ErrorAction SilentlyContinue)

    foreach ($proj in $allProjFiles) {
        if (-not $processedProjectPaths.Contains($proj.FullName)) {
            Write-Host "  -> Found Standalone Project: $($proj.Name)" -ForegroundColor Gray
            Get-ProjectReferencesFromFile -ProjectPath $proj.FullName -SolutionName "(Standalone / No Solution)" -SolutionPath ""
        }
    }
}

Write-Host "`nTotal Report Records Generated: $($reportRows.Count)" -ForegroundColor Green

# 4. Export logic
if ($reportRows.Count -eq 0) {
    Write-Warning "No solution or project reference data was found."
    return
}

# Determine export method based on OutputPath and ExportFormat
$extension = [System.IO.Path]::GetExtension($OutputPath).ToLower()
if ([string]::IsNullOrWhiteSpace($extension)) {
    $OutputPath += ".xlsx"
    $extension = ".xlsx"
}

$ResolvedOutputPath = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine((Get-Location).Path, $OutputPath))
$outputDir = [System.IO.Path]::GetDirectoryName($ResolvedOutputPath)
if (-not (Test-Path -Path $outputDir)) {
    $null = New-Item -ItemType Directory -Path $outputDir -Force
}

$exportedSuccessfully = $false

# Helper to export using ImportExcel module
function Export-ViaImportExcel {
    param ($Rows, $Path)
    if (Get-Module -ListAvailable -Name ImportExcel) {
        Import-Module ImportExcel -ErrorAction SilentlyContinue
        $Rows | Export-Excel -Path $Path -WorksheetName "Project References" -AutoSize -BoldTopHeader -TableStyle Medium2 -Show:$false
        return $true
    }
    return $false
}

# Helper to export using Excel COM automation
function Export-ViaExcelCOM {
    param ($Rows, $Path)
    try {
        $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)
        $worksheet.Name = "Project References"

        # Write Headers
        $properties = $Rows[0].psobject.Properties.Name
        for ($c = 0; $c -lt $properties.Count; $c++) {
            $cell = $worksheet.Cells.Item(1, $c + 1)
            $cell.Value2 = $properties[$c]
            $cell.Font.Bold = $true
            $cell.Interior.ColorIndex = 15 # Light Gray header
        }

        # Write Rows
        for ($r = 0; $r -lt $Rows.Count; $r++) {
            $rowObj = $Rows[$r]
            for ($c = 0; $c -lt $properties.Count; $c++) {
                $propName = $properties[$c]
                $val = $rowObj.$propName
                $worksheet.Cells.Item($r + 2, $c + 1).Value2 = [string]$val
            }
        }

        # Auto-fit columns
        $usedRange = $worksheet.UsedRange
        $null = $usedRange.Columns.AutoFit()

        # Save workbook
        if (Test-Path -Path $Path) { Remove-Item -Path $Path -Force }
        $workbook.SaveAs($Path, 51) # 51 = xlOpenXMLWorkbook (.xlsx)
        $workbook.Close($false)
        $excel.Quit()

        # Release COM objects
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()

        return $true
    }
    catch {
        Write-Verbose "COM Excel Automation unavailable or failed: $_"
        return $false
    }
}

# Helper to export CSV
function Export-ViaCSV {
    param ($Rows, $Path)
    $csvPath = [System.IO.Path]::ChangeExtension($Path, ".csv")
    $Rows | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Host "Report exported to Excel-compatible CSV: $csvPath" -ForegroundColor Cyan
    return $true
}

# Execute export based on preferences
if ($ExportFormat -eq "CSV" -or $extension -eq ".csv") {
    $exportedSuccessfully = Export-ViaCSV -Rows $reportRows -Path $ResolvedOutputPath
}
else {
    # Try ImportExcel module first
    Write-Host "Attempting Excel export via ImportExcel module..." -ForegroundColor DarkGray
    if (Export-ViaImportExcel -Rows $reportRows -Path $ResolvedOutputPath) {
        Write-Host "Report exported to Excel (.xlsx) using ImportExcel module." -ForegroundColor Green
        $exportedSuccessfully = $true
    }
    else {
        # Try COM object next
        Write-Host "Attempting Excel export via Excel COM Object..." -ForegroundColor DarkGray
        if (Export-ViaExcelCOM -Rows $reportRows -Path $ResolvedOutputPath) {
            Write-Host "Report exported to Excel (.xlsx) using Excel COM Automation." -ForegroundColor Green
            $exportedSuccessfully = $true
        }
        else {
            # Fallback to CSV
            Write-Host "Excel modules/COM not detected. Falling back to CSV export..." -ForegroundColor Yellow
            $exportedSuccessfully = Export-ViaCSV -Rows $reportRows -Path $ResolvedOutputPath
        }
    }
}

Write-Host "`nProcess Completed Successfully!" -ForegroundColor Green
