Write-Host "====================================================" -ForegroundColor Cyan
Write-Host "=== STARTING CAD WORKFLOW AUTOMATION ORCHESTRATOR ===" -ForegroundColor Cyan
Write-Host "====================================================" -ForegroundColor Cyan

# --- CONFIGURATION PATHS ---
$CsvPath         = "C:\Users\pbowen\Downloads\TEST_ENVIORMENT_AUTO_SCRIPT\parts.csv"
$LspPath         = "U:\Engineering\CAD Services\Engineering Reference\CLBO_ps1_DIRECTORY\ColorToLayer.lsp"
$SeedPath        = "U:\Engineering\CAD Services\Engineering Reference\CLBO_ps1_DIRECTORY\SPC_Seed.dwg"
$AcadConsolePath = "C:\Program Files\Autodesk\AutoCAD 2026\accoreconsole.exe"
$AcadGuiPath     = "C:\Program Files\Autodesk\AutoCAD 2026\acad.exe"
# ---------------------------

$ArchiveDir = Join-Path (Get-Location).Path "_PROCESSED_DXF_ARCHIVE"
if (-not (Test-Path $ArchiveDir)) { New-Item -ItemType Directory -Path $ArchiveDir -Force | Out-Null }

function Test-ValidDwg($p) { (Test-Path $p) -and ((Get-Item $p).Length -gt 1024) }

if (-not (Test-Path $SeedPath)) {
    Write-Host "ERROR: Seed DWG not found at $SeedPath. Build it first. Script aborted." -ForegroundColor Red
    Exit
}

# 1. LOAD AND PARSE SPREADSHEET DATA
$PartLookup = @{}
if (Test-Path $CsvPath) {
    Import-Csv -Path $CsvPath | ForEach-Object {
        $part = $_.PartNumber.Trim()

        $qtyValue = "1"
        if ($_.Quantity) { $qtyValue = $_.Quantity.Trim() }
        elseif ($_.Quanity) { $qtyValue = $_.Quanity.Trim() }

        $PartLookup[$part] = @{
            Quantity  = $qtyValue
            Thickness = $_.Thickness.Trim()
            Material  = $_.Material.Trim()
        }
    }
    Write-Host "Successfully loaded data for $($PartLookup.Count) parts from CSV." -ForegroundColor Green
} else {
    Write-Host "ERROR: CSV file not found at $CsvPath. Script aborted." -ForegroundColor Red
    Exit
}

$escapedLspPath  = $LspPath.Replace('\', '/')
$escapedSeedPath = $SeedPath.Replace('\', '/')

# GENERATE SILENT HASHTAG-TO-DASH REPLACER
$h2dLspPath = Join-Path $env:TEMP "HashToDash.lsp"
$escapedH2dPath = $h2dLspPath.Replace('\', '/')
$h2dLspContent = @"
(defun c:HashToDash ( / ss i ent oldStr newStr)
  (if (setq ss (ssget "X" '((0 . "TEXT,MTEXT"))))
    (progn
      (setq i 0)
      (while (< i (sslength ss))
        (setq ent (entget (ssname ss i)))
        (setq oldStr (cdr (assoc 1 ent)))
        (setq newStr oldStr)
        (while (vl-string-search "#" newStr)
          (setq newStr (vl-string-subst "-" "#" newStr))
        )
        (if (/= oldStr newStr)
          (entmod (subst (cons 1 newStr) (assoc 1 ent) ent))
        )
        (setq i (1+ i))
      )
    )
  )
  (princ)
)
"@
$h2dLspContent | Out-File -FilePath $h2dLspPath -Encoding ascii -Force


# 2. SCAN NUMBERED DIRECTORIES
Get-ChildItem -Directory | Where-Object { $_.Name -match '^(\d+)' -and $_.Name -notmatch '^\d+-[a-zA-Z]' } | ForEach-Object {
    $folder = $_
    Write-Host "`nScanning Folder: $($folder.Name)" -ForegroundColor Yellow

    $files = Get-ChildItem -Path $folder.FullName -Filter "*.dxf" -File
    if ($files.Count -eq 0) {
        Write-Host "  -> No DXF files found in this folder." -ForegroundColor DarkGray
        return
    }

    foreach ($file in $files) {
        $originalName = $file.Name
        $workingName  = $file.BaseName

        Write-Host "  Processing: $originalName..." -ForegroundColor White

        # 3. PRE-SCAN FOR BEVEL INDICATORS
        $fileContent = [System.IO.File]::ReadAllText($file.FullName)
        $isBeveled = $fileContent -match '(?i)\b([KV]|BEVEL)\b'

        # 4. DATA LOOKUP & FILENAME CLEANUP
        $partData = $PartLookup[$workingName]
        if ($partData) {
            $quantity = $partData.Quantity
            if ($partData.Thickness -match '^\d*\.\d+$') {
                $thickStr = ([double]$partData.Thickness * 1000).ToString("F0")
            } else {
                $thickStr = $partData.Thickness
            }
            $targetFolderName = "${thickStr}-$($partData.Material)"
        } else {
            $quantity = "1"
            $targetFolderName = "Unsorted"
        }

        # Safe filename string replacements
        if ($workingName -match '#') { $workingName = $workingName -replace '#', '-' }
        if ($workingName -match '\s+_') { $workingName = $workingName -replace '\s+(_)', '$1' }

        $workingName = $workingName -replace '_\d+-[A-Za-z0-9-]+_\d+$', '' 
        $workingName = $workingName -replace '_\d+_\d+$', ''               

        # Append the correct target folder and quantity
        $workingName = "${workingName}_${targetFolderName}_$quantity"

        $targetFolderPath = Join-Path (Get-Location).Path $targetFolderName
        if (-not (Test-Path $targetFolderPath)) {
            New-Item -ItemType Directory -Path $targetFolderPath -Force | Out-Null
        }
        $finalDwgPath  = Join-Path $targetFolderPath "${workingName}.dwg"
        $escapedDwgPath = $finalDwgPath.Replace('\', '/')

        $ScrPath = Join-Path $env:TEMP "automation_job.scr"

        # 5. SHARED PREAMBLE (Layers & Lisp Imports)
        $InjectLayers = @(
            "FILEDIA", "0",
            "SECURELOAD", "0",
            "-INSERT", "`"$escapedSeedPath`"", "0,0", "1", "1", "0",
            "ERASE", "L", "",
            "(load `"$escapedLspPath`")",
            "ColorToLayer",
            "(load `"$escapedH2dPath`")",
            "HashToDash"
        )

        # 6. DECISION ENGINE (GATEKEEPER)
        if ($isBeveled) {
            Write-Host "    [!] BEVEL DETECTED. Loading into AutoCAD for manual check..." -ForegroundColor Magenta

            # custom FINISH command changed to _.CLOSE instead of _.QUIT
            $finishLisp = "(defun c:FINISH () (setvar `"FILEDIA`" 0) (if (findfile `"$escapedDwgPath`") (command `"_.SAVEAS`" `"2018`" `"$escapedDwgPath`" `"Y`") (command `"_.SAVEAS`" `"2018`" `"$escapedDwgPath`")) (command `"_.CLOSE`") (princ))"

            $ScrContent = $InjectLayers + @(
                $finishLisp,
                "SECURELOAD", "1",
                "FILEDIA", "1",
                "_ALERT", "`"Review the part. When finished, type FINISH in the command line and press Enter to automatically save and sort.`""
            )
            $ScrContent | Out-File -FilePath $ScrPath -Encoding ascii -Force

            # Removed -Wait and .WaitForExit() methods so PowerShell can move to the file watcher
            Start-Process $AcadGuiPath -ArgumentList "`"$($file.FullName)`" /b `"$ScrPath`"" 

            Write-Host "    --> Waiting for you to type FINISH in AutoCAD..." -ForegroundColor Cyan
            
            # THE FILE WATCHER: Loops silently until the DWG physically exists
            while (-not (Test-ValidDwg $finalDwgPath)) {
                Start-Sleep -Milliseconds 500
            }

            # Give AutoCAD a half-second to fully release the file lock after closing the tab
            Start-Sleep -Milliseconds 500

            Move-Item -Path $file.FullName -Destination (Join-Path $ArchiveDir $file.Name) -Force -ErrorAction SilentlyContinue
            Write-Host "    [SUCCESS] Saved and sorted to $targetFolderName" -ForegroundColor Green
            
        } else {
            Write-Host "    [+] Clean Part. Running in background core console..." -ForegroundColor Green

            $ScrContent = $InjectLayers + @(
                "_SAVEAS", "2018", "`"$escapedDwgPath`"",
                "SECURELOAD", "1",
                "FILEDIA", "1",
                "_QUIT", "Y"
            )
            $ScrContent | Out-File -FilePath $ScrPath -Encoding ascii -Force

            Start-Process $AcadConsolePath -ArgumentList "/i `"$($file.FullName)`" /s `"$ScrPath`"" -Wait -NoNewWindow
            Start-Sleep -Seconds 1

            if (Test-ValidDwg $finalDwgPath) {
                Move-Item -Path $file.FullName -Destination (Join-Path $ArchiveDir $file.Name) -Force -ErrorAction SilentlyContinue
            } else {
                Write-Host "    [X] ERROR: DWG missing or too small. Original DXF kept." -ForegroundColor Red
            }
        }
    }
}

if (Test-Path (Join-Path $env:TEMP "automation_job.scr")) { Remove-Item (Join-Path $env:TEMP "automation_job.scr") -Force }
if (Test-Path $h2dLspPath) { Remove-Item $h2dLspPath -Force }

Write-Host "`n====================================================" -ForegroundColor Cyan
Write-Host "===   ALL PARTS PROCESSED, SORTED, AND SAVED!    ===" -ForegroundColor Cyan
Write-Host "===   Originals archived in _PROCESSED_DXF_ARCHIVE ==" -ForegroundColor Cyan
Write-Host "====================================================" -ForegroundColor Cyan