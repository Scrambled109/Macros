param(
    [string]$AcadConsolePath = "C:\Program Files\Autodesk\AutoCAD 2026\accoreconsole.exe"
)

Write-Host "====================================================" -ForegroundColor Cyan
Write-Host "=== STARTING CAD WORKFLOW AUTOMATION ORCHESTRATOR ===" -ForegroundColor Cyan
Write-Host "====================================================" -ForegroundColor Cyan

# --- CONFIGURATION PATHS ---
$CsvPath         = Join-Path $PSScriptRoot "parts.csv"
$LspPath         = Join-Path $PSScriptRoot "ColortoLayer.lsp"
$SeedPath        = Join-Path $PSScriptRoot "SPC_Seed.dwg"

# How long (seconds) to wait for a headless (accoreconsole) job before giving up on it.
$ConsoleTimeoutSec = 180
# ---------------------------

$ArchiveDir = Join-Path (Get-Location).Path "_PROCESSED_DXF_ARCHIVE"
if (-not (Test-Path $ArchiveDir)) { New-Item -ItemType Directory -Path $ArchiveDir -Force | Out-Null }

$LogDir = Join-Path (Get-Location).Path "_ORCHESTRATOR_LOGS"
if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }

# --- HELPERS ------------------------------------------------------------------

function Test-ValidDwg($p) { (Test-Path $p) -and ((Get-Item $p).Length -gt 1024) }

# Normalize a thickness cell to thousandths-of-an-inch (the "500" in 500-HSLA-65).
# Handles plain decimals (0.5 -> 500) AND fractions, because shop travelers list
# stock as "1/2", "3/8", "1-1/2", etc. A raw "1/2" left untouched would put a
# forward slash into the folder/DWG name, and AutoCAD's SAVEAS reads that slash
# as a directory separator -> "path does not exist" and the whole script desyncs.
function Convert-ThicknessToMils($raw) {
    if ([string]::IsNullOrWhiteSpace($raw)) { return $raw }
    $t = $raw.Trim()
    # Pure decimal inches: 0.5, .375  ->  thousandths
    if ($t -match '^\d*\.\d+$') {
        return ([double]$t * 1000).ToString("F0")
    }
    # Fraction, with an optional whole part: 1/2, 3/8, 1-1/2, 1 1/2
    if ($t -match '^(?:(\d+)[-\s])?(\d+)\s*/\s*(\d+)$') {
        $whole = if ($matches[1]) { [double]$matches[1] } else { 0 }
        $num   = [double]$matches[2]
        $den   = [double]$matches[3]
        if ($den -ne 0) {
            return (($whole + ($num / $den)) * 1000).ToString("F0")
        }
    }
    # Anything else (already-formatted "500", odd data): leave as-is.
    return $t
}

# Strip characters that are illegal in Windows filenames or that AutoCAD's
# SAVEAS would misinterpret as path separators. Belt-and-suspenders so no
# stray '/' or ':' from spreadsheet data can ever break the save again.
function Get-SafeName($name) {
    if ($null -eq $name) { return $name }
    return ($name -replace '[\\/:*?"<>|]', '-')
}

# Wait until a file has fully finished being written: it must exist, be a sane
# size, keep the SAME size across two checks, and be openable with NO other
# process holding a lock. This replaces the old fixed Start-Sleep guesses and
# stops us from moving/archiving while AutoCAD is still flushing the DWG.
function Wait-FileStable($path, $timeoutSec = 30) {
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $lastLen = -1
    while ($sw.Elapsed.TotalSeconds -lt $timeoutSec) {
        if (Test-Path $path) {
            $len = (Get-Item $path).Length
            if ($len -gt 1024 -and $len -eq $lastLen) {
                try {
                    $fs = [System.IO.File]::Open($path, 'Open', 'Read', 'None')
                    $fs.Close()
                    return $true   # size settled AND no lock held -> safe to touch
                } catch { }        # still locked; keep waiting
            }
            $lastLen = $len
        }
        Start-Sleep -Milliseconds 400
    }
    return $false
}

# Decide whether a DXF needs the bevel filename marker by looking ONLY at text
# stored in TEXT / MTEXT entities. The old version regex-matched the entire raw
# DXF, so any stray 'K' or 'V' in a style name, block name, or structural field
# added an unnecessary marker. In ASCII DXF every group code sits on its
# own line followed by its value line, so we walk the pairs, track the current
# entity type from group 0, and only test the group 1 / 3 strings of text
# entities against the bevel flags.
function Test-BeveledDxf($path) {
    try { $lines = [System.IO.File]::ReadAllLines($path) } catch { return $false }
    $curType = ''
    $textChunks = [System.Collections.Generic.List[string]]::new()
    for ($i = 0; $i -lt ($lines.Count - 1); $i += 2) {
        $code = $lines[$i].Trim()
        $val  = $lines[$i + 1]
        if ($code -eq '0') {
            if (($curType -eq 'TEXT' -or $curType -eq 'MTEXT') -and (($textChunks -join '') -match '(?i)\b(BEVEL(ED)?|BVL|CHAMFER|SNIP(E|ED|ING)?|BACK\s*GOUGE|GOUGE)\b|\b(K|V|RV)\b|\b(K|V|RV)\s*[-–—:/]?\s*\d+(\.\d+)?\b|(-{1,2}>|<{1,2}-|[←↑→↓↖↗↘↙⇐⇑⇒⇓➔➜➤►◄△▽])')) { return $true }
            $curType = $val.Trim().ToUpper()
            if ($curType -eq 'LEADER' -or $curType -eq 'MLEADER' -or $curType -eq 'MULTILEADER') { return $true }
            $textChunks.Clear()
        } elseif (($code -eq '1' -or $code -eq '3') -and ($curType -eq 'TEXT' -or $curType -eq 'MTEXT')) {
            $textChunks.Add($val)
        }
    }
    return (($curType -eq 'TEXT' -or $curType -eq 'MTEXT') -and (($textChunks -join '') -match '(?i)\b(BEVEL(ED)?|BVL|CHAMFER|SNIP(E|ED|ING)?|BACK\s*GOUGE|GOUGE)\b|\b(K|V|RV)\b|\b(K|V|RV)\s*[-–—:/]?\s*\d+(\.\d+)?\b|(-{1,2}>|<{1,2}-|[←↑→↓↖↗↘↙⇐⇑⇒⇓➔➜➤►◄△▽])'))
}

# --- PRE-FLIGHT VALIDATION ----------------------------------------------------
# Fail loudly and early if the pieces AutoCAD needs aren't where we think they
# are, instead of launching AutoCAD against a missing seed / lisp / exe.
$fatal = $false
if (-not (Test-Path $SeedPath))        { Write-Host "ERROR: Seed DWG not found at $SeedPath" -ForegroundColor Red;        $fatal = $true }
if (-not (Test-Path $LspPath))         { Write-Host "ERROR: ColorToLayer LISP not found at $LspPath" -ForegroundColor Red; $fatal = $true }
if (-not (Test-Path $AcadConsolePath)) { Write-Host "ERROR: accoreconsole.exe not found at $AcadConsolePath" -ForegroundColor Red; $fatal = $true }
if ($fatal) { Write-Host "One or more required paths are missing. Fix the CONFIGURATION PATHS section. Script aborted." -ForegroundColor Red; Exit 1 }

# 1. LOAD AND PARSE SPREADSHEET DATA
# CSV is OPTIONAL (per the README): without it, parts still convert, they just
# land in an "Unsorted" folder with a quantity of 1 instead of being renamed.
$PartLookup = @{}
if (Test-Path $CsvPath) {
    $csvRows = @(Import-Csv -Path $CsvPath)
    $csvHeaders = @($csvRows | Select-Object -First 1 | ForEach-Object { $_.PSObject.Properties.Name })
    $missingHeaders = @('PartNumber', 'Thickness', 'Material') | Where-Object { $_ -notin $csvHeaders }
    if ($missingHeaders.Count -gt 0) {
        Write-Host "ERROR: parts.csv is missing required column(s): $($missingHeaders -join ', ')" -ForegroundColor Red
        Exit 1
    }

    $csvRows | ForEach-Object {
        $part = ([string]$_.PartNumber).Trim()
        if ([string]::IsNullOrWhiteSpace($part)) { return }

        $qtyValue = "1"
        if ($_.Quantity) { $qtyValue = $_.Quantity.Trim() }
        elseif ($_.Quanity) { $qtyValue = $_.Quanity.Trim() }

        $PartLookup[$part] = @{
            Quantity  = $qtyValue
            Thickness = ([string]$_.Thickness).Trim()
            Material  = ([string]$_.Material).Trim()
        }
    }
    Write-Host "Successfully loaded data for $($PartLookup.Count) parts from CSV." -ForegroundColor Green
} else {
    Write-Host "WARNING: CSV not found at $CsvPath. Parts will be converted but left in 'Unsorted' with qty 1." -ForegroundColor Yellow
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

# Running tallies for the end-of-run summary. Script-scoped so the increments
# inside the ForEach-Object pipeline block below actually persist out here.
$script:cleanCount = 0
$script:bevelCount = 0
$script:failCount  = 0

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

        # 3. PRE-SCAN FOR BEVEL INDICATORS (text entities only)
        $isBeveled = Test-BeveledDxf $file.FullName

        # 4. DATA LOOKUP & FILENAME CLEANUP
        $partData = $PartLookup[$workingName]
        if ($partData) {
            $quantity = $partData.Quantity
            $thickStr = Convert-ThicknessToMils $partData.Thickness
            $targetFolderName = "${thickStr}-$($partData.Material)"
        } else {
            $quantity = "1"
            $targetFolderName = "Unsorted"
        }

        # Guarantee the folder/DWG name can never carry a path-breaking char.
        $targetFolderName = Get-SafeName $targetFolderName

        # Safe filename string replacements
        if ($workingName -match '#') { $workingName = $workingName -replace '#', '-' }
        if ($workingName -match '\s+_') { $workingName = $workingName -replace '\s+(_)', '$1' }

        $workingName = $workingName -replace '_\d+-[A-Za-z0-9-]+_\d+$', ''
        $workingName = $workingName -replace '_\d+_\d+$', ''

        # Append the correct target folder and quantity
        $workingName = "${workingName}_${targetFolderName}_$quantity"
        $workingName = Get-SafeName $workingName
        if ($isBeveled) { $workingName = "${workingName}(B)" }

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

        # 6. PROCESS EVERY FILE HEADLESS. Bevel parts follow the same fast path
        # but keep their visible (B) filename marker for downstream identification.
        if ($isBeveled) {
            Write-Host "    [B] Bevel detected. Marking output with (B) and running in Core Console..." -ForegroundColor Magenta
        } else {
            Write-Host "    [+] Standard part. Running in Core Console..." -ForegroundColor Green
        }

        $saveLisp = "(if (findfile `"$escapedDwgPath`") (command `"_.SAVEAS`" `"2018`" `"$escapedDwgPath`" `"Y`") (command `"_.SAVEAS`" `"2018`" `"$escapedDwgPath`"))"
        $ScrContent = $InjectLayers + @(
            $saveLisp,
            "SECURELOAD", "1",
            "FILEDIA", "1",
            "_QUIT", "Y"
        )
        $ScrContent | Out-File -FilePath $ScrPath -Encoding ascii -Force

        # Capture stdout/stderr so a failed headless job is diagnosable. Watch
        # the exit code and enforce a timeout so one wedged file cannot stall
        # the whole batch.
        $logPath = Join-Path $LogDir "$($file.BaseName).log"
        $errPath = Join-Path $LogDir "$($file.BaseName).err.log"
        $proc = Start-Process $AcadConsolePath `
            -ArgumentList "/i `"$($file.FullName)`" /s `"$ScrPath`"" `
            -NoNewWindow -PassThru `
            -RedirectStandardOutput $logPath -RedirectStandardError $errPath

        if (-not $proc.WaitForExit($ConsoleTimeoutSec * 1000)) {
            try { $proc.Kill() } catch { }
            Write-Host "    [X] TIMEOUT after ${ConsoleTimeoutSec}s. Killed accoreconsole. See $logPath. Original DXF kept." -ForegroundColor Red
            $script:failCount++
            continue
        }

        if (($proc.ExitCode -eq 0) -and (Wait-FileStable $finalDwgPath)) {
            Move-Item -Path $file.FullName -Destination (Join-Path $ArchiveDir $file.Name) -Force -ErrorAction SilentlyContinue
            Write-Host "    [SUCCESS] Converted to $([System.IO.Path]::GetFileName($finalDwgPath))" -ForegroundColor Green
            if ($isBeveled) { $script:bevelCount++ } else { $script:cleanCount++ }
        } else {
            Write-Host "    [X] ERROR: exit=$($proc.ExitCode); DWG missing or too small. Original DXF kept. Log: $logPath" -ForegroundColor Red
            $script:failCount++
        }
    }
}

if (Test-Path (Join-Path $env:TEMP "automation_job.scr")) { Remove-Item (Join-Path $env:TEMP "automation_job.scr") -Force }
if (Test-Path $h2dLspPath) { Remove-Item $h2dLspPath -Force }

Write-Host "`n====================================================" -ForegroundColor Cyan
Write-Host "===   ALL PARTS PROCESSED, SORTED, AND SAVED!    ===" -ForegroundColor Cyan
Write-Host ("===   Standard: {0,-3}  Bevel-marked: {1,-3}  Failed: {2,-3} ==" -f $cleanCount, $bevelCount, $failCount) -ForegroundColor Cyan
Write-Host "===   Originals archived in _PROCESSED_DXF_ARCHIVE ==" -ForegroundColor Cyan
if ($failCount -gt 0) { Write-Host "===   Some parts FAILED - see _ORCHESTRATOR_LOGS   ==" -ForegroundColor Red }
Write-Host "====================================================" -ForegroundColor Cyan
