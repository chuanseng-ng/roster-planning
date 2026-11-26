# Builds dist\PHRosterPlanner.xlsm by driving desktop Excel over COM.
# Run:  powershell -ExecutionPolicy Bypass -File build\build.ps1
# Exit codes: 0 self-test PASS, 1 precondition failed, 2 self-test FAIL,
#             3 Excel could not run the self-test, 4 unexpected build error.

$ErrorActionPreference = 'Stop'

$repoRoot   = Split-Path -Parent $PSScriptRoot
$vbaDir     = Join-Path $repoRoot 'src\vba'
$fixDir     = Join-Path $repoRoot 'fixtures\normal_year'
$testCsv    = Join-Path $repoRoot 'fixtures\testsheet.csv'
$distDir    = Join-Path $repoRoot 'dist'
$outPath    = Join-Path $distDir 'PHRosterPlanner.xlsm'
$invariant  = [Globalization.CultureInfo]::InvariantCulture

# ---------------------------------------------------------------- step 1
$vbomOk = $true
try {
    $prop = Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security' -Name 'AccessVBOM' -ErrorAction Stop
    if ($prop.AccessVBOM -ne 1) { $vbomOk = $false }
}
catch {
    Write-Host 'WARNING: could not read the AccessVBOM setting. Continuing; the module import will fail with a clear error if access is not trusted.'
}
if (-not $vbomOk) {
    Write-Host 'Excel does not trust access to the VBA project object model.'
    Write-Host 'Enable it: open Excel, File > Options > Trust Center > Trust Center Settings >'
    Write-Host 'Macro Settings, tick "Trust access to the VBA project object model", click OK,'
    Write-Host 'close Excel, then run this script again.'
    exit 1
}

# ---------------------------------------------------------------- helpers
function Get-CsvWidth([string]$line) {
    $n = 1
    $inQ = $false
    foreach ($ch in $line.ToCharArray()) {
        if ($ch -eq '"') { $inQ = -not $inQ }
        elseif (($ch -eq ',') -and (-not $inQ)) { $n++ }
    }
    return $n
}

# Returns @{ Width = n; Rows = array of string[] } for every line of the CSV
# (the header line included). Real CSV parsing; UTF-8.
function Read-CsvRows([string]$path) {
    $lines = @(Get-Content -Path $path -Encoding UTF8)
    $first = ''
    foreach ($l in $lines) {
        if ($l.Trim().Length -gt 0) { $first = $l; break }
    }
    $width = Get-CsvWidth $first
    $names = @()
    for ($i = 1; $i -le $width; $i++) { $names += ('C' + $i) }
    $objs = @($lines | ConvertFrom-Csv -Header $names)
    $rows = New-Object System.Collections.ArrayList
    foreach ($o in $objs) {
        $arr = New-Object 'string[]' $width
        for ($i = 1; $i -le $width; $i++) {
            $v = $o.('C' + $i)
            if ($null -eq $v) { $v = '' }
            $arr[$i - 1] = [string]$v
        }
        [void]$rows.Add($arr)
    }
    return @{ Width = $width; Rows = $rows }
}

# Sets a COM property through IDispatch directly. PowerShell's own COM binder remembers the
# type of the first value assigned to a property name (for example text) and then refuses
# other types (numbers, arrays) with a cast error; this call avoids that binder.
function Set-ComProp($obj, [string]$name, $value) {
    $comArgs = New-Object 'object[]' 1
    $comArgs[0] = $value
    [void]$obj.GetType().InvokeMember($name, [System.Reflection.BindingFlags]::SetProperty, $null, $obj, $comArgs)
}

# Writes a 2-D block cell by cell, top-left at ($row1, $col1). Empty entries are skipped.
# (Assigning the whole array to Range.Value2 fails on some PowerShell/Excel combinations.)
function Write-Block($ws, [int]$row1, [int]$col1, $data, [int]$rows, [int]$cols) {
    for ($r = 0; $r -lt $rows; $r++) {
        for ($c = 0; $c -lt $cols; $c++) {
            $v = $data[$r, $c]
            if ($null -ne $v) {
                $cell = $ws.Cells.Item($row1 + $r, $col1 + $c)
                Set-ComProp $cell 'Value2' $v
            }
        }
    }
}

function Col-Letter([int]$n) {
    return [string][char](64 + $n)
}

function Write-Header($ws, [string[]]$names) {
    for ($c = 1; $c -le $names.Length; $c++) {
        $ws.Cells.Item(1, $c).NumberFormat = '@'
        $ws.Cells.Item(1, $c).Value2 = $names[$c - 1]
    }
    $addr = 'A1:' + (Col-Letter $names.Length) + '1'
    $ws.Range($addr).Font.Bold = $true
}

# Loads a headed CSV into the sheet from row 2. $dateCols are 1-based column numbers.
function Load-Table($ws, [string]$csvPath, [int]$width, [int[]]$dateCols) {
    $csv = Read-CsvRows $csvPath
    $count = $csv.Rows.Count - 1
    for ($c = 1; $c -le $width; $c++) {
        if ($dateCols -contains $c) { $fmt = 'yyyy-mm-dd' } else { $fmt = '@' }
        $ws.Columns.Item($c).NumberFormat = $fmt
    }
    if ($count -lt 1) { return }
    $data = New-Object 'object[,]' $count, $width
    for ($r = 0; $r -lt $count; $r++) {
        $src = $csv.Rows.Item($r + 1)
        for ($c = 1; $c -le $width; $c++) {
            $t = $src[$c - 1]
            if ($dateCols -contains $c) {
                if ($t.Trim().Length -gt 0) {
                    $data[$r, ($c - 1)] = [datetime]::ParseExact($t.Trim(), 'yyyy-MM-dd', $invariant).ToOADate()
                }
            }
            else {
                if ($t.Length -gt 0) { $data[$r, ($c - 1)] = $t }
            }
        }
    }
    $addr = 'A2:' + (Col-Letter $width) + [string]($count + 1)
    Write-Block $ws 2 1 $data $count $width
}

function Add-Button($ws, [string]$cell, [double]$width, [double]$height, [string]$caption, [string]$macro) {
    $anchor = $ws.Range($cell)
    $shp = $ws.Shapes.AddFormControl(0, $anchor.Left, $anchor.Top, $width, $height)
    $shp.OnAction = $macro
    $shp.Placement = 3
    $tf = $shp.TextFrame
    $ch = $tf.Characters()
    $ch.Text = $caption
    Release-Com $ch
    Release-Com $tf
    Release-Com $shp
    Release-Com $anchor
}

function Release-Com($obj) {
    if ($null -ne $obj) {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($obj) } catch { }
    }
}

# ---------------------------------------------------------------- step 2
$exitCode = 0
$xl = $null
$wb = $null
$sheets = $null
$vbp = $null
$vbcomps = $null
$sheetObjs = @{}
$sheetNames = @('Settings', 'Staff', 'Holidays', 'LastYear', 'Leave', 'Roster', 'Review', 'Summary', 'Approvals', 'Test')

try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.EnableEvents = $false

    # ------------------------------------------------------------ step 3
    $wbs = $xl.Workbooks
    $wb = $wbs.Add()
    Release-Com $wbs
    $sheets = $wb.Worksheets
    # remove leftover default sheets first so the new ones follow in order
    while ($sheets.Count -gt 1) {
        $extra = $sheets.Item($sheets.Count)
        $extra.Delete()
        Release-Com $extra
    }
    $first = $sheets.Item(1)
    $first.Name = 'Settings'
    $sheetObjs['Settings'] = $first
    for ($i = 1; $i -lt $sheetNames.Length; $i++) {
        $last = $sheets.Item($sheets.Count)
        $new = $sheets.Add([Type]::Missing, $last)
        $new.Name = $sheetNames[$i]
        $sheetObjs[$sheetNames[$i]] = $new
        Release-Com $last
    }

    # ------------------------------------------------------------ step 4
    Write-Header $sheetObjs['Settings']  @('Key', 'Value')
    Write-Header $sheetObjs['Staff']     @('Name', 'Title', 'Senior', 'Role', 'Cannot do')
    Write-Header $sheetObjs['Holidays']  @('Holiday name', 'Date', 'Added')
    Write-Header $sheetObjs['LastYear']  @('Holiday name', 'Date', 'Staff name', 'Slot')
    Write-Header $sheetObjs['Leave']     @('Staff name', 'From date', 'To date', 'Shift')
    Write-Header $sheetObjs['Roster']    @('Holiday', 'Date', 'Day 1', 'Day 2', 'Evening', 'Night 1', 'Night 2', 'Standby D/E', 'Standby N')
    Write-Header $sheetObjs['Review']    @('Date', 'Group', 'Slot', 'Rule not met', 'Reason', 'Candidates', 'Approve')
    Write-Header $sheetObjs['Summary']   @('Staff name', 'Active slots', 'Eligible dates', 'Ratio', 'Standby slots')
    Write-Header $sheetObjs['Approvals'] @('Date', 'Slot', 'Staff name', 'Rules broken')

    # Settings rows 2-13: values are written as text (the VBA reads them with CStr)
    $ws = $sheetObjs['Settings']
    $ws.Columns.Item(1).NumberFormat = '@'
    $ws.Columns.Item(2).NumberFormat = '@'
    $setCsv = Read-CsvRows (Join-Path $fixDir 'settings.csv')
    $nSet = $setCsv.Rows.Count - 1
    if ($nSet -ne 12) { throw ('settings.csv should hold 12 keys but holds ' + $nSet) }
    $setData = New-Object 'object[,]' $nSet, 2
    for ($r = 0; $r -lt $nSet; $r++) {
        $src = $setCsv.Rows.Item($r + 1)
        $setData[$r, 0] = $src[0]
        $setData[$r, 1] = $src[1]
    }
    Write-Block $ws 2 1 $setData $nSet 2

    # ------------------------------------------------------------ step 5
    Load-Table $sheetObjs['Staff']    (Join-Path $fixDir 'staff.csv')    5 @()
    Load-Table $sheetObjs['Holidays'] (Join-Path $fixDir 'holidays.csv') 3 @(2)
    Load-Table $sheetObjs['LastYear'] (Join-Path $fixDir 'lastyear.csv') 4 @(2)
    Load-Table $sheetObjs['Leave']    (Join-Path $fixDir 'leave.csv')    4 @(2, 3)
    $sheetObjs['Holidays'].Cells.Item(1, 5).NumberFormat = '@'
    $sheetObjs['Holidays'].Cells.Item(1, 5).Value2 = 'Sample dates - verify against MOM'

    # ------------------------------------------------------------ step 6
    $ws = $sheetObjs['Test']
    $ws.Cells.NumberFormat = '@'
    $tcsv = Read-CsvRows $testCsv
    $tRows = $tcsv.Rows.Count
    $tWidth = $tcsv.Width
    if ($tRows -gt 0) {
        $tData = New-Object 'object[,]' $tRows, $tWidth
        for ($r = 0; $r -lt $tRows; $r++) {
            $src = $tcsv.Rows.Item($r)
            for ($c = 0; $c -lt $tWidth; $c++) {
                if ($src[$c].Length -gt 0) { $tData[$r, $c] = $src[$c] }
            }
        }
        $tAddr = 'A1:' + (Col-Letter $tWidth) + [string]$tRows
        Write-Host 'Loading the Test sheet (this can take a minute)...'
        Write-Block $ws 1 1 $tData $tRows $tWidth
    }

    # ------------------------------------------------------------ step 7
    $vbp = $wb.VBProject
    $vbcomps = $vbp.VBComponents
    $basFiles = @(Get-ChildItem -Path $vbaDir -Filter '*.bas' | Sort-Object Name)
    foreach ($f in $basFiles) {
        try {
            $imported = $vbcomps.Import($f.FullName)
            Release-Com $imported
            Write-Host ('Imported ' + $f.Name)
        }
        catch {
            Write-Host ('FAILED to import ' + $f.FullName)
            Write-Host 'If the message mentions programmatic access, enable "Trust access to the VBA project object model" (Trust Center > Macro Settings).'
            throw
        }
    }

    # ------------------------------------------------------------ step 9 (before buttons so button positions are final)
    foreach ($name in $sheetNames) {
        if (($name -eq 'Approvals') -or ($name -eq 'Test')) { continue }
        $ws = $sheetObjs[$name]
        $ws.Activate()
        $win = $null
        try {
            $win = $xl.ActiveWindow
            $win.FreezePanes = $false
            $win.SplitColumn = 0
            $win.SplitRow = 1
            $win.FreezePanes = $true
        }
        catch {
            Write-Host ('Note: could not freeze the header row on sheet ' + $name + ': ' + $_.Exception.Message)
        }
        Release-Com $win
        $used = $ws.UsedRange
        $cols = $used.Columns
        [void]$cols.AutoFit()
        Release-Com $cols
        Release-Com $used
    }

    # ------------------------------------------------------------ step 8
    $ws = $sheetObjs['Settings']
    Add-Button $ws 'D2' 130 24 'Generate roster'  'modMain.GenerateRoster'
    Add-Button $ws 'D4' 130 24 'Start new year'   'modMain.StartNewYear'
    Add-Button $ws 'D6' 130 24 'Run self-test'    'modMain.RunSelfTest'
    Add-Button $sheetObjs['Holidays'] 'E3' 150 24 'Expand long weekends' 'modMain.ExpandLongWeekends'
    Add-Button $sheetObjs['Review']   'K1' 130 24 'Apply approvals'      'modMain.ApplyApprovals'

    # ------------------------------------------------------------ step 9 (hide)
    $sheetObjs['Settings'].Activate()
    Set-ComProp $sheetObjs['Approvals'] 'Visible' ([int]0)
    Set-ComProp $sheetObjs['Test'] 'Visible' ([int]0)

    # ------------------------------------------------------------ step 10
    if (-not (Test-Path -Path $distDir)) {
        [void](New-Item -ItemType Directory -Path $distDir)
    }
    if (Test-Path -Path $outPath) {
        Remove-Item -Path $outPath -Force
    }
    $wb.SaveAs($outPath, 52)
    Write-Host ('Saved ' + $outPath)

    # ------------------------------------------------------------ step 11
    $result = $null
    $xl.Visible = $true
    try {
        $result = $xl.Run('modSelfTest.SelfTestSilent')
    }
    catch {
        Write-Host ('The self-test could not run: ' + $_.Exception.Message)
        Write-Host 'Open dist\PHRosterPlanner.xlsm, press Alt+F11, then Debug > Compile VBAProject to find the line'
        $exitCode = 3
    }
    if ($exitCode -eq 0) {
        $text = [string]$result
        Write-Host $text
        if ($text.StartsWith('PASS', [StringComparison]::Ordinal)) { $exitCode = 0 } else { $exitCode = 2 }
    }
}
catch {
    Write-Host ('Build failed: ' + $_.Exception.Message)
    $exitCode = 4
}
finally {
    try { if ($null -ne $wb) { $wb.Close($false) } } catch { }
    try { if ($null -ne $xl) { $xl.Quit() } } catch { }
    foreach ($k in @($sheetObjs.Keys)) { Release-Com $sheetObjs[$k] }
    Release-Com $vbcomps
    Release-Com $vbp
    Release-Com $sheets
    Release-Com $wb
    Release-Com $xl
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

exit $exitCode
