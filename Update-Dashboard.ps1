$ErrorActionPreference = 'Stop'

$dashboardDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = Split-Path -Parent (Split-Path -Parent $dashboardDir)
$sourceExcel = Get-ChildItem -LiteralPath $projectDir -Filter '*.xlsx' -File | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $sourceExcel) { throw "No .xlsx source file found in $projectDir" }
$sourceExcelPath = $sourceExcel.FullName
$targetExcel = Join-Path $dashboardDir 'masterplan.xlsx'

Write-Host ''
Write-Host '=== Update Dashboard v15 ==='
Write-Host ('Excel     : {0}' -f $sourceExcelPath)
Write-Host ('Dashboard : {0}' -f $dashboardDir)
Write-Host ''

if (-not (Test-Path -LiteralPath $sourceExcelPath)) {
    throw "Source Excel not found: $sourceExcelPath"
}

Write-Host '[1/3] Copying latest Excel into Dashboard...'
Copy-Item -LiteralPath $sourceExcelPath -Destination $targetExcel -Force
$src = Get-Item -LiteralPath $sourceExcelPath
$dst = Get-Item -LiteralPath $targetExcel
if ($src.Length -ne $dst.Length) {
    throw 'Copied Excel size does not match source.'
}
Write-Host ('[OK] masterplan.xlsx updated ({0:N1} MB)' -f ($dst.Length / 1MB))
Write-Host ('     Last modified: {0}' -f $src.LastWriteTime)

Write-Host '[2/3] Checking JavaScript syntax...'
$node = Get-Command node -ErrorAction SilentlyContinue
if ($node) {
    & node --check (Join-Path $dashboardDir 'app.js')
    if ($LASTEXITCODE -ne 0) { throw "app.js syntax check failed with exit code $LASTEXITCODE" }
    Write-Host '[OK] app.js syntax passed.'
} else {
    Write-Host '[INFO] Node.js not found; JavaScript syntax check skipped.'
}

Write-Host '[3/3] Checking Git publish status...'
$git = Get-Command git -ErrorAction SilentlyContinue
$gitDir = Join-Path $dashboardDir '.git'
if (-not $git) {
    Write-Host '[INFO] Git not found. Local Dashboard update is complete; GitHub push skipped.'
} elseif (-not (Test-Path -LiteralPath $gitDir)) {
    Write-Host '[INFO] This Dashboard folder is not a Git repository yet. Local update is complete; GitHub push skipped.'
} else {
    Push-Location $dashboardDir
    try {
        & git add masterplan.xlsx app.js index.html data.json photos
        $changes = & git status --porcelain
        if (-not $changes) {
            Write-Host '[INFO] No dashboard changes to publish.'
        } else {
            $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            & git commit -m "Update dashboard $stamp"
            if ($LASTEXITCODE -ne 0) { throw "git commit failed with exit code $LASTEXITCODE" }
            & git push
            if ($LASTEXITCODE -ne 0) { throw "git push failed with exit code $LASTEXITCODE" }
            Write-Host '[OK] Changes pushed to GitHub.'
        }
    } finally {
        Pop-Location
    }
}

Write-Host ''
Write-Host '=== Finished ==='
Write-Host 'Workflow: Save Excel -> double-click Update Dashboard.bat'
