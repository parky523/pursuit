# Pursuit one-time setup. Right-click > "Run with PowerShell" (or: powershell -ExecutionPolicy Bypass -File setup.ps1)
# - creates config.json and Job_Search_Tracker.xlsx from the templates (never overwrites existing ones)
# - creates Desktop, Start Menu and autostart shortcuts to Pursuit.exe
# - optionally opens the firewall for the phone app (asks for admin)
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $here
$exe = Join-Path $here 'Pursuit.exe'

if (-not (Test-Path 'config.json')) { Copy-Item 'config.example.json' 'config.json'; Write-Host 'Created config.json' }
else { Write-Host 'config.json already exists - left as is' }

$cfg = Get-Content 'config.json' -Raw | ConvertFrom-Json
$tracker = $cfg.tracker
if (-not [System.IO.Path]::IsPathRooted($tracker)) { $tracker = Join-Path $here $tracker }
if (-not (Test-Path $tracker)) { Copy-Item 'Job_Search_Tracker_TEMPLATE.xlsx' $tracker; Write-Host "Created tracker: $tracker" }
else { Write-Host "Tracker found: $tracker" }

if (-not (Test-Path $exe)) {
    Write-Host ''
    Write-Host 'Pursuit.exe not found. Download it from the GitHub Releases page into this folder,' -ForegroundColor Yellow
    Write-Host 'or run build_exe.bat (needs Python). Then run this setup again for the shortcuts.' -ForegroundColor Yellow
    Read-Host 'Press Enter to close'
    exit
}

$sh = New-Object -ComObject WScript.Shell
$start = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs'
$targets = @(
    (Join-Path ([Environment]::GetFolderPath('Desktop')) 'Pursuit.lnk'),
    (Join-Path $start 'Pursuit.lnk'),
    (Join-Path $start 'Startup\Pursuit.lnk')
)
foreach ($t in $targets) {
    $l = $sh.CreateShortcut($t)
    $l.TargetPath = $exe; $l.WorkingDirectory = $here; $l.IconLocation = "$exe,0"
    $l.Description = 'Pursuit - job hunt tracker'; $l.Save()
}
Write-Host 'Shortcuts created: Desktop, Start Menu, Startup (opens at login)'

$ans = Read-Host 'Allow the phone app through Windows Firewall now? (needs admin) [y/N]'
if ($ans -match '^[yY]') {
    $rule = "New-NetFirewallRule -DisplayName 'Pursuit phone app (8787)' -Direction Inbound -Action Allow -Protocol TCP -LocalPort 8787 -Program '$exe' -Profile Private,Public"
    Start-Process powershell -Verb RunAs -Wait -ArgumentList '-NoProfile', '-Command', $rule
    Write-Host 'Firewall rule added.'
}
Start-Process $exe
Write-Host 'Done - Pursuit is starting.'
Read-Host 'Press Enter to close'
