$ErrorActionPreference = "Stop"
$BaseUrl = "https://raw.githubusercontent.com/murmurano-star/marketplace-review-monitor/main/launcher"
$InstallDir = Join-Path $env:LOCALAPPDATA "MarketplaceReviewMonitor"
$Desktop = [Environment]::GetFolderPath("Desktop")
$ShortcutPath = Join-Path $Desktop "Мониторинг отзывов GIGAS и SOFT99.lnk"

New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null

$Files = @(
    "server.py",
    "index.html",
    "start_launcher.ps1",
    "README.md"
)

foreach ($File in $Files) {
    $Uri = "$BaseUrl/$File"
    $Destination = Join-Path $InstallDir $File
    Invoke-WebRequest -Uri $Uri -OutFile $Destination -UseBasicParsing
}

$StartScript = Join-Path $InstallDir "start_launcher.ps1"
$Shell = New-Object -ComObject WScript.Shell
$Shortcut = $Shell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
$Shortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$StartScript`""
$Shortcut.WorkingDirectory = $InstallDir
$Shortcut.Description = "Ручной запуск мониторинга отзывов GIGAS и SOFT99"
$Shortcut.IconLocation = "$env:SystemRoot\System32\shell32.dll,220"
$Shortcut.Save()

Add-Type -AssemblyName PresentationFramework
$Answer = [System.Windows.MessageBox]::Show(
    "Приложение установлено и ярлык создан на рабочем столе.`n`nЗапустить сейчас?",
    "Мониторинг отзывов",
    "YesNo",
    "Information"
)
if ($Answer -eq "Yes") {
    & $StartScript
}
