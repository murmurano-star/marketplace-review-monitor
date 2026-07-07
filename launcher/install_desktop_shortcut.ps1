$ErrorActionPreference = "Stop"
$LauncherDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$StartScript = Join-Path $LauncherDir "start_launcher.ps1"
$Desktop = [Environment]::GetFolderPath("Desktop")
$ShortcutPath = Join-Path $Desktop "Мониторинг отзывов GIGAS и SOFT99.lnk"

if (-not (Test-Path $StartScript)) {
    throw "Не найден файл запуска: $StartScript"
}

$Shell = New-Object -ComObject WScript.Shell
$Shortcut = $Shell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
$Shortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$StartScript`""
$Shortcut.WorkingDirectory = $LauncherDir
$Shortcut.Description = "Ручной запуск мониторинга отзывов GIGAS и SOFT99"
$Shortcut.IconLocation = "$env:SystemRoot\System32\shell32.dll,220"
$Shortcut.Save()

Add-Type -AssemblyName PresentationFramework
[System.Windows.MessageBox]::Show(
    "Ярлык создан на рабочем столе:`n$ShortcutPath",
    "Мониторинг отзывов",
    "OK",
    "Information"
) | Out-Null
