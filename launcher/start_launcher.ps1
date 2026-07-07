$ErrorActionPreference = "Stop"
$LauncherDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Url = "http://127.0.0.1:8765"

function Test-LocalPort {
    try {
        $client = New-Object System.Net.Sockets.TcpClient
        $result = $client.BeginConnect("127.0.0.1", 8765, $null, $null)
        $connected = $result.AsyncWaitHandle.WaitOne(500)
        if ($connected) { $client.EndConnect($result) }
        $client.Close()
        return $connected
    } catch {
        return $false
    }
}

if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
    Add-Type -AssemblyName PresentationFramework
    [System.Windows.MessageBox]::Show(
        "Не установлен GitHub CLI. Установите его с https://cli.github.com/ и повторите запуск.",
        "Мониторинг отзывов",
        "OK",
        "Error"
    ) | Out-Null
    Start-Process "https://cli.github.com/"
    exit 1
}

& gh auth status 2>$null | Out-Null
if ($LASTEXITCODE -ne 0) {
    Start-Process powershell.exe -Wait -ArgumentList @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-Command", "gh auth login --web"
    )
    & gh auth status 2>$null | Out-Null
    if ($LASTEXITCODE -ne 0) { exit 1 }
}

if (-not (Test-LocalPort)) {
    $pythonCommand = Get-Command py -ErrorAction SilentlyContinue
    $arguments = @()
    if ($pythonCommand) {
        $pythonExe = $pythonCommand.Source
        $arguments = @("-3", (Join-Path $LauncherDir "server.py"))
    } else {
        $pythonCommand = Get-Command python -ErrorAction SilentlyContinue
        if (-not $pythonCommand) {
            Add-Type -AssemblyName PresentationFramework
            [System.Windows.MessageBox]::Show(
                "Python 3 не найден. Установите Python с https://www.python.org/downloads/ и повторите запуск.",
                "Мониторинг отзывов",
                "OK",
                "Error"
            ) | Out-Null
            Start-Process "https://www.python.org/downloads/"
            exit 1
        }
        $pythonExe = $pythonCommand.Source
        $arguments = @((Join-Path $LauncherDir "server.py"))
    }

    Start-Process -FilePath $pythonExe -ArgumentList $arguments -WorkingDirectory $LauncherDir -WindowStyle Hidden
    Start-Sleep -Seconds 2
}

Start-Process $Url
