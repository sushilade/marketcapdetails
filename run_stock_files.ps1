$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$LogRoot = Join-Path $ProjectRoot "logs"
$HolidayFile = Join-Path $ProjectRoot "market_holidays.txt"

New-Item -ItemType Directory -Force -Path $LogRoot | Out-Null

function Write-RunLog {
    param([string]$Message)
    $LogFile = Join-Path $LogRoot ("stock_runner_" + (Get-Date -Format "yyyyMMdd") + ".log")
    $Line = "[" + (Get-Date -Format "yyyy-MM-dd HH:mm:ss") + "] " + $Message
    Add-Content -Path $LogFile -Value $Line
    Write-Host $Line
}

function Get-MarketHolidays {
    if (-not (Test-Path $HolidayFile)) {
        return @()
    }

    $Dates = @()
    foreach ($Line in Get-Content -Path $HolidayFile) {
        $Match = [regex]::Match($Line, "^\s*(\d{4}-\d{2}-\d{2})")
        if ($Match.Success) {
            $Dates += $Match.Groups[1].Value
        }
    }
    return $Dates
}

function Get-PythonExecutable {
    $Candidates = @("py", "python")
    foreach ($Candidate in $Candidates) {
        try {
            $Command = Get-Command $Candidate -ErrorAction Stop
            return $Command.Source
        }
        catch {
        }
    }

    throw "Python executable not found. Install Python or add python.exe to PATH."
}

function Invoke-Script {
    param(
        [string]$ScriptName,
        [string]$PythonExe,
        [string]$RunId
    )

    $ScriptPath = Join-Path $ProjectRoot $ScriptName
    if (-not (Test-Path $ScriptPath)) {
        throw "Script not found: $ScriptPath"
    }

    $Stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $StdoutLog = Join-Path $LogRoot ($ScriptName -replace "\.py$", "_${Stamp}_stdout.log")
    $StderrLog = Join-Path $LogRoot ($ScriptName -replace "\.py$", "_${Stamp}_stderr.log")
    $Arguments = "-3 `"$ScriptPath`""

    Write-RunLog "Starting $ScriptName"
    $Process = Start-Process -FilePath $PythonExe -ArgumentList $Arguments -WorkingDirectory $ProjectRoot -Wait -PassThru -RedirectStandardOutput $StdoutLog -RedirectStandardError $StderrLog -NoNewWindow

    if ($Process.ExitCode -ne 0) {
        Write-RunLog "Failed $ScriptName with exit code $($Process.ExitCode)"
        throw "Script failed: $ScriptName"
    }

    Write-RunLog "Completed $ScriptName"
}

$Today = Get-Date
$TodayKey = $Today.ToString("yyyy-MM-dd")
$MarketHolidays = Get-MarketHolidays

if ($Today.DayOfWeek -eq [DayOfWeek]::Saturday -or $Today.DayOfWeek -eq [DayOfWeek]::Sunday) {
    Write-RunLog "Skipped: weekend"
    exit 0
}

if ($MarketHolidays -contains $TodayKey) {
    Write-RunLog "Skipped: market holiday $TodayKey"
    exit 0
}

$RunId = Get-Date -Format "yyyyMMdd_HHmmss"
$PythonExe = Get-PythonExecutable

Write-RunLog "Market working day detected: $TodayKey"
Invoke-Script -ScriptName "apinse.py" -PythonExe $PythonExe -RunId $RunId
Invoke-Script -ScriptName "websitedevlopment.py" -PythonExe $PythonExe -RunId $RunId

Write-RunLog "All scripts completed successfully"
