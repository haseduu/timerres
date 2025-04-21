param(
    [int]$Samples = 30,
    [double]$Increment = 0.002,
    [switch]$QuickTest
)

function Test-AdminPrivileges {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Error "Administrator privileges are required to run this script."
        exit 1
    }
}

function Get-StartupFolder {
    return [Environment]::GetFolderPath("Startup")
}

function Create-StartupShortcut($Resolution) {
    $WshShell = New-Object -ComObject WScript.Shell
    $ShortcutPath = Join-Path (Get-StartupFolder) "SetTimerResolution.lnk"
    $TargetPath = Join-Path $PSScriptRoot "SetTimerResolution.exe"
    $Arguments = "--resolution $($Resolution * 1E4) --no-console"

    $Shortcut = $WshShell.CreateShortcut($ShortcutPath)
    $Shortcut.TargetPath = $TargetPath
    $Shortcut.Arguments = $Arguments
    $Shortcut.WorkingDirectory = $PSScriptRoot
    $Shortcut.WindowStyle = 7
    $Shortcut.Save()

    Write-Host "Shortcut created to apply $($Resolution)ms at startup." -ForegroundColor Green
}

function Format-Duration($Seconds) {
    if ($Seconds -ge 60) {
        $Minutes = [math]::Floor($Seconds / 60)
        $RemainingSeconds = $Seconds % 60
        return "$Minutes min $RemainingSeconds sec"
    } else {
        return "$Seconds sec"
    }
}

function Show-Progress($Current, $Total, $BarWidth = 40, $Activity = "Processing") {
    $PercentComplete = [math]::Min(100, [math]::Floor(($Current / $Total) * 100))
    $CompletedBlocks = [math]::Floor(($BarWidth * $PercentComplete) / 100)
    $RemainingBlocks = $BarWidth - $CompletedBlocks
    
    # Clear the entire line before writing
    Write-Host -NoNewLine "`r" + (" " * (80)) + "`r"
    
    # Build progress bar
    $ProgressBar = "["
    if ($CompletedBlocks -gt 0) {
        $ProgressBar += "=" * $CompletedBlocks
    }
    
    if ($CompletedBlocks -lt $BarWidth) {
        $ProgressBar += ">"
        $RemainingBlocks--
    }
    
    if ($RemainingBlocks -gt 0) {
        $ProgressBar += " " * $RemainingBlocks
    }
    
    $ProgressBar += "] $($PercentComplete)% - $Activity"
    
    # Write progress bar
    Write-Host -NoNewLine "`r$ProgressBar"
    
    # If we're at 100%, add a newline
    if ($PercentComplete -eq 100) {
        Write-Host ""
    }
}

function Enable-GlobalResolutionSupport {
    $RegistryPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\kernel"
    $ValueName = "GlobalTimerResolutionRequests"
    $ExpectedValue = 1

    $CurrentValue = Get-ItemProperty -Path $RegistryPath -Name $ValueName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty $ValueName -ErrorAction SilentlyContinue

    if ($CurrentValue -ne $ExpectedValue) {
        Write-Host "`nActivating global timer resolution support..." -ForegroundColor Cyan
        try {
            Set-ItemProperty -Path $RegistryPath -Name $ValueName -Value $ExpectedValue -Type DWord -Force
            Write-Host "Registry key applied successfully. System restart required." -ForegroundColor Green
            Start-Sleep -Seconds 2
            shutdown /r /t 5 /c "Restarting to apply global timer configuration"
            exit
        } catch {
            Write-Error "Error applying registry key. Run the script as administrator."
            exit 1
        }
    }
}

function Download-RequiredExecutable($FileName, $Url) {
    $OutputPath = Join-Path $PSScriptRoot $FileName
    
    Write-Host "Downloading $FileName from GitHub..." -ForegroundColor Yellow
    try {
        $WebClient = New-Object System.Net.WebClient
        $WebClient.DownloadFile($Url, $OutputPath)
        Write-Host "$FileName downloaded successfully." -ForegroundColor Green
        return $true
    } catch {
        Write-Error "Failed to download $FileName. Error: $_"
        return $false
    }
}

function Assert-ExecutablesExist {
    $RequiredFiles = @{
        "SetTimerResolution.exe" = "https://github.com/valleyofdoom/TimerResolution/releases/download/SetTimerResolution-v1.0.0/SetTimerResolution.exe"
        "MeasureSleep.exe" = "https://github.com/valleyofdoom/TimerResolution/releases/download/MeasureSleep-v1.0.0/MeasureSleep.exe"
    }
    
    $AllFilesExist = $true
    
    foreach ($File in $RequiredFiles.Keys) {
        $FilePath = Join-Path $PSScriptRoot $File
        if (-not (Test-Path $FilePath)) {
            Write-Warning "Required file '$File' not found in the current directory."
            $DownloadConfirm = Read-Host "Would you like to download $File from GitHub? (y/n)"
            
            if ($DownloadConfirm -match '^[yY]') {
                $DownloadSuccess = Download-RequiredExecutable -FileName $File -Url $RequiredFiles[$File]
                if (-not $DownloadSuccess) {
                    $AllFilesExist = $false
                }
            } else {
                Write-Warning "Make sure both executables are in the same folder as this script."
                $AllFilesExist = $false
            }
        }
    }
    
    if (-not $AllFilesExist) {
        exit 1
    }
}

function Get-BenchmarkParameters {
    Write-Host "`n=== Timer Resolution Benchmark Tool ===`n" -ForegroundColor Cyan

    # Get Samples parameter
    $SamplesInput = Read-Host "Enter number of samples per test (recommended: 30)"
    if ([string]::IsNullOrWhiteSpace($SamplesInput)) {
        Write-Host "Using default value: $Samples samples" -ForegroundColor Yellow
    } else {
        try {
            $Samples = [int]$SamplesInput
            if ($Samples -le 0) {
                throw "Samples must be greater than zero."
            }
        } catch {
            Write-Warning "Invalid input. Using default value: $Samples samples"
        }
    }

    # Get Increment parameter
    $IncrementInput = Read-Host "Enter increment value in ms (recommended: 0.002)"
    if ([string]::IsNullOrWhiteSpace($IncrementInput)) {
        Write-Host "Using default value: $Increment ms" -ForegroundColor Yellow
    } else {
        try {
            $Increment = [double]$IncrementInput
            if ($Increment -le 0) {
                throw "Increment must be greater than zero."
            }
        } catch {
            Write-Warning "Invalid input. Using default value: $Increment ms"
        }
    }

    # Ask for Quick Test
    if (-not $QuickTest) {
        $QuickTestInput = Read-Host "Run in Quick Test mode? (y/n)"
        if ($QuickTestInput -match '^[yY]') {
            $QuickTest = $true
            $Samples = 10
            Write-Host "Running in Quick Test mode with $Samples samples per resolution." -ForegroundColor Yellow
        }
    } else {
        $Samples = 10
        Write-Host "Running in Quick Test mode with $Samples samples per resolution." -ForegroundColor Yellow
    }

    $InitialResolution = Read-Host "Enter initial resolution in ms (recommended: 0.5)"
    $FinalResolution = Read-Host "Enter final resolution in ms (recommended: 0.6)"

    try {
        [double]$Start = $InitialResolution
        [double]$End = $FinalResolution

        if ($Start -le 0 -or $End -le 0) {
            throw "Resolution values must be greater than zero."
        }
        if ($Start -ge $End) {
            throw "Final resolution must be greater than initial resolution."
        }
        return @{ Start = $Start; End = $End; Samples = $Samples; Increment = $Increment }
    } catch {
        Write-Error "Invalid input. Please enter valid numbers."
        exit 1
    }
}

function Run-Benchmark($Parameters) {
    $Start = $Parameters.Start
    $End = $Parameters.End
    $Samples = $Parameters.Samples
    $Increment = $Parameters.Increment

    $Iterations = [math]::Ceiling(($End - $Start) / $Increment) + 1
    $EstimatedTime = Format-Duration ($Iterations * 2 * $Samples / 10)

    Write-Host "`nBenchmark Configuration:" -ForegroundColor Cyan
    Write-Host " - Resolution Range: $Start ms to $End ms"
    Write-Host " - Increment: $Increment ms"
    Write-Host " - Samples per Test: $Samples"
    Write-Host " - Total Tests: $Iterations"
    Write-Host " - Estimated Time: $EstimatedTime"

    $Confirmation = Read-Host "`nStart benchmark? (y/n)"
    if ($Confirmation -notmatch '^[yY]') {
        Write-Host "Benchmark canceled." -ForegroundColor Yellow
        exit 0
    }

    Set-Location $PSScriptRoot
    $ResultsFile = Join-Path $PSScriptRoot "benchmark_results.csv"
    "RequestedResolutionMs,AverageDeltaMs,StandardDeviationMs,ConsistencyPercentage,OverallScore" | Out-File $ResultsFile

    $StartTime = Get-Date
    $TestCounter = 0
    $TotalTests = $Iterations * 3

    $BestResults = @{
        LowestDelta = @{ Value = [double]::MaxValue; Resolution = $null }
        LowestStdDev = @{ Value = [double]::MaxValue; Resolution = $null }
        HighestConsistency = @{ Value = 0; Resolution = $null }
        HighestOverallScore = @{ Value = 0; Resolution = $null }
    }

    for ($CurrentResolution = $Start; $CurrentResolution -le $End; $CurrentResolution += $Increment) {
        $CurrentResolution = [math]::Round($CurrentResolution, 6)
        $TestCounter++

        Write-Host "`n`nTesting Resolution: $($CurrentResolution)ms... ($TestCounter of $Iterations)" -ForegroundColor Cyan

        Stop-Process -Name "SetTimerResolution" -ErrorAction SilentlyContinue
        Start-Process ".\SetTimerResolution.exe" -ArgumentList @("--resolution", ($CurrentResolution * 1E4), "--no-console")
        Start-Sleep 1

        $AllDeltas = @()
        $AllStdDevs = @()

        for ($Run = 1; $Run -le 3; $Run++) {
            $RunActivity = "Resolution $($CurrentResolution)ms - Run $Run of 3"
            Show-Progress ($Run + ($TestCounter - 1) * 3) $TotalTests 40 $RunActivity
            
            $Output = .\MeasureSleep.exe --samples $Samples
            $OutputLines = $Output -split "`n"

            $Average = $null
            $StdDev = $null

            foreach ($Line in $OutputLines) {
                if ($Line -match "Avg: ([\d\.]+)") { $Average = [double]$Matches[1] }
                if ($Line -match "STDEV: ([\d\.]+)") { $StdDev = [double]$Matches[1] }
            }

            if ($null -ne $Average -and $null -ne $StdDev) {
                $AllDeltas += $Average
                $AllStdDevs += $StdDev
                Write-Host "  Run $Run Result: Avg=$($Average)ms, STDEV=$($StdDev)ms" -ForegroundColor Gray
            }
        }

        # Show complete progress for this resolution
        Show-Progress ($TestCounter * 3) $TotalTests 40 "Resolution $($CurrentResolution)ms completed"

        $AverageDelta = ($AllDeltas | Measure-Object -Average).Average
        $AverageStdDev = ($AllStdDevs | Measure-Object -Average).Average
        $Consistency = if ($AverageDelta -gt 0) { 100 * (1 - [Math]::Min(1, $AverageStdDev / $AverageDelta)) } else { 0 }

        $NormalizedDelta = 1 - [Math]::Min(1, $AverageDelta / 2)
        $NormalizedStdDev = 1 - [Math]::Min(1, $AverageStdDev / 0.5)
        $NormalizedConsistency = $Consistency / 100
        $OverallScore = [math]::Round(($NormalizedDelta * 0.5) + ($NormalizedStdDev * 0.25) + ($NormalizedConsistency * 0.25) * 100, 2)

        "$CurrentResolution, $AverageDelta, $AverageStdDev, $Consistency, $OverallScore" | Out-File $ResultsFile -Append

        if ($AverageDelta -lt $BestResults.LowestDelta.Value) {
            $BestResults.LowestDelta.Value = $AverageDelta
            $BestResults.LowestDelta.Resolution = $CurrentResolution
        }
        if ($AverageStdDev -lt $BestResults.LowestStdDev.Value) {
            $BestResults.LowestStdDev.Value = $AverageStdDev
            $BestResults.LowestStdDev.Resolution = $CurrentResolution
        }
        if ($Consistency -gt $BestResults.HighestConsistency.Value) {
            $BestResults.HighestConsistency.Value = $Consistency
            $BestResults.HighestConsistency.Resolution = $CurrentResolution
        }
        if ($OverallScore -gt $BestResults.HighestOverallScore.Value) {
            $BestResults.HighestOverallScore.Value = $OverallScore
            $BestResults.HighestOverallScore.Resolution = $CurrentResolution
        }

        Stop-Process -Name "SetTimerResolution" -ErrorAction SilentlyContinue
    }

    # Rest of the function remains the same...
    $EndTime = Get-Date
    $TotalExecutionTime = Format-Duration ($($EndTime - $StartTime).TotalSeconds)

    Write-Host "`n`n=== Benchmark Complete ===`n" -ForegroundColor Green
    Write-Host "Total Execution Time: $TotalExecutionTime" -ForegroundColor Cyan
    Write-Host "Results saved to: $ResultsFile" -ForegroundColor Cyan

    Write-Host "`nBest Results by Criterion:" -ForegroundColor Yellow
    Write-Host " - Lowest Average Delta: $($BestResults.LowestDelta.Resolution) ms (Average: $([math]::Round($BestResults.LowestDelta.Value, 3)) ms)"
    Write-Host " - Most Stable (Lowest Standard Deviation): $($BestResults.LowestStdDev.Resolution) ms (STDEV: $([math]::Round($BestResults.LowestStdDev.Value, 3)) ms)"
    Write-Host " - Highest Consistency: $($BestResults.HighestConsistency.Resolution) ms (Consistency: $([math]::Round($BestResults.HighestConsistency.Value, 1))%)"
    Write-Host " - Highest Overall Score: $($BestResults.HighestOverallScore.Resolution) ms (Score: $($BestResults.HighestOverallScore.Value))" -ForegroundColor Green

    Write-Host "`nRecommendation:" -ForegroundColor Cyan
    Write-Host "The resolution with the highest overall score is $($BestResults.HighestOverallScore.Resolution) ms."
    Write-Host "This resolution offers a good balance between low latency and consistent timing."

    return $BestResults.HighestOverallScore.Resolution
}

function Apply-StartupResolution($RecommendedResolution) {
    $ApplyChoice = Read-Host "`nDo you want to apply this resolution ($RecommendedResolution ms) at system startup? (y/n)"
    if ($ApplyChoice -match '^[yY]') {
        Create-StartupShortcut $RecommendedResolution
        Start-Process ".\SetTimerResolution.exe" -ArgumentList @("--resolution", $RecommendedResolution, "--no-console")
    } else {
        $CustomChoice = Read-Host "Apply a different resolution at startup? (y/n)"
        if ($CustomChoice -match '^[yY]') {
            try {
                $CustomResolutionInput = Read-Host "Enter your preferred resolution in ms"
                [double]$CustomResolution = $CustomResolutionInput
                if ($CustomResolution -le 0) {
                    throw "Resolution must be greater than zero."
                }
                Create-StartupShortcut $CustomResolution
                Start-Process ".\SetTimerResolution.exe" -ArgumentList @("--resolution", $CustomResolution, "--no-console")
            } catch {
                Write-Warning "Invalid input. No startup action performed."
            }
        } else {
            Write-Host "No startup action performed." -ForegroundColor Yellow
        }
    }
}

Test-AdminPrivileges
Enable-GlobalResolutionSupport
Assert-ExecutablesExist
$BenchmarkParameters = Get-BenchmarkParameters
$BestResolution = Run-Benchmark $BenchmarkParameters
Apply-StartupResolution $BestResolution
exit 0