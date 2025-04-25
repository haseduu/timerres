param(
    [int]$Samples = 30,
    [double]$Increment = 0.002,
    [switch]$QuickTest
)

function Test-AdminPrivileges {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        if ($global:Language -eq "PT-BR") {
            Write-Error "Privilegios de administrador sao necessarios para executar este script."
        } else {
            Write-Error "Administrator privileges are required to run this script."
        }
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

    if ($global:Language -eq "PT-BR") {
        Write-Host "Atalho criado para aplicar $($Resolution)ms na inicializacao." -ForegroundColor Green
    } else {
        Write-Host "Shortcut created to apply $($Resolution)ms at startup." -ForegroundColor Green
    }
}

function Format-Duration($Seconds) {
    if ($Seconds -ge 60) {
        $Minutes = [math]::Floor($Seconds / 60)
        $RemainingSeconds = $Seconds % 60
        if ($global:Language -eq "PT-BR") {
            return "$Minutes min $RemainingSeconds seg"
        } else {
            return "$Minutes min $RemainingSeconds sec"
        }
    } else {
        if ($global:Language -eq "PT-BR") {
            return "$Seconds seg"
        } else {
            return "$Seconds sec"
        }
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
        if ($global:Language -eq "PT-BR") {
            Write-Host "`nO suporte para resolucao global de timer precisa ser ativado." -ForegroundColor Cyan
            Write-Host "Isso requer uma alteracao no registro e reinicializacao do sistema." -ForegroundColor Yellow
            Write-Host "AVISO: A funcionalidade de resolucao de timer nao funcionara corretamente sem esta alteracao." -ForegroundColor Red
            
            $Confirmation = Read-Host "Deseja aplicar a alteracao no registro agora? (s/n)"
            
            if ($Confirmation -match '^[sS]') {
                try {
                    Set-ItemProperty -Path $RegistryPath -Name $ValueName -Value $ExpectedValue -Type DWord -Force
                    Write-Host "Chave de registro aplicada com sucesso." -ForegroundColor Green
                    
                    $RestartConfirmation = Read-Host "E necessaria uma reinicializacao do sistema para que as alteracoes entrem em vigor. Reiniciar agora? (s/n)"
                    
                    if ($RestartConfirmation -match '^[sS]') {
                        Write-Host "O sistema sera reiniciado em 5 segundos..." -ForegroundColor Yellow
                        Start-Sleep -Seconds 5
                        shutdown /r /t 5 /c "Reiniciando para aplicar configuracao global de timer"
                        exit
                    } else {
                        Write-Host "Por favor, reinicie seu sistema manualmente para que as alteracoes entrem em vigor." -ForegroundColor Yellow
                        Write-Host "AVISO: As funcoes de resolucao de timer nao funcionarao corretamente ate a reinicializacao do sistema." -ForegroundColor Red
                        $ContinueAnyway = Read-Host "Continuar com o benchmark mesmo assim? (s/n)"
                        if ($ContinueAnyway -notmatch '^[sS]') {
                            Write-Host "Saindo do script. Execute novamente apos a reinicializacao do sistema." -ForegroundColor Cyan
                            exit 0
                        }
                    }
                } catch {
                    Write-Error "Erro ao aplicar a chave de registro. Execute o script como administrador."
                    exit 1
                }
            } else {
                Write-Host "Alteracao de registro recusada." -ForegroundColor Yellow
                Write-Host "AVISO: As funcoes de resolucao de timer nao funcionarao corretamente sem esta alteracao." -ForegroundColor Red
                $ContinueAnyway = Read-Host "Continuar com o benchmark mesmo assim? (Nao recomendado) (s/n)"
                if ($ContinueAnyway -notmatch '^[sS]') {
                    Write-Host "Saindo do script." -ForegroundColor Cyan
                    exit 0
                }
            }
        } else {
            Write-Host "`nGlobal timer resolution support needs to be activated." -ForegroundColor Cyan
            Write-Host "This requires a registry change and system restart." -ForegroundColor Yellow
            Write-Host "WARNING: The timer resolution functionality will not work properly without this change." -ForegroundColor Red
            
            $Confirmation = Read-Host "Do you want to apply the registry change now? (y/n)"
            
            if ($Confirmation -match '^[yY]') {
                try {
                    Set-ItemProperty -Path $RegistryPath -Name $ValueName -Value $ExpectedValue -Type DWord -Force
                    Write-Host "Registry key applied successfully." -ForegroundColor Green
                    
                    $RestartConfirmation = Read-Host "A system restart is required for changes to take effect. Restart now? (y/n)"
                    
                    if ($RestartConfirmation -match '^[yY]') {
                        Write-Host "System will restart in 5 seconds..." -ForegroundColor Yellow
                        Start-Sleep -Seconds 5
                        shutdown /r /t 5 /c "Restarting to apply global timer configuration"
                        exit
                    } else {
                        Write-Host "Please restart your system manually for changes to take effect." -ForegroundColor Yellow
                        Write-Host "WARNING: Timer resolution functions will not work properly until system restart." -ForegroundColor Red
                        $ContinueAnyway = Read-Host "Continue with the benchmark anyway? (y/n)"
                        if ($ContinueAnyway -notmatch '^[yY]') {
                            Write-Host "Exiting script. Please run again after system restart." -ForegroundColor Cyan
                            exit 0
                        }
                    }
                } catch {
                    Write-Error "Error applying registry key. Run the script as administrator."
                    exit 1
                }
            } else {
                Write-Host "Registry change declined." -ForegroundColor Yellow
                Write-Host "WARNING: Timer resolution functions will not work properly without this change." -ForegroundColor Red
                $ContinueAnyway = Read-Host "Continue with the benchmark anyway? (Not recommended) (y/n)"
                if ($ContinueAnyway -notmatch '^[yY]') {
                    Write-Host "Exiting script." -ForegroundColor Cyan
                    exit 0
                }
            }
        }
    }
}

function Download-RequiredExecutable($FileName, $Url) {
    $OutputPath = Join-Path $PSScriptRoot $FileName
    
    if ($global:Language -eq "PT-BR") {
        Write-Host "Baixando $FileName do GitHub..." -ForegroundColor Yellow
    } else {
        Write-Host "Downloading $FileName from GitHub..." -ForegroundColor Yellow
    }
    
    try {
        $WebClient = New-Object System.Net.WebClient
        $WebClient.DownloadFile($Url, $OutputPath)
        
        if ($global:Language -eq "PT-BR") {
            Write-Host "$FileName baixado com sucesso." -ForegroundColor Green
        } else {
            Write-Host "$FileName downloaded successfully." -ForegroundColor Green
        }
        return $true
    } catch {
        if ($global:Language -eq "PT-BR") {
            Write-Error "Falha ao baixar $FileName. Erro: $_"
        } else {
            Write-Error "Failed to download $FileName. Error: $_"
        }
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
            if ($global:Language -eq "PT-BR") {
                Write-Warning "Arquivo necessario '$File' nao encontrado no diretorio atual."
                $DownloadConfirm = Read-Host "Gostaria de baixar $File do GitHub? (s/n)"
                
                if ($DownloadConfirm -match '^[sS]') {
                    $DownloadSuccess = Download-RequiredExecutable -FileName $File -Url $RequiredFiles[$File]
                    if (-not $DownloadSuccess) {
                        $AllFilesExist = $false
                    }
                } else {
                    Write-Warning "Certifique-se que ambos os executaveis estejam na mesma pasta que este script."
                    $AllFilesExist = $false
                }
            } else {
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
    }
    
    if (-not $AllFilesExist) {
        exit 1
    }
}

function Set-Language {
    Write-Host "Select Language / Selecione o Idioma:" -ForegroundColor Cyan
    Write-Host "1. English"
    Write-Host "2. Portuguese (PT-BR) / Portugues (PT-BR)"
    
    $LanguageChoice = Read-Host "Enter your choice / Digite sua escolha (1/2)"
    
    switch ($LanguageChoice) {
        "2" {
            $global:Language = "PT-BR"
            Write-Host "Idioma definido para Portugues (PT-BR)." -ForegroundColor Green
        }
        default {
            $global:Language = "EN"
            Write-Host "Language set to English." -ForegroundColor Green
        }
    }
}

function Apply-CustomResolution {
    if ($global:Language -eq "PT-BR") {
        try {
            $CustomResolution = Read-Host "Digite a resolucao desejada em ms (ex: 0.5)"
            [double]$CustomResolution = $CustomResolution
            
            if ($CustomResolution -le 0) {
                throw "A resolucao deve ser maior que zero."
            }
            
            Create-StartupShortcut $CustomResolution
            Start-Process ".\SetTimerResolution.exe" -ArgumentList @("--resolution", ($CustomResolution * 1E4), "--no-console")
            Write-Host "Resolucao $CustomResolution ms aplicada com sucesso!" -ForegroundColor Green
        } catch {
            Write-Warning "Entrada invalida. Use um numero valido maior que zero."
            return $false
        }
    } else {
        try {
            $CustomResolution = Read-Host "Enter your desired resolution in ms (ex: 0.5)"
            [double]$CustomResolution = $CustomResolution
            
            if ($CustomResolution -le 0) {
                throw "Resolution must be greater than zero."
            }
            
            Create-StartupShortcut $CustomResolution
            Start-Process ".\SetTimerResolution.exe" -ArgumentList @("--resolution", ($CustomResolution * 1E4), "--no-console")
            Write-Host "Resolution $CustomResolution ms applied successfully!" -ForegroundColor Green
        } catch {
            Write-Warning "Invalid input. Please use a valid number greater than zero."
            return $false
        }
    }
    
    return $true
}

function Get-BenchmarkParameters {
    if ($global:Language -eq "PT-BR") {
        Write-Host "`n=== Ferramenta de Benchmark de Resolucao de Timer ===`n" -ForegroundColor Cyan
        
        # Ask if user wants to skip benchmark
        $SkipBenchmark = Read-Host "Deseja pular o benchmark e aplicar uma resolucao personalizada? (s/n)"
        if ($SkipBenchmark -match '^[sS]') {
            $Success = Apply-CustomResolution
            if ($Success) {
                exit 0
            }
        }

        # Get Samples parameter
        $SamplesInput = Read-Host "Informe o numero de amostras por teste (recomendado: 30)"
        if ([string]::IsNullOrWhiteSpace($SamplesInput)) {
            Write-Host "Usando valor padrao: $Samples amostras" -ForegroundColor Yellow
        } else {
            try {
                $Samples = [int]$SamplesInput
                if ($Samples -le 0) {
                    throw "O numero de amostras deve ser maior que zero."
                }
            } catch {
                Write-Warning "Entrada invalida. Usando valor padrao: $Samples amostras"
            }
        }

        # Get Increment parameter
        $IncrementInput = Read-Host "Informe o valor do incremento em ms (recomendado: 0.002)"
        if ([string]::IsNullOrWhiteSpace($IncrementInput)) {
            Write-Host "Usando valor padrao: $Increment ms" -ForegroundColor Yellow
        } else {
            try {
                $Increment = [double]$IncrementInput
                if ($Increment -le 0) {
                    throw "O incremento deve ser maior que zero."
                }
            } catch {
                Write-Warning "Entrada invalida. Usando valor padrao: $Increment ms"
            }
        }

        # Ask for Quick Test
        if (-not $QuickTest) {
            $QuickTestInput = Read-Host "Executar no modo de Teste Rapido? (s/n)"
            if ($QuickTestInput -match '^[sS]') {
                $QuickTest = $true
                $Samples = 10
                Write-Host "Executando no modo de Teste Rapido com $Samples amostras por resolucao." -ForegroundColor Yellow
            }
        } else {
            $Samples = 10
            Write-Host "Executando no modo de Teste Rapido com $Samples amostras por resolucao." -ForegroundColor Yellow
        }

        $InitialResolution = Read-Host "Informe a resolucao inicial em ms (recomendado: 0.5)"
        $FinalResolution = Read-Host "Informe a resolucao final em ms (recomendado: 0.6)"
    } else {
        Write-Host "`n=== Timer Resolution Benchmark Tool ===`n" -ForegroundColor Cyan
        
        # Ask if user wants to skip benchmark
        $SkipBenchmark = Read-Host "Would you like to skip the benchmark and apply a custom resolution? (y/n)"
        if ($SkipBenchmark -match '^[yY]') {
            $Success = Apply-CustomResolution
            if ($Success) {
                exit 0
            }
        }

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
    }

    try {
        [double]$Start = $InitialResolution
        [double]$End = $FinalResolution

        if ($Start -le 0 -or $End -le 0) {
            if ($global:Language -eq "PT-BR") {
                throw "Os valores de resolucao devem ser maiores que zero."
            } else {
                throw "Resolution values must be greater than zero."
            }
        }
        if ($Start -ge $End) {
            if ($global:Language -eq "PT-BR") {
                throw "A resolucao final deve ser maior que a resolucao inicial."
            } else {
                throw "Final resolution must be greater than initial resolution."
            }
        }
        return @{ Start = $Start; End = $End; Samples = $Samples; Increment = $Increment }
    } catch {
        if ($global:Language -eq "PT-BR") {
            Write-Error "Entrada invalida. Por favor, insira numeros validos."
        } else {
            Write-Error "Invalid input. Please enter valid numbers."
        }
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

    if ($global:Language -eq "PT-BR") {
        Write-Host "`nConfiguracao do Benchmark:" -ForegroundColor Cyan
        Write-Host " - Alcance de Resolucao: $Start ms ate $End ms"
        Write-Host " - Incremento: $Increment ms"
        Write-Host " - Amostras por Teste: $Samples"
        Write-Host " - Total de Testes: $Iterations"
        Write-Host " - Tempo Estimado: $EstimatedTime"

        $Confirmation = Read-Host "`nIniciar benchmark? (s/n)"
        if ($Confirmation -notmatch '^[sS]') {
            Write-Host "Benchmark cancelado." -ForegroundColor Yellow
            exit 0
        }
    } else {
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

        if ($global:Language -eq "PT-BR") {
            Write-Host "`n`nTestando Resolucao: $($CurrentResolution)ms... ($TestCounter de $Iterations)" -ForegroundColor Cyan
        } else {
            Write-Host "`n`nTesting Resolution: $($CurrentResolution)ms... ($TestCounter of $Iterations)" -ForegroundColor Cyan
        }

        Stop-Process -Name "SetTimerResolution" -ErrorAction SilentlyContinue
        Start-Process ".\SetTimerResolution.exe" -ArgumentList @("--resolution", ($CurrentResolution * 1E4), "--no-console")
        Start-Sleep 1

        $AllDeltas = @()
        $AllStdDevs = @()

        for ($Run = 1; $Run -le 3; $Run++) {
            if ($global:Language -eq "PT-BR") {
                $RunActivity = "Resolucao $($CurrentResolution)ms - Execucao $Run de 3"
            } else {
                $RunActivity = "Resolution $($CurrentResolution)ms - Run $Run of 3"
            }
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
                if ($global:Language -eq "PT-BR") {
                    Write-Host "  Resultado da Execucao $Run - Media=$($Average)ms, DESVPAD=$($StdDev)ms" -ForegroundColor Gray
                } else {
                    Write-Host "  Run $Run Result: Avg=$($Average)ms, STDEV=$($StdDev)ms" -ForegroundColor Gray
                }
                            }
                        }

        # Show complete progress for this resolution
        if ($global:Language -eq "PT-BR") {
            Show-Progress ($TestCounter * 3) $TotalTests 40 "Resolucao $($CurrentResolution)ms concluida"
        } else {
            Show-Progress ($TestCounter * 3) $TotalTests 40 "Resolution $($CurrentResolution)ms completed"
        }

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

    $EndTime = Get-Date
    $TotalExecutionTime = Format-Duration ($($EndTime - $StartTime).TotalSeconds)

    if ($global:Language -eq "PT-BR") {
        Write-Host "`n`n=== Benchmark Concluido ===`n" -ForegroundColor Green
        Write-Host "Tempo Total de Execucao: $TotalExecutionTime" -ForegroundColor Cyan
        Write-Host "Resultados salvos em: $ResultsFile" -ForegroundColor Cyan

        Write-Host "`nMelhores Resultados por Criterio:" -ForegroundColor Yellow
        Write-Host " - Menor Media Delta: $($BestResults.LowestDelta.Resolution) ms (Media: $([math]::Round($BestResults.LowestDelta.Value, 3)) ms)"
        Write-Host " - Mais Estavel (Menor Desvio Padrao): $($BestResults.LowestStdDev.Resolution) ms (DESVPAD: $([math]::Round($BestResults.LowestStdDev.Value, 3)) ms)"
        Write-Host " - Maior Consistencia: $($BestResults.HighestConsistency.Resolution) ms (Consistencia: $([math]::Round($BestResults.HighestConsistency.Value, 1))%)"
        Write-Host " - Maior Pontuacao Geral: $($BestResults.HighestOverallScore.Resolution) ms (Pontuacao: $($BestResults.HighestOverallScore.Value))" -ForegroundColor Green

        Write-Host "`nRecomendacao:" -ForegroundColor Cyan
        Write-Host "A resolucao com a maior pontuacao geral e $($BestResults.HighestOverallScore.Resolution) ms."
        Write-Host "Esta resolucao oferece um bom equilibrio entre baixa latencia e temporizacao consistente."
    } else {
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
    }

    return $BestResults.HighestOverallScore.Resolution
}

function Apply-StartupResolution($RecommendedResolution) {
    if ($global:Language -eq "PT-BR") {
        $ApplyChoice = Read-Host "`nDeseja aplicar esta resolucao ($RecommendedResolution ms) na inicializacao do sistema? (s/n)"
        if ($ApplyChoice -match '^[sS]') {
            Create-StartupShortcut $RecommendedResolution
            Start-Process ".\SetTimerResolution.exe" -ArgumentList @("--resolution", ($RecommendedResolution * 1E4), "--no-console")
        } else {
            $CustomChoice = Read-Host "Aplicar uma resolucao diferente na inicializacao? (s/n)"
            if ($CustomChoice -match '^[sS]') {
                try {
                    $CustomResolutionInput = Read-Host "Digite sua resolucao preferida em ms"
                    [double]$CustomResolution = $CustomResolutionInput
                    if ($CustomResolution -le 0) {
                        throw "A resolucao deve ser maior que zero."
                    }
                    Create-StartupShortcut $CustomResolution
                    Start-Process ".\SetTimerResolution.exe" -ArgumentList @("--resolution", ($CustomResolution * 1E4), "--no-console")
                } catch {
                    Write-Warning "Entrada invalida. Nenhuma acao de inicializacao realizada."
                }
            } else {
                Write-Host "Nenhuma acao de inicializacao realizada." -ForegroundColor Yellow
            }
        }
    } else {
        $ApplyChoice = Read-Host "`nDo you want to apply this resolution ($RecommendedResolution ms) at system startup? (y/n)"
        if ($ApplyChoice -match '^[yY]') {
            Create-StartupShortcut $RecommendedResolution
            Start-Process ".\SetTimerResolution.exe" -ArgumentList @("--resolution", ($RecommendedResolution * 1E4), "--no-console")
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
                    Start-Process ".\SetTimerResolution.exe" -ArgumentList @("--resolution", ($CustomResolution * 1E4), "--no-console")
                } catch {
                    Write-Warning "Invalid input. No startup action performed."
                }
            } else {
                Write-Host "No startup action performed." -ForegroundColor Yellow
            }
        }
    }
}

# Initialize global language variable
$global:Language = "EN"

# First ask for language preference
Set-Language

# Continue with script execution
Test-AdminPrivileges
Enable-GlobalResolutionSupport
Assert-ExecutablesExist
$BenchmarkParameters = Get-BenchmarkParameters
$BestResolution = Run-Benchmark $BenchmarkParameters
Apply-StartupResolution $BestResolution
exit 0
