# TimerRes

## An All-in-One Timer Resolution Benchmark and Setup Tool

TimerRes is a comprehensive PowerShell script for optimizing your Windows system's timer resolution. It helps identify and apply the optimal resolution value for your specific hardware, improving system responsiveness and gaming performance.

## Features

- **Automated Benchmarking**: Tests a range of timer resolution values to find the optimal setting for your system
- **Scientific Approach**: Measures actual sleep precision across multiple samples to ensure reliable results
- **Automatic Setup**: Creates startup shortcuts to apply your optimal resolution at system boot
- **Registry Configuration**: Enables Windows global timer resolution support automatically
- **Comprehensive Results**: Provides detailed analysis of each resolution's performance characteristics
- **Self-contained**: Automatically downloads required components if needed

## Requirements

- Windows 10/11
- PowerShell 5.1 or later
- Administrator privileges
- Internet connection (only if executables need to be downloaded)

## Quick Start

1. Right-click the script and select "Run with PowerShell as Administrator"
2. Follow the prompts to set benchmark parameters
3. Wait for the benchmark to complete
4. Apply the recommended resolution at startup when prompted

## How It Works

TimerRes runs a comprehensive benchmark across a range of timer resolution values. For each resolution:

1. The Windows timer resolution is set using SetTimerResolution.exe
2. Multiple test runs with MeasureSleep.exe measure the actual sleep precision
3. Statistical analysis determines the optimal value based on:
   - Average delta (difference between requested and actual sleep time)
   - Standard deviation (stability/consistency of sleep precision)
   - Overall score (weighted combination of metrics)

The script provides detailed results and automatically offers to apply the best-performing resolution at system startup.

## Advanced Usage

You can run the script with custom parameters:

```powershell
.\TimerRes.ps1 -Samples 50 -Increment 0.001 -QuickTest
```

Parameters:
- **Samples**: Number of measurements per resolution (default: 30)
- **Increment**: Step size between tested resolutions in ms (default: 0.002)
- **QuickTest**: Run a faster benchmark with fewer samples (default: false)

## Understanding Results

- **Lowest Average Delta**: Resolution with the smallest average difference between requested and actual sleep time
- **Most Stable**: Resolution with the smallest standard deviation (most consistent timing)
- **Highest Consistency**: Resolution with the best ratio between delta and standard deviation
- **Highest Overall Score**: Resolution with the best balanced performance across all metrics

## Credits

- **SetTimerResolution.exe & MeasureSleep.exe**: Developed by [Amit](https://github.com/valleyofdoom)

## License

This script is released under the MIT License. The bundled executables may have their own licensing terms.

## Disclaimer

Modifying system timer resolution could affect power consumption and battery life on laptop devices. Use at your own risk.
