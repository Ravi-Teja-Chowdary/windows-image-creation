# Use Server Core 2022 for stability and smaller size
FROM mcr.microsoft.com/windows/servercore:ltsc2022

WORKDIR /install

# Copy installation files
COPY ./MSOffice ./OfficeInstall
COPY ConfigureDCOM.ps1 .

# Switch to PowerShell for easier logic and logging
SHELL ["powershell", "-Command", "$ErrorActionPreference = 'Stop'; $ProgressPreference = 'SilentlyContinue';"]

RUN Write-Host '--- Stage 1: Creating Configuration XML ---'; `
    $xml = '<Configuration><Display Level="none" CompletionNotice="no" SuppressModal="yes" AcceptEula="yes" /></Configuration>'; `
    $xml | Out-File -FilePath .\silent_config.xml -Encoding utf8; `
    `
    Write-Host '--- Stage 2: Starting Office Setup ---'; `
    $logPath = 'C:\install\office_log.txt'; `
    # We use /config for older setup.exe or /configure for ODT setup.exe
    # Adding /log to capture internal setup details
    $proc = Start-Process -FilePath .\OfficeInstall\setup.exe `
        -ArgumentList '/config .\silent_config.xml', '/log', $logPath `
        -PassThru; `
    `
    Write-Host 'Waiting for installation to complete (Timeout: 20 mins)...'; `
    # Monitor the process. If it hangs, we catch it.
    $timeout = 1200; $timer = 0; `
    while (!$proc.HasExited -and $timer -lt $timeout) { `
        Start-Sleep -Seconds 30; $timer += 30; `
        Write-Host "Still installing... ($timer seconds elapsed)"; `
    } `
    `
    if (!$proc.HasExited) { `
        Write-Host 'ERROR: Installation timed out. Printing logs...'; `
        if (Test-Path $logPath) { Get-Content $logPath } `
        Stop-Process -Id $proc.Id -Force; `
        exit 1; `
    } `
    `
    Write-Host '--- Stage 3: Installation Finished ---'; `
    Write-Host "Exit Code: $($proc.ExitCode)"; `
    if ($proc.ExitCode -ne 0) { `
        if (Test-Path $logPath) { Get-Content $logPath } `
        exit 1; `
    }

# Final cleanup of install files to keep image size down
RUN Remove-Item -Recurse -Force ./OfficeInstall, ./silent_config.xml
