FROM mcr.microsoft.com/windows/servercore:ltsc2022

WORKDIR /install

# Copy installation files
COPY ./MSOffice ./OfficeInstall
COPY ConfigureDCOM.ps1 .

# Switch to PowerShell
SHELL ["powershell", "-Command", "$ErrorActionPreference = 'Stop'; $ProgressPreference = 'SilentlyContinue';"]

# Every line below MUST end with a backtick (`) to stay part of the RUN command
RUN Write-Output '--- STAGE: CONFIGURING XML ---'; `
    $xml = '<Configuration><Display Level="none" CompletionNotice="no" SuppressModal="yes" AcceptEula="yes" /></Configuration>'; `
    $xml | Out-File -FilePath .\silent_config.xml -Encoding utf8; `
    Write-Output '--- STAGE: STARTING INSTALLATION ---'; `
    $logPath = 'C:\install\office_log.txt'; `
    $proc = Start-Process -FilePath .\OfficeInstall\setup.exe `
        -ArgumentList '/config .\silent_config.xml', '/log', $logPath `
        -PassThru; `
    $timeout = 1200; $timer = 0; `
    while (!$proc.HasExited -and $timer -lt $timeout) { `
        Start-Sleep -Seconds 30; $timer += 30; `
        Write-Warning "Installation in progress... ($timer seconds elapsed)"; `
    } `
    if (!$proc.HasExited) { `
        Write-Error 'ERROR: Installation timed out. Log Contents:'; `
        if (Test-Path $logPath) { Get-Content $logPath | ForEach-Object { Write-Error $_ } } `
        Stop-Process -Id $proc.Id -Force; `
        exit 1; `
    } `
    Write-Output "Finished with Exit Code: $($proc.ExitCode)"; `
    if ($proc.ExitCode -ne 0) { `
        if (Test-Path $logPath) { Get-Content $logPath | ForEach-Object { Write-Error $_ } } `
        exit 1; `
    }

# Cleanup
RUN Remove-Item -Recurse -Force ./OfficeInstall, ./silent_config.xml
