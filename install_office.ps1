$ErrorActionPreference = 'Stop'
Write-Output '--- STAGE: CONFIGURING XML ---'
# Legacy Bootstrapper XML format
$xml = @"
<Configuration>
    <Display Level="none" CompletionNotice="no" SuppressModal="yes" AcceptEula="yes" />
    <Setting Id="SETUP_REBOOT" Value="Never" />
    <OptionState Id="GWW_Common" State="Local" Children="force" />
</Configuration>
"@
$xml | Out-File -FilePath .\silent_config.xml -Encoding utf8

Write-Output '--- STAGE: STARTING INSTALLATION ---'
$logPath = 'C:\install\office_log.txt'

# For 'Bootstrapper', we use /config. 
# We also use 'Start-Process' but we will manually poll the log file.
$proc = Start-Process -FilePath .\OfficeInstall\setup.exe `
    -ArgumentList "/config .\silent_config.xml" `
    -PassThru

$timeout = 1800; $timer = 0
Write-Output "Watching log file for activity..."

while (!$proc.HasExited -and $timer -lt $timeout) {
    Start-Sleep -Seconds 30; $timer += 30
    
    # Try to read the log file WHILE it is installing to see the real-time error
    if (Test-Path $logPath) {
        $lastLine = Get-Content $logPath -Tail 1
        Write-Warning "[$timer s] Current Log: $lastLine"
    } else {
        Write-Warning "[$timer s] Installation in progress... (Log file not created yet)"
    }
}

if (!$proc.HasExited) {
    Write-Error "!! TIMEOUT !! Setup is stuck. Final 20 lines of log:"
    if (Test-Path $logPath) { Get-Content $logPath -Tail 20 | Write-Host }
    Stop-Process -Id $proc.Id -Force
    exit 1
}

Write-Output "Finished with Exit Code: $($proc.ExitCode)"
if ($proc.ExitCode -ne 0 -and $proc.ExitCode -ne 3010) {
    Write-Error "Setup failed. Checking logs..."
    if (Test-Path $logPath) { Get-Content $logPath -Tail 50 | Write-Host }
    exit 1
}
