$configPath = "C:\install\configuration.xml"

Write-Host "--- STAGE: CREATING CONFIG ---"
# Optimized for the Bootstrapper/Standard Setup
$xmlContent = @"
<Configuration>
  <Display Level="None" CompletionNotice="No" SuppressModal="Yes" AcceptEula="Yes" />
  <Setting Id="SETUP_REBOOT" Value="Never" />
  <Setting Id="REBOOT" Value="ReallySuppress"/>
</Configuration>
"@
$xmlContent | Out-File -FilePath $configPath -Encoding utf8

Write-Host "--- STAGE: STARTING INSTALLATION ---"
# Using /config for the Bootstrapper setup.exe
$process = Start-Process -FilePath ".\OfficeInstall\setup.exe" -ArgumentList "/config `"$configPath`"" -PassThru -NoNewWindow -Wait

if ($process.ExitCode -ne 0) {
    Write-Host "--- ERROR: Setup failed with exit code $($process.ExitCode) ---"
    exit 1
}

Write-Host "--- SUCCESS: Office Installed ---"
