$configPath = "C:\install\configuration.xml"

Write-Host "--- STAGE: CREATING ODT CONFIG ---"
# This XML is specifically designed for the Office Deployment Tool
$xmlContent = @"
<Configuration>
  <Add SourcePath="C:\install\OfficeInstall" OfficeClientEdition="64">
    <Product ID="O365ProPlusRetail">
      <Language ID="en-us" />
    </Product>
  </Add>
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="AUTOACTIVATE" Value="1" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Property Name="SharedComputerLicensing" Value="1" />
</Configuration>
"@
$xmlContent | Out-File -FilePath $configPath -Encoding utf8

Write-Host "--- STAGE: STARTING ODT INSTALLATION ---"
# Note: Using /configure instead of /config
$process = Start-Process -FilePath ".\OfficeInstall\setup.exe" -ArgumentList "/configure `"$configPath`"" -PassThru -NoNewWindow -Wait

if ($process.ExitCode -ne 0) {
    Write-Host "--- ERROR: Setup failed with exit code $($process.ExitCode) ---"
    exit 1
}

Write-Host "--- SUCCESS: Office Installed ---"
