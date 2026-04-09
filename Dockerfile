FROM mcr.microsoft.com/windows/servercore:ltsc2022

# Set PowerShell as the default shell for all RUN commands
# This prevents the "< was unexpected" error by avoiding CMD's parser
SHELL ["powershell", "-Command", "$ErrorActionPreference = 'Stop'; $ProgressPreference = 'SilentlyContinue';"]

# Set working directory
WORKDIR /install

# Copy Office setup and scripts from the build context
COPY ./MSOffice ./OfficeInstall
COPY ConfigureDCOM.ps1 .

# Generate a silent config and Install Excel
# We use ascii encoding to ensure the installer can read the XML properly
RUN $xml = '<Configuration><Display Level="none" CompletionNotice="no" SuppressModal="yes" AcceptEula="yes" /></Configuration>'; \
    $xml | Out-File -FilePath .\silent_config.xml -Encoding ascii; \
    Write-Host "Starting Office Installation..."; \
    $process = Start-Process -FilePath .\OfficeInstall\setup.exe -ArgumentList '/config .\silent_config.xml' -Wait -PassThru; \
    if ($process.ExitCode -ne 0) { throw "Installation failed with exit code $($process.ExitCode)" }

# Run the DCOM and User configuration
RUN .\ConfigureDCOM.ps1

# Cleanup installer files to keep image size down
RUN Remove-Item -Recurse -Force ./OfficeInstall

CMD ["powershell"]
