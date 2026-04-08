FROM mcr.microsoft.com/windows/servercore:ltsc2022

# Copy Office setup and scripts
WORKDIR /install
COPY ./MSOffice ./OfficeInstall
COPY ConfigureDCOM.ps1 .

# Generate a silent config and Install Excel
RUN powershell -Command \
    $xml = '<Configuration><Display Level=\"none\" CompletionNotice=\"no\" SuppressModal=\"yes\" AcceptEula=\"yes\" /></Configuration>'; \
    $xml | Out-File -FilePath .\silent_config.xml -Encoding utf8; \
    Start-Process -FilePath .\OfficeInstall\setup.exe -ArgumentList '/config .\silent_config.xml' -Wait

# Run the DCOM and User configuration
RUN powershell -File C:/install/ConfigureDCOM.ps1

# Cleanup installer files to keep image size down
RUN powershell -Command Remove-Item -Recurse -Force C:/install/OfficeInstall

CMD ["powershell"]
