FROM mcr.microsoft.com/windows/servercore:ltsc2022

WORKDIR /install

# 1. Copy everything needed
COPY ./MSOffice ./OfficeInstall
COPY ConfigureDCOM.ps1 .
COPY install_office.ps1 .

# 2. Run the Installation script
RUN powershell -Command New-Item -Path 'C:\Users\ContainerUser\AppData\Local\Temp' -ItemType Directory -Force
RUN powershell -ExecutionPolicy Bypass -File .\install_office.ps1

# 3. Run the DCOM Configuration script
RUN powershell -ExecutionPolicy Bypass -File .\ConfigureDCOM.ps1

# 4. Final Cleanup
RUN powershell -Command Remove-Item -Recurse -Force ./OfficeInstall, ./install_office.ps1, ./ConfigureDCOM.ps1, ./silent_config.xml
