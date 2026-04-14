FROM mcr.microsoft.com/windows/servercore:ltsc2022

WORKDIR /install

# 1. Install necessary features for Office DCOM and Setup
RUN powershell -Command \
    Install-WindowsFeature Net-Framework-45-Core; \
    Install-WindowsFeature Web-WebServer; `
    Set-ExecutionPolicy Bypass -Scope Process -Force

# 2. Copy installation files
COPY ./MSOffice ./OfficeInstall
COPY ConfigureDCOM.ps1 .
COPY install_office.ps1 .

# 3. Run the installation
# We increase the priority of the process to ensure it gets enough CPU cycles
RUN powershell -NoProfile -ExecutionPolicy Bypass -File .\install_office.ps1

# 4. Run DCOM Configuration
RUN powershell -NoProfile -ExecutionPolicy Bypass -File .\ConfigureDCOM.ps1

# 5. Final Cleanup
RUN powershell -Command Remove-Item -Recurse -Force ./OfficeInstall, ./install_office.ps1, ./ConfigureDCOM.ps1, ./configuration.xml

CMD ["powershell"]
