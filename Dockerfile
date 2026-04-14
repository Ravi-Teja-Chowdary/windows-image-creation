FROM mcr.microsoft.com/windows/servercore:ltsc2022

WORKDIR /install

# 1. Prepare Environment
RUN powershell -Command \
    Install-WindowsFeature Net-Framework-45-Core; \
    Set-ExecutionPolicy Bypass -Scope Process -Force

# 2. Copy installation files (Now verified by the GH Action step)
COPY ./MSOffice ./OfficeInstall
COPY ConfigureDCOM.ps1 .
COPY install_office.ps1 .

# 3. Run the installation
RUN powershell -NoProfile -ExecutionPolicy Bypass -File .\install_office.ps1

# 4. Run DCOM Configuration
RUN powershell -NoProfile -ExecutionPolicy Bypass -File .\ConfigureDCOM.ps1

# 5. Final Cleanup
RUN powershell -Command Remove-Item -Recurse -Force ./OfficeInstall, ./install_office.ps1, ./ConfigureDCOM.ps1, ./configuration.xml

CMD ["powershell"]
