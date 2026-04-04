FROM mcr.microsoft.com/powershell:7.4-ubuntu-22.04

# Ensure $env:TEMP is set (Linux base image does not set it)
ENV TEMP=/tmp
ENV TMPDIR=/tmp

# Install OS-level dependencies:
#   dnsutils — provides dig for DMARC/SPF/DKIM DNS lookups in Invoke-MailSecurityAudit
RUN apt-get update && apt-get install -y --no-install-recommends dnsutils && rm -rf /var/lib/apt/lists/*

# Install PowerShell modules in a single layer to reduce image size.
# Graph SDK sub-modules are listed individually rather than the full
# Microsoft.Graph meta-package to keep the image smaller.
SHELL ["pwsh", "-Command"]
RUN Set-PSRepository -Name PSGallery -InstallationPolicy Trusted; \
    Install-Module -Name Az.Accounts                                -Scope AllUsers -Force; \
    Install-Module -Name Az.KeyVault                                -Scope AllUsers -Force; \
    Install-Module -Name Microsoft.Graph.Authentication             -Scope AllUsers -Force; \
    Install-Module -Name Microsoft.Graph.Applications               -Scope AllUsers -Force; \
    Install-Module -Name Microsoft.Graph.DeviceManagement           -Scope AllUsers -Force; \
    Install-Module -Name Microsoft.Graph.DeviceManagement.Enrollment -Scope AllUsers -Force; \
    Install-Module -Name Microsoft.Graph.Devices.CorporateManagement -Scope AllUsers -Force; \
    Install-Module -Name Microsoft.Graph.Groups                     -Scope AllUsers -Force; \
    Install-Module -Name Microsoft.Graph.Identity.DirectoryManagement -Scope AllUsers -Force; \
    Install-Module -Name Microsoft.Graph.Identity.SignIns            -Scope AllUsers -Force; \
    Install-Module -Name Microsoft.Graph.Reports                    -Scope AllUsers -Force; \
    Install-Module -Name Microsoft.Graph.Users                      -Scope AllUsers -Force; \
    Install-Module -Name ExchangeOnlineManagement                   -Scope AllUsers -Force

WORKDIR /app
COPY . /app/

ENTRYPOINT ["pwsh", "-NoProfile", "-File", "/app/container-entrypoint.ps1"]
