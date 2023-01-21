# Install chocolatey and allow scripts to run
Set-ExecutionPolicy Bypass -Scope LocalMachine -Force
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
choco feature enable -n=allowGlobalConfirmation

# Install packages using chocolatey
choco install thunderbird --ignore-checksum
choco install python --ignore-checksum

# Install VSCode and extensions for displaying PDFs and CSVs
Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force; 
Install-Script Install-VSCode -Force; 
Install-VSCode.ps1 -AdditionalExtensions 'tomoki1207.pdf', 'ms-python.python', 'GrapeCity.gc-excelviewer'

# Deploy settings to associate PDFs and CSVs with the previously installed extensions in VSCode
$VSCODESETTINGS = @"
{
    "workbench.editorAssociations": {
        "*.pdf": "pdf.preview",
        "*.csv": "gc-excelviewer-csv-editor"
    },
    "window.newWindowDimensions": "maximized"
}
"@

$VSCODESETTINGS | Out-File C:\Users\WDAGUtilityAccount\AppData\Roaming\Code\User\Settings.json -Encoding ASCII

# Function to create a registry key if it doesn't exist, and set a value for that key
# This is used for adjusting the taskbar layout and other settings within the sandbox
function Regedit($path, $name, $value){    
    IF(!(Test-Path -Path $path)) { 
        $basePath = $path.SubString(0, $path.LastIndexOf("\"))
        $key = $path.split("\")[-1]
        New-Item -Path $basePath -Name $key
    } 
    Set-ItemProperty -Path $path -Name $Name -Value $value 
}

# Define a taskbar layout
$LAYOUT_START_MENU_TASKBAR_BLANK = @"
<?xml version="1.0" encoding="utf-8"?>
<LayoutModificationTemplate
    xmlns="http://schemas.microsoft.com/Start/2014/LayoutModification"
    xmlns:defaultlayout="http://schemas.microsoft.com/Start/2014/FullDefaultLayout"
    xmlns:start="http://schemas.microsoft.com/Start/2014/StartLayout"
    xmlns:taskbar="http://schemas.microsoft.com/Start/2014/TaskbarLayout"
    Version="1">
  <CustomTaskbarLayoutCollection PinListPlacement="Replace">
    <defaultlayout:TaskbarLayout>
      <taskbar:TaskbarPinList>
        <taskbar:DesktopApp DesktopApplicationLinkPath="C:\Program Files\Mozilla Thunderbird\thunderbird.exe"/>
        <taskbar:DesktopApp DesktopApplicationLinkPath="C:\Windows\explorer.exe"/>
        <taskbar:DesktopApp DesktopApplicationLinkPath="C:\Program Files\Microsoft VS Code\Code.exe"/>
        <taskbar:DesktopApp DesktopApplicationLinkPath="%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe"/>
      </taskbar:TaskbarPinList>
    </defaultlayout:TaskbarLayout>
  </CustomTaskbarLayoutCollection>
</LayoutModificationTemplate>
"@

# Function to apply the taskbar layout
function ApplyTaskbarLayout(){
    $Layout = $LAYOUT_START_MENU_TASKBAR_BLANK

    # Creates the blank layout file
    $Layout | Out-File C:\Users\WDAGUtilityAccount\Desktop\StartLayout.xml -Encoding ASCII

    # Assign the start layout and force it to apply with "LockedStartLayout" at both the machine and user level
    $regAliases = @("HKLM", "HKCU")
    foreach ($regAlias in $regAliases){
        $basePath = $regAlias + ":\SOFTWARE\Policies\Microsoft\Windows"
        $keyPath = $basePath + "\Explorer" 
        Regedit -Path $keyPath -Name "LockedStartLayout" -Value 1
        Regedit -Path $keyPath -Name "StartLayoutFile" -Value "C:\Users\WDAGUtilityAccount\Desktop\StartLayout.xml"
    }

    # Set taskbar alignment left
    REG ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /f /v TaskbarAl /t REG_DWORD /d 0

    # Disable chat icon
    REG ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /f /v TaskbarMn /t REG_DWORD /d 0

    # Disable task view icon
    REG ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /f /v ShowTaskViewButton /t REG_DWORD /d 0
}

ApplyTaskbarLayout

# Disable W11 context menus
REG ADD "HKEY_CURRENT_USER\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32" /f /ve

# Restart Explorer (necessary to load the new layout)
Stop-Process -name explorer

# Download the python script
$pythonScriptUrl = "https://raw.githubusercontent.com/robbycuenot/frequent-senders/main/frequent-senders.py"
Invoke-WebRequest -Uri $pythonScriptUrl -OutFile ~\Desktop\frequent-senders.py

# To allow commands such as 'pip' to be run within the same session, we need to
# refresh the environment variables. This is done by running the `refreshenv`
# command, which is defined in the `chocolateyProfile.psm1` module. However, this
# module is not loaded until the next terminal session, so we need to load it
# manually.

# Make `refreshenv` available right away, by defining the $env:ChocolateyInstall
# variable and importing the Chocolatey profile module.
$env:ChocolateyInstall = Convert-Path "$((Get-Command choco).Path)\..\.."   
Import-Module "$env:ChocolateyInstall\helpers\chocolateyProfile.psm1"

# refreshenv is now an alias for Update-SessionEnvironment
refreshenv

# Update pip and install dependencies
python -m pip install --upgrade pip
pip install tqdm
pip install matplotlib

# Change current and default directory to Desktop
cd ~\Desktop
echo "cd ~\Desktop" >> $PSHOME\Profile.ps1

Write-Host "Setup complete!"
