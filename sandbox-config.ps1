Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))

choco feature enable -n=allowGlobalConfirmation

choco install vscode --ignore-checksum

choco install thunderbird --ignore-checksum

choco install python --ignore-checksum

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
        <taskbar:DesktopApp DesktopApplicationLinkPath="C:\Program Files\Google\Chrome\Application\chrome.exe"/>
        <taskbar:DesktopApp DesktopApplicationLinkPath="C:\Windows\explorer.exe"/>
        <taskbar:DesktopApp DesktopApplicationLinkPath="C:\Program Files\Microsoft VS Code\Code.exe"/>
        <taskbar:DesktopApp DesktopApplicationLinkPath="%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe"/>
      </taskbar:TaskbarPinList>
    </defaultlayout:TaskbarLayout>
  </CustomTaskbarLayoutCollection>
</LayoutModificationTemplate>
"@

function TweakRegedit($path, $name, $value){    
    IF(!(Test-Path -Path $path)) { 
        $basePath = $path.SubString(0, $path.LastIndexOf("\"))
        $key = $path.split("\")[-1]
        New-Item -Path $basePath -Name $key
    } 
    Set-ItemProperty -Path $path -Name $Name -Value $value 
}

function TweakApplyStartTaskbarLayout(){

    

    $Layout = $LAYOUT_START_MENU_TASKBAR_BLANK

    #Delete layout file if it already exists
    If(Test-Path C:\Windows\StartLayout.xml)
    {
        Remove-Item C:\Users\WDAGUtilityAccount\Desktop\StartLayout.xml
    }

    #Creates the blank layout file
    $Layout | Out-File C:\Users\WDAGUtilityAccount\Desktop\StartLayout.xml -Encoding ASCII

    #Assign the start layout and force it to apply with "LockedStartLayout" at both the machine and user level
    $regAliases = @("HKLM", "HKCU")
    foreach ($regAlias in $regAliases){
        $basePath = $regAlias + ":\SOFTWARE\Policies\Microsoft\Windows"
        $keyPath = $basePath + "\Explorer" 
        TweakRegedit -Path $keyPath -Name "LockedStartLayout" -Value 1
        TweakRegedit -Path $keyPath -Name "StartLayoutFile" -Value "C:\Users\WDAGUtilityAccount\Desktop\StartLayout.xml"
    }

    #Set taskbar alignment left
    REG ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /f /v TaskbarAl /t REG_DWORD /d 0

    #Disable chat icon
    REG ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /f /v TaskbarMn /t REG_DWORD /d 0

    #Disable task view icon
    REG ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /f /v ShowTaskViewButton /t REG_DWORD /d 0

    #Disable W11 context menus
    REG ADD "HKEY_CURRENT_USER\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32" /f /ve

    #Restart Explorer, press the Windows key (necessary to load the new layout), and give it a few seconds to process
    Stop-Process -name explorer
}

TweakApplyStartTaskbarLayout

$pythonScriptUrl = "https://raw.githubusercontent.com/robbycuenot/frequent-senders/main/frequent-senders.py"
Invoke-WebRequest -Uri $pythonScriptUrl -OutFile ~\Desktop\frequent-senders.py

# Make `refreshenv` available right away, by defining the $env:ChocolateyInstall
# variable and importing the Chocolatey profile module.
# Note: Using `. $PROFILE` instead *may* work, but isn't guaranteed to.
$env:ChocolateyInstall = Convert-Path "$((Get-Command choco).Path)\..\.."   
Import-Module "$env:ChocolateyInstall\helpers\chocolateyProfile.psm1"

# refreshenv is now an alias for Update-SessionEnvironment
# (rather than invoking refreshenv.cmd, the *batch file* for use with cmd.exe)
# This should make git.exe accessible via the refreshed $env:PATH, so that it
# can be called by name only.
refreshenv

python -m pip install --upgrade pip

pip install tqdm

pip install matplotlib

cd ~\Desktop

echo "cd ~\Desktop" >> $PSHOME\Profile.ps1

Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope LocalMachine -Force

Write-Host "\nSetup complete!"
