try {
    $progressPreference = 'silentlyContinue'
    New-Item -Path $env:ProgramData\ -Name CustomScripts -ItemType Directory -Force -ErrorAction SilentlyContinue
    $InstallerFolder = $(Join-Path $env:ProgramData CustomScripts)
    
    Write-Host "Downloading WinGet and its dependencies..."
    Write-Information "Downloading WinGet and its dependencies..."
    $InstallerFolder = $(Join-Path $env:ProgramData CustomScripts)
        
    Invoke-WebRequest -Uri https://aka.ms/getwinget -OutFile $InstallerFolder\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle
    Invoke-WebRequest -Uri https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx -OutFile $InstallerFolder\Microsoft.VCLibs.x64.14.00.Desktop.appx
    Invoke-WebRequest -Uri https://github.com/microsoft/microsoft-ui-xaml/releases/download/v2.7.3/Microsoft.UI.Xaml.2.7.x64.appx -OutFile $InstallerFolder\Microsoft.UI.Xaml.2.7.x64.appx
    Add-AppxProvisionedPackage -Online -SkipLicense -PackagePath $InstallerFolder\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle -DependencyPackagePath $InstallerFolder\Microsoft.VCLibs.x64.14.00.Desktop.appx,$InstallerFolder\Microsoft.UI.Xaml.2.7.x64.appx
    
    $TestWinget = Get-AppxProvisionedPackage -Online | Where-Object {$_.DisplayName -eq "Microsoft.DesktopAppInstaller"}
    If ([Version]$TestWinGet.Version -gt "2022.506.16.0") 
	{
        Remove-Item -Path "$InstallerFolder\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle" -Force -ErrorAction Continue
        Remove-Item -Path "$InstallerFolder\Microsoft.VCLibs.x64.14.00.Desktop.appx" -Force -ErrorAction Continue
        Remove-Item -Path "$InstallerFolder\Microsoft.UI.Xaml.2.7.x64.appx" -Force -ErrorAction Continue
       
		Write-Host "WinGet is Installed" 
        exit 0
	}
    else {
        Write-Host "WinGet not Installed" 
        exit 1
    }
    }
catch
    {
    $errMsg = $_.Exception.Message
    write-host $errMsg
    exit 1
     }

