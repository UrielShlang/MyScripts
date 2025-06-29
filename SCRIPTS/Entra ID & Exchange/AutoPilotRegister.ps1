[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$env:Path += ";C:\Program Files\WindowsPowerShell\Scripts"
 Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned
Install-PackageProvider -Name NuGet -RequiredVersion 2.8.5.208 -Force
Set-psrepository -name psgallery -installationpolicy Trusted
 Install-Script -Name Get-WindowsAutopilotInfo -Force
 Get-WindowsAutopilotInfo -Online -assign
