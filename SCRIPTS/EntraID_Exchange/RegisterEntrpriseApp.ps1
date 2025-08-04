Connect-MgGraph -Scopes "Application.ReadWrite.All", "Directory.ReadWrite.All"
$ServicePrincipalID=@{
  "AppId" = "ec156f81-f23a-47bd-b16f-9fb2c66420f9"
  }
New-MgServicePrincipal -BodyParameter $ServicePrincipalId |
  Format-List id, DisplayName, AppId, SignInAudience