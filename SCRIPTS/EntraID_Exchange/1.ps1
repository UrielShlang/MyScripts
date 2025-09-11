Function TextMessage  {
 [CmdletBinding()]


$body = @{
     From = $Twilio_NUMBER;
     To = $User_NUMBER;
     Body = "Test" | Out-String
 }
  # Create a credential object for HTTP basic auth
 $p = $token | ConvertTo-SecureString -asPlainText -Force
 $credential = New-Object System.Management.Automation.PSCredential($sid, $p)
Invoke-WebRequest $url -Method Post -Credential $credential -Body $body -UseBasicParsing | ConvertFrom-Json | Select-Object sid, body  
}
TextMessage