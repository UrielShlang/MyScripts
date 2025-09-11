Function TextMessage  {
 [CmdletBinding()]
 Param
 (
     # My Twilio Number
     [Parameter(Mandatory=$false)]
     [ValidateNotNullOrEmpty()]
     [string]$Twilio_NUMBER='+18147184353',
     # My Test Number
     [Parameter(Mandatory=$false)]
     [ValidateNotNullOrEmpty()]
     [string]$User_NUMBER='',
     # Twilio Auth SID
     [Parameter(Mandatory=$false)]
     [ValidateNotNullOrEmpty()]
     [string]$sid = "",
     # Twilio Auth Token
     [Parameter(Mandatory=$false)]
     [ValidateNotNullOrEmpty()]
     [string]$token = "",
     # Twilio API endpoint and POST params
     [Parameter(Mandatory=$false)]
     [ValidateNotNullOrEmpty()]
     [string]$url = "https://api.twilio.com/2010-04-01/Accounts/$sid/Messages.json"
 )

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