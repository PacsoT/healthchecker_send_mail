[Net.ServicePointManager]::SecurityProtocol =[Net.SecurityProtocolType]::Tls12
del .\HealthChecker.ps1
del *.htm
del *html
del *.xml
del *.txt
wget https://github.com/microsoft/CSS-Exchange/releases/latest/download/HealthChecker.ps1 -OutFile .\HealthChecker.ps1

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn; 

Get-ExchangeServer | ?{$_.AdminDisplayVersion -Match "^Version 15"} | %{.\HealthChecker.ps1 -Server $_.Name}; .\HealthChecker.ps1 -BuildHtmlServersReport;



IF (@( Get-Content .\ExchangeAllServersReport.html | Where-Object { $_.Contains("class=""Red""")  } ).Count -gt 0)
{ 
  $From = "nagya@om.hu"
  $To = "kis@om.hu”
  $Subject = "New vulnerability found on your Exchange environment!"
  $Body = Get-Content .\ExchangeAllServersReport.html -Raw 
  $SMTPServer = "whosyourdaddy.com"
  $SMTPPort = "25"

  $password = ConvertTo-SecureString 'RepenisF9' -AsPlainText -Force
  $creds = New-Object System.Management.Automation.PSCredential ('mailer@exch_server.hu', $password)

  Send-MailMessage -From $From -to $To -Subject $Subject -Body $Body -BodyAsHtml -SmtpServer $SMTPServer -Port $SMTPPort -Encoding UTF8 

}
else
{ Write-host "All is well!"}
 
