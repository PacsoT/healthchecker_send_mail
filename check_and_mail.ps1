cls

$testing= $true;

$health_checker_file_location = ".\ExchangeAllServersReport.html"
$exceptions_file_location = ".\exceptions.csv"



[Net.ServicePointManager]::SecurityProtocol =[Net.SecurityProtocolType]::Tls12

del .\HealthChecker*.txt
del .\HealthChecker*.xml
del .\ExchangeAllServersReport*.*

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn; 

.\HealthChecker.ps1 -ScriptUpdateOnly
.\HealthChecker.ps1
.\HealthChecker.ps1 -BuildHtmlServersReport -HtmlReportFile $health_checker_file_location;




$the_text = Get-Content $health_checker_file_location -Raw 
$total_number_of_occurances = (Select-String -Path $health_checker_file_location -Pattern "<td class=""Red"">" -AllMatches).Matches.Count



$start_key="<td class=""Red"">"
$end_key="</td>"
$next_start_key_index = 0;
$next_end_key_index = 1;

$custom_certificate_warning_treshold = 10;     # The Health Checker cries about certificates 30 days before it expires, here we can plower that number...

$occurances = @();
$exceptions = @();
$found_errors = @();


For ($i=0; $i -lt $total_number_of_occurances; $i++)
 {
  $next_start_key_index = $the_text.IndexOf($start_key,$next_end_key_index)
  $next_end_key_index = $the_text.IndexOf($end_key,$next_start_key_index)
  $occurances += $the_text.Substring($next_start_key_index+$start_key.Length,$next_end_key_index-$next_start_key_index-$start_key.Length)
  
  if($testing)
   {
    $next_start_key_index;
    $next_end_key_index;
   } 
 } 


if($testing)
 {
  $occurances.Count
  $total_number_of_occurances;

  Foreach($occ in $occurances)
   {
    Write-host "Occurance"
    Write-host "-------------------"
    Write-host $occ -ForegroundColor Gray
    Write-host "===================" -ForegroundColor White
    Write-host
   }
 }

$exceptions_obj= Import-Csv -Path $exceptions_file_location
foreach ($exc in $exceptions_obj) {$exceptions+=$exc.Name}
$we_found_new_errors = $false;

For ($i=0; $i -lt $total_number_of_occurances; $i++)
 {
  if (-not("" -eq ($occurances[$i]))) # megvizsgáljuk, hogy üres string-e a $occ...
   {
    if (-not ($null -eq ($occurances[$i] -as [int])))      #megvizsgaljuk, hogy szám-e az $occ...
     {
      if (($occurances[$i] -as [int])-le $custom_certificate_warning_treshold)         #megvizsgáljuk, hogy alacsonyabb-e az érték, mint a mi egyedi érték
       {
        $we_found_new_errors=$true;
        $found_errors += $occurances[$i]
       }
     }
    else    # valami számmá nem alakítható érték van az $occ-ben...
     {
      if(-not($exceptions.Contains($occurances[$i])))        # Megvizsgáljuk, hogy az érték bvenne van-e a kivételek listájába...
       {
        $we_found_new_errors=$true;
        $found_errors += $occurances[$i]
        
       } 
     }
    }
 }
        
  
 
 
if($we_found_new_errors)
 {
  if($testing)
   {
    Write-host "We're gon' send an e-mail, because we found some new errors!"
   }

  $From = "user@comany.com"
  $Subject = "Cause for concern was found on the Exchange server"

  $SMTPServer = "localhost"
  $SMTPPort = "25"

  

  
  if($testing)
   {
    $to = "tamas.pacso@smp.hu"
   }
  else
   {
    $to = "support@silicondirect.net"
   }


   $body = @'
             <h2 style="font-family:Corbel"> <b>Hi!</b> Unfortunatelly I found some errors on the Exchange server! </h2>
             <p style="font-family:Corbel">Here is a list of the concerning messages I extracted from the Health Checker's HTML report...</p>
             <ul>
             
'@;

foreach($err in $found_errors)
 {
  $body+= "<li style=""font-family:'Century Gothic';color:#993333"">"+ $err +"</li>"
 }

 $body+="</ul>
  <p style=""font-family:Corbel; padding-top:30px; align:center"">And here is the complete list of texts we have choosen to ignore: </p>
  <ul>
  "


foreach($exc in $exceptions)
 {
  $body+= "<li style=""font-family:'Century Gothic';color:#444444"">"+ $exc +"</li>"
 }
 $body+="</ul>
 <p style=""font-family:Corbel; text-align: center;"">Attached you can find the complete report, please do have a look!</p>
 "

  
  Send-MailMessage -From $From -to $To -Subject $Subject -Body $body -BodyAsHtml -SmtpServer $SMTPServer -Port $SMTPPort  -Encoding UTF8 -Attachments $health_checker_file_location
 
 
 }

  






  #
 ###
 ###   Useful command repository
 ###
  #
  #
  ##  With this, we can export all occurances, so we can create the exceptions file...
  #
  ### $occurances |Select-Object -unique @{Name='Name';Expression={$_}} | Export-Csv -NoTypeInformation -Path $exceptions_file_location

  #
  #
  ##   Ha csak a riportoto akarjuk elküldeni
  #   
  ###  Send-MailMessage -From $From -to $To -Subject $Subject -Body $the_text -BodyAsHtml -SmtpServer $SMTPServer -Port $SMTPPort  -Encoding UTF8