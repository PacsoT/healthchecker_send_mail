param (
   [switch]$Testing = $false,
   [switch]$Verbose = $false,
   [switch]$DumpOccurances = $false
)


cls

$health_checker_file_location = ".\ExchangeAllServersReport.html"
$runlog_file_location= ".\check_and_mail.log"

$exceptions_file_location = ".\exceptions.csv"
if (-not(Test-Path $exceptions_file_location -PathType Leaf)) 
 {
  $DumpOccurances=$true
 } 

$log_entry  = Get-Date -Format "yyyy.mm.dd HH:mm:ss"
$log_entry  += ' Script started.';

Add-Content -Path $runlog_file_location -Value $log_entry

[Net.ServicePointManager]::SecurityProtocol =[Net.SecurityProtocolType]::Tls12

del .\HealthChecker*.txt
del .\HealthChecker*.xml
del .\ExchangeAllServersReport*.*

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn; 

.\HealthChecker.ps1 -ScriptUpdateOnly
[System.GC]::Collect()
.\HealthChecker.ps1
[System.GC]::Collect()
.\HealthChecker.ps1 -BuildHtmlServersReport -HtmlReportFile $health_checker_file_location;


$log_entry  = Get-Date -Format "yyyy.mm.dd HH:mm:ss"
$log_entry  += ' Health Checker script was updated then called, and finished.';
Add-Content -Path $runlog_file_location -Value $log_entry




$the_text = Get-Content $health_checker_file_location -Raw 
$total_number_of_occurances = (Select-String -Path $health_checker_file_location -Pattern "<td class=""Red"">" -AllMatches).Matches.Count



$start_key="<td class=""Red"">"
$end_key="</td>"
$next_start_key_index = 0;
$next_end_key_index = 1;

$custom_certificate_warning_treshold = 10;     # A health checker 30 nap után nyávog a tanusítvány lejárat előtt, de ezt mi itt felülbírálhatjuk...

$occurances = @();
$exceptions = @();
$found_errors = @();


For ($i=0; $i -lt $total_number_of_occurances; $i++)
 {
  $next_start_key_index = $the_text.IndexOf($start_key,$next_end_key_index)
  $next_end_key_index = $the_text.IndexOf($end_key,$next_start_key_index)
  $occurances += $the_text.Substring($next_start_key_index+$start_key.Length,$next_end_key_index-$next_start_key_index-$start_key.Length)
  
  if($Verbose)
   {
    Write-host "Found one between "$next_start_key_index, " and ", $next_end_key_index;
   }
 }


if($Verbose)
 {
  Write-host "We have this many Occurances: "$occurances.Count
  Foreach($occ in $occurances)
   {
    Write-host "Occurance"
    Write-host "-------------------"
    Write-host $occ -ForegroundColor Gray
    Write-host "===================" -ForegroundColor White
    Write-host
   }
 }

if($DumpOccurances)
 {
  $occurances |Select-Object -unique @{Name='Name';Expression={$_}} | Export-Csv -NoTypeInformation -Path $exceptions_file_location
  Write-host "I did not find the exceptions file, or I may have been instructed, to dump all occurances to the exceptions file... whatever."
  Write-Host "All occuraances have been dumped to: ",$exceptions_file_location
  Write-host
  Write-host "You may want to have a look at the file, modify what you want to modify, and run the script again."

  $log_entry  = 'Dumped the occurances into the exceptions file, exiting at: ';
  $log_entry += Get-Date -Format "dddd yyyy.mm.dd HH:mm"
  Add-Content -Path $runlog_file_location -Value $log_entry

  exit;
 }


$exceptions_obj= Import-Csv -Path $exceptions_file_location
foreach ($exc in $exceptions_obj) {$exceptions+=$exc.Name}
$we_found_new_errors = $false;

For ($i=0; $i -lt $total_number_of_occurances; $i++)
 {
  if (-not("" -eq ($occurances[$i]))) # Check if $occ is not empty
   {
    if (-not ($null -eq ($occurances[$i] -as [int])))      #check if $occ is a number
     {
      if (($occurances[$i] -as [int]) -le $custom_certificate_warning_treshold)         #check if $occ is smaller than our custom certificate warning number...
       {
        $we_found_new_errors=$true;
        $found_errors += $occurances[$i]
        If($Verbose)
         {
          Write-Host "This number is smaller than we would like it to have. We have an error!"
          Write-Host $occurances[$i] -ForegroundColor Gray 
          Write-host "===================" -ForegroundColor White
          Write-host
         }
       }
     }
    else    # We have something in $occ which cannot be coverted into a number...
     {
      if(-not($exceptions.Contains($occurances[$i])))        # Here is the ggod stuff, we chack if $occ is listed as an exception or not...
       {
        $we_found_new_errors=$true;
        $found_errors += $occurances[$i]
        If($Verbose)
         {
          Write-Host "We have something:" -ForegroundColor White
          Write-Host $occurances[$i] -ForegroundColor Gray 
          Write-host "===================" -ForegroundColor White
          Write-host
         }
       } 
     }
    }
 }
        
  
 
 
if($we_found_new_errors)
 {
  if($Verbose)
   {
    Write-host "We're gon' send an e-mail, because we found some new errors!"
   }

  $From = "adminreport@smp.hu"
  $Subject = "Cause for concern was found on the "+$env:COMPUTERNAME+" Exchange server"

  $SMTPServer = "localhost"
  $SMTPPort = "25"

  

  
  if($Testing)
   {
    $to = "tamas.pacso@xxxxxxxxxxxxx.hu"
   }
  else
   {
    $to = "support@xxxxxxxxxxxxxxxx.net"
   }

# From here we start assembling the e-mails HTML body.

   $body = "<h2 style=""font-family:Corbel""> <b>Hi!</b> Unfortunatelly I found some errors on the "+$env:COMPUTERNAME+" Exchange server! </h2>"
   $body+= "<p style=""font-family:Corbel"">Here is a list of the concerning messages I extracted from the Health Checker's HTML report...</p>"
   $body+= "  <ul>"
             


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
  $log_entry  =  Get-Date -Format "yyyy.mm.dd HH:mm:ss"
  $log_entry  += 'Sending e-mail, then exiting.';
  Add-Content -Path $runlog_file_location -Value $log_entry
  
  Send-MailMessage -From $From -to $To -Subject $Subject -Body $body -BodyAsHtml -SmtpServer $SMTPServer -Port $SMTPPort  -Encoding UTF8 -Attachments $health_checker_file_location
 
 
 } 
else
 {
  Write-Host "We did not find any errors on the exchange system, or all errors have been excused in the exceptions file. Hurray!"

  $log_entry   = Get-Date -Format "yyyy.mm.dd HH:mm:ss"
  $log_entry  += ' No errors, or all have been exused... exiting.';
  Add-Content -Path $runlog_file_location -Value $log_entry
 }

  






  #
 ###
 ###   Useful command repository. Stuff I wanted to keep, so I don't have to figure them out again...
 ###
  #
  #
  ##  With this, we can export all occurances from a report, For example if we want to create an exceptions file...
  #
  ### $occurances |Select-Object -unique @{Name='Name';Expression={$_}} | Export-Csv -NoTypeInformation -Path $exceptions_file_location

  #
  #
  ##   If we wanna send only the report, but not the summary.
  #   
  ###  Send-MailMessage -From $From -to $To -Subject $Subject -Body $the_text -BodyAsHtml -SmtpServer $SMTPServer -Port $SMTPPort  -Encoding UTF8