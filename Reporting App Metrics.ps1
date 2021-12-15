#################################################
####  Created and maintained by Zouhdid002   ####
#################################################


$head = @"
<style>
.styled-table1 {
    border-collapse: collapse;
    margin: 25px 0;
    font-size: 0.9em;
    font-family: verdana;
    min-width: 400px;
    box-shadow: 0 0 20px rgba(0, 0, 0, 0.15);
    width: 100%;
    text-align: center;
}
.styled-table1 thead tr {
    background-color: #16A085;
    color: #ffffff;
    text-align: center;
}
.styled-table1 th {
    background-color: #16A085;
    color: #ffffff;
    text-align: center;
}
.styled-table1 td {
    padding: 12px 15px;
    font-size: 11px;
    text-align: center;
}
.styled-table1 tbody tr {
    border-bottom: 1px solid #dddddd;
    text-align: center;
}
.styled-table2 {
    border-collapse: collapse;
    margin: 25px 0;
    font-size: 0.9em;
    font-family: verdana;
    min-width: 400px;
    box-shadow: 0 0 20px rgba(0, 0, 0, 0.15);
    width: 100%;
    text-align: left;
}
.styled-table2 thead tr {
    background-color: #ffffff;
    color: #16A085;
    margin: 25px 0;
    text-align: left;
}
.styled-table2 th {
    background-color: #ffffff;
    color: #16A085;
    margin: 25px 0;
    text-align: center;
}
.styled-table2 td {
    padding: 10px 10px;
    font-size: 11px;
    text-align: left;
}
.styled-table2 tbody tr {
    border-bottom: 1px solid #dddddd;
    text-align: left;
}
</style>
"@

$Cred = Get-Credential
$so = New-PSSessionOption -SkipRevocationCheck

$Apps = Get-Content Apps.json | ConvertFrom-Json
$head | Out-File HtmlReport.html
write-output "<table width=100% style='border-spacing: 20px 20px;'>" | Out-File HtmlReport.html -Append

ForEach ($App in $Apps) {

$App.App
$App.Url

$AppHtml = ""
$AppStatus = "OK"
$Warnings = 0

$AppHtml = "<tr valign='top'><td>"
$AppHtml += "<table class='styled-table1'>"
$AppHtml +=  "<th width=33%>Application</th><th width=33%>Status</th><th width=33%>Url</th>"
$AppHtml +=  "<tr><td>$($App.App)</td><td><img src='Images\OK.png' class='center' width='30' height='30'></td><td>"

 # Url Check
  $URL = $App.Url
    
      try {
              $request= Invoke-WebRequest -uri $URL -UseBasicParsing

              if ($request.StatusCode -eq "200") {
                        write-host "`n Site - $URL is up `n" -ForegroundColor green
                        $AppHtml +=  "<font color='green'>OK</font>" 
              }
              else {
                        write-host "`n Site - $URL is down `n" ` -ForegroundColor red
                        $AppHtml +=  "<font color='red'>NOTOK</font>"
                        $AppStatus = "NOTOK" 
              }

      } catch {
                    write-host "`n Site is not accessible.`n" ` -ForegroundColor red
                    $AppHtml +=  "<font color='red'>NOTOK</font>"
                    $AppStatus = "NOTOK"
      }

$AppHtml +=  "</td></tr>"
$AppHtml +=  "</table> </td> <td> <table class='styled-table2'>"
$AppHtml +=  "<th>Servers</th><th>CPU</th><th>RAM</th><th>Services</th><th>Disks</th>"

        ForEach ($Server in $App.Servers)
        {
            
            $Server.name           
            $Processor = $null
            $Memory = $null
            $RoundMemory = $null
            $Object = $null

            Try {

             
                # Disk Usage
                $cmd = { Get-WmiObject Win32_LogicalDisk -Filter "DriveType='3'"   | Select-Object DeviceID, @{N="UsedSpace";e={[math]::Round($($_.Size - $_.FreeSpace) / $_.Size * 100,1)}} | Sort-Object -Property 'Used Space (%)' -Descending }
                $Disks = Invoke-Command -ComputerName $Server.name -UseSSL -ScriptBlock $cmd -SessionOption $so -Credential $Cred 
                $Disks | foreach {if($_.UsedSpace -gt 90) {$Warnings = 1}}

                # Processor Check
                $cmd = { (Get-WmiObject  -Class win32_processor  | Measure-Object -Property LoadPercentage -Average | Select-Object Average).Average }
                $Processor = Invoke-Command -ComputerName $Server.name -UseSSL -ScriptBlock $cmd -SessionOption $so -Credential $Cred 

                # Memory Check
                $cmd = { Get-WmiObject  -Class win32_operatingsystem  }
                $Memory = Invoke-Command -ComputerName $Server.name -UseSSL -ScriptBlock $cmd -SessionOption $so -Credential $Cred 
                if ($Memory -ne $null) {
                $Memory = ((($Memory.TotalVisibleMemorySize - $Memory.FreePhysicalMemory)*100)/ $Memory.TotalVisibleMemorySize)
                $RoundMemory = [math]::Round($Memory, 2)
                } else {
                $RoundMemory = 0
                } 

                # Html 
                $AppHtml +=  "<tr valign='top'>"
                write-output "Server : $($Server.name)"
                $AppHtml +=  "<td text-align='center'>$($Server.name)</td>"                              
                write-output "CPU : $Processor%"
                if($Processor -gt 90) {
                $Warnings = 1
                $AppHtml +=  "<td text-align='center'><font color='red'>$Processor%</font></td>"
                } else {
                $AppHtml +=  "<td text-align='center'>$Processor%</td>"
                }
                write-output "RAM : $RoundMemory%"
                if($RoundMemory -gt 90) {
                $Warnings = 1
                $AppHtml +=  "<td text-align='center'><font color='red'>$RoundMemory%</font></td>"
                } else {
                $AppHtml +=  "<td text-align='center'>$RoundMemory%</td>"
                }

                # Services Check
                $Services = $Server.Services
                $AppHtml +=  "<td text-align='center'><ul text-align='left'>"
                write-output "Services :"
                $Services | foreach {
                    $Service = $_
                    $cmd = {
                        param($sn)
                        Get-WmiObject Win32_Service | Where {$_.Name -eq $sn } -ErrorAction Continue 
                    }
                    $Service_status = Invoke-Command -ComputerName $Server.name -UseSSL -ScriptBlock $cmd -SessionOption $so -Credential $Cred -ArgumentList $Service
                    if($Service_status.State -notcontains "Running") {
                    $AppStatus = "NOTOK"
                    $AppHtml += "<li>$($Service_status.Name) : <font color='red'>$($Service_status.State)</font></li>"
                    } else {
                    $AppHtml += "<li>$($Service_status.Name) : $($Service_status.State)</li>"
                    }
                    write-output "`t $($Service_status.Name) : $($Service_status.State)"
                    }
                $AppHtml +=  "</ul></td>"

                # Disks check
                write-output "Disks :"
                $AppHtml +=  "<td text-align='center'><ul text-align='left'>"
                $Disks | foreach { 
                    if($_.UsedSpace -gt 90) {
                    $Warnings = 1
                    $AppHtml +=  "<li>$($_.DeviceId) <font color='red'>$($_.UsedSpace)%</font></li>" 
                    } else {
                    $AppHtml +=  "<li>$($_.DeviceId) $($_.UsedSpace)%</li>"
                    }
                    write-output "`t $($_.DeviceId) $($_.UsedSpace)%"                    
                    }
                $AppHtml +=  "</ul></td>"
                $AppHtml +=  "</tr>"
            }
            Catch {
                Write-Host "There are errors for $($Server.name): "$_.Exception.Message
                $AppHtml +=  "<tr><td>$($Server.name)</td><td>!</td><td>!</td><td>!</td><td>$($_.Exception).Message</td></tr>"
                Continue
            }
        }

$AppHtml +=  "</table></td></tr>" 

# refresh html App Global Status
if ($AppStatus -eq "NOTOK") {
$AppHtml.Replace('OK.png','NOTOK.png')
} elseif ($Warnings -eq 1) {
$AppHtml.Replace('OK.png','WARN.png')
}

$AppHtml  | Out-File HtmlReport.html -Append

}
