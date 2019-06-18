#Removes old workstations from AD.

$SboxDomains = Import-Csv "D:\required_files\SBOXDomains.csv"

$credential= Import-Csv -Path "D:\required_files\SBOX.csv"

$un= $credential.username

$pw= $credential.key | ConvertTo-SecureString

For ($i = 0; $i -lt $SboxDomains.Count; $i ++){

    $sboxDomain = $SboxDomains[$i].Domain

    $sboxCredUN="$($sboxDomain)\$($un)"
    
    $connectionCredentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sboxCredUN, $pw

    $aDSession = New-PSSession -ComputerName $SboxDomains[$i].ServerFQDN -Credential $connectionCredentials -Authentication Kerberos

    ####### Put in scriptblock

    Invoke-Command -Session $aDSession -ScriptBlock {

        $domain = (Get-ADDomain).DnsRoot

        $basebou  =  (Get-ADDomain).DistinguishedName

        # Move ADComputer Objects out of Computers OU and into Workstation OU

        if($domain -ne "interbrand.internal"){

            $searchMoveb =  ("CN=Computers," + $basebou)
            $destMoveb =  ("OU=Workstations," + $basebou)

        

            $computersToMove = (Get-ADComputer -Filter * -SearchBase $searchMoveb).DistinguishedName
        
            if($computersToMove){

                for ($j = 0; $j -lt $computersToMove.Count; $j ++) {
                    Write-Host "Moving $($computersToMove[$j]) to Workstations OU."

	                Move-ADObject $computersToMove[$j] -TargetPath $destMoveb
                }

            }

        }

        # Process workstation deletion

        $DaysInactiveDelete = 121

        $time2die = (Get-Date).Adddays(-($DaysInactiveDelete))

        $workstationsToDelete = @()

        if($domain -eq "interbrand.internal"){

            $workstationOU = "OU=Workstations,OU=NorthAmerica,DC=interbrand,DC=internal"
            $workstationManualUpdateOU = "OU=Workstations_Manual_Updates,OU=NorthAmerica,DC=interbrand,DC=internal"


        }else{

            $workstationOU =  ("OU=Workstations," + $basebou)
            $workstationManualUpdateOU =  ("OU=Workstations_Manual_Updates," +  $basebou)
       
        }
        
        $workstations = Get-ADComputer -Filter { LastLogonTimeStamp -lt $time2die } -SearchBase $workstationOU -Properties LastLogonTimeStamp,LastLogon,whenChanged | ForEach-Object {$workstationsToDelete += $_;}

        $manualUpdateWorkstations = Get-ADComputer -Filter { LastLogonTimeStamp -lt $time2die } -SearchBase $workstationManualUpdateOU -Properties LastLogonTimeStamp,LastLogon,whenChanged | ForEach-Object {$workstationsToDelete += $_;}

        if($domain -eq "ny.ddb.com" -or $domain -eq "interbrand.internal"){

            if($domain -eq "interbrand.internal"){

                $workstationPilotOU = "OU=Workstations_Pilot,OU=NorthAmerica,DC=interbrand,DC=internal"
                $workstationTestOU = "OU=Workstations_Test,OU=NorthAmerica,DC=interbrand,DC=internal"

            }else{

                $workstationPilotOU = ("OU=Workstations_Pilot," +  $basebou)
                $workstationTestOU =  ("OU=Workstations_Test," +  $basebou)

            }

            $pilotWorkstations = Get-ADComputer -Filter { LastLogonTimeStamp -lt $time2die } -SearchBase $workstationPilotOU -Properties LastLogonTimeStamp,LastLogon,whenChanged | ForEach-Object {$workstationsToDelete += $_;}

            $testWorkstations = Get-ADComputer -Filter { LastLogonTimeStamp -lt $time2die } -SearchBase $workstationTestOU -Properties LastLogonTimeStamp,LastLogon,whenChanged | ForEach-Object {$workstationsToDelete += $_;}
        }

        for ($j = 0; $j -lt $workstationsToDelete.Count; $j ++){
            
            $workstationDN = $workstationsToDelete[$j].DistinguishedName

            Write-Host "Deleting $workstationDN" -ForegroundColor Yellow

            Get-ADComputer -Identity $workstationDN |  Remove-ADObject -Recursive -Confirm:$false
        }
    }
}

$_Domains = @()

Function OfficeLocations(){
    
    for ($i = 0; $i -lt $SboxDomains.Count; $i ++){

        $global:_Domains += "<li>$($SboxDomains[$i].Domain)</li>"
    } 
}

OfficeLocations

$smtpserver= "100.117.4.120"
$from= "PowerShell Server Reporting <PSService@ad.corp>"
#ADAdmins@ddb.com
$to= "ADAdmins@ddb.com"
$subject= "DDB NA Workstation Cleanup"
$body= "
    <h1>DDB NA Workstation Cleanup</h1>
    <p>Workstations not logged into after 120 days have been removed from AD in the following domains:</p>
    <br/>
    <ul>
        $_Domains
    </ul>
"

# Send the mail message

Send-MailMessage -To $to -From $from -Subject $subject -Body $body -BodyAsHtml -SmtpServer $smtpserver -Priority High