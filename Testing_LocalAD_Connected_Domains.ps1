#This script is meant to test the connections to the SBOX domains and their servers. If a connection fails, an email will be sent. Tested daily.

Import-Module ActiveDirectory

$SBOXDomains= Import-csv "D:\required_files\SBOXDomains.csv"

#Set "True" for testing. Leave blank otherwise.
$testing="True"

#Array for export
$SBOXNADELUSERS=@()

#For testing one office, gets first unit of array
#$TargetDomain=$SBOXDomains[3]

$ConfirmADConn = @()

$ErrorEmail = @()

$credential= Import-Csv -Path "D:\required_files\SBOX.csv"

$un= $credential.username

$pw= $credential.key | ConvertTo-SecureString



Foreach($TargetDomain in $SBOXDomains){

    $ContinuedOperation = "Normal"

    $Domain = $TargetDomain.Domain

    $credun="$($Domain)\$($un)"
    
    $DomainCredentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $credun, $pw
    
    #Write-Host "Testing New-PsSession to $Domain"

    New-PSSession -ComputerName $TargetDomain.ServerFQDN -Credential $DomainCredentials -Authentication Kerberos
   

    #TESTING
    $TestData=$NULL

    if(($testing) -eq "True"){
        Write-Host "Testing connection to $($Domain)."
        $TestData= (Get-ADUser -Filter {UserPrincipalName -like "Daniel.frank*"} -Server $Targetdomain.ServerIP -Credential $DomainCredentials -Properties *).UserPrincipalName

        if(($TestData) -eq $NULL){
            $TestData="Error. Please check connection for $($Domain)"
            $ErrorEmail+=$TestData
        }
        
        $ConfirmADConn+=$TestData
        
    }



    Get-PSSession | Remove-PSSession

}

$ConfirmADConn

if($ErrorEmail -ne $NULL){

    #Prepare email for error reports
    $smtpserver= "100.117.4.120"
    $from= "PowerShell Server Reporting <noreply@PSService.local>"
    $to= "daniel.frank@ddb.com"
    $subject= "SBOX Reporting Connection Issues Detected!"
    $body= "
        <p style='text-indent:1em;'>Errors have been detected during the testing of the SBOX server connections from the PowerShell scritping server at 100.116.4.118. Please look into these immediately.
        </br>
        $(foreach($line in $erroremail){
        $line
        '<br/>'
        })
        </p>"

    # Send the mail message

    Send-MailMessage -SmtpServer $smtpserver -From $from -To $to -Subject $subject -Body $body -BodyAsHtml -Priority High

    }
    <#

    Scratch

   $cred = Get-Credential
   $nycred = Get-Credential
   
   New-PSSession -ComputerName DNYNYMVPRIDCN04.NY.DDB.COM -Credential $cred -Authentication Kerberos
   Get-ADUser -Filter {UserPrincipalName -like "Daniel.frank*"} -Server DNYNYMVPRIDCN04.NY.DDB.COM -Credential $nycred -Properties Userprincipalname


    #>