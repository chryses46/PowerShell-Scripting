# This program provides reports for the following:
#     Audit logs for the past 60 days (default) Per mailbox
#     Mailbox Inbox Rules per Mailbox
#     MSOL and EXOL Data per user:
#         Email  
#         UserPrincipalName
#         LastPasswordChangeTimestamp
#         ForwardingSmtpAddress
#         DeliverToMailboxAndForward
#

# ENTERS EXCHANGE ONLINE AND OFFICE 365 POWERSHELL CONNECTION

$CreateEXOPSSession = (Get-ChildItem -Path $env:userprofile `
-Filter CreateExoPSSession.ps1 -Recurse -ErrorAction SilentlyContinue -Force | `
 Select -Last 1).DirectoryName
. "$CreateEXOPSSession\CreateExoPSSession.ps1"

Connect-EXOPSSession

Connect-MsolService

# SET DOMAIN
$Domain=(Get-msoldomain | Where-Object {$_.isDefault -eq $true}).Name

# CHOOSE FROM A SINGLE USER OR A LIST
$Choice= Read-Host "Please enter the number 1 for one user or 2 to import a list"
$Choosen=$null
While($choosen -eq $null){
    if($choice -eq 1){
        $Choosen= 1
    }elseif($choice -eq 2){
        $Choosen= 2
    }else{
        $Choosen =$null
        $Choice= Read-Host "Please enter (1) user or (2) list."
    }
}

# ENTER CUSTOM AUDIT DATE RANGE IF DESIRED. DEFAULT IS 60 DAYS.
 While($CustDate -eq $null){
    $CustDateQuestion= Read-Host "Would you like to enter a custom date range for the audit report?"
    $CustDate=$null
    if($CustDateQuestion -eq "Yes"){
        $CustDate = 1
    }elseif($CustDateQuestion -eq "No"){
        $CustDate = 2
    }else{
        $CustDate = $null
        $CustDateQuestion= Read-Host "Please respond 'Yes' or 'No'"
        }
    }

    if($CustDate -eq 1){
    
        $StartDateDate = Read-Host "Please enter start date. Ex: 03/23/2018"
        $EndDateDate= Read-Host "Please enter end date. Ex: 04/06/2018"
    
    }

    if($CustDate -eq 2){
    
        Write-Host "Start and end dates set to default (Past 30 days)"
        $StartDateDate = (Get-Date).AddMonths(-1).ToShortDateString()
        $EndDateDate = (Get-Date).ToShortDateString()
    
    }

# SINGLE USER MODE
if($Choosen -eq 1){

    $UserSMTP= Read-Host "Please enter user email"
    $OutPath = Read-Host "Please enter the root directory for export"
    $OutDir = "$OutPath\$userSMTP"
    

    If (!(Test-Path $OutDir))
       {
        New-Item -ItemType Directory -Path $OutDir -ErrorAction SilentlyContinue | Out-Null -Verbose
       }

    # PROCESS AUDIT LOGS FOR A SINGLE USER
    $AuditLogPath = "$OutDir\auditlog_$($userSMTP).csv"
       
    # Start date and end date to pull audit logs for
    # EXAMPLE DATE FORMAT:  $StartDate = "03/23/2018 00:00"
    # EXAMPLE DATE FORMAT:  $EndDate = "04/06/2018 23:59"
    #
    $StartDate= $StartDateDate +" 00:00"
    $EndDate= $EndDateDate+" 23:59"

    # Generate a random session id
    $sessionID = -join ((48..57) + (97..122) | Get-Random -Count 32 | % {[char]$_})

    # Arbitrary initial values just to enter the loop cleanly
    $beforesize = 0
    $aftersize = 1 

    # Retrieve the maximum number of records per request (50,000 while using ReturnLargeSet) until it is all captured
    while ( $beforesize -ne $aftersize) {
        $beforesize = (Get-Item $AuditLogPath  -ErrorAction SilentlyContinue).Length
        Search-UnifiedAuditLog `
            -UserIds $userSMTP `
            -StartDate $StartDate `
            -EndDate $EndDate `
            -SessionId $sessionID `
            -SessionCommand ReturnLargeSet | `
                Select-Object `
                    CreationDate,UserIDs,Operations,AuditData | `
                        Export-Csv -Path $AuditLogPath -NoTypeInformation -Append
        $aftersize = (Get-Item $AuditLogPath).Length 
    }

     # FILE FOR MSOL DATA (ONE USER)
    $userData = "$OutDir\365_Forensics_$userSMTP.csv"

    # USER INBOX REPORT
    $userInboxRules = "$OutDir\365_InboxRules_$userSMTP.csv"

    # GET RECENT PASSWORD RESET
    $PW= Get-MsolUser -UserPrincipalName $userSMTP |Select UserPrincipalName,LastPasswordChangeTimestamp

    # GET MAILBOX FORWARDING RULES
    $FWD= Get-Mailbox $userSMTP | Select UserPrincipalName,ForwardingSmtpAddress,DeliverToMailboxAndForward

    # GET MAILBOX RULES AND THEIR DETAILS, IF ANY
    $RULE= Get-InboxRule -Mailbox $userSMTP | Select Name,Description,Enabled,Priority,ForwardTo,ForwardAsAttachmentTo,RedirectTo,DeleteMessage 
    $RULE | Export-Csv -path $userInboxRules -NoTypeInformation
    
    $CustomUserProperties = [Ordered]@{

        Email=$userSMTP;   
        UserPrincipalName=$PW.UserPrincipalName;
        LastPasswordChangeTimestamp=$PW.LastPasswordChangeTimestamp;
        ForwardingSmtpAddress=$FWD.ForwardingSmtpAddress;
        DeliverToMailboxAndForward=$FWD.DeliverToMailboxAndForward;

    }

    $CollectedProps=[pscustomobject]$CustomUserProperties
    $CollectedProps | Export-Csv $userData -NoTypeInformation


}

# LIST MODE
if($Choosen -eq 2){

    $SMTPList = Read-Host "Please enter full path of User_SMTP.csv"
    $OutPath = Read-Host "Please enter the root directory for export"
    $CollectedSMTP = Import-Csv -Path $SMTPList
    $OutDir = "$OutPath\$domain"
    
    If (!(Test-Path $OutDir))
    {
        New-Item -ItemType Directory -Path $OutDir -ErrorAction SilentlyContinue | Out-Null -Verbose
    }

    # PROCESS LIST OF SMTP ADDRESSES FOR AUDIT LOGS

    foreach($user in $CollectedSMTP){

        $userSMTP = $user.Mail
        $AuditLogPath = "$OutDir\auditlog_$($userSMTP).csv"
 
        If (!(Test-Path $AuditLogPath))
        {
            New-Item -ItemType Directory -Path $OutDir -ErrorAction SilentlyContinue | Out-Null -Verbose
        }
 

   
        # Start date and end date to pull audit logs for
        # EXAMPLE DATE FORMAT:  $StartDate = "03/23/2018 00:00"
        # EXAMPLE DATE FORMAT:  $EndDate = "04/06/2018 23:59"
        #
        $StartDate= $StartDateDate + " 00:00"
        $EndDate= $EndDateDate+" 23:59"

        # Generate a random session id
        $sessionID = -join ((48..57) + (97..122) | Get-Random -Count 32 | % {[char]$_})

        # Arbitrary initial values just to enter the loop cleanly
        $beforesize = 0
        $aftersize = 1 

        # Retrieve the maximum number of records per request (50,000 while using ReturnLargeSet) until it is all captured
        while ( $beforesize -ne $aftersize) {
            $beforesize = (Get-Item $AuditLogPath -ErrorAction SilentlyContinue).Length
            Search-UnifiedAuditLog `
                -UserIds $userSMTP `
                -StartDate $StartDate `
                -EndDate $EndDate `
                -SessionId $sessionID `
                -SessionCommand ReturnLargeSet | `
                    Select-Object `
                        CreationDate,UserIDs,Operations,AuditData | `
                            Export-Csv -Path $AuditLogPath -NoTypeInformation -Append
            $aftersize = (Get-Item $AuditLogPath).Length 
        }


    }

     # FILE FOR MSOL DATA (MANY USERS)
    $collectedUserData = "$OutDir\365_Forensics_$Domain.csv"

    # PUT DATA FOR MANY USERS IN A LIST
    $CollectedUserPropsList=@()

    foreach($user in $CollectedSMTP){
        $UserSMTP = $user.Mail

        # GET RECENT PASSWORD RESET
        $PW= Get-MsolUser -UserPrincipalName $userSMTP |Select UserPrincipalName,LastPasswordChangeTimestamp

        # GET MAILBOX FORWARDING RULES
        $FWD= Get-Mailbox $userSMTP | Select UserPrincipalName,ForwardingSmtpAddress,DeliverToMailboxAndForward

        # USER INBOX REPORT
        $userInboxRules = "$OutDir\365_InboxRules_$userSMTP.csv"

        # GET MAILBOX RULES AND THEIR DETAILS, IF ANY
        $RULE= Get-InboxRule -Mailbox $userSMTP | Select Name,Description,Enabled,Priority,ForwardTo,ForwardAsAttachmentTo,RedirectTo,DeleteMessage 
        $RULE | Export-Csv -path $userInboxRules -NoTypeInformation
    
        $CustomUserProperties = [Ordered]@{

            Email=$userSMTP;   
            UserPrincipalName=$PW.UserPrincipalName;
            LastPasswordChangeTimestamp=$PW.LastPasswordChangeTimestamp;
            ForwardingSmtpAddress=$FWD.ForwardingSmtpAddress;
            DeliverToMailboxAndForward=$FWD.DeliverToMailboxAndForward;

        }

        $CollectedProps=[pscustomobject]$CustomUserProperties
        $CollectedUserPropsList+=$CollectedProps
       
    }

    $CollectedUserPropsList | Export-Csv -Path $collectedUserData -NoTypeInformation 

}

Get-PSSession | Remove-PSSession
