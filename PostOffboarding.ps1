# SETTINGS

$Global:transcriptPath = "\\Location\Offboarding\POST-Offboarding\Audit\" # Set audit file location
$Global:Path = "\\location\Offboarding\POST-Offboarding\Data\OffboardedUsersToProcess.csv" # Master file of users to process (Check function creates it initially)
$Global:MasterFileBackup = "\\location\Offboarding\POST-Offboarding\Data\Backup\OffboardedUsersToProcess_" # path needs to end in \OffboardedUsersToProcess_ as file is named based on date
$Global:BackupsPath = "\\location\Offboarding\POST-Offboarding\Data\Backup\"
$Global:emailSmtpServer = "domain-com.mail.protection.outlook.com" # 365 domain used to send the emails
$Global:emailTo = "email@domain.com" # email sent to this address
$Global:emailFrom = "email@domain.com" # email sent from this address
$Global:PreSharedKeyLocation = "\\location\Offboarding\POST-Offboarding\Secure\PreShareKeyEncypted.txt"
$Global:365AdminAccount = "admin@domain.onmicrosoft.com"
$Global:emailBodyError = @"
Hello,<br>
<br> 
This warning has been sent as the post-offboarding master file size is outside of scope (200KB - 1MB).<br>
 <br>
Please check the file on the NAS, it is located: \\Location\Offboarding\POST-Offboarding\Data<br>
<br>
Thank you,<br>
<br>
Technology Operations Team
"@ # Email message for error with the master file
$emailBCC = "email@domain.com" # BCC report emails to this address (IT)
$emailBodyINI = @"
$managerName,<br>
<br>
This is an automated email to inform you of changes to the settings and forwarding rules for the mailbox $selecteduserEmail that will take place on <u><b>$($date)</b></u>.<br>
<br>
Overview of upcoming changes:<br>
<ul>
<li>All access to this mailbox will be removed.</li>
<li>The mailbox will be archived and the email address will become invalid.</li>
<li>Email sent to $selecteduserEmail will no longer forward to you or anybody else.</li>
<li>Any emails sent to $selecteduserEmail will bounce, and the sender will receive a bounce back notification.
</ul>
If you would like to request an extension of access to this mailbox, please open a new request on the Service Center <a href="https://link.com">here</a>.<br>
<br>
Thank you,<br>
<br>
Technology Operations Team
"@ # Email notification for initial pre post-offboarding process (iniNOTIFY)
$emailBodyNOTIFY = @"
$managerName,<br>
<br>
This is an automated email to inform you of changes to the settings and forwarding rules for the mailbox $selecteduserEmail that will take place on <u><b>$($date)</b></u>.<br>
<br>
Overview of upcoming changes:<br>
<ul>
<li>All access to this mailbox will be removed.</li>
<li>The mailbox will be archived and the email address will become invalid.</li>
<li>Email sent to $selecteduserEmail will no longer forward to you or anybody else.</li>
<li>Any emails sent to $selecteduserEmail will bounce, and the sender will receive a bounce back notification.
</ul>
If you would like to request an extension of access to this mailbox, please open a new request on the Service Center <a href="https://link.com">here</a>.<br>
<br>
Thank you,<br>
<br>
Technology Operations Team
"@ # Email notification for pre post-offboarding process (NOTIFY)


function Load-Module ($m) {

    # If module is imported say that and do nothing
    if (Get-Module | Where-Object { $_.Name -eq $m }) {
        write-host "Module $m is already imported."
    }
    else {

        # If module is not imported, but available on disk then import
        if (Get-Module -ListAvailable | Where-Object { $_.Name -eq $m }) {
            Import-Module $m -Verbose
            cls
        }
        else {

            # If module is not imported, not available on disk, but is in online gallery then install and import
            if (Find-Module -Name $m | Where-Object { $_.Name -eq $m }) {
                Install-Module -Name $m -Force -Verbose -Scope CurrentUser
                Import-Module $m -Verbose
                cls
            }
            else {

                # If the module is not imported, not available and not in the online gallery then abort
                write-host "Module $m not imported, not available and not in an online gallery, exiting."
                EXIT 1
            }
        }
    }
}

function Transcribe { 
    #transcript logging
    $currentAdmin = $env:Username
    $adminPC = $env:ComputerName
    $date = Get-Date -f yyyy-MM-dd_hh-mm-ss
    $transcriptFile = $transcriptPath + $currentAdmin + "_" + "$adminPC" + "_" + $date + ".txt" 
    Start-Transcript -Path $transcriptFile -noclobber
}

# Start transcript
Transcribe

Function Check(){
$DisabledUsers = Get-AzureADUser -Filter "accountEnabled eq false" -All $true
$SharedMailboxes = Get-Mailbox –ResultSize Unlimited –RecipientTypeDetails SharedMailbox

$UsersToProcess = New-Object System.Collections.ArrayList
$UsersToProcessUPN = New-Object System.Collections.ArrayList
$UsersToProcessFW = New-Object System.Collections.ArrayList

# Checking AD disabled accounts against shared mailboxes
foreach($DisabledUser in $DisabledUsers){
    foreach($SharedMailbox in $SharedMailboxes){
        if($DisabledUser.UserPrincipalName -eq $SharedMailbox.UserPrincipalName){
            $UsersToProcess.Add($DisabledUser) | out-null
            $UsersToProcessUPN.Add($DisabledUser.UserPrincipalName) | out-null
        }
    }
}
# From those accounts getting 
Foreach($UserToProcessUPN in $UsersToProcessUPN){
    $info = Get-Mailbox $UserToProcessUPN | Select-Object UserPrincipalName,ForwardingSmtpAddress,DeliverToMailboxAndForward
    $UsersToProcessFW.Add($info) | out-null
}


$UsersToProcessUPN | Sort-Object
$UsersToProcess | Export-Csv -Path "C:\out\OffboardedUsersToProcess.csv" -NoTypeInformation
$arrayObjects = $UsersToProcessUPN | ForEach-Object { 
     [PSCustomObject]@{'Value' = $_} 
}
$arrayObjects | Export-Csv -Path "C:\out\OffboardedUsersToProcessUPN.csv" -NoTypeInformation
$UsersToProcessFW | Export-Csv -Path "C:\out\OffboardedUsersToProcessFW.csv" -NoTypeInformation
} #Only used for initial discovery
Function Load-Changes(){
    [System.Collections.ArrayList]$Global:Users = Import-Csv $Path

    $filecheck = dir $Path
    if(($filecheck.Length -le 200000) -or ($filecheck.Length -ge 1024000))
    {
        # Notify the team to check the master file
        $emailSubject = "[WARNING] Check POST-Offboarding master file"
        Send-MailMessage -To $emailTo -From $emailFrom -Subject $emailSubject -Body $emailBodyError -SmtpServer $emailSmtpServer -BodyAsHtml
    }
}
Function Date-Check($u){
    $Global:today = $(get-date).Ticks
    $when = ([DateTime]$u.datetotakeaction).Ticks
    if($when -lt $today){
        return $true
        }else{
        return $false}
}
Function iniNotify($u){
    # Notification
    $selecteduserEmail = $u.Mail
    $selecteduserName = $u.DisplayName
    $selecteduserFirstName = $u.GivenName
    $selecteduserSurname = $u.Surname
    $emailTo = $u.SupervisorToNotify
    $date = (get-date).AddDays(7).ToString('MM/dd/yyyy')
    $managerUser = ($u.SupervisorToNotify).Split("@")[0]
    $managerName = (Get-ADUser -Identity $managerUser).GivenName	

    $emailSubject = "[IT Notice] Upcoming change to $($selecteduserName)'s mailbox"
        Send-MailMessage -To $emailTo -cc $emailBCC -From $emailFrom -Subject $emailSubject -Body $emailBodyINI -SmtpServer $emailSmtpServer -BodyAsHtml
        $u.DateToTakeAction = (get-date).AddDays(4)
        $u.ActionToTake = "NOTIFY"
    }
Function Notify($u){
    # Delete all calendar events as this is happening 27 days after offboarding and manager was informed during offboarding
    $ureport = (Remove-CalendarEvents –Identity $u.Mail -CancelOrganizedMeetings -QueryStartDate (Get-Date) -QueryWindowInDays 365 -Confirm:$false)
    if($ureport -ne $null)
    {        
        Write-Output "<<<---- " + $u.MailNickName + " --->>>"            
        Write-Output $ureport
        $ureport = $null
    }

    # Notification
    $selecteduserEmail = $u.Mail
    $selecteduserName = $u.DisplayName
    $selecteduserFirstName = $u.GivenName
    $selecteduserSurname = $u.Surname
    $emailTo = $u.SupervisorToNotify
    $date = (get-date).AddDays(3).ToString('MM/dd/yyyy')
    $managerUser = ($u.SupervisorToNotify).Split("@")[0]
    $managerName = (Get-ADUser -Identity $managerUser).GivenName	

    $emailSubject = "[IT Notice] Upcoming change to $($selecteduserName)'s mailbox"
        Send-MailMessage -To $emailTo -cc $emailBCC -From $emailFrom -Subject $emailSubject -Body $emailBodyNOTIFY -SmtpServer $emailSmtpServer -BodyAsHtml
        $u.DateToTakeAction = (get-date).AddDays(3)
        $u.ActionToTake = "ARCHIVE"
    }
Function Extend($u){
    $addTime = $(get-date).AddDays(23)
    $u.datetotakeaction = $addTime
    $u.actiontotake = "iniNOTIFY"
}
Function Archive($u){
    # Remove access to this mailbox
    $accesslist = (Get-MailboxPermission -Identity $u.Mail | Select User | Where-Object {($_.user -like '*@*')}).User #changed $u.UserPrincileName to $u.Mail
    if($accesslist -ne $null)
    {
        foreach($member in $accesslist)
        {
            Write-Host "---> Removing "$member" permissions from "$($u.UserPrincipalName)""

            Remove-MailboxPermission -Identity $u.Mail -User $member -AccessRights FullAccess -Confirm:$false
        }
    }

    # Disable email forwarding 
    $FwInfo = Get-Mailbox $u.UserPrincipalName | Select UserPrincipalName,ForwardingSmtpAddress,DeliverToMailboxAndForward
    if($FwInfo.ForwardingSmtpAddress -ne $null){
        set-Mailbox -Identity $u.UserPrincipalName -ForwardingSMTPAddress $null
    }
    
    # Append "-archive" to the end of the shared mailbox's email address (e.g., john.doe-archive@domain.com)
    $global:newEmail = (Get-ADUser $u.MailNickName).UserPrincipalName
    $newEmail = $newEmail.Split("@")
    $newEmail[0] = $newEmail[0] + "-archive"
    $newEmail = $newEmail[0] + "@" + $newEmail[1]
    set-aduser $u.MailNickName -emailaddress $newEmail
    set-ADUser $u.MailNickName -Add @{msExchHideFromAddressLists = "TRUE" } -Confirm:$false
    $u.Mail = $newEmail
    # Write-Host "Email set to : $($newEmail)" -ForegroundColor Green
    
    # Append "ARCHIVE" to the shared mailbox's last name field. (e.g., "John Doe ARCHIVE")
    $u.Surname = ($_.Surname + " ARCHIVE")
    $user = Get-ADUser $u.MailNickName -properties proxyaddresses
    Get-ADuser $user.DistinguishedName | ForEach-Object {
    Set-ADuser $_ -Surname ($_.Surname + " ARCHIVE")
    Set-ADUser $_ -UserPrincipalName $newEmail
    Rename-ADObject $_ -NewName "$($_.GivenName) $($_.Surname) ARCHIVE"
    $u.DisplayName = "$($_.GivenName) $($_.Surname) ARCHIVE"
    $u.MailNickName = $newEmail.Split("@")[0]
    }

    # Remove all email aliases from the mailbox
    #$user = Get-ADUser $u.OnPremisesSecurityIdentifier -properties proxyaddresses
    $user.proxyaddresses | ForEach-Object {
        Get-ADuser $u.OnPremisesSecurityIdentifier | Set-ADuser -remove @{proxyaddresses = $_ }
        # Write-Host "Removing "$_" from " $User.name "'s Proxy address" -ForegroundColor Green
    }
    $u.ActionToTake = "COMPLETED"
}
Function Save-Changes($u){
    #Save changes
    $u | Export-Csv -Path $Path -NoTypeInformation
    
    #Save backup and delete any backup file older than 90 days
    $BackupPath = $MasterFileBackup + (get-date -format "MM_dd_yyyy") + ".csv"
    $u | Export-Csv -Path $BackupPath -NoTypeInformation

    #Parameters
    $Path = $BackupsPath # Path where the file is located 
    $Days = "90" # Number of days before current date
 
    #Calculate Cutoff date
    $CutoffDate = (Get-Date).AddDays(-$Days)
     
    #Get All Files modified more than the last 90 days
    Get-ChildItem -Path $Path -Recurse -File | Where-Object { $_.LastWriteTime -lt $CutoffDate } | Remove-Item –Force -Verbose
}

Load-Module "AzureAD"
Load-Module "PSExcel"
Load-Module "ActiveDirectory"

$PreSharedKeyEncrypted = Get-Content -Path $PreSharedKeyLocation | ConvertTo-SecureString
$Cred = New-Object System.Management.Automation.PsCredential($365AdminAccount,$PreSharedKeyEncrypted)

Connect-AzureAD -Credential $Cred
Connect-ExchangeOnline -Credential $Cred

# Action will be based on ActionToTake attribute
# If value is ARCHIVE: Proceed with the extended offboarding procedure. No notification to anyone is needed.
# If value is LEAVE: Do not perform any actions.
# If value is NOTIFY: Notify the individual indicated in column Z that the email forwarding from the user in column X will cease to be forwarded on June 30. The notification to the individuals should indicate that if they do not need the email, no response is needed. But if they need emails to continue to be forwarded, they can request a 30 day extension.
# If value is REVIEW: Do not perform any actions. Further review is required.
# If value is EXTEND: Extends the forward for another 30 days and changes next state to NOTIFY.

Load-Changes

foreach($User in $Users)
{
    if(Date-Check($User)){
        switch($User.ActionToTake)
        {
            "ARCHIVE" { Archive($User)
            Write-Host "Running Archive on "$User.Mail}
            "iniNOTIFY" { iniNOTIFY($User)
            Write-Host "Running iniNOTIFY (7 days) on "$User.Mail}
            "LEAVE" {} # No Action
            "NOTIFY" { Notify($User)
            Write-Host "Running Notify (3 days) on "$User.Mail}
            "REVIEW" {} # No Action
            "EXTEND" { Extend($User)
            Write-Host "Running Extend on "$User.Mail} # Extends the time the forward is on for another 30 days from now
            "COMPLETED" {} # No Action
        }
    }
}

Save-Changes($Users)