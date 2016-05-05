param (
    [Parameter (Mandatory=$False)]
    [alias("SkipBE")]
    [boolean]$SkipBackup,

    [Parameter (Mandatory=$False)]
    [alias("SkipPST")]
    [boolean]$SkipMailExport

)

<#
.Name
    Process-ExEmpolyee.ps1

.Description
    Connects to various systems to process an ex-employee.  Systems include AD, Sharepoint, and Call Manager/Unity.  Incorporates Write-Log-COPY.ps1 for logging.  Both scripts can be found at \\<server>\<share>\Ex-Employee.
    All logging goes to \\<server>\<share>\Ex-Employee\<username>.log

.Notes
    Created by:  <redacted>
    Created on:  03/15/2016

.TODO
    Add multi-computer management (e.g. if ex-employee had more than 1 device)
    Add mailbox and email forwarding management (if reasonably possible)
    Add TL;DR to email
    Separate each section to it's own function/script/module to run independently from the rest of the script
    Add skip options for sections of the script so if something fails, it can be re-run with only the applicable sections.

#>

#Import Logging function
.'\Write-Log.ps1'



#Ignore certificate errors
add-type @"
    using System.Net;
    using System.Security.Cryptography.X509Certificates;
    public class TrustAllCertsPolicy : ICertificatePolicy {
        public bool CheckValidationResult(
            ServicePoint srvPoint, X509Certificate certificate,
            WebRequest request, int certificateProblem) {
            return true;
        }
    }
"@
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy

########################################################################################################################################
# Define global variables and get user information
########################################################################################################################################

#Get ex-employee's username
$Username = Read-Host "Enter username of Ex-Employee"

#Get ex-employee's computer name
$Computer = Read-Host "Enter the computer name of the device the user used"

#Get domain admin creds.  Needs to be in domain\<username> format in order to launch another powershell session as another user
$Creds = Get-Credential -Message "Enter your domain admin credentials WITH the domain (e.g <domain>\<username>"

#Get power user credentials.  Needs to not include the domain since the call manager web requests can't handle it with the domain
$PUCreds = Get-Credential -Message "Enter your power user credentials WITHOUT the domain"

#Get AD user object
$User = Get-ADUser $Username -Properties *

#Parse out AD user object properties
$fullname = $User.DisplayName
$userprincipalname = $user.UserPrincipalName
$UserEmail=$User.EmailAddress
$UserPrimaryGroup = $User.PrimaryGroup
$Firstname = $user.GivenName
$Lastname = $user.Surname

#Set logging path
$path = "\\<sever>\Share\Logs\$username.log"

#Get Date
$date = get-date -UFormat "%Y%m%d"

#Define the mail server
$SMTP = "<SMTP server>"

#Define the subject of the email messages
$subject = "Ex-Employee $fullname"

#Get Manager Details
try{

    $Manager = Get-ADUser $user.manager -Properties *
    $ManagerEmail = $Manager.emailaddress
    $ManagerName = $Manager.Name
    $ManagerDetails = $true

}
catch{

    Write-Log "There is no manager defined for this user" -Path $path -Level Warn
    $ManagerDetails = $false

}

#Get the IT user that is running this script's email address
$ITUser = Get-ADUser $env:USERNAME -Properties *
$ITEmail = $ITUser.emailaddress

########################################################################################################################################
# Export mailbox to .PST file in users's folder
########################################################################################################################################

if ($SkipMailExport -eq $False){
    Write-Log "Starting mailbox export" -Path $path

    #Call the Export-PST script to backup the .OST to a .PST in the \\<server>\<share>\<user>\ folder.  Script can be run independently and will return $true if completed and $false if failed.  Since we're using Start-Process, though, we'll need to check the exit codes
    $PSTExport = Start-process powershell.exe -Credential $creds -ArgumentList "\\<server>\<share>\Export-PST.ps1 -Path $Path -username $username" -Wait -PassThru


    if($PSTExport.ExitCode -eq 0){
        Write-Log "Successfully exported mailbox" -Path $Path
    }
    else{
        Write-log "There was an error during the mailbox export process.  Please check and run the export-pst.ps1 script again if the mailbox wasn't exported" -Path $path -Level Error

        #Define the body of the message
        $body = @"

    There was a problem with the ex-employee mailbox export.  The script has exited and will not continue.  Please export the mailbox manually and try again.
"@

        Send-MailMessage -To $ITEmail -Subject $subject -From $ITEmail -Body $body -SmtpServer $SMTP
        break
    }
}


########################################################################################################################################
# Back up the user's machine and user folder on the file share
########################################################################################################################################

if($SkipBackup -eq $False){
    write-log "Starting the backup process" -path $path

    $BESession = New-PSSession -ComputerName <server> -Credential $Creds
    $BEresult = Invoke-Command -Session $BESession -ArgumentList $Computer,$date,$username -ScriptBlock {

        #Import the computer name and ate
        param($computer,$date, $username)

        #Import the backupexec cmdlets
        Import-Module bemcli

        #Set count, agent status, and overall result variables to use later
        $count = 0
        $Result = $null
        $AgentJobStatus = $null

        #Tell BE to use tape slot 23 for this job
        $tape = get-bestorage -name "IBM 0003 [0023..0023]"

        #Tell BE to define the media set as "ExEmployees"
        $mediaset = get-bemediaset -name "ExEmployees"
  
        #Define what directories BE needs to backup
        $CDrive = new-befilesystemselection -path "C:\" -pathisdirectory -recurse     
        $UserFolder = new-befilesystemselection -path "D:\users\$Username\" -pathisdirectory -recurse 

        #Give the jobs a unique name
        $CDriveJobName = "ex_" + $Username + "_" + $date  
        $ShareJobName = "ex_" + $Username + "_FileShare_" + $date

        #Install the BE Agent and reboot the machine if needed
        $Agent = Install-BEWindowsAgentServer -Name $computer -LogonAccount backupexecsvc -restartautomaticallyifnecessary

        <# NOTE:
            There are 2 different gets that you can use:  get-bejob and get-bejobhistory.  The former will grab what's currently running or scheduled to run.
            The latter only finds what's already completed/failed.  If the job is a "backup", however, rather than an install, you can pull back the status of a completed job if it was recent
            If the backup jobtype is "Install", get-bejob will NOT be able to find the status once it's completed/failed
         #> 

        #Check job history for the installation.  Timeout after 30 minutes if nothing is found.  
        While (($AgentJobStatus -eq $null) -and ($count -le 60)){
        
            $AgentJobStatus = (get-bejobhistory $Agent).jobstatus
            Start-Sleep -Seconds 30
            $count++

        } 

        #If the agent successfully installed, run the machine backup and the user folder backup (\\<server>\<share>\<username>)
        if ($AgentJobStatus -eq "Succeeded"){
        
            #Define both the machine and share jobs
            $PCJob = Submit-BEOnetimeBackupJob -agentserver $computer -name $CDriveJobName -storage $tape -tapestoragemediaset $mediaset -backupsetdescription "Ex-Employee $username" -filesystemselection $CDrive
            $ShareJob = Submit-BEOnetimeBackupJob -agentserver <file server> -name $ShareJobName -storage $tape -tapestoragemediaset $mediaset -backupsetdescription "Ex-Employee $Username FileShare" -filesystemselection $UserFolder

            #Get the status of the jobs
            $PCJobStatus = $PCJob.status
            $ShareJobStatus = $ShareJob.status

            #Loop until the status is succeeded for both jobs or one of them errors out.
            while (($PCJobStatus -ne "Succeeded")-and ($ShareJobStaus -ne "Succeeded")){

                Start-Sleep -Seconds 300
                $PCJobStatus = (get-bejob $PCJob).jobstatus
                $ShareJobStatus = (get-bejob $ShareJobName).jobstatus

                #Check the status to make sure neither of the jobs failed.  If so, break the loop
                if (($PCJobStatus -eq "Error") -or ($ShareJobStatus -eq "Error")){
                    $Result = "Failed"
                    break
                }
        
            }

            #If the agent installation and the jobs finished successfully, $result will still be $null.  Change it to success
            if ($result = $null){
                $Result = "Success"
            }

        }


        else{

            $result =  "Failed"

        }

        return $Result
    }

    if ($BEresult -eq "Success"){
        Write-Log "Successfully backed up the machine and share" -Path $path
    }
    else{
        Write-Log "An error occurred during the backup process.  Please check and make sure that backup is taken before running the script again" -Path $path -Level Error

        #Define the body of the message
        $body = @"

    There was a problem with the ex-employee backup.  The script has exited and will not continue.  Check backup server for the backups.
"@

        Send-MailMessage -To $ITEmail -Subject $subject -From $ITEmail -Body $body -SmtpServer $SMTP
        break
     }

    Remove-PSSession $BESession
}

########################################################################################################################################
# Disable account and move to Disabled Users
########################################################################################################################################

Write-Log "Disabling $fullname's account..." -Path $path

#Check if User is disabled
if ($User.Enabled){

    try{

        Disable-ADAccount -Identity $Username -Credential $Creds
        Write-Log "Successfully disabled the account" -Path $path

    }
    catch{

        Write-Log "Failed to disable the account" -Level Error -Path $path

    }
}

else {

    Write-Log "User is already disabled" -path $path

}

#Check to see if user is a FTE
if ($user.PrimaryGroup -eq "<distinguished name of group"){

    $DisabledUsersOU = "<distinguished name of OU>"
    Write-Log "User is a FTE" -Path $path
    Write-log "Moving user to Disabled Users..." -Path $path

    #Check to see if the user is already part of the disabled users OU
    if ($user.DistinguishedName -like "*Disabled Users*"){

        Write-Log "User is already part of Disabled Users OU" -Path $path

    }
    else{

        Try{

            Move-ADObject $user.distinguishedName -TargetPath $DisabledUsersOU
            Write-Log "Successfully moved account to Disabled Users" -Path $path

        }
        Catch{

            Write-Log "Couldn't move AD account" -Path $path -Level Error

        }   
    }     
}
#Check to see if user is a contract employee
elseif ($user.PrimaryGroup -eq "<distinguished name of group"){

    Write-Log "User is a contractor" -Path $path
    Write-Log "Checking to see if a local DISABLED folder exists..." -Path $path

    #Find the ex-employee's OU
    $AccountOU = $user.distinguishedName -split ','
    $AccountOU = $AccountOU[2..($AccountOU.Count)] -join ','

    #Append DISABLED and see if OU Exists
    $DisabledOU = "OU=DISABLED,"
    $DisabledOU += $AccountOU

    Try{

        $temp = Get-ADOrganizationalUnit $DisabledOU
        Write-Log "Found local DISABLED OU" -Path $path
        Write-Log "Moving user to $DisabledOU..." -Path $path

        Try{

            Move-ADObject $user.distinguishedName -TargetPath $DisabledOU -Credential $Creds
            Write-Log "Successfully moved user to $DisabledOU" -Path $path

        }
        Catch{
e
            Write-Log "Couldn't move AD account" -Path $path -Level Error

        } 
               
    }
    Catch{

        Write-Log "Couldn't find local DISABLED OU.  Will not move account." -Path $path -Level Error

    }
    
}
else{

    Write-Log "Coudn't determine if user was a FTE or Contract employee.  Will not move AD user" -Path $path -Level Error

}


########################################################################################################################################
# Change employee list on Sharepoint
########################################################################################################################################

Write-Log "Modifying the employee list on Sharepoint..." -Path $path

#Start a remote powershell session to the sharepoint server and pass it the first and last name
$SPSession = New-PSSession -ComputerName <sharepoint server> -Credential $creds
$spresult = Invoke-Command -Session $SPSession -ArgumentList $Firstname,$Lastname,$path -ScriptBlock {
    
    param($Firstname,$Lastname,$path)
    
    #Add sharepoint functionality to remote shell
    Add-PSSnapin microsoft.sharepoint.powershell

    #Drill down to the new hire list
    $spweb = get-spweb "http://example.domain.com/sites/Human Resources"
    $splist = $spweb.Lists["Employees"]

    #Get items that have both the first and last names in sharepoint
    $spitem = $splist.getitems() | where {($_.Name -eq $Lastname) -and ($_.Xml -like "*$Firstname*")}  
    
    #Single items that are returned from the above where clause are SPListItem types.  If more than 1 item is returned, the type is "SPListItemCollection"
    #We only want to update a single record
    try{

        if($spitem.gettype().Name -eq "SPListItem"){

            $spitem["Employee Status"] = "X-Employee"
            $Spitem.update()
            return "Updated the Employee record"

        }

        else{

            return "More than 1 entry found.  Cannot update the List" 

        }
    }

    catch{

        return "Couldn't find a new hire record for $Lastname, $Firstname"
    }
}

Write-Log $spresult -Path $path

#Destroy the remote session
Remove-PSSession $SPSession


########################################################################################################################################
# Remove Security Groups/DLs
########################################################################################################################################

Write-Log "Removing security groups from user account..."

Try{

    #Get all groups except Domain Users and Contractors
    $ADgroups = Get-ADPrincipalGroupMembership -Identity $user -Credential $creds | where {($_.Name -ne "Domain Users") -and ($_.Name -ne "Contractors Group")}
    
    #Remove the groups
    Remove-ADPrincipalGroupMembership -credential $creds -Identity "$user" -MemberOf $ADgroups -Confirm:$false

    Write-Log "Successfully removed the security groups" -Path $path

}
catch{

    Write-Log "Couldn't remove groups from the account"  -Path $path -Level Error

}


########################################################################################################################################
# Send email to Manager
########################################################################################################################################

if($ManagerDetails){

    #Define the subject of the message
    $subject = "Ex-Employee $fullname"

    #Define the body of the message
    $body = @"
Hello, 

We are currently processing $Firstname's Ex-Employee ticket.  Would you like his/her email to be forwarded to you or any other member of your team?  Also, do you need any of the files from his/her user folder?  You can find the folder here: \\<server>\<share>\$username.

Thanks,
IT
"@

    try{

        Write-Log "Sending email to $ManagerEmail..." -Path $path
        Send-MailMessage -To $ManagerEmail -Subject $subject -From $ITEmail -Body $body -SmtpServer $SMTP
        Write-Log "Successfully sent email to $ManagerEmail" -Path $path

    }
    catch{

        Write-Log "Failed to send email to $ManagerEmail" -Path $path -Level Error

    }
}
else{

    Write-Log "No Manager details were found earlier, so an email cannot be sent."  -Path $path -Level Warn

}


########################################################################################################################################
# Remove from Lync
########################################################################################################################################

Write-Log "Removing ex-employee from Lync..."

#Connect to lync server and import the session.  Import is needed as the ocspowershell endpoint has a restricted languagemode enabled and I couldn't figure out how to elevate those permissions in a remote session
$lyncsession = New-PSSession -Credential $Creds -ConnectionUri "https://lyncserver.domain.com/ocspowershell"
Import-PSSession $lyncsession

try{

    disable-csuser -identity $userprincipalname
    Write-log "Successfully disabled the user in Lync" -path $path

}
catch{

    Write-log "Couldn't disable the user in Lync" -Path $path -Level Error

}

Remove-PSSession $lyncsession


########################################################################################################################################
# Clear Managed By
########################################################################################################################################

Write-Log "Clearing Manager field from user's AD object" -Path $path

try{
    Set-ADUser $user -Manager $null -Credential $Creds
    Write-Log "Successfully cleared Manager field from user's AD object" -Path $path
}
catch{
    Write-log "Couldn't clear out the Manager field in AD" -Path $path -Level Error
}


########################################################################################################################################
# Remove Voicemail
########################################################################################################################################

#Set the route partition that your lines are on
$routepartition = ""

#Get the ex-employee's unity object
$unityresult = Invoke-WebRequest -Uri "https://<server>/vmrest/users?query=(alias is $username)" -Credential $PUCreds

#Convert the response to an XML doc
$xmldoc = [xml]$unityresult.Content

#Make sure the returned XML document isn't empty
if ($xmldoc.users.IsEmpty){

    write-log "Couldnt find a voice mailbox" -path $path -level Error

}
else{

    $OID = $xmldoc.users.user.ObjectId

    #Delete the user from Unity
    $delresult = Invoke-WebRequest -Uri "https://<server>/vmrest/users/$OID" -Credential $PUCreds -Method Delete
    
    if($delresult.StatusCode -eq 204){

        write-log "Successfully deleted the user from Unity" -path $path

    }
    else{

        Write-log "Voice Mailbox may not have been deleted." -path $path -level warn

    }
}


########################################################################################################################################
# Change Directory number and device description from Call Manager to Available
########################################################################################################################################

#Define the SOAP call for the listLine function.  Gets all lines in call manager.  NOTE:  You'll need to change the routepartition name in the $listlinebody variable below
$listLinebody = @"
<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:ns="http://www.cisco.com/AXL/API/10.5">
   <soapenv:Header/>
   <soapenv:Body>
      <ns:listLine sequence="?">
         <searchCriteria>
            <routePartitionName>$routepartition</routePartitionName>
         </searchCriteria>
         <returnedTags uuid="?">
            <!--Optional:-->
            <pattern>?</pattern>
            <!--Optional:-->
            <description>?</description>
            <associatedDevices>
               <!--Zero or more repetitions:-->
               <device>?</device>
            </associatedDevices>
         </returnedTags>
      </ns:listLine>
   </soapenv:Body>
</soapenv:Envelope>
"@

try{

    #Call the listLine function
    $listLineresponse = Invoke-WebRequest -Uri "https://<server>:8443/axl/" -Credential $PUCreds -Method Post -Body $listLinebody -ContentType "text/xml"
    
    #Convert the HTTP response to an XML document
    $listLinexml = [xml]$listLineresponse.Content

    Write-Log "Successfully retrieved a list of all lines in the route partition" -Path $path

}
catch{

    Write-Log "Couldn't talk to the webservice to run the 'listLine' call" -Path $path -level Error

}

$pattern=$null

#Loop through the XML document, find any lines with a description that contains "<Firstname> <Lastname>", and grab the extension.  Is insensitive of case
foreach($node in $listLinexml.SelectNodes("//line")){

    if($node.Description -like "*$firstname $lastname*"){

        $pattern += @($node.pattern)

    }
}

#If no results are found, stop
if($pattern.length -eq 0){

    Write-Log "No line found for '$firstname $Lastname'" -Path $path -Level Warn

}

#If only 1 result is found, process the rest
elseif ($pattern.length -eq 1){
    
    Write-Log "Found 1 extension matching '*$fistname $Lastname*'" -Path $path

    #Define the SOAP XML for the getLine call
    $getLinebody = @"
        <soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:ns="http://www.cisco.com/AXL/API/10.5">
            <soapenv:Header/>
            <soapenv:Body>
                <ns:getLine sequence="?">
                    <!--You have a CHOICE of the next 2 items at this level-->
                    <pattern>$pattern</pattern>
                    <!--Optional:-->
                    <routePartitionName uuid="?">$routepartition</routePartitionName>
                    <!--Optional:-->
                    <returnedTags uuid="?">
                    <associatedDevices>
                        <!--Zero or more repetitions:-->
                        <device>?</device>
                    </associatedDevices>
                    </returnedTags>
                </ns:getLine>
            </soapenv:Body>
        </soapenv:Envelope>
"@

    try{

        #Call the getLine request
        $getLineresponse = Invoke-WebRequest -Uri "https://<server>:8443/axl/" -Credential $cred -Method Post -Body $getLinebody -ContentType "text/xml"

        #Convert the HTTP resonse to XML
        $getLinexml = [xml]$getLineresponse.Content

        #Get all the devices listed in the XML doc
        $devices = $getLinexml.SelectNodes("//device")

        Write-Log "Successfully retrieved a list of all devices with the extension '$pattern'" -Path $path

    }
    catch{

        Write-Log "Couldn't retrieve list of devices with extension '$pattern'" -Path $path -Level Error
        
    }

    #Define the updateLine SOAP XML to clear out the line details
    $updateLinebody = @"
        <soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:ns="http://www.cisco.com/AXL/API/10.5">
           <soapenv:Header/>
           <soapenv:Body>
              <ns:updateLine sequence="?">
                 <pattern>$pattern</pattern>
                 <!--Optional:-->
                 <routePartitionName uuid="?">$routepartition</routePartitionName>
                 <!--Optional:-->
                 <description>Available</description>
                 <!--Optional:-->
                 <alertingName>Available</alertingName>
                 <!--Optional:-->
                 <asciiAlertingName>Available</asciiAlertingName>
                 <!--Optional:-->
                 <active>true</active>
              </ns:updateLine>
           </soapenv:Body>
        </soapenv:Envelope>
"@

    try{

        #Call the updateLine SOAP request
        $updateLineResult = Invoke-WebRequest -Uri "https://<server>:8443/axl/" -Credential $cred -Method Post -Body $updateLinebody -ContentType "text/xml"

        Write-Log "Successfully cleared out the details for extension '$pattern'" -Path $path
    }
    catch{

        Write-Log "Couldn't clear the line details for extension '$pattern'" -path $path -Level Error

    }
}

#If too many results were found, don't continue
else{

    Write-log "Found too many lines that contain description '$firstname $lastname'.  Can't update Call Manager." -path $path -level Error

}

$Deviceprofilearray = $null
if($devices){
    
    #Loop through the devices and get the name of the device (e.g. SEP5597969T0C60)
    foreach($device in $devices){
        $deviceName = $device.'#text'
            
        #Define the SOAP XML of the getDeviceProfile call.  This is used to get the configuration of the physical phone before we modify it.
        $getDeviceProfilebody = @"
        <soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:ns="http://www.cisco.com/AXL/API/10.5">
            <soapenv:Header/>
            <soapenv:Body>
                <ns:getDeviceProfile sequence="?">
                    <!--You have a CHOICE of the next 2 items at this level-->
                    <name>$deviceName</name>
                    <returnedTags ctiid="?" uuid="?">
                    <!--Optional:-->
                    <name>?</name>
                    <!--Optional:-->
                    <description>?</description>
                    <!--Optional:-->
                    <lines>
                        <!--You have a CHOICE of the next 2 items at this level-->
                        <!--Zero or more repetitions:-->
                        <line ctiid="?" uuid="?">
                            <!--Optional:-->
                            <index>?</index>
                            <!--Optional:-->
                            <label>?</label>
                            <!--Optional:-->
                            <display>?</display>
                            <!--Optional:-->
                            <dirn uuid="?">
                                <!--Optional:-->
                                <pattern>?</pattern>
                                <!--Optional:-->
                                <routePartitionName uuid="?">?</routePartitionName>
                            </dirn>
                            <!--Optional:-->
                            <ringSetting>?</ringSetting>
                            <!--Optional:-->
                            <consecutiveRingSetting>?</consecutiveRingSetting>
                            <!--Optional:-->
                            <ringSettingIdlePickupAlert>?</ringSettingIdlePickupAlert>
                            <!--Optional:-->
                            <ringSettingActivePickupAlert>?</ringSettingActivePickupAlert>
                            <!--Optional:-->
                            <displayAscii>?</displayAscii>
                            <!--Optional:-->
                            <e164Mask>?</e164Mask>
                            <!--Optional:-->
                            <dialPlanWizardId>?</dialPlanWizardId>
                            <!--Optional:-->
                            <mwlPolicy>?</mwlPolicy>
                            <!--Optional:-->
                            <maxNumCalls>?</maxNumCalls>
                            <!--Optional:-->
                            <busyTrigger>?</busyTrigger>
                            <!--Optional:-->
                            <callInfoDisplay>
                                <!--Optional:-->
                                <callerName>?</callerName>
                                <!--Optional:-->
                                <callerNumber>?</callerNumber>
                                <!--Optional:-->
                                <redirectedNumber>?</redirectedNumber>
                                <!--Optional:-->
                                <dialedNumber>?</dialedNumber>
                            </callInfoDisplay>
                            <!--Optional:-->
                            <recordingProfileName uuid="?">?</recordingProfileName>
                            <!--Optional:-->
                            <monitoringCssName uuid="?">?</monitoringCssName>
                            <!--Optional:-->
                            <recordingFlag>?</recordingFlag>
                            <!--Optional:-->
                            <audibleMwi>?</audibleMwi>
                            <!--Optional:-->
                            <speedDial>?</speedDial>
                            <!--Optional:-->
                            <partitionUsage>?</partitionUsage>
                            <!--Optional:-->
                            <associatedEndusers>
                                <!--Zero or more repetitions:-->
                                <enduser>
                                <!--Optional:-->
                                <userId>?</userId>
                                </enduser>
                            </associatedEndusers>
                            <!--Optional:-->
                            <missedCallLogging>?</missedCallLogging>
                            <!--Optional:-->
                            <recordingMediaSource>?</recordingMediaSource>
                        </line>
                        <!--Zero or more repetitions:-->
                        <lineIdentifier>
                            <!--Optional:-->
                            <directoryNumber>?</directoryNumber>
                            <!--Optional:-->
                            <routePartitionName>?</routePartitionName>
                        </lineIdentifier>
                    </lines>
                    </returnedTags>
                </ns:getDeviceProfile>
            </soapenv:Body>
        </soapenv:Envelope>
"@

        try{

            #Call the getDeviceProfile method and store the XML results in an array
            $Deviceprofilearray += @((Invoke-WebRequest -Uri "https://<server>:8443/axl/" -Credential $cred -Method Post -Body $getDeviceProfilebody -ContentType "text/xml").content)

            Write-Log "Successfully retrieved device details for '$deviceName'" -Path $path

        }
        catch{

            Write-Log "Couldn't retrieve device details for '$deviceName'" -Path $path -Level Error

        }

    }
    
    #Loop through the device array    
    foreach($Deviceprofile in $Deviceprofilearray){
        
        #Convert the profile to an XML doc    
        $xmldoc = [xml]$Deviceprofile
        
        #Get lines that don't match the ex-employee's extension. This is used to send to updateDeviceProfile in case there are multiple extensions tied to the ex-employee's phone (e.g. front desk/Customer support/etc. and their own line)
        $xmlOtherlinesConfig = $xmldoc.SelectNodes("//line") | where {$_.dirn.pattern -ne $pattern}
            
        $lines = $null

        #Loop through the other lines to get their configuration
        foreach($config in $xmlOtherlinesConfig){
                
            $label = $Config.label
            $index = $Config.index
            $display = $Config.display
            $DisplayAscii = $Config.displayAscii
            $extension = $Config.dirn.pattern
            $PartitionName = $Config.dirn.routePartitionName.'#text'
            $dirnuuid = $config.dirn.uuid
            $partitionuuid = $config.dirn.routePartitionName.uuid

            #Define the lines section of the XML to be used later with the updateDeviceProfile call
            $lines += @"
                <line ctiid="?">
                    <index>$index</index>
                    <label>$label</label>
                    <display>$display</display>
                    <dirn uuid="$dirnuuid">
                        <pattern>$extension</pattern>
                        <routePartitionName uuid="$partitionuuid">$PartitionName</routePartitionName>
                    </dirn>
                    <displayAscii>$DisplayAscii</displayAscii>
                </line>
"@
        } 
        
        #Get the current device name 
        $devicename = $xmldoc.SelectNodes("//name").'#text'

        #Define the SOAP XML of the updateDeviceProfile method
        $updatephonebody = @"
        <soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:ns="http://www.cisco.com/AXL/API/10.5">
            <soapenv:Header/>
            <soapenv:Body>
                <ns:updateDeviceProfile sequence="?">
                    <name>$deviceName</name>
                    <description>Available</description>
                    <lines>
                    $lines
                    </lines>
                </ns:updateDeviceProfile>
            </soapenv:Body>
        </soapenv:Envelope>
"@          

        try{

            #Call the updateDeviceProfile method
            Invoke-WebRequest -Uri "https://<server>:8443/axl/" -Credential $cred -Method Post -Body $updatephonebody -ContentType "text/xml"

            Write-Log "Successfully update '$deviceName'" -Path $path

        }
        catch{

            Write-Log "Couldn't update '$deviceName'" -Path $path -Level Error

        }
    }
}


########################################################################################################################################
# Remove from computer WSUS
########################################################################################################################################

Write-Log "Removing computer from WSUS" -Path $path

try{
    $wsus = Get-WsusServer -name "<server name>" -port <port number>
    $client = $wsus.GetComputerTargetByName("$Computer.<domain>.com")
    $client.Delete()
    Write-Log "Successfully deleted the computer from WSUS" -Path $path
}
catch{
    Write-Log "Couldn't delete computer from WSUS" -Path $path -Level Error
}


########################################################################################################################################
# Delete Computer Object
########################################################################################################################################

Write-Log "Deleting computer from AD" -Path $path

try{
    Remove-ADComputer -credential $creds -Identity $computer -Confirm:$false
    Write-Log "Successfully deleted computer from AD" -Path $path
}
catch{
    Write-Log "Couldn't delete computer from AD" -Path $path -Level Error
}


########################################################################################################################################
# Email log to yourself when the script is completed
########################################################################################################################################


#Define the body of the message
$body = @"

The Ex-employee script has finished.  Please see attachment for the log file
"@

Send-MailMessage -To $ITEmail -Subject $subject -From $ITEmail -Body $body -SmtpServer $SMTP -Attachments $Path
