param(
    [Parameter (Mandatory=$True)]
    [alias("LogPath")]
    [string]$Path,

    [Parameter (Mandatory=$True)]
    [string]$username
)

#Load the logging script
.'\Write-Log-COPY.ps1'

#Add the Exchange snap in
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010

Write-Log "Setting permissions for Exchange Trusted Subsystem to be able to write to user's  folder" -Path $Path

try{
    #Set the permissions for the user's  folder to allow the Exchange server to push the PST file
    $ACL = Get-Acl "\\<server>\<share>\$username\"
    $Ar = New-Object System.Security.AccessControl.FileSystemAccessRule("Exchange Trusted Subsystem", "Modify", "ContainerInherit,ObjectInherit", "None", "Allow")
    $ACL.SetAccessRule($Ar)
    Set-Acl "\\<server>\<share>\$username\" $ACL
    Write-Log "Successfully set permissions" -Path $Path
}
catch{
    Write-Log "Couldn't set permissions on user's folder.  Check that 'Exchange Trusted Subsystem' has modify rights to user's folder" -Path $Path
    break
}

Write-Log "Starting Mailbox export" -Path $Path

#Start the export and store the request
$export = New-MailboxExportRequest -Mailbox $username -FilePath \\<server>\<share>\$username\$username.pst

#If an error occurred during the export request, $export will be null
if($export -ne $null){
    
    while ($exportStatus -ne "Completed"){

        $time = get-date -DisplayHint Time
        $exportStatus = (Get-MailboxExportRequest $export).status
        Write-Log "PST Export Status: $exportStatus - $time" -Path $Path
        Start-sleep -seconds 30
    }

    Write-Log "Completed export at $time" -Path $Path

    
    $completed = $true

}
else {

    $completed = $false

}

#Return status of the job
return $completed
