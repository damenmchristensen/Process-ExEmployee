# Process-ExEmployee
Ex-Employee Powershell script

Connects to various systems to process an ex-employee.  Systems include AD, Sharepoint, and Call Manager/Unity.  Incorporates Write-Log-COPY.ps1 for logging.  It will need to be run as administrator since it spins off another shell to launch the export-pst script.  In my environment, I need to run the mailbox export as a different user and New-MailboxEpoxrtRequest can't accept alternate credentials.  You shouldn't need to do this if you run the export-pst script separately.  I've incorporated a -SkipBE and a -SkipPST switches that will not execute that part of the code.  Simply supply a value of $True to whichever of those 2 you don't want to do. 

This is an early draft and is not very pretty right now.  It was also something that I used to learn on so please keep that in mind.  Alos, please read over the code before you run this in any environment.  I will not be held responsible for anyone destroying accounts, systems, etc.
