<#
Synopsis: Script that automatically does a backup and restore of the interbase Database, the variables prior to
          Start-Transcript -Path $ConsoleTranscipt line has to be set as per your environment
	  General Process for script is as per below
	  
	  Main Functions
	  1.1 Stop All Services 
	  1.2 Start Database File Level Copy (robocopy standard, deletes files in destination folder.)
	  1.3 Start MT32 Folder File Level Copy (Robocopy for changes only)
	  1.4 Start Interbase Services
	  1.5 Start GBAK of Interbase Backup
	  1.6 Keep checking the file size of the Database as the GBAK is run
	  1.7 Start Restore
	  
	  Helper Function
	  2.1 Check-ServiceStatus --> Check the Services status
	  2.2 Mailer --> Email Out 
	  2.3 GStat-Check --> Check Datastable Statistics
	  2.4 checkCurrentFileSize --> Since visually cannot see backup or restore process at times, this iterates as backup or restore file size increases
	  2.5 checkLogFile --> Looks at Logs file specified (used in mailer for messages)
	  2.6 Start-Transcript --> Captures Log of The Script Run in destination specified
	  
	  Program Run
	  3.1 Functions are called in "# Program Run Steps"
	  3.2 Uncomment the Mailer function as required.
#>

#Global Environment Variables
# Blat Program is a emailing program though no longer used by myself.
$env:PATH = $env:PATH + "C:\Program Files\BLAT;" + "C:\Program Files\Embarcadero\InterBase\bin"
$env:ISC_USER="username" # Set database username here
$env:ISC_PASSWORD="password" # Set Database PAssword here

#Email Variables
$smtpserver="<your mailhost>.<yourdomain>.com" #
$emailTo="xxxx@xxxx.net.nz"

#Live Datababse File & FolderPath
$Live_DBCopySource_fldr="F:\MT32\Data\"
$Live_MT32_Path="F:\MT32\Data\MT32.IB"
$Live_BLOB_Path="F:\MT32\Data\BLOB.IB"
$Live_MT32_Fldr="E:\MT32\"

#File Level Copy Destinations
$Dstn_Copy_Fldr_DB="F:\MT32\LAST_GOOD_DBS"
$Dstn_Copy_Fldr_MT32="E:\MT32_COPY\"

#Database Backup Destinations via GBAK
$setDate=(get-date -f yyyy-MMM-dd) #Sets date variable permanently so that if you run over midnight can reference same file path
$bkup_Destination_Fldr="G:\MT32_Backups\Database"
$MT32_bkup_destination="G:\MT32_Backups\Database\MT32_$($setDate).ibk"
$BLOB_bkup_destination="G:\MT32_Backups\Database\BLOB_$($setDate).ibk"

#Logging Paths
$Log_Fldr_Path="G:\MT32_Backups\Logs\" #Update only Log Folder Path with trailing backslash
$log_Robocopy_DB_Path= -join ($Log_Fldr_Path,"robcopy_DB_$(get-date -f yyyy-MMM-dd-@-HHmm).log")
$log_Robocopy_MT32_Path= -join ($Log_Fldr_Path,"robcopy_MT32_$(get-date -f yyyy-MMM-dd-@-HHmm).log")
$MT32_bkup_log = -join ($Log_Fldr_Path, "MT32_Backup_log_$(get-date -f yyyy-MMM-dd-@-HHmm).log")
$BLOB_bkup_log = -join ($Log_Fldr_Path, "BLOB_Backup_log_$(get-date -f yyyy-MMM-dd-@-HHmm).log")
$MT32_restore_log= -join ($Log_Fldr_Path, "MT32_Restore_log_$(get-date -f yyyy-MMM-dd-@-HHmm).log")
$BLOB_restore_log= -join ($Log_Fldr_Path, "BLOB_Restore_log_$(get-date -f yyyy-MMM-dd-@-HHmm).log")
$ConsoleTranscipt= -join ($Log_Fldr_Path, "Console_Transcript_$(get-date -f yyyy-MMM-dd-@-HHmm).log")

#List all Medtech Services including Interbase and third party services that access the database
$medtechServices=@('medtechss','drinfo','MedtechGP2GP','IBG_MedTech_IB11','IBS_MedTech_IB11','Tomcat5','ShadowProtectSvc','HealthOneDataExtractor', 'HealthOneAutoUpdate')

#Starts to capture the transcript of the console session in the path specified.
Start-Transcript -Path $ConsoleTranscipt 
#=============================================================

#Block or Unblock Port 3050
function toggleFirewallPortBlock{
    param([string]$NewStatus)

    If($NewStatus -eq "Block")  {
        Write-Host "Creating Rule to Block 3050"
        netsh advfirewall firewall add rule name="Block Interbase 3050" dir=in action=block protocol=TCP localport=3050
    }
    Elseif($NewStatus -eq "Unblock")
    {
        Write-Host "Delete the Block Rule for 3050"
        netsh advfirewall firewall delete rule name="Block Interbase 3050" protocol=tcp localport=3050
    }
}

#Check Current Medtech related Services Running
function Check-ServiceStatus
{
    get-service -name $medtechServices | Sort-Object Status
    Write-host " "
}

# Stops all Services attached to Medtech
function Stop-AllServices{
    Check-ServiceStatus
    write-host -nonewline "Current Service status above, continue STOP all services? (Y/N)"
    $response = read-host     #Checks&Balances
    if ( $response -eq "Y" ){ #Checks&Balances
        get-service -name $medtechServices | stop-service
        Check-ServiceStatus
    } #Checks&Balances
}

#Starts all Services related to Medtech
function Start-AllServices{
    Check-ServiceStatus
    write-host -nonewline "Current Service status above, continue START all services? (Y/N)"
    $response = read-host
    if ( $response -eq "Y" ){
        get-service -name $medtechServices | start-service
        Check-ServiceStatus
    }
}

# Start Interbase Services
function Start-IBServices{
    Check-ServiceStatus
    write-host -nonewline "Current Service status above, press Y to start IB Services? (Y/N)"
    $response = read-host
    if ( $response -eq "Y" ){
        Write-Host "Starting IB Services"
        get-service -name IBG_MedTech_IB11,IBS_MedTech_IB11 | start-service
        Check-ServiceStatus
    }
}

#GStat to check the database
function GStat-Check{
    Write-Host "MT32 GSTAT"
    gstat -h $Live_MT32_Path | Out-Default #Out-Default Forces output to appear in transcript
    Write-Host " "
    Write-Host "BLOB GSTAT"
    gstat -h $Live_BLOB_Path | Out-Default #Out-Default Forces output to appear in transcript
}

<#
Does file level copy of the Databases from the live source to destination)
This Does cover the scenario if the Database Copy Destination is in the same Sub-Folder as the Live
Ensure $Dstn_Copy_Fldr_DB is not in sub-folder of $Live_DBCopySource_fldr
#>
function File-CopyDatabases{
    Write-Host "Below Are the Files in the Destinaton Database File Level Copy Folder"
    Get-ChildItem -Path $Dstn_Copy_Fldr_DB -Recurse
    
    $Daysback = "0" # Delete files older than -2 days from current time
    $CurrentDate = Get-Date
    $DatetoDelete = $CurrentDate.AddDays($Daysback)
    
    write-host -nonewline "Delete the old files before new copy? (Y/N)"
    $response = read-host #Checks&Balances
    if ( $response -eq "Y" ){ #Checks&Balances
        Write-Host "Deleting Files older than " $DatetoDelete "...."
       
        Get-ChildItem -Path $Dstn_Copy_Fldr_DB -Include *.* -Recurse | `
        Where-Object { $_.LastWriteTime -lt $DatetoDelete } | `
        foreach { $_.Delete()}
        
        Stop-AllServices
        Write-Host "Starting Copy"
        ROBOCOPY $Live_DBCopySource_fldr $Dstn_Copy_Fldr_DB /tee /MIR /log+:$log_Robocopy_DB_Path
    } #Checks&Balances
}

#Do a file level Copy of the MT32 Folder.
function File-CopyMT32_Folder{

    Write-host "Robocopy Source: " $Live_MT32_Fldr "to Destination: " $Dstn_Copy_Fldr_MT32 "? (Y/N)"
    Write-Host "Note that that Robocopy will purge anything that does not mirror the Source"
    
    $response = read-host
    if ( $response -eq "Y" ){
        
        #Purge Delete Any Content in Destination Folder
        #Write-Host "Deleting Old MT32 Backup Folder Contents...."
        #Get-ChildItem -Path $Dstn_Copy_Fldr_MT32 -Recurse | Remove-Item -force -recurse
        
        Write-Host 
        Write-Host "Starting Copy of MT32 Live Folder,Press any key to continue ..."
        $x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        
        Write-Host "Copying Current MT32 Folder as a File Level Backup"
        robocopy $Live_MT32_Fldr $Dstn_Copy_Fldr_MT32 /E /MIR /tee /log+:$log_Robocopy_MT32_Path
    }
}

#Using the GBAK command backs up the database from the Live to destination path variables
function BackupDatabases{
    Write-host "Starting BACKUP of Databases from Live!!!"
    Write-Host
    Write-Host "MT32 Live Path "$Live_MT32_Path
    Write-Host "MT32 Backup Destination "$MT32_bkup_destination
    Write-Host 
    Write-Host "BLOB Live Path "$Live_BLOB_Path
    Write-Host "BLOB Backup Destination "$BLOB_bkup_destination
    Write-Host 
    Write-Host "Do you want to continue with the backup? (Y/N)"
    $response = read-host
    
    if ( $response -eq "Y" ){
        Start-IBServices
        Start-Job –Name GbakBackupMT32 -ScriptBlock {
            gbak.exe -b -v -ig -se localhost/MedTech_IBXE7:service_mgr $args[0] $args[1] -y $args[2]
            } -ArgumentList @($Live_MT32_Path, $MT32_bkup_destination, $MT32_bkup_log)
        
        Start-Sleep -s 10
        checkCurrentFileSize -FolderPath $bkup_Destination_Fldr -DBType "MT32" -ProcessType "Backing up"
        Get-Job | Wait-Job
        
        Start-Job -Name GbakBackupBLOB -ScriptBlock {
            gbak.exe -b -v -ig -se localhost/MedTech_IBXE7:service_mgr $args[0] $args[1] -y $args[2]
            } -ArgumentList @($Live_BLOB_Path, $BLOB_bkup_destination, $BLOB_bkup_log)
        
        Start-Sleep -s 10    
        checkCurrentFileSize -FolderPath $bkup_Destination_Fldr -DBType "BLOB" -ProcessType "Backing up"
        Get-Job | Wait-Job
    }
}

#Restore the Datbases backed up from source to destination.
function RestoreDatabases{
    
    Write-host "Starting Restore of Databases from Backup!!!"
    Write-Host "MT32 Live Path "$Live_MT32_Path
    Write-Host "MT32 Backup Source "$MT32_bkup_destination
    Write-Host 
    Write-Host "BLOB Live Path "$Live_BLOB_Path
    Write-Host "BLOB Backup Source "$BLOB_bkup_destination
    Write-Host 
    Write-Host "Do you want to continue with the backup? (Y/N)"
    $response = read-host
    
    if ( $response -eq "Y" ){
        Start-IBServices
        Start-Job –Name gbak-RESTORE-MT32 -ScriptBlock {
            gbak.exe -rep -v -se localhost/MedTech_IBXE7:service_mgr $args[0] $args[1] -y $args[2]
            # 'gbak.exe' $args[0]
            } -ArgumentList @($MT32_bkup_destination, $Live_MT32_Path, $MT32_restore_log)
        
        Start-Sleep -s 10
        checkCurrentFileSize -FolderPath $Live_DBCopySource_fldr -DBType "MT32" -ProcessType "Restoring"
        
        Get-Job | Wait-Job
        
        Start-Job -Name GbakBackupBLOB -ScriptBlock {
            gbak.exe -rep -v -se localhost/MedTech_IBXE7:service_mgr $args[0] $args[1] -y $args[2]
            } -ArgumentList @($BLOB_bkup_destination, $Live_BLOB_Path, $BLOB_restore_log)
         
        Start-Sleep -s 10 
        checkCurrentFileSize -FolderPath $Live_DBCopySource_fldr -DBType "BLOB" -ProcessType "Restoring"
        Get-Job | Wait-Job
    }
}

<#FileSize Check
	.Finds the last modified file in the folder specified and returns its size
	.Keeps returning the file size until it stops changing
	.Covers requirement of ensuring that backup or restore file is size is changing.
#>
function checkCurrentFileSize{
    param([string]$FolderPath, [string]$DBType, [string]$ProcessType )
    Start-Sleep -s 3
    $CurrentFilePath = Get-ChildItem "$FolderPath" | Sort {$_.LastWriteTime} | select -last 1 | % {$_.FullName}
    
    Write-Host $CurrentFilePath
     
    $previousSize = (Get-item $CurrentFilePath).length / 1MB
    Write-Host "Script checks every 10seconds for file size"
    Write-Host $ProcessType $DBType "database current size is" $previousSize "MB"
    Start-Sleep -s 10
    while ($previousSize -ne ((Get-item $CurrentFilePath).length / 1MB)){
        $previousSize = (Get-item $CurrentFilePath).length / 1MB
        Start-Sleep -s 10
        Write-Host $ProcessType $DBType "database current size is" $previousSize
    }
}

function Mailer ($message, $subject) 
<# This is a simple function that that sends a message. 
The variables defined below can be passed as parameters by taking them out  
and putting then in the parentheseis above. 
 
i.e. "Function Mailer ($subject)" 
 
#> 
{ 
$emailFrom="medtecb-BandR@medicalcentre.co.nz" 
$emailTo="user@domain.com"
$smtpserver="smtp.server.com"
$smtp=new-object Net.Mail.SmtpClient($smtpServer) 
$smtp.Send($emailFrom, $emailTo, $subject, $message) 
} 

function checkLogFile{
    param([string]$FileToCheck)
    Write-Host "-===========Log File to Check: ===========-"
    Write-Host $File2Check
    Write-Host
    $a = ((Get-Content -Path $FileToCheck) -join "`n")
    return $a
}

#============================================================================================================
# Program Run Steps
GStat-Check
Check-ServiceStatus
File-CopyMT32_Folder
#Mailer -message (checkLogFile $log_Robocopy_MT32_Path) -subject "Robocopy MT32 Log $(get-date -f yyyy-MMM-dd-@-HHmm)"
#Mailer -message (checkLogFile $ConsoleTranscipt) -subject "Script Transcript $(get-date -f yyyy-MMM-dd-@-HHmm)"

Stop-AllServices
#Mailer -message (checkLogFile $ConsoleTranscipt) -subject "Script Transcript $(get-date -f yyyy-MMM-dd-@-HHmm)"

toggleFirewallPortBlock -NewStatus "Block"
#Mailer -message (checkLogFile $ConsoleTranscipt) -subject "Script Transcript $(get-date -f yyyy-MMM-dd-@-HHmm)"

Check-ServiceStatus
#Mailer -message (checkLogFile $ConsoleTranscipt) -subject "Script Transcript $(get-date -f yyyy-MMM-dd-@-HHmm)"

File-CopyDatabases
#Mailer -message (checkLogFile $log_Robocopy_MT32_Path) -subject "Robocopy DB Log  $(get-date -f yyyy-MMM-dd-@-HHmm)"
#Mailer -message (checkLogFile $ConsoleTranscipt) -subject "Script Transcript $(get-date -f yyyy-MMM-dd-@-HHmm)"

BackupDatabases
Mailer -message (checkLogFile $MT32_bkup_log) -subject "MT32 Backup Log $(get-date -f yyyy-MMM-dd-@-HHmm)"
Mailer -message (checkLogFile $BLOB_bkup_log) -subject "BLOB Backup Log $(get-date -f yyyy-MMM-dd-@-HHmm)"
#Mailer -message (checkLogFile $ConsoleTranscipt) -subject "Script Transcript $(get-date -f yyyy-MMM-dd-@-HHmm)"

RestoreDatabases
#Mailer -message (checkLogFile $MT32_restore_log) -subject "MT32 Restore Log $(get-date -f yyyy-MMM-dd-@-HHmm)"
#Mailer -message (checkLogFile $BLOB_restore_log) -subject "BLOB Restore Log $(get-date -f yyyy-MMM-dd-@-HHmm)"
#Mailer -message (checkLogFile $ConsoleTranscipt) -subject "Script Transcript $(get-date -f yyyy-MMM-dd-@-HHmm)"

toggleFirewallPortBlock -NewStatus "Unblock"
Mailer -message (checkLogFile $ConsoleTranscipt) -subject "Script Transcript $(get-date -f yyyy-MMM-dd-@-HHmm)"

#============================================================================================================

# Pausing Script
Write-Host
Write-Host
Write-Host
Stop-Transcript
Write-Host "Press any key to EXIT ..."
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
