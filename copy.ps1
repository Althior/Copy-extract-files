<#
 #   Script Powershell V2
 #   => Gets the first .zip archive in the $path directory
 #   => Backup the files in the $dest directory
 #   => Extract archive in the $dest directory, overwriting existing files
 #   => Print output in log file transcript.log
 #   => logoff the current rdp session
 #
 #>

# import functions to manage Zip
$RootFolder = Split-Path -Path $MyInvocation.MyCommand.Path  
. "$RootFolder\zip.ps1"


$path = "C:\rep" #"\\tsclient\Copy"
$dest = "C:\dest"
$File = Get-Childitem "$path\*.zip" | Select-Object -First 1
$date=$File.LastWriteTime
$log = $File -replace 'zip','log'
$todayDate = $((get-date).tostring(‘HH,mm,ss_dd-MM-yyyy’))

# Logs
$pathlog = "$RootFolder\Logs"
If (! (Test-Path $pathlog)) { new-item $pathlog -ItemType Directory }
$tlog = "$pathlog\$TodayDate-copy.log"
if (! (Test-Path $tlog)) {new-item $tlog -type file}

# Start transcription in .\copy.log file
Start-Transcript -path $tlog



Write-Host "[$todayDate] Starting execution"

# if there is no Zip archive in $path folder
if ((!"$File") -or "$File" -eq ' ') 
{
    Write-Host "[$todayDate] Zip Archive not present in $path directory"
}
else 
{
    # if $log is not already here
    if (! (Test-Path $log))
    {
        Write-Host "[$todayDate] Create $log"
        new-item $log -type file
        ADD-content -path $log -value "Log for $File : LastWrite"
    }

    # get last line of log file
    $dateOld = (Get-Content -Path $log)[-1]

    # Check difference between LastWriteTime of the file and LastWriteTime in the log
    $diff = NEW-TIMESPAN –Start $dateOld –End $date
    if ($diff.Seconds -eq 0)
    {
        Write-Host "[$todayDate] Nothing to do, everything already up to date"
    } 
    else 
    {
        Write-Host "[$todayDate] File $File have been mofified:"    
        
        Write-Host "[$todayDate] Create Backup"
        Zip -Sourcedir $dest -Destination $dest
        
        Write-Host "[$todayDate] Copy to destination"
        UnZip -File "$File" -Destination $dest
        
        # Set the LastWriteTime in log
        Write-Host "[$todayDate] Write in $log"
        ADD-content -path $log -value $date
    }
}



# Logoff current session
function Get-TSSessions {
    # parse results from Qwinsta into object
    qwinsta /server:"localhost" | ForEach-Object {$_.Trim() -replace "\s+",","} | ConvertFrom-Csv
}
# Get ID for session of type 'rdpwd'
$ID = Get-TSSessions | ? { $_.type -eq 'rdpwd' }
Write-Host "[$todayDate] Logoff $ID"
Write-Host "[$todayDate] Done"
Stop-Transcript

#Logoff $ID.ID /server:localhost