<#
Script Powershell V1
=> Gets the first .zip archive in the $path directory
=> Backup the files in the $dest directory
=> Extract archive in the $dest directory, overwriting existing files
=> Print output in log file exec.log

#>
$RootFolder = Split-Path -Path $MyInvocation.MyCommand.Path  
$path = "C:\rep"
$dest = "C:\dest"
$File = get-childitem "$path\*.zip" | Select-Object -First 1
$date=$File.LastWriteTime
$log = $File -replace 'zip','log'

    # Create a Zip archive named Backup-HH,mm,ss-dd-MM-yyyy.zip
    # $sourcedir: folder to zip
    # $destination: location of the new zip
    function ZipBackup($sourcedir, $destination)
    {   
        [System.Reflection.Assembly]::LoadFrom("$Rootfolder\Ionic.Zip.dll") | Out-Null;
        $namezip = "Backup-$((get-date).tostring(‘HH,mm,ss-dd-MM-yyyy’)).zip"    
        $zipfile = new-object Ionic.Zip.ZipFile
        $e = $zipfile.AddSelectedFiles("name != Backup-*.zip", $sourcedir, "backup", $true)
        # Create folder Backup in $destination if not here
        If (! (Test-Path "$destination\Backup")) { new-item "$destination\Backup" -ItemType Directory }
        $zipfile.Save($destination + "\Backup\$namezip");
        $zipfile.Dispose()
        
    }

    function UnZip($file, $destination)
    {
        $shell = new-object -com shell.application
        $zip = $shell.NameSpace($file)
        foreach($item in $zip.items())
        {
            $shell.Namespace($destination).copyhere($item, 0x10)
        }
    }






# if $sclog is not already here
$sclog = "$Rootfolder\exec.log"
if (! (Test-Path $sclog)) {new-item $sclog -type file}
ADD-content -path $sclog -value "$((get-date).tostring(‘HH,mm,ss-dd-MM-yyyy’)): Starting execution"

# if there is no Zip archive in $path folder
if ((!"$File") -or "$File" -eq ' ') {
    ADD-content -path $sclog -value "$((get-date).tostring(‘HH,mm,ss-dd-MM-yyyy’)): Zip Archive not present in $path directory"

} else {

    # if $log is not already here
    if (! (Test-Path $log))
    {
        ADD-content -path $sclog -value "$((get-date).tostring(‘HH,mm,ss-dd-MM-yyyy’)): Create $log"
        new-item $log -type file
        ADD-content -path $log -value "Log for $File : LastWrite"
    }

    # get last line of log file
    $dateOld = (Get-Content -Path $log)[-1]

    # Check difference between LastWriteTime of the file and LastWriteTime in the log
    $diff = NEW-TIMESPAN –Start $dateOld –End $date
    if ($diff.Seconds -eq 0)
    {
        ADD-content -path $sclog -value "$((get-date).tostring(‘HH,mm,ss-dd-MM-yyyy’)): Nothing to do, everything already up to date"
    } else {
        ADD-content -path $sclog -value "$((get-date).tostring(‘HH,mm,ss-dd-MM-yyyy’)): File $File have been mofified:"    
        
        ADD-content -path $sclog -value "$((get-date).tostring(‘HH,mm,ss-dd-MM-yyyy’)): Create Backup"
        ZipBackup -Sourcedir $dest -Destination $dest
        
        ADD-content -path $sclog -value "$((get-date).tostring(‘HH,mm,ss-dd-MM-yyyy’)): Copy to destination"
        UnZip -File "$File" -Destination $dest
        
        # Set the LastWriteTime in log
        ADD-content -path $sclog -value "$((get-date).tostring(‘HH,mm,ss-dd-MM-yyyy’)): Write in $log"
        ADD-content -path $log -value $date
    }
}

 ADD-content -path $sclog -value "$((get-date).tostring(‘HH,mm,ss-dd-MM-yyyy’)): Done `n"
 ADD-content -path $sclog -value "`n"