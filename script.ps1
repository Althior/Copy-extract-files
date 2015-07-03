$path = "C:\rep\"
$name = "test"
$File=Get-ChildItem "$path$name.zip"
$date=$File.LastWriteTime
$log = "$path$name.log"
$dest = "C:\dest\"

function ZipBackup( $name, $sourcedir )
{    
    [System.Reflection.Assembly]::LoadFrom("Ionic.Zip.dll");
    $dirtozip = "c:\rep"
    $namezip = "Backup-$((get-date).tostring(‘HH,mm,ss-dd-MM-yyyy’)).zip"    
    $zipfile = new-object Ionic.Zip.ZipFile
    $e = $zipfile.AddDirectory($dest, $name)
    $zipfile.Save($namezip);
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



$dateOld = (Get-Content -Path $log)[-1]
$diff = NEW-TIMESPAN –Start $dateOld –End $date

# Check difference between LastWriteTime of the file and LastWriteTime in the log
if ($diff.Seconds -eq 0)
{
    Write-Host "Nothing to do, everything already up to date"
} else {
    Write-Host "File $name.zip have been mofified:"    
    
    Write-Host "Create Backup"
    
    ZipBackup -Name "Backup" -Sourcedir $dest
    
    Write-Host "Copy to dest"
    UnZip -File "$path$name.zip" -Destination $dest
}

