$RootFolder = Split-Path -Path $MyInvocation.MyCommand.Path  

# UnZip $file to $destination
# Overwrite existing files
function UnZip($file, $destination)
{
    # Check params    
    if ($file -eq $null -or $destination -eq $null)
    {
        throw [System.Exception] "UnZip: parameter(s) missing."
    }
    if (!(Test-Path $file)) 
    {
       throw [System.IO.FileNotFoundException] "UnZip: $file not found."
    }
    if (!(Test-Path $destination))
    {
       throw [System.IO.FileNotFoundException] "UnZip: $destination not found."
    }
    
    $shell = new-object -com shell.application
    $zip = $shell.NameSpace($file)
    
    # Check if archive is correct
    $zipi = $zip.items() | select @{n='Fullname'; e={$_.name}}
    if ($zipi -eq $null) 
    {
        throw [System.Exception] "UnZip: $file is empty or is not a valid Zip archive."
    }
    
    foreach($item in $zip.items())
    {
        $shell.Namespace($destination).copyhere($item, 0x10)
    }
}

# Create a Zip archive
# $sourcedir: folder to zip
# $destination: location of the new zip
function Zip($name, $sourcedir, $destination)
{   
    try {[System.Reflection.Assembly]::LoadFrom("$Rootfolder\Ionic.Zip.dll") | Out-Null;}
    catch { throw [System.Exception] "Zip: Load Assembly Fail."; break}
    $zipfile = new-object Ionic.Zip.ZipFile
    $e = $zipfile.AddDirectory($sourcedir)
    #$e = $zipfile.AddSelectedFiles("name = *", $sourcedir, $true)
    $zipfile.Save("$destination\$name");
    $zipfile.Dispose()      
      
}


Zip "test.zip" "C:\rep" "C:\rep"