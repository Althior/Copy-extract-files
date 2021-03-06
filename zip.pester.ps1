$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path).Replace(".Tests.", ".")
. "$here\$sut"

Describe "Zip" {

# Set temp environment
$path = (Split-Path $here)+"\Tests"
If (! (Test-Path $path)) { new-item $path -ItemType Directory }
Zip -name "empty.zip" -sourcedir $path -destination $path

new-item "$($path)\test_file1.txt" -type file -Force
ADD-content -path "$($path)\test_file1.txt" -value "sfsgdgdfgd"
new-item "$($path)\test_file2.txt" -type file -Force

Zip -name "ok.zip" -sourcedir $path -destination $path

Write-host (get-childitem $path)


    Context "UnZip function" { 
     
        It "Fails if 1st param is absent" {
            {Unzip -Destination $path}  | Should throw "UnZip: parameter(s) missing."
        }
        It "Fails if 2nd param is absent" {
            {Unzip -File "$path\toto.zip"}  | Should throw "UnZip: parameter(s) missing."
        }
        It "Fails if both param are absent" {
            {Unzip}  | Should throw "UnZip: parameter(s) missing."
        }
        It "Fails if source file is not found" {
            {Unzip -File "$here\toto.zip" -Destination $path}  | Should throw "UnZip: $here\toto.zip not found."
        }
        It "Fails if destination folder is not found" {
            {Unzip -File "$path\ok.zip" -Destination "C:\dedes64fs"}  | Should throw "UnZip: C:\dedes64fs not found."
        }
        It "Fails if zip archive is not correct" {
            {Unzip -File "$path\empty.zip" -Destination $path}  | Should throw "UnZip: $path\empty.zip is empty or is not a valid Zip archive."
        }
        It "Pass if everything is OK" {
            {Unzip -File "$path\ok.zip" -Destination $path}  | Should not throw
        }
    }
    
    <#Context "Zip function" { 
    function Loadfrom {}
    Mock LoadFrom {throw [System.Exception] "Zip: Load Assembly Fail."}
    Write-host (get-childitem $path)
        It "Fails if assembly file is absent" {
            {Zip -name "ZIP1.zip" -sourcedir $path -destination $path}  | Should throw "Zip: Load Assembly Fail."
        }
    }#>

# Flush temp directory
get-childitem $path | remove-item -recurse -force
}