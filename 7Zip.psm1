function Extract-7Zip {
        [CmdletBinding()]
        Param(
            [Parameter(Mandatory=$true)]
            [string]$Zip, 
            [string]$ExtractPath, 
            [string]$Password, 
            [switch]$Force,
            [switch]$Continue,
            [string]$Exclude
        )

        Begin{
            Write-Debug "`$ExtractPath: $ExtractPath"
            #Get information about the sorce directory to call later
            $FilesToZip = Get-Item $Zip -ErrorAction Continue
            $Parrent = $Zip.Parent.Name
            
            # Look for the 7zip executable.
            $pathTo32Bit7Zip = "C:\Program Files (x86)\7-Zip\7z.exe"
            $pathTo64Bit7Zip = "C:\Program Files\7-Zip\7z.exe"
            if (Test-Path $pathTo64Bit7Zip) { $pathTo7ZipExe = $pathTo64Bit7Zip } 
            elseif (Test-Path $pathTo32Bit7Zip) { $pathTo7ZipExe = $pathTo32Bit7Zip }
            elseif (Test-Path $env:temp\7Za.exe) { $pathTo7ZipExe = "$env:temp\7Za.exe"}
            else { 
                #This will download 7za to execute more complex zip commands.
                $url = "http://d.7-zip.org/a/7za920.zip"
                $output = "$env:Temp\7za920.zip"

                $wc = New-Object System.Net.WebClient
                $wc.DownloadFile($url, $output)

                #Expand the download
                $shell = New-Object -com shell.application
                $zip = $shell.NameSpace($output)
                foreach($item in $zip.items()){
                    $shell.Namespace($env:Temp).copyhere($item,16)
                }

                $pathTo7ZipExe = Join-Path $env:Temp "7za.exe"
             }
    
            Write-Debug "`$pathTo7ZipExe: $pathTo7ZipExe"

            #Arguments
            # Create the arguments to use to zip up the files.
            # Command-line argument syntax can be found at: http://www.dotnetperls.com/7-zip-examples
            $arguments = "x `"$FilesToZip`" -o`"$ExtractPath`""
            #EndArguments

            #Password
            if (!([string]::IsNullOrEmpty($Password))) { $arguments += " -p$Password" }
            #EndPassword

            #Force
            if ($Force) { $arguments += " -aoa" }
            #EndForce


            #Exclude
            $Excludes = $Exclude -split ","
            Clear-Variable Exclude
            if (!([string]::IsNullOrEmpty($Excludes))) { 
                Foreach ($item in $Excludes) {
                    $Exclude += $item -replace [Regex]::Escape("$item")," -xr!$item"    
                }
            $arguments += "$Exclude"
            }
            #EndExclude
        }

        Process{
            Try{
                Write-Output "Extracting: $FilesToZip"

                Write-Debug "Start-Process -FilePath $pathTo7ZipExe -ArgumentList $arguments -PassThru -Wait -WindowStyle Hidden"
                Write-Verbose "Start-Process -FilePath $pathTo7ZipExe -ArgumentList $arguments -PassThru -Wait -WindowStyle Hidden"

                $p = Start-Process -FilePath $pathTo7ZipExe -ArgumentList $arguments -PassThru -Wait -WindowStyle Hidden
            }
    
            Catch{
                Write-Output "There was an error zipping the file: $($_.Exception)"
                if ($Continue -ne $true) {
                    throw "There was a problem creating the zip file '$ZipFilePath'."
                }
              #Break
            }
        }

        End{
            If($?){
                Write-Output " Extract Completed Successfully."
            }
        }
    }

function New-7Zip {
        [CmdletBinding()]
        Param(
            [Parameter(Mandatory=$true)]
            [string]$Zip, 
            [string]$PathtoZip, 
            [string]$Password, 
            [switch]$Recurse,
            [switch]$Continue,
            [string]$Exclude
        )

        Begin{
            Write-Debug "`$PathtoZip: $PathtoZip"
            
            # Look for the 7zip executable.
            $pathTo32Bit7Zip = "C:\Program Files (x86)\7-Zip\7z.exe"
            $pathTo64Bit7Zip = "C:\Program Files\7-Zip\7z.exe"
            if (Test-Path $pathTo64Bit7Zip) { $pathTo7ZipExe = $pathTo64Bit7Zip } 
            elseif (Test-Path $pathTo32Bit7Zip) { $pathTo7ZipExe = $pathTo32Bit7Zip }
            elseif (Test-Path $env:temp\7Za.exe) { $pathTo7ZipExe = "$env:temp\7Za.exe"}
            else { 
                #This will download 7za to execute more complex zip commands.
                $url = "http://d.7-zip.org/a/7za920.zip"
                $output = "$env:Temp\7za920.zip"

                $wc = New-Object System.Net.WebClient
                $wc.DownloadFile($url, $output)

                #Expand the download
                $shell = New-Object -com shell.application
                $zip = $shell.NameSpace($output)
                foreach($item in $zip.items()){
                    $shell.Namespace($env:Temp).copyhere($item,16)
                }

                $pathTo7ZipExe = Join-Path $env:Temp "7za.exe"
             }
    
            Write-Debug "`$pathTo7ZipExe: $pathTo7ZipExe"

            #Arguments
            # Create the arguments to use to zip up the files.
            # Command-line argument syntax can be found at: http://www.dotnetperls.com/7-zip-examples
            $arguments = "a `"$Zip`" `"$PathtoZip`""
            #EndArguments

            #Password
            if (!([string]::IsNullOrEmpty($Password))) { $arguments += " -p$Password" }
            #EndPassword

            #Recurse
            if ($Recurse) { $arguments += " -r" }
            #EndRecurse


            #Exclude
            $Excludes = $Exclude -split ","
            Clear-Variable Exclude
            if (!([string]::IsNullOrEmpty($Excludes))) { 
                Foreach ($item in $Excludes) {
                    $Exclude += $item -replace [Regex]::Escape("$item")," -xr!$item"    
                }
            $arguments += "$Exclude"
            }
            #EndExclude
        }

        Process{
            Try{
                Write-Output "Compressing: $FilesToZip"

                Write-Debug "Start-Process -FilePath $pathTo7ZipExe -ArgumentList $arguments -PassThru -Wait -WindowStyle Hidden"
                Write-Verbose "Start-Process -FilePath $pathTo7ZipExe -ArgumentList $arguments -PassThru -Wait -WindowStyle Hidden"

                $p = Start-Process -FilePath $pathTo7ZipExe -ArgumentList $arguments -PassThru -Wait -WindowStyle Hidden
            }
    
            Catch{
                Write-Output "There was an error zipping the file: $($_.Exception)"
                if ($Continue -ne $true) {
                    throw "There was a problem creating the zip file '$ZipFilePath'."
                }
              #Break
            }
        }

        End{
            If($?){
                Write-Output " Compression Completed Successfully."
            }
        }
    }