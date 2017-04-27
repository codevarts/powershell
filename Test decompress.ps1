Add-Type -AssemblyName System.IO.Compression.FileSystem

function Unzip
{
    param(
            [Parameter(Mandatory=$True)][string]$fileToUnZip,
            [Parameter(Mandatory=$True)][string]$outputPath
         )

    $fileToUnZip = $(Resolve-Path $fileToUnZip)
    $outputPath = $(Resolve-Path $outputPath)

    Write-Host("In: $fileToUnZip`r`nOut: $outputPath")

    $folderName = Split-Path $fileToUnZip -Leaf
    Write-Host($folderName)
    
    Write-Host([io.path]::GetFileNameWithoutExtension($fileToUnZip))
    
    # Check it temp extract directory exists otherwise create it
    if (!(Test-Path -PathType Container C:\TempExtracted))
    {
      New-Item -ItemType Directory -Name TempExtracted -Path 'C:\'  
    }

    # Temp Directory to extract to
    $temp = 'C:\TempExtracted\'

    # Extract the Zipped files to the TempExtract Directory
    [System.IO.Compression.ZipFile]::ExtractToDirectory($fileToUnZip, $temp)
    Write-Host("Unzipping $fileToUnZip to temp directory")

    # Get newest file in directory
    $folderToCopy = Get-ChildItem -Path $temp | Sort-Object LastWriteTime | Select-Object -Last 1

    # Copy new file and overwrite existing file
    Copy-Item -Force -Recurse ($temp + $folderToCopy + "\") -Destination $outputPath
    Write-Host("File Copied")
}
