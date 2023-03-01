# Check if the current user is an administrator
$currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
$isAdmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if ($isAdmin) {
    Write-Host "This script is running with administrative privileges."
}
else {
    Write-Host "This script requires administrative privileges to run."
    Exit
}

# Set the download link and local file path
$downloadUrl = "https://datadesignsa-my.sharepoint.com/personal/a_patsouros_datadesignsa_onmicrosoft_com/_layouts/15/download.aspx?SourceUrl=%2Fpersonal%2Fa%5Fpatsouros%5Fdatadesignsa%5Fonmicrosoft%5Fcom%2FDocuments%2FDeploys%2FOCX%2Ezip"
$filePath = "$env:USERPROFILE\Downloads\OCX.zip"

# Create a WebClient object to download the file
$client = New-Object System.Net.WebClient

# Download the file and save it to the specified path
$client.DownloadFile($downloadUrl, $filePath)

# Check if the file was downloaded successfully
if (-not (Test-Path $filePath)) {
    # If the file doesn't exist, output an error message and stop the script
    Write-Host "Error: Failed to download the file."
    Exit
}

# Extract the file to a folder with the same name as the file
$extractPath = Split-Path -Path $filePath -Parent
Expand-Archive -Path $filePath -DestinationPath $extractPath

# Output a message indicating success
Write-Host "File downloaded and extracted successfully."


# Set the source and destination folder paths
$sourcePath = "$env:USERPROFILE\Downloads\OCX"
$destPath = "C:\Windows\SysWow64"

# Get all files with .OCX extension from the source folder
$files = Get-ChildItem -Path $sourcePath -Filter "*.OCX"

# Copy the files to the destination folder
foreach ($file in $files) {
    Copy-Item -Path $file.FullName -Destination $destPath
}

# Register the copied files using regsvr32.exe
foreach ($file in $files) {
    $regsvr32Path = Join-Path $env:SystemRoot "SysWOW64\regsvr32.exe"
    $arg = "/s $($file.FullName)"
    & $regsvr32Path $arg
}

# Display a message for success
Write-Host "All files copied and registered successfully."
