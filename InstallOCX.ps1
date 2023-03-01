#SET THE POLICY BELOW BEFORE RUNNING THE SCRIPT
#Set-ExecutionPolicy -ExecutionPolicy Unrestricted
#RESET the policy back to default after.
#Set-ExecutionPolicy -ExecutionPolicy Restricted

#Get-ExecutionPolicy -List
#
#   * Restricted/Default:
#    This is the default policy. You may not launch any script. The only possibility is to use PowerShell in ‘interactive mode’.
#   * Allsigned:
#    All the scripts that should be executed on the machine must be signed by a ‘Trusted Publisher’.
#   * RemoteSigned:
#    This concerns only the scripts that have been downloaded from the internet. These scripts must be signed by a ‘Trusted Publisher’.
#   * Unrestricted:
#    No constraints on the execution of scripts. All the scripts will be executed if you accept a warning message. I do not recommend it in a production environment.
#   * Bypass:
#    No blockage, no warning message. Everything is executed without control.
#   * Undefined:
#    Removes the currently assigned execution policy from the current scope. This parameter will not remove an execution policy that is set in a Group Policy scope.


# Check if the current user is an administrator
$currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
$isAdmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if ($isAdmin) {
    Write-Host "This script is running with administrative privileges."
	if (Get-Module -ListAvailable -Name 7Zip4PowerShell) {
    Write-Host "The 7Zip4PowerShell module is installed."
} else {
    Write-Host "The 7Zip4PowerShell module is not installed."
	Install-Module 7Zip4PowerShell -Scope CurrentUser -Force -Verbose
}

    
}
else {
    Write-Host "This script requires administrative privileges to run."
    Exit
}

# Set the download link and local file path
$downloadUrl = "http://datadesign.gr/Content/dist/OCX.zip"
$filePath = "$env:USERPROFILE\Downloads\OCX.zip"

Start-BitsTransfer -Source $downloadUrl -Destination $filePath

# Check if the file was downloaded successfully
if (-not (Test-Path $filePath)) {
    # If the file doesn't exist, output an error message and stop the script
    Write-Host "Error: Failed to download the file."
    Exit
}

# Extract the file to a folder with the same name as the file
$extractPath = "$env:USERPROFILE\Downloads\OCX\"

if (Test-Path $extractPath) {
   
    Write-Host "Folder Exists"
    # Perform Delete file from folder operation
}
else
{
  
    #PowerShell Create directory if not exists
    New-Item $extractPath -ItemType Directory
    Write-Host "Folder Created successfully"

}
Write-Host "Attempting to expand from $filepath to $extractPath"
#Expand-Archive -Path $filePath -DestinationPath $extractPath

Expand-7Zip -ArchiveFileName $filePath -TargetPath $extractPath -Verbose


# Output a message indicating success
Write-Host "File downloaded and extracted successfully."

# Set the source and destination folder paths
$sourcePath = "$env:USERPROFILE\Downloads\OCX\"
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
	Write-Host "Registered: $file"
}

# Display a message for success
Write-Host "All files copied and registered successfully."
