Write-Host "Auto Installer Script v0.8"
#SET THE POLICY BELOW BEFORE RUNNING THE SCRIPT
#Set-ExecutionPolicy -ExecutionPolicy Unrestricted
#Set-ExecutionPolicy -ExecutionPolicy Restricted
Start-Sleep -Seconds 5
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

if (-not [Environment]::Is64BitOperatingSystem)
{
    Write-Error "This script requires a 64-bit operating system."
    Exit
}

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
$downloadUrl = "http://datadesign.gr/Content/dist/OCX.zip"
$fontsUrl = "http://datadesign.gr/Content/dist/fonts.zip"
$filePath = "$env:USERPROFILE\Downloads\OCX.zip"
$fontsPath = "$env:USERPROFILE\Downloads\fonts.zip"

Start-BitsTransfer -Source $downloadUrl -Destination $filePath
Start-BitsTransfer -Source $fontsUrl -Destination $fontsPath

# Check if the file was downloaded successfully
if (-not (Test-Path $filePath)) {
    # If the file doesn't exist, output an error message and stop the script
    Write-Host "Error: Failed to download the file."
    Exit
}

# Extract the file to a folder with the same name as the file
$extractPath = "$env:USERPROFILE\Downloads\OCX\"
$extractPath2 = "$env:USERPROFILE\Downloads\fonts\"

if (Test-Path $extractPath) {
   
    Write-Host "Folder Exists"
    # Perform Delete file from folder operation
}
else
{
  
    #PowerShell Create directory if not exists
    New-Item $extractPath -ItemType Directory
    Write-Host "Folder $extractPath Created successfully"
}

if (Test-Path $extractPath2) {
   
    Write-Host "Folder Exists"
    # Perform Delete file from folder operation
}
else
{
  
    #PowerShell Create directory if not exists
    New-Item $extractPath2 -ItemType Directory
    Write-Host "Folder $extractPath2 Created successfully"
}
######## OCX ########
Write-Host "Attempting to expand from $filepath to $extractPath"
try {
	Expand-Archive -Path $filePath -DestinationPath $extractPath
    } catch {
    Write-Host "Extraction failed. Trying alternative method..."
    	if (Get-Module -ListAvailable -Name 7Zip4PowerShell) {
    		Write-Host "The 7Zip4PowerShell module is installed."
	} else {
    		Write-Host "The 7Zip4PowerShell module is not installed."
		Install-Module 7Zip4PowerShell -Scope CurrentUser -Force -Verbose
	}    
	Expand-7Zip -ArchiveFileName $filePath -TargetPath $extractPath -Verbose
}
######## FONTS ########
Write-Host "Attempting to expand from $fontsPath to $extractPath2"
try {
	Expand-Archive -Path $fontsPath -DestinationPath $extractPath2
    } catch {
    Write-Host "Extraction failed. Trying alternative method..."
    # Your alternative extraction method here
    	if (Get-Module -ListAvailable -Name 7Zip4PowerShell) {
    		Write-Host "The 7Zip4PowerShell module is installed."
	} else {
    		Write-Host "The 7Zip4PowerShell module is not installed."
		Install-Module 7Zip4PowerShell -Scope CurrentUser -Force -Verbose
	}    
	Expand-7Zip -ArchiveFileName $fontsPath -TargetPath $extractPath2 -Verbose
}
# Output a message indicating success
Write-Host "Files downloaded and extracted successfully."

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

######## Installing fonts ########
$fonts = Get-ChildItem -Path $extractPath2 -Filter "*.ttf"
Write-Host "Installing Fonts"
foreach ($fontFile in $fonts) {
    $fontFilePath = $fontFile.FullName
    $fontDestinationPath = $env:windir + "\Fonts\" + $fontFile.Name
    #New-Item -ItemType File $fontDestinationPath -Force | Out-Null
    #Copy-Item $fontFilePath $fontDestinationPath -Force
    #Invoke-Item $fontDestinationPath
    	if(Test-Path -Path $fontDestinationPath){
		$FontFile.Name + " already installed"
	}
	else {
		"Installing font " + $fontFile.Name
		#Extract Font information for Reqistry
		$ShellFolder = (New-Object -COMObject Shell.Application).Namespace($fontsPath)
		$ShellFile = $shellFolder.ParseName($fontFile.name)
		$ShellFileType = $shellFolder.GetDetailsOf($shellFile, 2)
		#Set the $FontType Variable
		If ($ShellFileType -Like '*TrueType font file*') {$FontType = '(TrueType)'}
		#Update Registry and copy font to font directory
		$RegName = $shellFolder.GetDetailsOf($shellFile, 21) + ' ' + $FontType
		New-ItemProperty -Name $RegName -Path "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Fonts" -PropertyType string -Value $fontFile.name -Force | out-null
		Copy-item $fontFile.FullName -Destination $systemFontsPath
		"Done"
	}
}
Write-Host "Done installing fonts!"
########################################
# Display a message for success
Write-Host "All files copied and registered successfully."
######## Deleting not needed OCX ########
Write-Host "Deleting archive: $filePath"
Remove-Item -Path $filePath -Recurse -Force
Write-Host "Deleting Folder: $extractPath"
Remove-Item -Path $extractPath -Recurse -Force
######## Deleting not needed fonts ########
Write-Host "Deleting archive: $filePath"
Remove-Item -Path $fontsPath -Recurse -Force
Write-Host "Deleting Folder: $extractPath2"
Remove-Item -Path $extractPath2 -Recurse -Force
$userInput = Read-Host "Create ODBC Record? 'yes' to continue"
If ($userInput -eq "yes") {
	$dsnName = "DIAGBASE"
	$driverName = "SQL Server"
	$databaseName = "DIAGBASE"
	$odbc = Read-Host "Please input the server ip/name"
	if ($odbc -ne $null -and $odbc -ne "") {
		if (Get-OdbcDsn -Name $dsnName -ErrorAction SilentlyContinue) {
		   Remove-OdbcDsn -Name $dsnName -DsnType "User" -Confirm:$false
		}
		Add-OdbcDsn -Name $dsnName -DriverName $driverName -DsnType "User" -SetPropertyValue @("Server=$odbc", "Trusted_Connection=No", "Database=$dsnName")
	}
}
Write-Host "All Done."
