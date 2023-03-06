$url = "https://api.github.com/repos/rustdesk/rustdesk/releases/latest"
$response = Invoke-RestMethod -Uri $url

$v64 = ""
$v32 = ""
$is64 = [Environment]::Is64BitOperatingSystem
foreach ($asset in $response.assets) {
    #Write-Host $asset.browser_download_url
    if ($asset.browser_download_url.EndsWith("windows_x32-portable.zip")) {
        $v32 = $asset.browser_download_url
    }
    if ($asset.browser_download_url.EndsWith("windows_x64-portable.zip")) {
        $v64 = $asset.browser_download_url
    }
}
# -------------------
Write-Host "Version for 64bit : $v64" 
Write-Host "Version for 32bit : $v32"
if ($is64) {
    Write-Host "We shall use 64bit version"
    $use = $v64
    }
    else {
    Write-Host "We shall use 32bit version"
    $use = $v32
    }

$currentUserFolder = [Environment]::GetFolderPath("UserProfile")
$tempFolderPath = Join-Path $currentUserFolder "TempRustDesk"
$outputFile = Join-Path $tempFolderPath "RustDesk.zip"

if (-not (Test-Path $tempFolderPath -PathType Container)) {
    New-Item -ItemType Directory -Path $tempFolderPath
    Start-BitsTransfer -Source $use -Destination $outputFile 
    Expand-Archive -Path $outputFile -DestinationPath $tempFolderPath -Force
    $exeFile = Get-ChildItem -Path $tempFolderPath -Filter Rust*.exe | Select-Object -First 1
    Invoke-Item $exeFile.FullName
}
else { 
    Write-Host "Searching in: $tempFolderPath"
    $exeFile = Get-ChildItem -Path $tempFolderPath -Filter Rust*.exe | Select-Object -First 1
    if ($exeFile -eq $null) { 
        if (Test-Path $outputFile) { 
            Expand-Archive -Path $outputFile -DestinationPath $tempFolderPath -Force
            $exeFile = Get-ChildItem -Path $tempFolderPath -Filter Rust*.exe | Select-Object -First 1
            Invoke-Item $exeFile.FullName
        }
        else {
            Start-BitsTransfer -Source $use -Destination $outputFile 
            Expand-Archive -Path $outputFile -DestinationPath $tempFolderPath -Force
            $exeFile = Get-ChildItem -Path $tempFolderPath -Filter Rust*.exe | Select-Object -First 1
            Invoke-Item $exeFile.FullName
        }
    }
    else {
        Invoke-Item $exeFile.FullName
    }
}
Write-Host "Script Finished."
