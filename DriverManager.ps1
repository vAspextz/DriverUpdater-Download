# Ensure running as Admin
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltinRole] "Administrator")) {
    Write-Warning "Please run this script as Administrator."
    exit
}

function CleanName($name) {
    return ($name -replace '[^\w\-]', '_')
}

function InstallDriversFromFolder($folderPath) {
    $installers = Get-ChildItem -Path $folderPath -Include *.exe, *.msi, *.inf -Recurse -ErrorAction SilentlyContinue
    if ($installers.Count -eq 0) {
        Write-Warning "No installer files found in $folderPath"
        return
    }
    foreach ($file in $installers) {
        Write-Host "Installing $($file.Name)..."
        try {
            switch ($file.Extension.ToLower()) {
                ".msi" { Start-Process $file.FullName -ArgumentList "/quiet /norestart" -Wait }
                ".exe" { Start-Process $file.FullName -ArgumentList "/S /quiet /norestart" -Wait }
                ".inf" { Start-Process "pnputil.exe" -ArgumentList "/add-driver `"$($file.FullName)`" /install" -Wait }
            }
            Write-Host "Installed $($file.Name) successfully."
        } catch {
            Write-Warning "Failed to install $($file.Name): $_"
        }
    }
}

function UnzipAndInstallAndCopyDrivers($mbFolderName, $usbDriverPath) {
    $userDownloads = [Environment]::GetFolderPath("UserProfile") + "\Downloads"
    $unzippedRoot = "$userDownloads\Unzipped"

    if (!(Test-Path $unzippedRoot)) {
        New-Item -ItemType Directory -Path $unzippedRoot | Out-Null
    }

    $zipFiles = Get-ChildItem -Path $userDownloads -Filter *.zip -ErrorAction SilentlyContinue
    if ($zipFiles.Count -eq 0) {
        Write-Host "No ZIP files found in Downloads. Skipping ZIP processing."
        return
    }

    foreach ($zip in $zipFiles) {
        $zipName = [System.IO.Path]::GetFileNameWithoutExtension($zip.Name)
        $tempExtractFolder = Join-Path $unzippedRoot $zipName

        try {
            Expand-Archive -Path $zip.FullName -DestinationPath $tempExtractFolder -Force
            Remove-Item $zip.FullName -Force
        } catch {
            Write-Warning "Failed to unzip $($zip.Name). Skipping."
            continue
        }

        Write-Host "Installing drivers from: $tempExtractFolder"
        InstallDriversFromFolder $tempExtractFolder

        # Ensure USB motherboard driver folder exists
        if (!(Test-Path $usbDriverPath)) {
            New-Item -ItemType Directory -Path $usbDriverPath -Force | Out-Null
            Write-Host "Created USB motherboard folder: $usbDriverPath"
        }

        # Merge unzipped content into USB motherboard folder
        Write-Host "Merging $tempExtractFolder into USB folder: $usbDriverPath"
        Copy-Item "$tempExtractFolder\*" -Destination $usbDriverPath -Recurse -Force

        # Clean up
        Remove-Item $tempExtractFolder -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function InstallGPUDrivers {
    $videoControllers = Get-WmiObject Win32_VideoController | Where-Object { $_.Name -ne $null }
    if ($videoControllers.Count -eq 0) {
        Write-Host "No video controllers detected."
        return
    }

    foreach ($gpu in $videoControllers) {
        $gpuName = $gpu.Name
        Write-Host "`nDetected GPU: $gpuName"

        if ($gpuName -match "NVIDIA") {
            $gpuFolder = "NVIDIA"
        } elseif ($gpuName -match "AMD" -or $gpuName -match "Radeon") {
            $gpuFolder = "AMD"
        } elseif ($gpuName -match "Intel") {
            $gpuFolder = "Intel"
        } else {
            Write-Warning "Unknown GPU vendor. Skipping driver install for: $gpuName"
            continue
        }

        $downloadsPath = Join-Path ([Environment]::GetFolderPath("UserProfile") + "\Downloads\GPUDrivers") $gpuFolder
        $usbPath = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) ("GPUDrivers\" + $gpuFolder)

        if (Test-Path $downloadsPath) {
            Write-Host "Installing $gpuFolder GPU drivers from Downloads..."
            InstallDriversFromFolder $downloadsPath

            if (!(Test-Path $usbPath)) {
                Write-Host "Copying GPU drivers to USB: $usbPath"
                Copy-Item $downloadsPath -Destination $usbPath -Recurse -Force
            } else {
                Write-Host "$gpuFolder drivers already exist on USB. Skipping copy."
            }
        } elseif (Test-Path $usbPath) {
            Write-Host "Installing $gpuFolder GPU drivers from USB..."
            InstallDriversFromFolder $usbPath
        } else {
            Write-Warning "No $gpuFolder GPU drivers found in Downloads or USB."
        }
    }
}

# --- Main Logic ---

try {
    $mb = Get-WmiObject Win32_BaseBoard
    $manufacturer = $mb.Manufacturer.Trim()
    $product = $mb.Product.Trim()
} catch {
    Write-Warning "Could not detect motherboard info. Exiting."
    exit
}

$mbFolderName = CleanName("$manufacturer" + "_" + "$product")
Write-Host "`nDetected motherboard: $manufacturer $product"
Write-Host "Folder name will be: $mbFolderName"

$usbRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$usbDriverPath = Join-Path $usbRoot $mbFolderName
$downloadsUnzippedPath = Join-Path ([Environment]::GetFolderPath("UserProfile") + "\Downloads\Unzipped") $mbFolderName

# Step 1: Always try to unzip/install from ZIPs in Downloads
UnzipAndInstallAndCopyDrivers -mbFolderName $mbFolderName -usbDriverPath $usbDriverPath

# Step 2: Check if matching folder exists in Downloads\Unzipped or USB
if (Test-Path $downloadsUnzippedPath) {
    Write-Host "`nFound driver folder in Downloads\Unzipped. Installing drivers..."
    InstallDriversFromFolder $downloadsUnzippedPath
} elseif (Test-Path $usbDriverPath) {
    Write-Host "`nFound driver folder on USB. Installing drivers..."
    InstallDriversFromFolder $usbDriverPath
} else {
    Write-Warning "`nNo driver folder found for motherboard: $mbFolderName"
    Write-Host "Please add drivers manually or ensure ZIPs exist in Downloads."
}

# Step 3: Detect and install GPU drivers
Write-Host "`nChecking for GPU drivers..."
InstallGPUDrivers

Write-Host "`nAll done. Script complete."



