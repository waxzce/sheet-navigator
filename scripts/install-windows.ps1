# Sheet Navigator - Windows install via Trusted Catalog
#
# Use this when "Upload My Add-in" is missing in Excel's My Add-ins dialog,
# which happens when your M365 tenant admin has disabled user-side sideload.
#
# Trusted Catalogs is a separate Office feature that often remains enabled.
# This script:
#   1. Creates a folder in your user profile
#   2. Downloads the manifest into it
#   3. Shares the folder as a personal SMB share (no admin required)
#   4. Registers that share as a Trusted Add-in Catalog in HKCU (no admin)
#
# After running, restart Excel completely, then go to
# Insert > My Add-ins > SHARED FOLDER tab > Sheet Navigator > Add.
#
# Run with:
#   irm https://sheet-navigator.waxzce.org/install-windows.ps1 | iex
#
# Or download and run:
#   powershell -ExecutionPolicy Bypass -File install-windows.ps1

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "Sheet Navigator - Windows install via Trusted Catalog" -ForegroundColor Cyan
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host ""

$manifestUrl = "https://sheet-navigator.waxzce.org/manifest.xml"
$folder = Join-Path $env:USERPROFILE "OfficeAddins\SheetNavigator"
$shareName = "SheetNavigator"
$officeVersion = "16.0"  # works for Office 2016, 2019, 2021, M365

# 1. Create folder and download manifest
Write-Host "[1/4] Creating folder $folder..."
New-Item -ItemType Directory -Path $folder -Force | Out-Null

Write-Host "[2/4] Downloading manifest..."
Invoke-WebRequest -Uri $manifestUrl -OutFile (Join-Path $folder "sheet-navigator.manifest.xml") -UseBasicParsing

# 2. Share the folder. New-SmbShare normally needs admin but some Windows builds
#    allow personal shares. If it fails, we fall back to the manual UI path.
Write-Host "[3/4] Sharing folder as \\$env:COMPUTERNAME\$shareName..."
$existingShare = Get-SmbShare -Name $shareName -ErrorAction SilentlyContinue
if (-not $existingShare) {
    try {
        New-SmbShare -Name $shareName -Path $folder -ReadAccess "$env:USERDOMAIN\$env:USERNAME" -ErrorAction Stop | Out-Null
        Write-Host "  -> share created."
    } catch {
        Write-Host ""
        Write-Host "Could not create SMB share automatically (admin rights needed)." -ForegroundColor Yellow
        Write-Host "Manual fallback (no admin needed):" -ForegroundColor Yellow
        Write-Host "  1. Open File Explorer at: $folder" -ForegroundColor Yellow
        Write-Host "  2. Right-click the folder > Properties > Sharing tab" -ForegroundColor Yellow
        Write-Host "  3. Click Share..." -ForegroundColor Yellow
        Write-Host "  4. Add your own user with Read permission, click Share" -ForegroundColor Yellow
        Write-Host "  5. Copy the path shown (something like \\PC-NAME\Users\...)" -ForegroundColor Yellow
        Write-Host "  6. Re-run this script (it will skip step 1-2 and only do the registry)" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Then re-run me with: irm https://sheet-navigator.waxzce.org/install-windows.ps1 | iex" -ForegroundColor Yellow
        exit 1
    }
} else {
    Write-Host "  -> share already exists, reusing."
}

# 3. Register as Trusted Catalog in HKCU (no admin required)
Write-Host "[4/4] Registering as Trusted Add-in Catalog..."
$uncPath = "\\$env:COMPUTERNAME\$shareName"
$catalogsRoot = "HKCU:\Software\Microsoft\Office\$officeVersion\WEF\TrustedCatalogs"

# Remove any existing entry pointing to the same path (avoid duplicates)
if (Test-Path $catalogsRoot) {
    Get-ChildItem $catalogsRoot | ForEach-Object {
        $entry = Get-ItemProperty -Path $_.PSPath -Name "Url" -ErrorAction SilentlyContinue
        if ($entry -and $entry.Url -eq $uncPath) {
            Remove-Item -Path $_.PSPath -Recurse -Force
        }
    }
}

$id = [guid]::NewGuid().ToString("D")
$reg = "$catalogsRoot\$id"
New-Item -Path $reg -Force | Out-Null
New-ItemProperty -Path $reg -Name "Id" -Value $id -PropertyType String -Force | Out-Null
New-ItemProperty -Path $reg -Name "Url" -Value $uncPath -PropertyType String -Force | Out-Null
New-ItemProperty -Path $reg -Name "Flags" -Value 1 -PropertyType DWord -Force | Out-Null
Write-Host "  -> catalog registered at $uncPath"

Write-Host ""
Write-Host "Done!" -ForegroundColor Green
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Cyan
Write-Host "  1. Close Excel completely (use Task Manager if it's hanging)"
Write-Host "  2. Reopen Excel"
Write-Host "  3. Open any workbook"
Write-Host "  4. Insert > My Add-ins > SHARED FOLDER tab (top of the dialog)"
Write-Host "  5. Click Sheet Navigator > Add"
Write-Host ""
Write-Host "If 'Shared Folder' tab doesn't appear, your tenant has also blocked"
Write-Host "Trusted Catalogs. In that case, only M365 admin centralized"
Write-Host "deployment will work - contact your IT admin."
Write-Host ""
