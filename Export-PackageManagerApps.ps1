# Export Package Manager Apps to CSV
# Each CSV contains reinstallation-critical information at the beginning

[CmdletBinding()]
param(
    [Parameter(Position = -1)]
    [Alias("o", "output")]
    [string]$OutputDirectory = "./output",

    [Parameter(Position = -1)]
    [switch]$Help
)

# Show help
if ($Help) {
    Write-Host "Package Manager Apps Exporter" -ForegroundColor Cyan
    Write-Host "==============================" -ForegroundColor Cyan
    Write-Host "このスクリプトは各パッケージマネージャーからインストール済みアプリの情報をCSVで出力します。`n"

    Write-Host "使用方法:" -ForegroundColor Yellow
    Write-Host "  .\Export-PackageManagerApps.ps1                    # ./outputディレクトリに出力"
    Write-Host "  .\Export-PackageManagerApps.ps1 -o C:\MyExports    # 指定ディレクトリに出力"
    Write-Host "  .\Export-PackageManagerApps.ps1 --output .\data    # 指定ディレクトリに出力"
    Write-Host "  .\Export-PackageManagerApps.ps1 -Help              # このヘルプを表示`n"

    Write-Host "再インストールコマンドサンプル:" -ForegroundColor Yellow
    Write-Host "  Microsoft Store:" -ForegroundColor Green
    Write-Host "    # PackageFamilyNameを使用"
    Write-Host "    Add-AppxPackage -Register `"C:\Program Files\WindowsApps\[PackageFullName]\AppxManifest.xml`" -DisableDevelopmentMode"
    Write-Host "    # または Microsoft Store から手動インストール`n"

    Write-Host "  Winget:" -ForegroundColor Green
    Write-Host "    winget install [PackageId]"
    Write-Host "    winget install Microsoft.PowerToys"
    Write-Host "    winget install --id [PackageId] --source [Source]`n"

    Write-Host "  Scoop:" -ForegroundColor Green
    Write-Host "    scoop install [Name]"
    Write-Host "    scoop install git"
    Write-Host "    scoop install [Bucket]/[Name]  # bucketが異なる場合`n"

    Write-Host "  Chocolatey:" -ForegroundColor Green
    Write-Host "    choco install [PackageId]"
    Write-Host "    choco install [PackageId] --version [Version]  # 特定バージョン"
    Write-Host "    choco install googlechrome -y`n"

    return
}

# Create output directory if it doesn't exist
if (!(Test-Path $OutputDirectory)) {
    New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
}

Write-Host "Output directory: $OutputDirectory" -ForegroundColor Green

# Check for package manager installation
Write-Host "`n各パッケージマネージャーの確認中..." -ForegroundColor Cyan
$scoopAvailable = Get-Command scoop -ErrorAction SilentlyContinue
$chocoAvailable = Get-Command choco -ErrorAction SilentlyContinue
$wingetAvailable = Get-Command winget -ErrorAction SilentlyContinue

Write-Host "Winget: " -NoNewline -ForegroundColor White
if ($wingetAvailable) {
    Write-Host "インストール済み" -ForegroundColor Green
} else {
    Write-Host "未インストール" -ForegroundColor Red
}

Write-Host "Scoop: " -NoNewline -ForegroundColor White
if ($scoopAvailable) {
    Write-Host "インストール済み" -ForegroundColor Green
} else {
    Write-Host "未インストール" -ForegroundColor Red
}

Write-Host "Chocolatey: " -NoNewline -ForegroundColor White
if ($chocoAvailable) {
    Write-Host "インストール済み" -ForegroundColor Green
} else {
    Write-Host "未インストール" -ForegroundColor Red
}

# 1. Microsoft Store Apps
Write-Host "`nMicrosoft Storeアプリを処理中..." -ForegroundColor Cyan
try {
    $storeApps = Get-AppxPackage | Where-Object {
        $_.SignatureKind -eq "Store" -and !$_.IsFramework
    } | ForEach-Object {
        $manifest = try {
            (Get-AppxPackageManifest $_.PackageFullName).Package.Properties.DisplayName
        } catch {
            $_.Name
        }

        [PSCustomObject]@{
            # Reinstall critical info first
            PackageFamilyName = $_.PackageFamilyName  # Primary reinstall identifier
            PackageFullName = $_.PackageFullName      # Alternative identifier

            # Additional info
            Name = $_.Name
            DisplayName = $manifest
            Version = $_.Version
            Publisher = $_.Publisher
            Architecture = $_.Architecture
            InstallLocation = $_.InstallLocation
        }
    }

    if ($storeApps.Count -gt 0) {
        $storeApps | Export-Csv -Path "$OutputDirectory\MicrosoftStore_Apps.csv" -NoTypeInformation -Encoding UTF8
        Write-Host "  → $($storeApps.Count)個のアプリを検出してCSVに出力しました" -ForegroundColor Yellow
    } else {
        Write-Host "  → インストール済みアプリが見つかりませんでした" -ForegroundColor Yellow
    }
} catch {
    Write-Host "Error exporting Microsoft Store apps: $_" -ForegroundColor Red
}

# 2. Winget Apps
Write-Host "`nWingetアプリを処理中..." -ForegroundColor Cyan
if ($wingetAvailable) {
    try {
        $wingetApps = @()

        # Use winget list to get better version info
        $wingetOutput = winget list --disable-interactivity 2>$null
        $lines = $wingetOutput | Where-Object { $_ -and $_.Trim() -ne "" }

        $headerFound = $false
        $separatorFound = $false

        foreach ($line in $lines) {
            # Skip until we find the header with "名前" or "Name"
            if (!$headerFound -and ($line -match "名前.*ID.*バージョン" -or $line -match "Name.*Id.*Version")) {
                $headerFound = $true
                continue
            }

            # Skip the separator line with dashes
            if ($headerFound -and !$separatorFound -and $line -match "^-+") {
                $separatorFound = $true
                continue
            }

            # Process data lines after separator
            if ($separatorFound -and $line.Trim() -ne "") {
                # Try different parsing strategies
                $name = ""
                $packageId = ""
                $version = ""
                $source = "winget"

                # Strategy 1: Split by multiple spaces (2 or more)
                $parts = $line -split '\s{2,}' | Where-Object { $_.Trim() -ne "" }

                if ($parts.Count -ge 2) {
                    $name = $parts[0].Trim()
                    $packageId = $parts[1].Trim()

                    if ($parts.Count -ge 3) {
                        $version = $parts[2].Trim()
                    }

                    if ($parts.Count -ge 4) {
                        $source = $parts[3].Trim()
                    }

                    # Additional validation: packageId should contain dots or be clearly identifiable
                    if ($packageId -and ($packageId -match '\.' -or $packageId -match '^[A-Za-z0-9\-_]+$')) {
                        # Skip MS Store apps and system components
                        if ($source -ne "msstore" -and $packageId -notmatch "^Microsoft\.(UI\.|VCLibs|WindowsTerminal$|Winget)") {
                            $wingetApps += [PSCustomObject]@{
                                # Reinstall critical info first
                                PackageId = $packageId
                                Source = $source

                                # Additional info
                                Name = $name
                                Version = if ($version -eq "< 1.0" -or $version -eq "" -or $version -match "^\s*$") { "latest" } else { $version }
                            }
                        }
                    }
                }
            }
        }


        if ($wingetApps.Count -gt 0) {
            $wingetApps | Export-Csv -Path "$OutputDirectory\Winget_Apps.csv" -NoTypeInformation -Encoding UTF8
            Write-Host "  → $($wingetApps.Count)個のアプリを検出してCSVに出力しました" -ForegroundColor Yellow
        } else {
            Write-Host "  → インストール済みアプリが見つかりませんでした" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "Error exporting Winget apps: $_" -ForegroundColor Red
    }
} else {
    Write-Host "Wingetが未インストールのため、スキップします" -ForegroundColor Yellow
}

# 3. Scoop Apps
Write-Host "`nScoopアプリを処理中..." -ForegroundColor Cyan
if ($scoopAvailable) {
    try {
        # Get Scoop root path
        $scoopRoot = if ($env:SCOOP) { $env:SCOOP } else { "$env:USERPROFILE\scoop" }
        $appsPath = Join-Path $scoopRoot "apps"
        $scoopApps = @()

        if (Test-Path $appsPath) {
            $appFolders = Get-ChildItem -Path $appsPath -Directory -ErrorAction SilentlyContinue
            foreach ($appFolder in $appFolders) {
                $appName = $appFolder.Name

                # Try to get version from current version folder
                $manifestPath = Join-Path $appFolder.FullName "current\manifest.json"

                $version = ""
                $bucket = "main"  # Default bucket

                # Try to read manifest for version and bucket info
                if (Test-Path $manifestPath) {
                    try {
                        $manifest = Get-Content $manifestPath -Raw | ConvertFrom-Json
                        $version = if ($manifest.version) { $manifest.version } else { "" }
                    } catch {}
                }

                # If no manifest, try to get version from folder name
                if (!$version) {
                    $versionFolders = Get-ChildItem -Path $appFolder.FullName -Directory | Where-Object { $_.Name -ne "current" }
                    if ($versionFolders) {
                        # Get the most recent version folder
                        $version = ($versionFolders | Sort-Object Name -Descending)[0].Name
                    }
                }

                # Try to determine the actual bucket using scoop info
                try {
                    $infoOutput = scoop info $appName 2>$null | Out-String
                    if ($infoOutput -match "Bucket:\s*(.+)") {
                        $bucket = $matches[1].Trim()
                    } elseif ($infoOutput -match "Manifest:\s*.+\\buckets\\(.+?)\\") {
                        $bucket = $matches[1].Trim()
                    }
                } catch {
                    # If scoop info fails, try alternative method
                    try {
                        # Check if the app exists in known buckets by looking at scoop home
                        $scoopHome = if ($env:SCOOP_GLOBAL) { $env:SCOOP_GLOBAL } else { $scoopRoot }
                        $bucketsPath = Join-Path $scoopHome "buckets"

                        if (Test-Path $bucketsPath) {
                            $bucketFolders = Get-ChildItem -Path $bucketsPath -Directory -ErrorAction SilentlyContinue
                            foreach ($bucketFolder in $bucketFolders) {
                                $appManifestPath = Join-Path $bucketFolder.FullName "$appName.json"
                                if (Test-Path $appManifestPath) {
                                    $bucket = $bucketFolder.Name
                                    break
                                }
                            }
                        }
                    } catch {
                        # Keep default "main" if all methods fail
                    }
                }

                # Special handling for scoop itself
                if ($appName -eq "scoop") {
                    $bucket = "main"
                    if (!$version) {
                        try {
                            $scoopVersionOutput = scoop --version 2>$null
                            if ($scoopVersionOutput -match "v?(\d+\.\d+[\.\d]*)") {
                                $version = $matches[1]
                            }
                        } catch {}
                    }
                }

                $scoopApps += [PSCustomObject]@{
                    # Reinstall critical info first
                    Name = $appName
                    Source = $bucket

                    # Additional info
                    Version = $version
                }
            }
        }

        if ($scoopApps.Count -gt 0) {
            $scoopApps | Export-Csv -Path "$OutputDirectory\Scoop_Apps.csv" -NoTypeInformation -Encoding UTF8
            Write-Host "  → $($scoopApps.Count)個のアプリを検出してCSVに出力しました" -ForegroundColor Yellow
        } else {
            Write-Host "  → インストール済みアプリが見つかりませんでした" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "Error exporting Scoop apps: $_" -ForegroundColor Red
    }
} else {
    Write-Host "Scoopが未インストールのため、スキップします" -ForegroundColor Yellow
}

# 4. Chocolatey Apps
Write-Host "`nChocolateyアプリを処理中..." -ForegroundColor Cyan
if ($chocoAvailable) {
    try {
        $chocoApps = @()

        # Get local packages only
        $chocoList = choco list --local-only -r 2>$null
        foreach ($line in $chocoList) {
            if ($line -match '^([^|]+)\|(.+)$') {
                $packageId = $Matches[1].Trim()
                $version = $Matches[2].Trim()

                # Skip Chocolatey's own packages
                if ($packageId -notmatch '^chocolatey') {
                    # Try to get additional info using different methods
                    $title = ""
                    $summary = ""
                    $author = ""

                    try {
                        # Method 1: Try standard info command
                        $infoOutput = choco info $packageId --local-only 2>$null | Out-String

                        # Parse the non-machine readable output
                        if ($infoOutput -match "Title:\s*(.+)") {
                            $title = $matches[1].Trim()
                        }
                        if ($infoOutput -match "Summary:\s*(.+)") {
                            $summary = $matches[1].Trim()
                        }
                        if ($infoOutput -match "Author:\s*(.+)") {
                            $author = $matches[1].Trim()
                        }

                        # If that didn't work, try machine readable format
                        if ([string]::IsNullOrWhiteSpace($title) -and [string]::IsNullOrWhiteSpace($summary)) {
                            $infoMachineOutput = choco info $packageId --local-only -r 2>$null
                            foreach ($infoLine in $infoMachineOutput) {
                                if ($infoLine -match '^Title:\s*(.+)$') {
                                    $title = $matches[1].Trim()
                                } elseif ($infoLine -match '^Summary:\s*(.+)$') {
                                    $summary = $matches[1].Trim()
                                } elseif ($infoLine -match '^Author:\s*(.+)$') {
                                    $author = $matches[1].Trim()
                                }
                            }
                        }

                        # If still empty, try search for more info
                        if ([string]::IsNullOrWhiteSpace($title)) {
                            try {
                                $searchOutput = choco search $packageId --exact 2>$null | Out-String
                                if ($searchOutput -match "$packageId\s+(.+?)\s+\d") {
                                    $title = $matches[1].Trim()
                                }
                            } catch {
                                # Ignore search errors
                            }
                        }
                    } catch {
                        # If info command fails, just use package ID as title
                        $title = $packageId
                    }

                    # Use packageId as title if still empty
                    if ([string]::IsNullOrWhiteSpace($title)) {
                        $title = $packageId
                    }

                    $chocoApps += [PSCustomObject]@{
                        # Reinstall critical info first
                        PackageId = $packageId    # Primary reinstall identifier
                        Version = $version         # Specific version for exact reinstall

                        # Additional info
                        Title = $title
                        Summary = $summary
                        Author = $author
                    }
                }
            }
        }

        if ($chocoApps.Count -gt 0) {
            $chocoApps | Export-Csv -Path "$OutputDirectory\Chocolatey_Apps.csv" -NoTypeInformation -Encoding UTF8
            Write-Host "  → $($chocoApps.Count)個のパッケージを検出してCSVに出力しました" -ForegroundColor Yellow
        } else {
            Write-Host "  → インストール済みパッケージが見つかりませんでした" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "Error exporting Chocolatey apps: $_" -ForegroundColor Red
    }
} else {
    Write-Host "Chocolateyが未インストールのため、スキップします" -ForegroundColor Yellow
}

# Summary
Write-Host "`n処理完了" -ForegroundColor Green
Write-Host "出力されたファイル:" -ForegroundColor Cyan

$csvFiles = @("MicrosoftStore_Apps.csv", "Winget_Apps.csv", "Scoop_Apps.csv", "Chocolatey_Apps.csv")
$outputFiles = @()
foreach ($fileName in $csvFiles) {
    $filePath = Join-Path $OutputDirectory $fileName
    if (Test-Path $filePath) {
        $outputFiles += $fileName
        Write-Host "  $fileName" -ForegroundColor Yellow
    }
}

if ($outputFiles.Count -eq 0) {
    Write-Host "  出力されたファイルはありません" -ForegroundColor Yellow
}
