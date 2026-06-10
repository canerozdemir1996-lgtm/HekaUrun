param(
    [switch]$Clean,
    [switch]$WithLocalSettings,
    [switch]$WithIndex,
    [switch]$WithoutLocalData
)

$ErrorActionPreference = "Stop"

if ($Clean) {
    Remove-Item -LiteralPath ".\build", ".\dist" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath ".\UrunYonetimMasasi_v3.spec" -Force -ErrorAction SilentlyContinue
}

python -m pip install -r requirements.txt

$pyInstallerArgs = @(
    "--noconfirm",
    "--clean",
    "--windowed",
    "--hidden-import",
    "pyodbc",
    "--name",
    "UrunYonetimMasasi_v3"
)
if (Test-Path -LiteralPath ".\assets\app_icon.ico") {
    $pyInstallerArgs += @("--icon", ".\assets\app_icon.ico")
}
$pyInstallerArgs += "modern_app.py"
python -m PyInstaller @pyInstallerArgs

$distRoot = ".\dist\UrunYonetimMasasi_v3"
$settingsDestination = ".\dist\UrunYonetimMasasi_v3\settings.json"
$indexDestination = ".\dist\UrunYonetimMasasi_v3\product_index.sqlite"
$renameLogDestination = ".\dist\UrunYonetimMasasi_v3\rename_log.jsonl"
$materialsRoot = Join-Path $distRoot "materials"

function Copy-BuildResource {
    param(
        [string]$SourcePath,
        [string]$DestinationPath,
        [switch]$Directory
    )

    if (-not (Test-Path -LiteralPath $SourcePath)) {
        return
    }

    if ($Directory) {
        Remove-Item -LiteralPath $DestinationPath -Recurse -Force -ErrorAction SilentlyContinue
        Copy-Item -LiteralPath $SourcePath -Destination $DestinationPath -Recurse -Force
        return
    }

    Copy-Item -LiteralPath $SourcePath -Destination $DestinationPath -Force
}

function Add-BundledMaterial {
    param(
        [string]$SourcePath,
        [string]$Bucket
    )

    if ([string]::IsNullOrWhiteSpace($SourcePath) -or -not (Test-Path -LiteralPath $SourcePath)) {
        return $null
    }

    $safeBucket = if ([string]::IsNullOrWhiteSpace($Bucket)) { "misc" } else { $Bucket }
    $targetDir = Join-Path $materialsRoot $safeBucket
    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
    $targetPath = Join-Path $targetDir (Split-Path -Path $SourcePath -Leaf)
    Copy-Item -LiteralPath $SourcePath -Destination $targetPath -Force
    return "materials/$safeBucket/" + (Split-Path -Path $SourcePath -Leaf)
}

function Resolve-BundleSourcePath {
    param(
        [string]$RawValue
    )

    if ([string]::IsNullOrWhiteSpace($RawValue)) {
        return $null
    }

    $trimmed = $RawValue.Trim()
    if ($trimmed -match '^(?i)https?://') {
        return $null
    }

    if ($trimmed.StartsWith("sqlite:///")) {
        $trimmed = $trimmed.Substring(10)
    }

    if ($trimmed.StartsWith("file:")) {
        return $null
    }

    try {
        return (Resolve-Path -LiteralPath $trimmed -ErrorAction Stop).Path
    } catch {
        return $null
    }
}

function Resolve-BundleFallbackPath {
    param(
        [string]$SettingName
    )
    return $null
}

if (-not $WithoutLocalData) {
    if (Test-Path -LiteralPath ".\settings.json") {
        Copy-BuildResource -SourcePath ".\settings.json" -DestinationPath $settingsDestination
    }
    if (Test-Path -LiteralPath ".\product_index.sqlite") {
        Copy-BuildResource -SourcePath ".\product_index.sqlite" -DestinationPath $indexDestination
    }
    if (Test-Path -LiteralPath ".\rename_log.jsonl") {
        Copy-BuildResource -SourcePath ".\rename_log.jsonl" -DestinationPath $renameLogDestination
    }
} else {
    Remove-Item -LiteralPath $settingsDestination -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $indexDestination -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $renameLogDestination -Force -ErrorAction SilentlyContinue
}

if ($WithLocalSettings -and (Test-Path -LiteralPath ".\settings.json")) {
    Copy-BuildResource -SourcePath ".\settings.json" -DestinationPath $settingsDestination
}
if ($WithIndex -and (Test-Path -LiteralPath ".\product_index.sqlite")) {
    Copy-BuildResource -SourcePath ".\product_index.sqlite" -DestinationPath $indexDestination
}

if (Test-Path -LiteralPath ".\assets") {
    $assetDestination = ".\dist\UrunYonetimMasasi_v3\assets"
    Copy-BuildResource -SourcePath ".\assets" -DestinationPath $assetDestination -Directory
}
# Shader files (source + compiled .qsb)
if (Test-Path -LiteralPath ".\shaders") {
    $shaderDest = ".\dist\UrunYonetimMasasi_v3\shaders"
    Copy-BuildResource -SourcePath ".\shaders" -DestinationPath $shaderDest -Directory
}

if ((-not $WithoutLocalData) -and (Test-Path -LiteralPath $settingsDestination)) {
    $settingsObject = Get-Content -LiteralPath $settingsDestination -Raw | ConvertFrom-Json
    $settingsObject.data_source = "sql"
    $settingsObject.excel_path = ""
    $settingsObject.sheet_name = ""
    $settingsObject.header_row = "1"

    $sqlConnection = [string]$settingsObject.sql_connection_string
    $lowered = $sqlConnection.ToLowerInvariant()
    $looksLikeSqlite = $lowered.StartsWith("file:") -or $lowered.StartsWith("sqlite:///") -or $lowered.EndsWith(".sqlite") -or $lowered.EndsWith(".sqlite3") -or $lowered.EndsWith(".db") -or $lowered.EndsWith(".db3")
    if ($looksLikeSqlite) {
        $sqliteSource = Resolve-BundleSourcePath -RawValue $sqlConnection
        if ($sqliteSource) {
            $bundledSqlite = Add-BundledMaterial -SourcePath $sqliteSource -Bucket "sqlite"
            if ($bundledSqlite) {
                $settingsObject.sql_connection_string = $bundledSqlite
            }
        }
    }

    $settingsObject | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $settingsDestination -Encoding UTF8
}

$runNote = @"
Programi buradan calistir:
UrunYonetimMasasi_v3.exe

Baska bilgisayara tasirken bu klasorun tamamini kopyala.
Bu build assets, shaders, QML transition ve mevcut yerel veri dosyalarini pakete dahil eder.
Varsayilan olarak settings.json, product_index.sqlite ve rename_log.jsonl varsa exe klasorune kopyalanir.
SQL sqlite dosyasi kullaniliyorsa materials klasorune kopyalanir ve settings.json paket ici goreli yola cevrilir.
Bu yerel verileri pakete koymadan build almak icin -WithoutLocalData kullan.

Not: PyInstaller'in build klasoru gecici calisma klasorudur. Oradaki exe calistirilmaz.
"@
$runNote | Set-Content -LiteralPath ".\dist\UrunYonetimMasasi_v3\BURADAKI_EXE_CALISTIR.txt" -Encoding UTF8

Remove-Item -LiteralPath ".\build" -Recurse -Force -ErrorAction SilentlyContinue

$exePath = (Resolve-Path -LiteralPath ".\dist\UrunYonetimMasasi_v3\UrunYonetimMasasi_v3.exe").Path
Write-Host "EXE hazir: $exePath"
Write-Host "Onemli: build klasoru gecicidir; programi sadece dist\UrunYonetimMasasi_v3 klasorunden calistir."

