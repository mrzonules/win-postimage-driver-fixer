#Requires -Version 5.1
<#
.SYNOPSIS
    Pull Surface driver INF packages from a share into a local staging folder,
    preserving directory structure. Designed to feed DRIVER_FIXER.ps1's _local\
    tree -- pnputil match-and-install consumes it.

.DESCRIPTION
    Walks the source tree looking for *.inf files. For each unique driver
    package (the folder that contains an INF), copies the ENTIRE folder and
    its descendants to the destination at the mirrored relative path. This
    captures companion files the INF references: .cat (signature), .sys
    (driver binary), .dll, firmware blobs (.bin/.hex/.ufp/.fwbin), etc.

    Deduplicates: if package B sits inside package A, only A is copied (B is
    pulled in by the recursive copy of A). Avoids double-copying nested INFs.

    Uses robocopy per package for long-path tolerance, retry on locked files,
    and resume-friendly behavior on flaky shares.

    NOTE: Only useful for EXTRACTED driver trees (INF + companions on disk).
    If the share has Surface DriverPack MSIs not yet extracted, run
    `msiexec /a Pack.msi /qn TARGETDIR=C:\Some\Path` first, then point this
    script at the extracted output. pnputil cannot consume raw MSIs.

.PARAMETER Source
    Source root, typically a UNC path to the share.
    E.g. \\server\share\<codename>\Drivers

.PARAMETER Destination
    Local destination, typically the _local\ folder on your USB.
    E.g. F:\_local

.PARAMETER DryRun
    Show what would be copied; don't actually copy.

.PARAMETER SkipExisting
    Skip packages whose destination folder already exists. Useful for
    resuming a partial pull without re-comparing every file.

.PARAMETER IncludeSize
    After each copy, measure the destination folder size and log it.
    Adds a little time per package on large trees -- off by default.

.EXAMPLE
    .\PULL_DRIVERS.ps1 -Source \\server\share\<codename> -Destination F:\_local

.EXAMPLE
    .\PULL_DRIVERS.ps1 -Source \\server\share\<codename> -Destination F:\_local -DryRun

.NOTES
    ENCODING: plain ASCII only. Same rule as DRIVER_FIXER.ps1 -- no em-dashes,
    smart quotes, or box-drawing characters. PowerShell 5.1 reads this file as
    Windows-1252 by default and will misinterpret UTF-8 bytes.
#>
param(
    [Parameter(Mandatory)] [string]$Source,
    [Parameter(Mandatory)] [string]$Destination,
    [switch]$DryRun,
    [switch]$SkipExisting,
    [switch]$IncludeSize
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $Source)) {
    throw "Source not found: $Source"
}
# IMPORTANT: use ProviderPath, not Path. On UNC sources Resolve-Path returns
# the provider-prefixed form "Microsoft.PowerShell.Core\FileSystem::\\host\..."
# via .Path -- which then mismatches against $_.Directory.FullName (bare UNC)
# in the relative-path Substring() call below. .ProviderPath strips the
# prefix and gives the bare filesystem path.
$Source      = (Resolve-Path -LiteralPath $Source).ProviderPath.TrimEnd('\')
$Destination = $Destination.TrimEnd('\')

New-Item -ItemType Directory -Force -Path $Destination | Out-Null

# Log lives at the destination root so it travels with the USB.
$LogFile = Join-Path $Destination ("PullDrivers_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))

function Log([string]$Msg) {
    $line = "{0}  {1}" -f (Get-Date -Format 'HH:mm:ss'), $Msg
    $line | Tee-Object -FilePath $LogFile -Append | Out-Host
}

Log "=== PULL_DRIVERS  |  $(Get-Date -Format 'yyyy-MM-dd HH:mm') ==="
Log "Source      : $Source"
Log "Destination : $Destination"
if ($DryRun)       { Log "Mode        : DRY RUN" }
if ($SkipExisting) { Log "Mode        : SkipExisting (resume)" }
Log ""
Log "Scanning for INF packages... (this can take a minute on a big share)"

# Find every INF. Long-path-tolerant on recent Win10/11 + PS 5.1.
# Using -ErrorAction SilentlyContinue because shares often have a few ACL-locked
# subdirs that throw and we don't want them to abort the whole scan.
$infs = @(Get-ChildItem -LiteralPath $Source -Recurse -Filter *.inf -File `
          -ErrorAction SilentlyContinue)
Log ("INFs found            : {0}" -f $infs.Count)
if ($infs.Count -eq 0) {
    Log "Nothing to copy. Is the source actually an extracted driver tree?"
    return
}

# Unique parent directories = candidate package roots.
$pkgDirs = @($infs | ForEach-Object { $_.Directory.FullName } |
             Select-Object -Unique)
Log ("Unique package dirs   : {0}" -f $pkgDirs.Count)

# Dedupe: drop deeper package roots that sit inside shallower ones. Sort by
# path length ascending; for each, skip if any already-kept dir is a prefix.
# Case-insensitive because Windows paths are.
$sorted = @($pkgDirs | Sort-Object { $_.Length })
$keep   = New-Object System.Collections.Generic.List[string]
foreach ($d in $sorted) {
    $nested = $false
    foreach ($k in $keep) {
        if ($d.StartsWith($k + '\', [StringComparison]::OrdinalIgnoreCase)) {
            $nested = $true; break
        }
    }
    if (-not $nested) { [void]$keep.Add($d) }
}
Log ("After nested dedupe   : {0}" -f $keep.Count)
Log ""

# Copy loop.
$copied     = 0
$skipped    = 0
$failed     = 0
$totalBytes = 0L
$i          = 0

foreach ($pkg in $keep) {
    $i++
    $rel  = $pkg.Substring($Source.Length).TrimStart('\')
    $dest = if ($rel) { Join-Path $Destination $rel } else { $Destination }

    Log ("[{0,4}/{1}] {2}" -f $i, $keep.Count, $rel)

    if ($SkipExisting -and (Test-Path -LiteralPath $dest)) {
        Log "          -> skip (exists)"
        $skipped++; continue
    }

    if ($DryRun) {
        Log "          -> would copy"
        $skipped++; continue
    }

    # robocopy flags:
    #   /E      recurse incl empty dirs
    #   /R:2    retry twice on transient failures (locked files, etc.)
    #   /W:2    2-sec wait between retries
    #   /NFL    no per-file listing
    #   /NDL    no per-directory listing
    #   /NJH    no job header
    #   /NJS    no job summary
    #   /NP     no progress bar (cleaner log)
    #   /XJ     skip junctions (avoid loops)
    & robocopy $pkg $dest /E /R:2 /W:2 /NFL /NDL /NJH /NJS /NP /XJ | Out-Null
    $rc = $LASTEXITCODE

    # robocopy exit codes: 0=no copy needed, 1=copy ok, 2=extra files in dest,
    # 3=copy + extra, 4=mismatched, 5..7=combinations. >=8 = real failure.
    if ($rc -ge 8) {
        Log ("          -> robocopy FAILED rc={0}" -f $rc)
        $failed++
    } else {
        if ($IncludeSize) {
            $sz = (Get-ChildItem -LiteralPath $dest -Recurse -File `
                   -ErrorAction SilentlyContinue |
                   Measure-Object -Property Length -Sum).Sum
            if ($sz) { $totalBytes += $sz; Log ("          -> ok ({0:N1} MB)" -f ($sz / 1MB)) }
            else     { Log "          -> ok" }
        } else {
            Log "          -> ok"
        }
        $copied++
    }
}

Log ""
Log "==================== SUMMARY ===================="
Log ("Packages copied  : {0}" -f $copied)
Log ("Packages skipped : {0}" -f $skipped)
Log ("Packages failed  : {0}" -f $failed)
if ($IncludeSize) {
    Log ("Total size       : {0:N1} MB" -f ($totalBytes / 1MB))
}
Log ("Log              : {0}" -f $LogFile)
