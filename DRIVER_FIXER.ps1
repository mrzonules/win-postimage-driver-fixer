#Requires -RunAsAdministrator
<#
.SYNOPSIS
    One-shot post-imaging driver remediation. Tuned for OOBE audit mode and
    "plain retail OS, fix afterward" workflows. USB-friendly: pre-staged NIC
    drivers, local driver repo as offline/online fallback, cumulative CSV log.

    Phase 0 : Audit-mode prep, service startup, pre-staged NIC/WiFi injection
    Phase 1 : Microsoft Update -- driver category, multi-pass
    Phase 2 : Per-device HWID lookup against Microsoft Update Catalog
    Phase 3 : Surface-likely pack (signal-based, not just exact model)
    Phase 4 : Local driver repository (USB \_local -- pnputil match-and-install)
    Phase 5 : Report + reboot handling

.PARAMETER AutoReboot
    Reboot automatically if WU or an MSI install requires it.

.PARAMETER Passes
    WU scan/install passes (default 2).

.PARAMETER MaxCacheMB
    Skip caching downloaded drivers larger than this (default 150 MB).
    Big Surface/OEM MSI packs are easy to re-pull; no point hoarding them.

.PARAMETER SkipMU
    Skip the Microsoft Update phase (use when you only want catalog + local).

.PARAMETER SurfaceHint
    Force a Surface model search string when the chassis is misreporting or
    you suspect a near-twin (e.g. -SurfaceHint "Surface Book 3"). Tried first
    in Phase 3 ahead of the auto-detected candidates.

.NOTES
    Expected USB layout (script lives at root of the USB):
        DRIVER_FIXER.ps1
        PULL_DRIVERS.ps1
        _network\        Pre-staged NIC/WiFi INFs -- gets the box online in audit
                         Seed with: Intel I225/I226, Realtek RTL8xxx, Killer,
                         Intel WiFi 6/6E/7, Qualcomm/Atheros, Mediatek
        _local\          Curated driver repo -- your stash of vendor packs
                         (Intel chipset/ME, Realtek audio, Surface MSIs,
                          ASUS/ASRock/Gigabyte chipset, etc.). Recursive INF
                          tree. pnputil match-and-install handles selection.
        _catalog\        Auto-populated. Per-HWID cache of catalog .cab files
                         under MaxCacheMB. Survives across runs/machines.
        DriverLog.csv    Auto-populated. Cumulative log across all machines.

    ENCODING: this file is plain ASCII only. Do not paste em-dashes, smart
    quotes, or box-drawing characters into it -- PowerShell 5.1 default file
    encoding is the system code page (typically Windows-1252), and any UTF-8
    multibyte chars without a BOM will be misread and break the parser.
#>
param(
    [switch]$AutoReboot,
    [int]$Passes     = 2,
    [int]$MaxCacheMB = 150,
    [switch]$SkipMU,
    [string]$SurfaceHint = ''
)

# StrictMode 2.0: catches uninitialized vars + bad property access. Note that
# v2.0 still rejects '.Count' on a bare string -- and PowerShell UNROLLS
# function output, so "return @('one')" hands the caller a scalar string.
# Fix is at the call site: any function return that should stay an array must
# be wrapped with @() by the consumer. Done throughout this file.
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# --- Paths -------------------------------------------------------------------
$CacheRoot  = Join-Path $PSScriptRoot '_catalog'
$NetRoot    = Join-Path $PSScriptRoot '_network'
$LocalRoot  = Join-Path $PSScriptRoot '_local'
$CsvLog     = Join-Path $PSScriptRoot 'DriverLog.csv'
$SessionLog = Join-Path $env:TEMP ("DriverFix_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))

New-Item -ItemType Directory -Force -Path $CacheRoot | Out-Null
# _local also needs to exist up front: Phase 2 extracts catalog cabs straight
# into it (per-HWID subfolders) so the tree grows into a portable driver repo
# across runs and machines. Phase 4's blanket pnputil pass then ingests it.
New-Item -ItemType Directory -Force -Path $LocalRoot | Out-Null

# --- Helpers -----------------------------------------------------------------
function Log([string]$Msg, [string]$Level = 'INFO') {
    $line = "{0}  [{1}]  {2}" -f (Get-Date -Format 'HH:mm:ss'), $Level.PadRight(4), $Msg
    $line | Tee-Object -FilePath $SessionLog -Append | Out-Host
}

# Safe count for anything that may be $null / scalar / COM under StrictMode v3+
function Safe-Count($obj) {
    if ($null -eq $obj) { return 0 }
    try { return [int]$obj.Count } catch { }
    return @($obj).Count
}

function Get-ImageState {
    try {
        $k = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Setup\State'
        return (Get-ItemProperty $k -ErrorAction Stop).ImageState
    } catch { return 'UNKNOWN' }
}

function Get-ProblemDevices {
    @(Get-PnpDevice -PresentOnly -ErrorAction SilentlyContinue |
        Where-Object { $_.Status -ne 'OK' -or $_.Problem -ne 0 })
}

function Get-HWIDs([string]$InstanceId) {
    try {
        $p = Get-PnpDeviceProperty -InstanceId $InstanceId `
             -KeyName 'DEVPKEY_Device_HardwareIds' -ErrorAction Stop
        return @($p.Data)
    } catch { return @() }
}

function Get-CompatIDs([string]$InstanceId) {
    try {
        $p = Get-PnpDeviceProperty -InstanceId $InstanceId `
             -KeyName 'DEVPKEY_Device_CompatibleIds' -ErrorAction Stop
        return @($p.Data)
    } catch { return @() }
}

function Snapshot {
    foreach ($d in Get-ProblemDevices) {
        # @() guards against PowerShell unrolling a single-element return into
        # a bare string, which would later break .Count under StrictMode.
        $ids     = @(Get-HWIDs    $d.InstanceId)
        $compats = @(Get-CompatIDs $d.InstanceId)
        [pscustomobject]@{
            FriendlyName = $d.FriendlyName
            Class        = $d.Class
            Problem      = $d.Problem
            InstanceId   = $d.InstanceId
            HWIDs        = $ids
            CompatIDs    = $compats
            BestHWID     = if ($ids) { $ids[0] } else { '' }
        }
    }
}

function Test-Network {
    try { [void][System.Net.Dns]::GetHostEntry("download.microsoft.com"); return $true }
    catch { return $false }
}

function Get-ComputerModel {
    try {
        $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
        return ("{0} {1}" -f $cs.Manufacturer, $cs.Model).Trim()
    } catch { return 'UNKNOWN' }
}

function Write-CsvLog([string]$HWID, [string]$Title, [string]$Source) {
    [pscustomobject]@{
        Timestamp = Get-Date -Format 'o'
        Host      = $env:COMPUTERNAME
        Model     = (Get-ComputerModel)
        HWID      = $HWID
        Title     = $Title
        Source    = $Source
    } | Export-Csv -Path $CsvLog -Append -NoTypeInformation -Force -ErrorAction SilentlyContinue
}

# Heuristic "is this thing Surface-like?" -- does NOT trust the Model string alone.
# Returns list of signal strings; empty list = not Surface-ish.
function Get-SurfaceSignals {
    $signals = @()
    try {
        $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
        if ($cs.Manufacturer -match 'Microsoft') { $signals += "CS.Mfr=$($cs.Manufacturer)" }
        if ($cs.Model        -match 'Surface')   { $signals += "CS.Model=$($cs.Model)" }
        if ($cs.SystemFamily -match 'Surface')   { $signals += "CS.Family=$($cs.SystemFamily)" }
    } catch { }
    try {
        $bb = Get-CimInstance Win32_BaseBoard -ErrorAction Stop
        if ($bb.Manufacturer -match 'Microsoft') { $signals += "BB.Mfr=$($bb.Manufacturer)" }
        if ($bb.Product      -match 'Surface')   { $signals += "BB.Product=$($bb.Product)" }
    } catch { }
    try {
        $bios = Get-CimInstance Win32_BIOS -ErrorAction Stop
        if ($bios.Manufacturer -match 'Microsoft') { $signals += "BIOS.Mfr=$($bios.Manufacturer)" }
    } catch { }
    # ACPI\MSHW* HWIDs are Surface-specific embedded controllers / sensor hubs.
    # Strong signal even when the Model string is blank or wrong.
    try {
        $mshw = @(Get-PnpDevice -PresentOnly -ErrorAction SilentlyContinue |
                  Where-Object { $_.InstanceId -match 'ACPI\\MSHW' })
        if ($mshw.Count -gt 0) { $signals += "ACPI\MSHW count=$($mshw.Count)" }
    } catch { }
    # Plain return -- caller is expected to wrap with @() to enforce array
    # shape. Mixing the comma operator here with @() at the call site would
    # double-wrap and produce a nested array that breaks -join / .Count.
    return $signals
}

# Build a ranked list of catalog search queries to try for Surface driver packs.
# Most-specific to most-generic. pnputil match-and-install (used in
# Install-SurfaceMsi below) means non-matching INFs are silently skipped, so
# trying a slightly wrong pack is safe -- not destructive.
function Get-SurfaceCandidates([string]$Hint) {
    $list = New-Object System.Collections.Generic.List[string]
    if ($Hint) { [void]$list.Add("$Hint drivers firmware") }
    try {
        $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
        if ($cs.Model -and $cs.Model -match 'Surface') {
            [void]$list.Add("$($cs.Model) drivers firmware")
            # "Surface Book 3" -> also try "Surface Book" as near-twin fallback
            $base = ($cs.Model -replace '\s+\d+\s*$','').Trim()
            if ($base -and $base -ne $cs.Model) { [void]$list.Add("$base drivers firmware") }
        }
        if ($cs.SystemFamily -and $cs.SystemFamily -match 'Surface' -and
            $cs.SystemFamily -ne $cs.Model) {
            [void]$list.Add("$($cs.SystemFamily) drivers firmware")
        }
    } catch { }
    try {
        $bb = Get-CimInstance Win32_BaseBoard -ErrorAction Stop
        if ($bb.Product -and $bb.Product -match 'Surface') {
            [void]$list.Add("$($bb.Product) drivers firmware")
        }
    } catch { }
    # Last-ditch generic -- catches recent packs even when nothing self-IDs
    [void]$list.Add('Surface drivers and firmware')
    return ($list | Select-Object -Unique)
}

# Admin-install extract (no actions run, no firmware applied), then pnputil
# match-and-install. pnputil /install ONLY installs INFs whose HardwareIDs
# match a present device, so feeding it a near-twin pack is safe: matching
# drivers come in, everything else is silently skipped.
function Install-SurfaceMsi([string]$MsiPath, [string]$ExtractDir) {
    New-Item -ItemType Directory -Force -Path $ExtractDir | Out-Null
    $p = Start-Process msiexec `
         -ArgumentList "/a `"$MsiPath`" /qn TARGETDIR=`"$ExtractDir`"" `
         -Wait -PassThru
    if ($p.ExitCode -ne 0) {
        Log "    msiexec /a failed RC=$($p.ExitCode)" 'WARN'
        return $false
    }
    $infs = @(Get-ChildItem $ExtractDir -Recurse -Filter *.inf -ErrorAction SilentlyContinue)
    if ($infs.Count -eq 0) { Log "    No INFs after extract" 'WARN'; return $false }
    Log "    Extracted $($infs.Count) INFs -- pnputil match-and-install"
    & pnputil /add-driver "$ExtractDir\*.inf" /subdirs /install 2>&1 |
        ForEach-Object { Log "    pnputil: $_" }
    return $true
}

# --- Catalog scraping --------------------------------------------------------
function Search-Catalog([string]$Query) {
    $uri = "https://www.catalog.update.microsoft.com/Search.aspx?q=" +
           [uri]::EscapeDataString($Query)
    try {
        $html = (Invoke-WebRequest $uri -UseBasicParsing -TimeoutSec 20).Content
        return @(
            [regex]::Matches($html,
                '<tr[^>]+id="([a-f0-9]{8}-(?:[a-f0-9]{4}-){3}[a-f0-9]{12})_') |
            ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique -First 5
        )
    } catch { Log "  Catalog search failed ($Query): $_" 'WARN'; return @() }
}

function Get-CatalogUrls([string[]]$Guids) {
    if (-not $Guids -or $Guids.Count -eq 0) { return @() }
    $payload = ($Guids | ForEach-Object {
        '{"size":"0","languages":"","uidInfo":"' + $_ + '","updateID":"' + $_ + '"}'
    }) -join ','
    try {
        $resp = Invoke-WebRequest `
            -Uri     "https://www.catalog.update.microsoft.com/DownloadDialog.aspx" `
            -Method  POST `
            -Body    "updateIDs=[$payload]" `
            -ContentType "application/x-www-form-urlencoded" `
            -UseBasicParsing -TimeoutSec 20
        return @(
            [regex]::Matches($resp.Content,
                'https?://[^\s''"<>]+\.(?:cab|msi)') |
            ForEach-Object { $_.Value } | Select-Object -Unique
        )
    } catch { Log "  DownloadDialog failed: $_" 'WARN'; return @() }
}

function Install-FromCab([string]$CabPath, [string]$ExtractDir) {
    New-Item -ItemType Directory -Force -Path $ExtractDir | Out-Null
    & expand.exe $CabPath -F:* $ExtractDir 2>&1 | Out-Null
    if (@(Get-ChildItem $ExtractDir -Recurse -Filter *.inf -ErrorAction SilentlyContinue).Count -gt 0) {
        & pnputil /add-driver "$ExtractDir\*.inf" /subdirs /install 2>&1 |
            ForEach-Object { Log "    pnputil: $_" }
        return $true
    }
    return $false
}

function Install-FromMsi([string]$MsiPath) {
    $p = Start-Process msiexec -ArgumentList "/i `"$MsiPath`" /quiet /norestart" -Wait -PassThru
    return $p
}

# --- PHASE 0: Audit/OOBE prep + NIC bootstrap --------------------------------
$model      = Get-ComputerModel
$imageState = Get-ImageState

Log ("=== DRIVER FIX  |  Host: {0}  |  {1} ===" -f $env:COMPUTERNAME, (Get-Date -Format 'yyyy-MM-dd HH:mm'))
Log ("Model: {0}" -f $model)
Log ("ImageState: {0}" -f $imageState)
Log ("Session log: {0}" -f $SessionLog)
if ($imageState -match 'AUDIT|RESEAL') { Log "Detected audit/OOBE mode -- running pre-online services bootstrap" }

# Ensure services that WU and BITS need -- commonly stopped in audit mode
foreach ($svc in @('wuauserv','bits','cryptsvc','TrustedInstaller')) {
    try {
        $s = Get-Service $svc -ErrorAction Stop
        if ($s.StartType -eq 'Disabled') { Set-Service $svc -StartupType Manual; Log "Enabled: $svc" }
        if ($s.Status     -ne 'Running') { Start-Service $svc; Log "Started: $svc" }
    } catch { Log "Could not start ${svc}: $_" 'WARN' }
}

$before = @(Snapshot)
Log ("Problem devices before: {0}" -f $before.Count)
$before | ForEach-Object { Log ("  [{0,2}] {1}" -f $_.Problem, $_.FriendlyName) 'WARN' }

if (-not (Test-Network)) {
    Log "No network -- injecting pre-staged NIC/WiFi drivers from $NetRoot..."
    if (Test-Path $NetRoot) {
        & pnputil /add-driver "$NetRoot\*.inf" /subdirs /install 2>&1 |
            ForEach-Object { Log "  $_" }
        & pnputil /scan-devices 2>&1 | Out-Null
        Start-Sleep 8   # give NDIS a moment to bind
        if (Test-Network) { Log "Network restored after NIC injection" }
        else              { Log "Still offline after NIC injection -- online phases will skip" 'WARN' }
    } else {
        Log "_network folder not found at $NetRoot" 'WARN'
        Log "Pre-populate it with Intel/Realtek/Killer/Mediatek/Qualcomm INFs" 'WARN'
    }
} else {
    Log "Network available"
}

# --- PHASE 1: Microsoft Update -----------------------------------------------
$rebootRequired = $false

if ($SkipMU) {
    Log "=== PHASE 1: Microsoft Update -- SKIPPED (-SkipMU) ==="
} elseif (-not (Test-Network)) {
    Log "=== PHASE 1: Microsoft Update -- SKIPPED (offline) ==="
} else {
    Log "=== PHASE 1: Microsoft Update ==="
    try {
        $svcMgr = New-Object -ComObject Microsoft.Update.ServiceManager
        [void]$svcMgr.AddService2("7971f918-a847-4430-9279-4a52d1efe18d", 7, "")
        Log "Microsoft Update service registered"
    } catch { Log "MU registration (non-fatal): $_" 'WARN' }

    for ($pass = 1; $pass -le $Passes; $pass++) {
        Log "WU pass $pass / $Passes"
        try {
            $sess       = New-Object -ComObject Microsoft.Update.Session
            $found      = $sess.CreateUpdateSearcher().Search("IsInstalled=0 and Type='Driver'").Updates
            $foundCount = Safe-Count $found
            Log "  Found: $foundCount"
            if ($foundCount -gt 0) {
                $coll = New-Object -ComObject Microsoft.Update.UpdateColl
                for ($i = 0; $i -lt $foundCount; $i++) {
                    $u = $found.Item($i)
                    if (-not $u.EulaAccepted) { $u.AcceptEula() }
                    [void]$coll.Add($u)
                    Log ("  + {0}" -f $u.Title)
                    Write-CsvLog '' $u.Title 'WU'
                }
                $dl = $sess.CreateUpdateDownloader(); $dl.Updates = $coll
                [void]$dl.Download()
                $inst = $sess.CreateUpdateInstaller(); $inst.Updates = $coll
                $ir   = $inst.Install()
                Log ("  RC={0}  Reboot={1}" -f $ir.ResultCode, $ir.RebootRequired)
                if ($ir.RebootRequired) { $rebootRequired = $true }
            }
        } catch { Log "  WU pass $pass error: $_" 'WARN' }
    }
}

# --- PHASE 2: Per-device HWID catalog lookup ---------------------------------
Log "=== PHASE 2: Per-device HWID catalog lookup ==="

if (Test-Network) {
    foreach ($dev in @(Snapshot)) {
        if (-not $dev.HWIDs -or @($dev.HWIDs).Count -eq 0) { continue }
        Log ("Device: [{0,2}] {1}" -f $dev.Problem, $dev.FriendlyName)

        $resolved = $false

        # Walk all HWIDs most-to-least specific, then compat IDs -- critical
        # for Surface devices and OEM-rebranded parts that do not self-ID.
        $searchIds = @($dev.HWIDs) + @($dev.CompatIDs) | Select-Object -Unique

        foreach ($hwid in $searchIds) {
            Log "  ID: $hwid"
            $safeName  = $hwid -replace '[\\/:*?"<>|]','_'
            $cacheDir  = Join-Path $CacheRoot $safeName

            # Extracted cabs land here -- per-HWID subfolder under _local so
            # the tree builds into a portable, pnputil-ingestible driver repo.
            $localPkg = Join-Path $LocalRoot $safeName

            # USB cache check first -- survives across machines
            $cached = Get-ChildItem $cacheDir -Filter *.cab -ErrorAction SilentlyContinue |
                      Select-Object -First 1
            if ($cached) {
                Log "  Cache hit: $($cached.Name)"
                if (Install-FromCab $cached.FullName $localPkg) {
                    Write-CsvLog $hwid $dev.FriendlyName "Cache:$($cached.Name)"
                    & pnputil /scan-devices 2>&1 | Out-Null
                    $resolved = $true; break
                }
            }

            # Try the full HWID first. If MUC returns nothing AND the HWID has
            # a "{GUID}\Name" shape (Surface SMF/SAM/battery/telemetry etc.),
            # retry with just the suffix -- MUC indexes those by friendly name,
            # not by the device-interface-class GUID. Empirically: full search
            # returns 0 hits for {C65C8174-...}\SurfaceSmfThermalClient, but
            # "SurfaceSmfThermalClient" alone returns the right cab.
            $guids = @(Search-Catalog $hwid)
            if ($guids.Count -eq 0 -and
                $hwid -match '^\{[0-9a-fA-F-]+\}\\(.+)$') {
                $suffix = $Matches[1]
                Log "  Retry without GUID prefix: $suffix"
                $guids = @(Search-Catalog $suffix)
            }
            if ($guids.Count -eq 0) { continue }

            $urls = @(Get-CatalogUrls $guids)
            if ($urls.Count -eq 0) { continue }

            foreach ($url in $urls) {
                $fname = [IO.Path]::GetFileName(($url -split '\?')[0])
                $tmp   = Join-Path $env:TEMP $fname
                Log "  Downloading: $fname"

                try { Invoke-WebRequest $url -OutFile $tmp -UseBasicParsing -TimeoutSec 90 }
                catch { Log "  Download failed: $_" 'WARN'; continue }

                $sizeMB = [math]::Round((Get-Item $tmp).Length / 1MB, 1)

                if ($sizeMB -le $MaxCacheMB) {
                    New-Item -ItemType Directory -Force -Path $cacheDir | Out-Null
                    Copy-Item $tmp (Join-Path $cacheDir $fname) -Force
                    Log "  Cached ($sizeMB MB)"
                } else {
                    Log "  Skipping cache -- $sizeMB MB > $MaxCacheMB MB threshold"
                }

                $ext = [IO.Path]::GetExtension($fname).ToLower()
                $ok  = $false

                switch ($ext) {
                    '.cab' {
                        # Extract straight into _local\<safe-hwid>\ so the INF
                        # tree persists and Phase 4 can re-ingest on future
                        # runs (and on other machines if the USB travels).
                        $ok = Install-FromCab $tmp $localPkg
                    }
                    '.msi' {
                        $p  = Install-FromMsi $tmp
                        $ok = $p.ExitCode -in @(0, 3010)
                        if ($p.ExitCode -eq 3010) { $rebootRequired = $true }
                    }
                }

                if ($ok) {
                    Log "  Installed: $fname"
                    Write-CsvLog $hwid $dev.FriendlyName "Catalog:$fname"
                    & pnputil /scan-devices 2>&1 | Out-Null
                    $resolved = $true; break
                }
            }
            if ($resolved) { break }
        }
        if (-not $resolved) {
            Log ("  Unresolved at HWID/Catalog: {0}" -f $dev.FriendlyName) 'WARN'
        }
    }
} else {
    Log "Skipping catalog lookup -- offline" 'WARN'
}

# --- PHASE 3: Surface-likely driver packs (heuristic, not just exact model) --
# This phase exists for the "looks like a Surface Book 3 but the Model string
# is wrong / blank / a near-twin SKU" case. We:
#   1. Decide Surface-ness from multiple signals (CS / BaseBoard / BIOS / MSHW
#      ACPI HWIDs), not from one Model string.
#   2. Build a ranked candidate list -- exact model, hint, family, stripped
#      generation ("Surface Book 3" -> "Surface Book"), and a generic fallback.
#   3. For each candidate, pull the top MSI from the catalog and ADMIN-EXTRACT
#      it (msiexec /a) -- this does NOT run install actions, so firmware update
#      payloads never fire. Then pnputil /install does HWID match-and-install
#      against present devices; non-matching INFs (including wrong-model
#      firmware INFs) are silently skipped. Safe to over-try.
Log "=== PHASE 3: Surface-likely driver packs ==="

$surfaceSignals = @(Get-SurfaceSignals)
if (-not (Test-Network)) {
    Log "Skipping -- offline"
} elseif ($surfaceSignals.Count -eq 0 -and -not $SurfaceHint) {
    Log "No Surface signals and no -SurfaceHint -- skipping"
} else {
    if ($surfaceSignals.Count -gt 0) {
        Log ("Surface signals: {0}" -f ($surfaceSignals -join '; '))
    } else {
        Log "No auto-detected signals, but -SurfaceHint provided -- proceeding"
    }
    $candidates = @(Get-SurfaceCandidates -Hint $SurfaceHint)
    Log ("Candidate queries ({0}):" -f $candidates.Count)
    $candidates | ForEach-Object { Log "  - $_" }

    $tried       = 0
    $maxAttempts = 3   # cap: top-specific, near-twin, generic
    foreach ($query in $candidates) {
        if ($tried -ge $maxAttempts) { break }
        Log "Trying: $query"
        $guids = @(Search-Catalog $query)
        if ($guids.Count -eq 0) { Log "  No catalog hits"; continue }

        # Catalog serves Surface packs as either .msi (admin-extract path)
        # or .cab (expand.exe path). Accept both and route by extension.
        # Prefer MSI when present -- whole-pack INF tree match-and-install.
        $allUrls = @(Get-CatalogUrls $guids)
        if ($allUrls.Count -eq 0) { Log "  No download URLs"; continue }

        $msiUrls = @($allUrls | Where-Object { $_ -match '\.msi(\?|$)' })
        $cabUrls = @($allUrls | Where-Object { $_ -match '\.cab(\?|$)' })
        $urls    = @($msiUrls) + @($cabUrls)
        if ($urls.Count -eq 0) { Log "  No installable URLs (.msi/.cab)"; continue }

        $url   = $urls[0]
        $fname = [IO.Path]::GetFileName(($url -split '\?')[0])
        $tmp   = Join-Path $env:TEMP $fname
        $extr  = Join-Path $env:TEMP ($fname + '_x')

        Log "  Downloading: $fname"
        try { Invoke-WebRequest $url -OutFile $tmp -UseBasicParsing -TimeoutSec 600 }
        catch { Log "  Download failed: $_" 'WARN'; continue }

        $ext = [IO.Path]::GetExtension($fname).ToLower()
        $ok  = $false
        switch ($ext) {
            '.msi' { $ok = Install-SurfaceMsi $tmp $extr }
            '.cab' { $ok = Install-FromCab    $tmp $extr }
        }
        if ($ok) {
            Write-CsvLog '' "Surface candidate '$query'" "SurfacePack:$fname"
            & pnputil /scan-devices 2>&1 | Out-Null
            $tried++
        }
    }
    if ($tried -eq 0) { Log "No Surface candidate produced an installable pack" 'WARN' }
}

# --- PHASE 4: Local repository fallback (USB \_local) ------------------------
Log "=== PHASE 4: Local driver repository ==="

if (Test-Path $LocalRoot) {
    $infCount = @(Get-ChildItem $LocalRoot -Recurse -Filter *.inf -ErrorAction SilentlyContinue).Count
    Log "Scanning $LocalRoot ($infCount INFs)"
    if ($infCount -gt 0) {
        # pnputil with /install only installs INFs that match present hardware,
        # so this is a safe blanket pass. /subdirs walks the tree.
        & pnputil /add-driver "$LocalRoot\*.inf" /subdirs /install 2>&1 |
            ForEach-Object { Log "  $_" }
        & pnputil /scan-devices 2>&1 | Out-Null
        Write-CsvLog '' 'Local repo blanket install' "Local:$LocalRoot"
    }
} else {
    Log "_local not present -- skipping. (Seed it with Intel/Realtek/OEM packs.)"
}

# --- PHASE 5: Report ---------------------------------------------------------
Log "=== PHASE 5: Report ==="

$after = @(Snapshot)
Log ("Before={0}  After={1}  Fixed={2}" -f $before.Count, $after.Count, ($before.Count - $after.Count))

if ($after.Count -gt 0) {
    Log "Remaining (manual follow-up):" 'WARN'
    foreach ($d in $after) {
        Log ("  [{0,2}] {1}" -f $d.Problem, $d.FriendlyName) 'WARN'
        if ($d.BestHWID) {
            $q = [uri]::EscapeDataString($d.BestHWID)
            Log "        HWID    : $($d.BestHWID)"
            Log "        Catalog : https://www.catalog.update.microsoft.com/Search.aspx?q=$q"
        }
    }
} else {
    Log "All problem devices resolved."
}

try { $rebootRequired = $rebootRequired -or (New-Object -ComObject Microsoft.Update.SystemInfo).RebootRequired }
catch { }

if ($rebootRequired) {
    if ($AutoReboot) { Log "Rebooting in 15s..."; Start-Sleep 15; Restart-Computer -Force }
    else             { Log "Reboot required -- re-run after reboot if devices remain" 'WARN' }
}

Log "Session log : $SessionLog"
if (Test-Path $CsvLog) { Log "Cumulative  : $CsvLog" }
