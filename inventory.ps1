$ErrorActionPreference = 'SilentlyContinue'

# -------------------- Utility --------------------
function SizeGB([decimal]$bytes) { if ($bytes -le 0) { return 0 } [math]::Round($bytes/1GB,2) }

function Test-IsAdmin {
    try {
        $id = [Security.Principal.WindowsIdentity]::GetCurrent()
        $p  = New-Object Security.Principal.WindowsPrincipal($id)
        return $p.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
    } catch { return $false }
}

# -------------------- Windows Version --------------------
function Get-WindowsVersion {
    $k = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
    $p = Get-ItemProperty -Path $k

    $build = [int]$p.CurrentBuild
    $ubr   = if ($p.UBR -ne $null) { [int]$p.UBR } else { 0 }
    $isWin11 = ($build -ge 22000)
    $family  = if ($isWin11) { 'Windows 11' } else { 'Windows 10' }

    $feature = if ($p.DisplayVersion) { $p.DisplayVersion } elseif ($p.ReleaseId) { $p.ReleaseId } else { '' }

    # Ricostruisci ProductName coerente se necessario
    $productName = $p.ProductName
    if ($productName -notmatch [regex]::Escape($family)) {
        $ed = $p.EditionID
        $short = switch -Regex ($ed) {
            'Professional' { 'Pro' }
            'Enterprise'   { 'Enterprise' }
            'Education'    { 'Education' }
            'Home'         { 'Home' }
            default        { $ed }
        }
        $productName = "$family $short".Trim()
    }

    [pscustomobject]@{
        ProductName    = $productName
        EditionID      = $p.EditionID
        DisplayVersion = $feature
        CurrentBuild   = $p.CurrentBuild
        UBR            = $ubr
        BuildFull      = ("{0}.{1}" -f $p.CurrentBuild, $ubr)
        Family         = $family
        IsWindows11    = $isWin11
    }
}

# -------------------- Office (build/canale) --------------------
function Get-OfficeInfo {
    # Click-to-Run (Microsoft 365 Apps / Office C2R)
    $ctr = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
    if (Test-Path $ctr) {
        $c = Get-ItemProperty $ctr
        return [pscustomobject]@{
            ChannelOrSKU = $c.ProductReleaseIds
            Architecture = $c.Platform
            Version      = $c.ClientVersionToReport
            InstallType  = 'ClickToRun'
        }
    }
    # MSI/legacy via Uninstall
    $uninstRoots = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    )
    $matches = foreach ($root in $uninstRoots) {
        if (Test-Path $root) {
            Get-ChildItem $root | ForEach-Object {
                $it = Get-ItemProperty $_.PsPath
                if ($it.DisplayName -match 'Microsoft (365|Office)') { $it }
            }
        }
    }
    $m = $matches | Sort-Object -Property DisplayVersion -Descending | Select-Object -First 1
    if ($m) {
        $arch = if ($m.DisplayName -match '64') { 'x64' } else { 'x86' }
        return [pscustomobject]@{
            ChannelOrSKU = $m.DisplayName
            Architecture = $arch
            Version      = $m.DisplayVersion
            InstallType  = 'MSI/Uninstall'
        }
    }
    return [pscustomobject]@{
        ChannelOrSKU = ''
        Architecture = ''
        Version      = ''
        InstallType  = 'NotFound'
    }
}

# -------------------- Hardware --------------------
function Get-CPU {
    $c = Get-CimInstance Win32_Processor | Select-Object -First 1 Name,NumberOfCores,NumberOfLogicalProcessors,MaxClockSpeed
    [pscustomobject]@{
        Name    = $c.Name
        Cores   = $c.NumberOfCores
        Threads = $c.NumberOfLogicalProcessors
        MHz     = $c.MaxClockSpeed
    }
}
function Get-RAMGB { $t = (Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory; SizeGB $t }
function Get-PhysicalDisks {
    Get-CimInstance Win32_DiskDrive | Select-Object @{n='Model';e={$_.Model}}, @{n='SizeGB';e={ SizeGB $_.Size }}
}
function Get-LogicalDisks {
    Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" |
        Select-Object DeviceID, @{n='SizeGB';e={ SizeGB $_.Size }}, @{n='FreeGB';e={ SizeGB $_.FreeSpace }}
}

# -------------------- Profili / Account --------------------
function Get-WindowsProfiles {
    Get-CimInstance Win32_UserProfile |
        Where-Object { $_.LocalPath -like 'C:\Users\*' -and $_.LocalPath -notmatch 'Default|Public|All Users' } |
        Sort-Object LastUseTime -Descending |
        Select-Object @{n='User';e={ Split-Path $_.LocalPath -Leaf }}, SID, @{n='LastUseTime';e={$_.ConvertToDateTime($_.LastUseTime)}}, Loaded
}

# Office identities da HKU (copre esecuzioni elevate/utente diverso se il profilo è caricato)
function Get-OfficeIdentities {
    $profiles = Get-CimInstance Win32_UserProfile |
        Where-Object { $_.LocalPath -like 'C:\Users\*' -and $_.SID -match '^S-1-5-21-' -and $_.Loaded -ne $false }

    foreach ($p in $profiles) {
        $sid  = $p.SID
        $user = Split-Path $p.LocalPath -Leaf
        $root = "Registry::HKEY_USERS\$sid\Software\Microsoft\Office\16.0\Common\Identity\Identities"
        if (Test-Path $root) {
            Get-ChildItem $root | ForEach-Object {
                $v = Get-ItemProperty $_.PsPath
                [pscustomobject]@{
                    WindowsUser = $user
                    DisplayName = $v.FriendlyName
                    Email       = $v.EmailAddress
                    Provider    = $v.ProviderName
                }
            }
        }
    }
}

function Find-EmailsInPath($path) {
    if (-not (Test-Path $path)) { return @() }
    $rx = '[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[A-Za-z]{2,}'
    $emails = New-Object System.Collections.Generic.HashSet[string]
    Get-ChildItem -Path $path -Recurse -File -Include *.json,*.log,*.txt -ErrorAction SilentlyContinue |
        ForEach-Object {
            try {
                $content = Get-Content $_.FullName -Raw -ErrorAction Stop
                [regex]::Matches($content, $rx) | ForEach-Object { [void]$emails.Add($_.Value.ToLower()) }
            } catch {}
        }
    return $emails.ToArray() | Sort-Object
}

# Teams: classico + nuovo (MSIX) per profili caricati
function Get-TeamsAccounts {
    $profiles = Get-CimInstance Win32_UserProfile |
        Where-Object { $_.LocalPath -like 'C:\Users\*' -and $_.Loaded -ne $false }

    $targets = @()
    foreach ($p in $profiles) {
        $up = $p.LocalPath
        # Teams classico
        $targets += Join-Path $up 'AppData\Roaming\Microsoft\Teams'
        $targets += Join-Path $up 'AppData\Local\Microsoft\Teams'
        # Teams nuovo (MSIX)
        $targets += Join-Path $up 'AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Roaming\Microsoft\Teams'
    }

    $rx = '[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[A-Za-z]{2,}'
    $emails = New-Object System.Collections.Generic.HashSet[string]

    foreach ($t in ($targets | Select-Object -Unique)) {
        if (-not (Test-Path $t)) { continue }
        $candidati = @(
            (Join-Path $t 'settings.json'),
            (Join-Path $t 'storage.json')
        ) | Where-Object { Test-Path $_ }
        $others = Get-ChildItem -Path $t -Recurse -File -Include *.json,*.log,*.txt -ErrorAction SilentlyContinue
        $files = @($candidati + $others | Select-Object -Unique)

        foreach ($f in $files) {
            try {
                $content = Get-Content $f.FullName -Raw -ErrorAction Stop
                foreach ($m in [regex]::Matches($content, $rx)) {
                    [void]$emails.Add($m.Value.ToLower())
                }
            } catch {}
        }
    }
    $emails.ToArray() | Sort-Object -Unique
}

# WAM/AAD (account moderni)
function Get-WAMAccounts {
    $profiles = Get-CimInstance Win32_UserProfile |
        Where-Object { $_.LocalPath -like 'C:\Users\*' -and $_.SID -match '^S-1-5-21-' -and $_.Loaded -ne $false }

    $out = @()
    foreach ($p in $profiles) {
        $sid  = $p.SID
        $user = Split-Path $p.LocalPath -Leaf

        $root1 = "Registry::HKEY_USERS\$sid\Software\Microsoft\IdentityCRL\StoredIdentities"
        if (Test-Path $root1) {
            Get-ChildItem $root1 -ErrorAction SilentlyContinue | ForEach-Object {
                try {
                    $v = Get-ItemProperty $_.PsPath
                    $email = $_.PSChildName
                    if ($email -notmatch '@') {
                        if ($v.PSObject.Properties.Name -contains 'Account') { $email = $v.Account }
                        if ($v.PSObject.Properties.Name -contains 'DisplayName' -and -not $email) { $email = $v.DisplayName }
                    }
                    if ($email) {
                        $out += [pscustomobject]@{
                            WindowsUser = $user; Source='IdentityCRL'; EmailOrUPN=$email; DisplayName=$v.DisplayName; AccountType=$v.AccountType
                        }
                    }
                } catch {}
            }
        }

        $root2 = "Registry::HKEY_USERS\$sid\Software\Microsoft\Windows\CurrentVersion\AAD\Identity\Cache"
        if (Test-Path $root2) {
            Get-ChildItem $root2 -ErrorAction SilentlyContinue | ForEach-Object {
                try {
                    $v = Get-ItemProperty $_.PsPath
                    $upn = $v.UserName
                    if ($upn) {
                        $out += [pscustomobject]@{
                            WindowsUser = $user; Source='AADCache'; EmailOrUPN=$upn; DisplayName=$v.FriendlyName; AccountType='AAD'
                        }
                    }
                } catch {}
            }
        }
    }
    $out | Where-Object { $_.EmailOrUPN -and $_.EmailOrUPN -match '@' } |
        Sort-Object WindowsUser, EmailOrUPN -Unique
}

# -------------------- Software installato --------------------
function Get-InstalledPrograms {
    param(
        [bool]$IncludeSystemComponents = $false,
        [bool]$IncludeUpdates = $false
    )

    $roots = @(
        @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall';             Scope='Machine'; Arch='x64' },
        @{ Path = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'; Scope='Machine'; Arch='x86' },
        @{ Path = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall';             Scope='User';    Arch=''    }
    )

    $list = @()
    foreach ($r in $roots) {
        if (-not (Test-Path $r.Path)) { continue }
        Get-ChildItem $r.Path -ErrorAction SilentlyContinue | ForEach-Object {
            try {
                $it = Get-ItemProperty $_.PsPath -ErrorAction Stop
                if (-not $it.DisplayName) { return }

                if (-not $IncludeSystemComponents) {
                    if ($it.PSObject.Properties.Name -contains 'SystemComponent' -and $it.SystemComponent -eq 1) { return }
                }
                if (-not $IncludeUpdates) {
                    if ($it.PSObject.Properties.Name -contains 'ReleaseType' -and $it.ReleaseType -match 'Update|Hotfix') { return }
                    if ($it.DisplayName -match '^(Security|Cumulative|Feature)\s+Update|Hotfix|Update for Microsoft Windows') { return }
                }

                $idate = $null
                if ($it.PSObject.Properties.Name -contains 'InstallDate') {
                    $raw = [string]$it.InstallDate
                    if ($raw -match '^\d{8}$') { $idate = [datetime]::ParseExact($raw,'yyyyMMdd',$null) }
                }

                $list += [pscustomobject]@{
                    Name            = $it.DisplayName
                    Version         = $it.DisplayVersion
                    Publisher       = $it.Publisher
                    InstallDate     = $idate
                    Scope           = $r.Scope
                    Arch            = $r.Arch
                    UninstallString = $it.UninstallString
                    RegistryKey     = $_.PsChildName
                }
            } catch {}
        }
    }

    $list | Sort-Object Name, Version, Publisher, Scope, Arch -Unique |
           Sort-Object Name, Version
}

function Get-StoreApps {
    param([bool]$AllUsers = $false)
    $params = @{}
    if ($AllUsers) { $params['AllUsers'] = $true }
    try {
        $apps = Get-AppxPackage @params -ErrorAction SilentlyContinue | Where-Object {
            $_.IsFramework -eq $false -and $_.Name -notmatch 'Microsoft.VCLibs|NET.Native|StorePurchaseApp|DesktopAppInstaller'
        } | Select-Object Name, PackageFullName, Version, Publisher, PublisherDisplayName
        $apps | Sort-Object Name, Version
    } catch { @() }
}

# -------------------- dsregcmd summary --------------------
function Get-DSRegSummary {
    $exe = "$env:SystemRoot\System32\dsregcmd.exe"
    if (Test-Path $exe) {
        $out = & $exe /status
        $keep = $out | Where-Object {
            $_ -match 'User|Tenant|Workplace|AzureAd|DomainJoin' -and $_ -notmatch 'Service'
        }
        return $keep
    }
}

# -------------------- Collect --------------------
$cpu       = Get-CPU
$ramGB     = Get-RAMGB
$pd        = Get-PhysicalDisks
$ld        = Get-LogicalDisks
$win       = Get-WindowsVersion
$office    = Get-OfficeInfo
$profiles  = Get-WindowsProfiles
$officeIds = Get-OfficeIdentities
$teams     = Get-TeamsAccounts
$wam       = Get-WAMAccounts
$dsreg     = Get-DSRegSummary

# Parametri inventario software
$IncludeSystemComponents = $false  # metti $true se vuoi anche componenti di sistema
$IncludeUpdates          = $false  # metti $true se vuoi includere aggiornamenti/Hotfix
$IncludeStoreApps        = $false  # metti $true per includere App MSIX/Store
$ExportCSV               = $false  # metti $true per creare CSV accanto al TXT

$programmi = Get-InstalledPrograms -IncludeSystemComponents:$IncludeSystemComponents -IncludeUpdates:$IncludeUpdates
$store = @()
if ($IncludeStoreApps) {
    $store = if (Test-IsAdmin) { Get-StoreApps -AllUsers:$true } else { Get-StoreApps -AllUsers:$false }
}

# -------------------- Render --------------------
$txt = New-Object System.Text.StringBuilder
$null = $txt.AppendLine("=== INVENTARIO PC ===")
$null = $txt.AppendLine("Data/Ora: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
$null = $txt.AppendLine()

# Hardware
$null = $txt.AppendLine("--- Hardware ---")
$null = $txt.AppendLine("CPU: $($cpu.Name) | Cores: $($cpu.Cores) | Threads: $($cpu.Threads) | MaxMHz: $($cpu.MHz)")
$null = $txt.AppendLine("RAM: $ramGB GB")
$null = $txt.AppendLine("Dischi fisici:")
foreach ($d in $pd) { $null = $txt.AppendLine("  - $($d.Model) : $($d.SizeGB) GB") }
$null = $txt.AppendLine("Volumi (logici):")
foreach ($v in $ld) { $null = $txt.AppendLine("  - $($v.DeviceID) : Size $($v.SizeGB) GB | Free $($v.FreeGB) GB") }
$null = $txt.AppendLine()

# Windows
$null = $txt.AppendLine("--- Windows ---")
$null = $txt.AppendLine("Edizione: $($win.ProductName) ($($win.EditionID))")
$null = $txt.AppendLine("Versione/Feature: $($win.DisplayVersion)")
$null = $txt.AppendLine("Build: $($win.BuildFull)")
$null = $txt.AppendLine()

# Office
$null = $txt.AppendLine("--- Office ---")
$null = $txt.AppendLine("InstallType: $($office.InstallType)")
$null = $txt.AppendLine("SKU/Channel: $($office.ChannelOrSKU)")
$null = $txt.AppendLine("Arch: $($office.Architecture)")
$null = $txt.AppendLine("Versione/Build: $($office.Version)")
$null = $txt.AppendLine()

# Profili Windows
$null = $txt.AppendLine("--- Profili Windows presenti ---")
foreach ($p in $profiles) { $null = $txt.AppendLine("  - $($p.User) | SID: $($p.SID) | LastUse: $($p.LastUseTime) | Loaded: $($p.Loaded)") }
$null = $txt.AppendLine()

# Account Office
$null = $txt.AppendLine("--- Account Office (HKU) ---")
if ($officeIds) { foreach ($i in $officeIds) { $null = $txt.AppendLine("  - $($i.DisplayName) <$($i.Email)> | $($i.Provider) | User:$($i.WindowsUser)") } }
else { $null = $txt.AppendLine("  (nessuno rilevato)") }
$null = $txt.AppendLine()

# Account Teams
$null = $txt.AppendLine("--- Account Teams (best-effort) ---")
if ($teams) { foreach ($t in $teams) { $null = $txt.AppendLine("  - $t") } }
else { $null = $txt.AppendLine("  (nessuno rilevato)") }
$null = $txt.AppendLine()

# Account moderni WAM/AAD
$null = $txt.AppendLine("--- Account moderni (WAM/AAD) ---")
if ($wam) {
    foreach ($w in $wam) { $null = $txt.AppendLine("  - $($w.EmailOrUPN) | User:$($w.WindowsUser) | Src:$($w.Source)") }
} else { $null = $txt.AppendLine("  (nessuno rilevato)") }
$null = $txt.AppendLine()

# Software installato
$null = $txt.AppendLine("--- Software installato (Uninstall) ---")
$null = $txt.AppendLine("Totale: $($programmi.Count)")
foreach ($a in $programmi) {
    $pub = if ($a.Publisher) { " | $($a.Publisher)" } else { '' }
    $arch= if ($a.Arch) { " | $($a.Arch)" } else { '' }
    $null = $txt.AppendLine(("  - {0} | {1}{2} | Scope:{3}{4}" -f $a.Name, $a.Version, $pub, $a.Scope, $arch))
}
$null = $txt.AppendLine()

# App MSIX/Store (opzionale)
if ($IncludeStoreApps) {
    $null = $txt.AppendLine("--- App MSIX/Store (Get-AppxPackage) ---")
    if ($store -and $store.Count -gt 0) {
        $null = $txt.AppendLine("Totale: $($store.Count)")
        foreach ($s in $store) {
            $pub = if ($s.PublisherDisplayName) { $s.PublisherDisplayName } else { $s.Publisher }
            $null = $txt.AppendLine(("  - {0} | {1} | {2}" -f $s.Name, $s.Version, $pub))
        }
    } else {
        $null = $txt.AppendLine("  (nessuna app rilevata o permessi mancanti)")
    }
    $null = $txt.AppendLine()
}

# Work/School / AAD (dsregcmd)
$null = $txt.AppendLine("--- Work/School / AAD (dsregcmd) ---")
if ($dsreg) { foreach ($l in $dsreg) { $null = $txt.AppendLine("  $l") } }
else { $null = $txt.AppendLine("  (dsregcmd non disponibile o nessun dato)") }

# -------------------- Output --------------------
$report = $txt.ToString()
$report | Out-Host

$base = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$nome = Read-Host "Inserisci nome per il file di log (es. MarioRossi)"
$nome = ($nome -replace '[^\w\-. ]','_')
if ([string]::IsNullOrWhiteSpace($nome)) { $nome = "report" }
$pathTxt = Join-Path $base ("{0}-log.txt" -f $nome)

$report | Out-File -FilePath $pathTxt -Encoding UTF8 -Force
Write-Host "Salvato: $pathTxt"

# Export CSV opzionale
if ($ExportCSV) {
    $csv1 = Join-Path $base ("{0}-software.csv" -f $nome)
    $programmi | Export-Csv -Path $csv1 -NoTypeInformation -Encoding UTF8
    Write-Host "Salvato: $csv1"
    if ($IncludeStoreApps -and $store -and $store.Count -gt 0) {
        $csv2 = Join-Path $base ("{0}-storeapps.csv" -f $nome)
        $store | Export-Csv -Path $csv2 -NoTypeInformation -Encoding UTF8
        Write-Host "Salvato: $csv2"
    }
}