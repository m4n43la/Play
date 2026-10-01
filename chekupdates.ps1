#requires -version 3.0

<#
.SYNOPSIS
    Inventaria el software instalado y detecta actualizaciones disponibles.

.DESCRIPTION
    Lee las entradas de desinstalacion del Registro y consulta WinGet y
    Chocolatey cuando estan disponibles. No instala, actualiza ni desinstala
    software.

    Compatible con Windows 7 SP1 o posterior y Windows PowerShell 3.0 o
    posterior. La consulta de actualizaciones depende de que el proveedor
    correspondiente este instalado y admita el sistema operativo.

.EXAMPLE
    .\chekupdates.ps1

.EXAMPLE
    .\chekupdates.ps1 -NoLog

.NOTES
    La salida de consola es un arreglo JSON de actualizaciones. De forma
    predeterminada, tambien se guardan el inventario, las actualizaciones y
    el registro en ProgramData.
#>

param(
    [switch]$NoLog
)

$ScriptVersion = '3.0.0'
$ErrorActionPreference = 'Continue'
$BaseDirectory = Join-Path $env:ProgramData 'SoftwareUpdate'
$LogFile = Join-Path $BaseDirectory 'SoftwareUpdate.log'
$InventoryFile = Join-Path $BaseDirectory 'SoftwareInventory.csv'
$UpdatesFile = Join-Path $BaseDirectory 'AvailableUpdates.csv'

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO'
    )

    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message

    if (-not $NoLog -and $script:OutputDirectoryAvailable) {
        try {
            Add-Content -LiteralPath $LogFile -Value $line -Encoding UTF8 -ErrorAction Stop
        }
        catch {
            Write-Verbose ('No se pudo escribir el registro: {0}' -f $_.Exception.Message)
        }
    }
}

function Write-Section {
    param([string]$Title)

    Write-Host ''
    Write-Host ('=' * 68)
    Write-Host ('  ' + $Title)
    Write-Host ('=' * 68)
}

function Get-SystemInfo {
    try {
        $os = Get-WmiObject -Class Win32_OperatingSystem -ErrorAction Stop
        return New-Object PSObject -Property @{
            ComputerName = $env:COMPUTERNAME
            Windows = $os.Caption
            Version = $os.Version
            Build = $os.BuildNumber
            Architecture = $os.OSArchitecture
            PowerShell = $PSVersionTable.PSVersion.ToString()
        }
    }
    catch {
        return New-Object PSObject -Property @{
            ComputerName = $env:COMPUTERNAME
            Windows = 'Desconocido'
            Version = 'Desconocida'
            Build = 'Desconocido'
            Architecture = 'Desconocida'
            PowerShell = $PSVersionTable.PSVersion.ToString()
        }
    }
}

function Get-UninstallRegistryEntries {
    $entries = New-Object System.Collections.ArrayList
    $hives = @(
        [Microsoft.Win32.RegistryHive]::LocalMachine,
        [Microsoft.Win32.RegistryHive]::CurrentUser
    )
    $views = @(
        [Microsoft.Win32.RegistryView]::Registry32,
        [Microsoft.Win32.RegistryView]::Registry64
    )
    $subKeyPath = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'

    foreach ($hive in $hives) {
        foreach ($view in $views) {
            $baseKey = $null
            $uninstallKey = $null
            try {
                $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey($hive, $view)
                $uninstallKey = $baseKey.OpenSubKey($subKeyPath)
                if ($null -eq $uninstallKey) {
                    continue
                }

                foreach ($subKeyName in $uninstallKey.GetSubKeyNames()) {
                    $applicationKey = $null
                    try {
                        $applicationKey = $uninstallKey.OpenSubKey($subKeyName)
                        if ($null -eq $applicationKey) {
                            continue
                        }

                        $displayName = [string]$applicationKey.GetValue('DisplayName', '')
                        if ([string]::IsNullOrWhiteSpace($displayName)) {
                            continue
                        }

                        $architecture = 'x86'
                        if ($view -eq [Microsoft.Win32.RegistryView]::Registry64) {
                            $architecture = 'x64'
                        }

                        [void]$entries.Add((New-Object PSObject -Property @{
                            Name = $displayName.Trim()
                            Version = [string]$applicationKey.GetValue('DisplayVersion', '')
                            Publisher = [string]$applicationKey.GetValue('Publisher', '')
                            InstallLocation = [string]$applicationKey.GetValue('InstallLocation', '')
                            SystemComponent = $applicationKey.GetValue('SystemComponent', 0)
                            ReleaseType = [string]$applicationKey.GetValue('ReleaseType', '')
                            ParentKeyName = [string]$applicationKey.GetValue('ParentKeyName', '')
                            RegistryHive = $hive.ToString()
                            Architecture = $architecture
                        }))
                    }
                    catch {
                        # Algunas claves de desinstalacion pueden no ser legibles.
                    }
                    finally {
                        if ($null -ne $applicationKey) {
                            $applicationKey.Dispose()
                        }
                    }
                }
            }
            catch {
                # La vista de 64 bits no existe en Windows de 32 bits.
            }
            finally {
                if ($null -ne $uninstallKey) {
                    $uninstallKey.Dispose()
                }
                if ($null -ne $baseKey) {
                    $baseKey.Dispose()
                }
            }
        }
    }

    return @($entries.ToArray())
}

function Get-SoftwareInventory {
    $applications = New-Object System.Collections.ArrayList
    $seen = @{}

    foreach ($entry in (Get-UninstallRegistryEntries)) {
        if ($entry.SystemComponent -eq 1 -or $entry.ReleaseType -or $entry.ParentKeyName) {
            continue
        }

        $name = ([string]$entry.Name).Trim()
        if ([string]::IsNullOrWhiteSpace($name)) {
            continue
        }

        $key = '{0}|{1}|{2}|{3}' -f $name, $entry.Version, $entry.Publisher, $entry.Architecture
        if ($seen.ContainsKey($key)) {
            continue
        }
        $seen[$key] = $true

        [void]$applications.Add((New-Object PSObject -Property @{
            Name = $name
            Version = ([string]$entry.Version).Trim()
            Publisher = ([string]$entry.Publisher).Trim()
            Architecture = $entry.Architecture
            RegistryHive = $entry.RegistryHive
            InstallLocation = ([string]$entry.InstallLocation).Trim()
        }))
    }

    return @($applications.ToArray() | Sort-Object Name, Version, Publisher)
}

function Export-CsvFile {
    param(
        [object[]]$Data,
        [string]$Path,
        [string[]]$Property
    )

    if (-not $script:OutputDirectoryAvailable) {
        return
    }

    try {
        if ($Data.Count -gt 0) {
            $Data | Select-Object -Property $Property | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
        }
        else {
            Set-Content -LiteralPath $Path -Value ($Property -join ',') -Encoding UTF8 -ErrorAction Stop
        }
        Write-Log ('Archivo guardado: {0}' -f $Path) 'SUCCESS'
    }
    catch {
        Write-Log ('No se pudo guardar {0}: {1}' -f $Path, $_.Exception.Message) 'WARN'
    }
}

function Get-ExecutablePath {
    param([string]$Name)

    if ($Name -ine 'winget') {
        $command = Get-Command -Name $Name -ErrorAction SilentlyContinue |
            Where-Object { $_.CommandType -eq 'Application' } |
            Select-Object -First 1
        if ($null -ne $command) {
            if ($command.Path) {
                return $command.Path
            }
            if ($command.Source) {
                return $command.Source
            }
        }
    }

    if ($Name -ieq 'choco') {
        $installRoots = New-Object System.Collections.ArrayList
        foreach ($root in @(
            $env:ChocolateyInstall,
            [Environment]::GetEnvironmentVariable('ChocolateyInstall', 'Machine'),
            (Join-Path $env:ProgramData 'chocolatey'),
            'C:\Chocolatey'
        )) {
            if (-not [string]::IsNullOrWhiteSpace([string]$root) -and
                -not $installRoots.Contains([string]$root)) {
                [void]$installRoots.Add([string]$root)
            }
        }

        try {
            $environmentKey = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey(
                'SYSTEM\CurrentControlSet\Control\Session Manager\Environment'
            )
            if ($null -ne $environmentKey) {
                $registryRoot = [string]$environmentKey.GetValue('ChocolateyInstall', '')
                $environmentKey.Dispose()
                if (-not [string]::IsNullOrWhiteSpace($registryRoot) -and
                    -not $installRoots.Contains($registryRoot)) {
                    [void]$installRoots.Add($registryRoot)
                }
            }
        }
        catch {}

        foreach ($root in $installRoots) {
            $candidate = Join-Path $root 'bin\choco.exe'
            if (Test-Path -LiteralPath $candidate -PathType Leaf) {
                return $candidate
            }
        }

        return $null
    }

    if ($Name -ieq 'winget') {
        $programFilesRoots = @(
            $env:ProgramW6432,
            $env:ProgramFiles,
            ${env:ProgramFiles(x86)}
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Select-Object -Unique

        foreach ($programFilesRoot in $programFilesRoots) {
            $windowsAppsPath = Join-Path $programFilesRoot 'WindowsApps'
            try {
                $packages = Get-ChildItem -LiteralPath $windowsAppsPath -Directory -Force `
                    -Filter 'Microsoft.DesktopAppInstaller_*' -ErrorAction Stop |
                    Sort-Object LastWriteTime -Descending
                foreach ($package in $packages) {
                    $candidate = Join-Path $package.FullName 'winget.exe'
                    if (Test-Path -LiteralPath $candidate -PathType Leaf) {
                        return $candidate
                    }
                }
            }
            catch {
                # WindowsApps puede no existir o no ser accesible en este contexto.
            }
        }

        $command = Get-Command -Name $Name -ErrorAction SilentlyContinue |
            Where-Object { $_.CommandType -eq 'Application' } |
            Select-Object -First 1
        if ($null -ne $command) {
            if ($command.Path) {
                return $command.Path
            }
            if ($command.Source) {
                return $command.Source
            }
        }

        $profilesRoot = Join-Path $env:SystemDrive 'Users'
        try {
            $profiles = Get-ChildItem -LiteralPath $profilesRoot -Directory -ErrorAction Stop
            foreach ($profile in $profiles) {
                $candidate = Join-Path $profile.FullName 'AppData\Local\Microsoft\WindowsApps\winget.exe'
                if (Test-Path -LiteralPath $candidate -PathType Leaf) {
                    return $candidate
                }
            }
        }
        catch {
            # El perfil de usuario puede no estar disponible para SYSTEM.
        }
    }

    return $null
}

function Get-WinGetUpdates {
    $path = Get-ExecutablePath 'winget'
    if (-not $path) {
        Write-Log 'WinGet no esta disponible.'
        return @()
    }

    Write-Log ('Consultando actualizaciones mediante WinGet desde: {0}' -f $path)
    try {
        $output = @(& $path list --upgrade-available --disable-interactivity 2>&1)
        $exitCode = $LASTEXITCODE
    }
    catch {
        Write-Log ('No se pudo ejecutar WinGet: {0}' -f $_.Exception.Message) 'WARN'
        return @()
    }

    if ($exitCode -ne 0) {
        Write-Log ('WinGet devolvio el codigo {0}.' -f $exitCode) 'WARN'
    }

    $updates = New-Object System.Collections.ArrayList
    foreach ($outputLine in $output) {
        $line = ([string]$outputLine).Trim()
        if (-not $line -or $line -match '^(Name\s+Id\s+Version|[-]{3,})' -or
            $line -match 'No installed package found|No applicable upgrade found') {
            continue
        }

        $match = [regex]::Match($line, '^(.*?)\s{2,}(\S+)\s{2,}(\S+)\s{2,}(\S+)(?:\s{2,}(\S+))?$')
        if (-not $match.Success) {
            continue
        }

        if ($match.Groups[2].Value -match '^(?i:Id)$') {
            continue
        }

        [void]$updates.Add((New-Object PSObject -Property @{
            Name = $match.Groups[1].Value.Trim()
            Id = $match.Groups[2].Value.Trim()
            InstalledVersion = $match.Groups[3].Value.Trim()
            AvailableVersion = $match.Groups[4].Value.Trim()
            Provider = 'WinGet'
        }))
    }

    return @($updates.ToArray())
}

function Get-ChocolateyUpdates {
    $path = Get-ExecutablePath 'choco'
    if (-not $path) {
        Write-Log 'Chocolatey no esta disponible.'
        return @()
    }

    Write-Log ('Consultando actualizaciones mediante Chocolatey desde: {0}' -f $path)
    try {
        $output = @(& $path outdated --limit-output 2>&1)
        $exitCode = $LASTEXITCODE
    }
    catch {
        Write-Log ('No se pudo ejecutar Chocolatey: {0}' -f $_.Exception.Message) 'WARN'
        return @()
    }

    if ($exitCode -ne 0 -and $exitCode -ne 2) {
        Write-Log ('Chocolatey devolvio el codigo {0}.' -f $exitCode) 'WARN'
    }

    $updates = New-Object System.Collections.ArrayList
    foreach ($outputLine in $output) {
        $parts = ([string]$outputLine).Trim().Split('|')
        if ($parts.Count -lt 3 -or [string]::IsNullOrWhiteSpace($parts[0]) -or
            [string]::IsNullOrWhiteSpace($parts[2])) {
            continue
        }

        [void]$updates.Add((New-Object PSObject -Property @{
            Name = $parts[0].Trim()
            Id = $parts[0].Trim()
            InstalledVersion = $parts[1].Trim()
            AvailableVersion = $parts[2].Trim()
            Provider = 'Chocolatey'
        }))
    }

    return @($updates.ToArray())
}

function Get-AllUpdates {
    $updates = New-Object System.Collections.ArrayList
    $seen = @{}

    foreach ($update in (Get-WinGetUpdates)) {
        $idKey = 'id:{0}' -f ([string]$update.Id).ToLowerInvariant()
        $nameKey = 'name:{0}' -f ([string]$update.Name).ToLowerInvariant()
        if (-not $seen.ContainsKey($idKey) -and -not $seen.ContainsKey($nameKey)) {
            $seen[$idKey] = $true
            $seen[$nameKey] = $true
            [void]$updates.Add($update)
        }
    }
    foreach ($update in (Get-ChocolateyUpdates)) {
        $idKey = 'id:{0}' -f ([string]$update.Id).ToLowerInvariant()
        $nameKey = 'name:{0}' -f ([string]$update.Name).ToLowerInvariant()
        if (-not $seen.ContainsKey($idKey) -and -not $seen.ContainsKey($nameKey)) {
            $seen[$idKey] = $true
            $seen[$nameKey] = $true
            [void]$updates.Add($update)
        }
    }

    $index = 1
    foreach ($update in $updates) {
        $update | Add-Member -MemberType NoteProperty -Name Index -Value $index -Force
        $index++
    }
    return @($updates.ToArray())
}

function Show-Updates {
    param([object[]]$Updates)

    Write-Section 'ACTUALIZACIONES DISPONIBLES'
    if ($Updates.Count -eq 0) {
        Write-Host 'No se encontraron actualizaciones mediante los proveedores disponibles.'
        return
    }

    $Updates | Format-Table Index, Name, Id, InstalledVersion, AvailableVersion, Provider -AutoSize
}

$script:OutputDirectoryAvailable = $false
try {
    if (-not (Test-Path -LiteralPath $BaseDirectory)) {
        New-Item -ItemType Directory -Path $BaseDirectory -Force -ErrorAction Stop | Out-Null
    }
    $script:OutputDirectoryAvailable = $true
}
catch {
    Write-Verbose ('No se pudo preparar el directorio de salida: {0}' -f $_.Exception.Message)
}

Write-Log 'Inicio del inventario y la comprobacion de actualizaciones.'

$inventory = @(Get-SoftwareInventory)
Write-Log ('Aplicaciones inventariadas: {0}' -f $inventory.Count)
Export-CsvFile -Data $inventory -Path $InventoryFile -Property @('Name', 'Version', 'Publisher', 'Architecture', 'RegistryHive', 'InstallLocation')

$updates = @(Get-AllUpdates)
Write-Log ('Actualizaciones detectadas: {0}' -f $updates.Count)
Export-CsvFile -Data $updates -Path $UpdatesFile -Property @('Index', 'Name', 'Id', 'InstalledVersion', 'AvailableVersion', 'Provider')

Write-Log 'Proceso terminado. No se ha instalado ni actualizado ningun software.'
ConvertTo-Json -InputObject @($updates | Select-Object Name, Id, InstalledVersion, AvailableVersion, Provider) -Depth 3
exit 0