param(
    [string]$Ids
)

$ErrorActionPreference = 'Continue'

$winget = Get-ChildItem `
    'C:\Program Files\WindowsApps' `
    -Filter 'winget.exe' `
    -Recurse `
    -ErrorAction SilentlyContinue |
    Select-Object -First 1 -ExpandProperty FullName

$resultados = @()

if (-not $winget) {
    [PSCustomObject]@{
        success = $false
        error = 'WINGET_NOT_AVAILABLE'
        resultados = @()
    } | ConvertTo-Json -Depth 5 -Compress

    exit 1
}

$idList = $Ids -split '\|'

foreach ($id in $idList) {

    $id = $id.Trim()

    if (-not $id) {
        continue
    }

    & $winget upgrade `
        --id $id `
        --exact `
        --silent `
        --accept-source-agreements `
        --accept-package-agreements `
        *> $null

    $exitCode = $LASTEXITCODE

    $resultados += [PSCustomObject]@{
        id = $id
        success = ($exitCode -eq 0)
        exitCode = $exitCode
    }
}

$globalSuccess = (
    $resultados.Count -gt 0 -and
    ($resultados | Where-Object { -not $_.success }).Count -eq 0
)

[PSCustomObject]@{
    success = $globalSuccess
    resultados = $resultados
} | ConvertTo-Json -Depth 5 -Compress
