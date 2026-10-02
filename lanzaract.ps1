param(
    [string]$Ids
)

$ErrorActionPreference = 'Continue'

# Buscar winget.exe
$winget = Get-ChildItem `
    'C:\Program Files\WindowsApps' `
    -Filter 'winget.exe' `
    -Recurse `
    -ErrorAction SilentlyContinue |
    Select-Object -First 1 -ExpandProperty FullName

if (-not $winget) {
    Write-Output 'WINGET_NOT_AVAILABLE'
    exit 1
}

# Separar IDs recibidos
$idList = $Ids -split '\|'

foreach ($id in $idList) {

    $id = $id.Trim()

    if (-not $id) {
        continue
    }

    Write-Output "Actualizando $id..."

    & $winget upgrade `
        --id $id `
        --exact `
        --silent `
        --accept-source-agreements `
        --accept-package-agreements

    Write-Output "Resultado $id : ExitCode=$LASTEXITCODE"
}
