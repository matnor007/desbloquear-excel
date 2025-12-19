param (
    [Parameter(Mandatory)]
    [string]$Arquivo
)

Write-Host "Arquivo recebido:" $Arquivo

if (!(Test-Path $Arquivo)) {
    Write-Error "Arquivo nao encontrado."
    exit
}

$dir = Split-Path $Arquivo
$nome = [System.IO.Path]::GetFileNameWithoutExtension($Arquivo)
$ext  = [System.IO.Path]::GetExtension($Arquivo)

$zipOriginal = Join-Path $dir "$nome.zip"
$temp = Join-Path $dir "temp_excel_$nome"
$saidaZip = Join-Path $dir "${nome}_DESBLOQUEADO.zip"
$saidaFinal = Join-Path $dir "${nome}_DESBLOQUEADO$ext"

Rename-Item $Arquivo $zipOriginal -Force
Expand-Archive $zipOriginal $temp -Force

Get-ChildItem "$temp\xl\worksheets\*.xml" -ErrorAction SilentlyContinue |
ForEach-Object {
    (Get-Content $_) -replace '<sheetProtection[^>]*/>', '' |
    Set-Content $_
}

Compress-Archive "$temp\*" $saidaZip -Force
Rename-Item $saidaZip $saidaFinal -Force

Remove-Item $temp -Recurse -Force
Rename-Item $zipOriginal $Arquivo -Force

Write-Host "Arquivo desbloqueado criado em:"
Write-Host $saidaFinal
