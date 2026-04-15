$publicPath = Join-Path $PSScriptRoot 'Public'
Get-ChildItem -Path $publicPath -Filter '*.ps1' -File | ForEach-Object { . $_.FullName }

$publicFunctions = Get-ChildItem -Path $publicPath -Filter '*.ps1' -File |
Select-Object -ExpandProperty BaseName

Export-ModuleMember -Function $publicFunctions