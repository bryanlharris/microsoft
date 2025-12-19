param(
    [string]$File
)

$word = New-Object -ComObject Word.Application
$word.Visible = $true
$word.Documents.Add() | Out-Null
$word.Documents.Open($File) | Out-Null
