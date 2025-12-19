param(
    [string]$File
)

Start-Process winword.exe
Start-Sleep -Milliseconds 2000

$word = [Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")
$word.Documents.Open($File) | Out-Null
