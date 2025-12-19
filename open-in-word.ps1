param(
    [string]$File
)

$scriptPath = $MyInvocation.MyCommand.Path
$openCommand = "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`" `"%1`""
$progId = "OpenInWord.File"
$wordExtensions = @(
    ".doc",
    ".docx",
    ".docm",
    ".dot",
    ".dotx",
    ".dotm",
    ".rtf"
)

New-Item -Path "HKCU:\Software\Classes\$progId\shell\open\command" -Force | Out-Null
Set-ItemProperty -Path "HKCU:\Software\Classes\$progId" -Name "(Default)" -Value "Open in Word" | Out-Null
Set-ItemProperty -Path "HKCU:\Software\Classes\$progId\shell\open\command" -Name "(Default)" -Value $openCommand | Out-Null

foreach ($extension in $wordExtensions) {
    New-Item -Path "HKCU:\Software\Classes\$extension" -Force | Out-Null
    Set-ItemProperty -Path "HKCU:\Software\Classes\$extension" -Name "(Default)" -Value $progId | Out-Null
}

Start-Process winword.exe
Start-Sleep -Milliseconds 2000

$word = [Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")
$word.Documents.Open($File) | Out-Null
