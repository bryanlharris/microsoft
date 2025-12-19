param()

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

$classesRoot = "HKCU:\Software\Classes"
$openInWordShellKeys = Get-ChildItem -Path $classesRoot -ErrorAction SilentlyContinue |
    Where-Object { Test-Path "$($_.PSPath)\shell\OpenInWord" }

foreach ($key in $openInWordShellKeys) {
    $shellKeyPath = "$($key.PSPath)\shell\OpenInWord"
    Remove-Item -Path $shellKeyPath -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
}

$progIdPath = "$classesRoot\$progId"
if (Test-Path $progIdPath) {
    Remove-Item -Path $progIdPath -Recurse -Force | Out-Null
}

foreach ($extension in $wordExtensions) {
    $extensionPath = "$classesRoot\$extension"
    if (Test-Path $extensionPath) {
        $defaultProgId = Get-ItemPropertyValue -Path $extensionPath -Name "(default)" -ErrorAction SilentlyContinue
        if ($defaultProgId -eq $progId) {
            Remove-ItemProperty -Path $extensionPath -Name "(default)" -ErrorAction SilentlyContinue | Out-Null
        }

        $extensionProgIdPath = "$extensionPath\$progId"
        if (Test-Path $extensionProgIdPath) {
            Remove-Item -Path $extensionProgIdPath -Recurse -Force | Out-Null
        }
    }
}
