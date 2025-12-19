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

$pathsToRemove = @(
    "HKCU:\Software\Classes\*\shell\OpenInWord",
    "HKCU:\Software\Classes\$progId"
)

foreach ($path in $pathsToRemove) {
    if (Test-Path $path) {
        Remove-Item -Path $path -Recurse -Force | Out-Null
    }
}

foreach ($extension in $wordExtensions) {
    $extensionPath = "HKCU:\Software\Classes\$extension"
    if (Test-Path $extensionPath) {
        Remove-Item -Path $extensionPath -Recurse -Force | Out-Null
    }
}
