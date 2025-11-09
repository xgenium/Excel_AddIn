Add-Type -AssemblyName System.Windows.Forms

$remoteUrl = "https://github.com/realkolobok/Excel_AddIn/raw/refs/heads/main/CustomAddIn_1.xlam"
$dest = "$env:APPDATA\Microsoft\AddIns\CustomAddIn_1.xlam"

if (Test-Path $dest) {
    $overwrite = [System.Windows.Forms.MessageBox]::Show(
        "The add-in file already exists at:`n$dest`nDo you want to overwrite it?",
        "Confirm Overwrite",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($overwrite -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-Host "Operation cancelled by user."
        exit
    }
}

Invoke-WebRequest $remoteUrl -OutFile $dest -UseBasicParsing
[System.Windows.Forms.MessageBox]::Show(
    "The add-in was installed successfully!",
    "Excel Add-In Installer",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information
)
