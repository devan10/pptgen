# install-gh.ps1
# Attempts to install GitHub CLI using winget, with simple fallbacks and verification.
param()

function Run-Command {
    param($cmd)
    Write-Host "Running: $cmd"
    $proc = Start-Process -FilePath pwsh -ArgumentList '-NoProfile','-Command',$cmd -Wait -Passthru -NoNewWindow -RedirectStandardOutput stdout.txt -RedirectStandardError stderr.txt
n    Get-Content stdout.txt -Raw | Write-Host
    Get-Content stderr.txt -Raw | Write-Host
}

Write-Host "Checking winget availability..."
if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
    Write-Error "winget not found. Please install App Installer or use the manual installer from https://github.com/cli/cli."
    exit 2
}

Write-Host "Attempting to install GitHub CLI via winget..."
$ids = @('GitHub.cli','GitHub.GitHubCLI','GitHubCLI.GitHubCLI')
$installed = $false
foreach ($id in $ids) {
    Write-Host "Trying package id: $id"
    try {
        & winget install --id $id -e --accept-package-agreements --accept-source-agreements
        $installed = $LASTEXITCODE -eq 0
    } catch {
        $installed = $false
    }
    if ($installed) { break }
}

if (-not $installed) {
    Write-Warning "Failed to install via winget. You can manually download installers from https://github.com/cli/cli/releases."
    exit 3
}

Write-Host "Verifying gh..."
try {
    gh --version
    Write-Host "gh installed successfully. Run 'gh auth login' to authenticate."
} catch {
    Write-Warning "gh installed but not found on PATH. Try restarting your terminal."
}
