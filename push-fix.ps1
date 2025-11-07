param(
  [string]$Message = "chore: export + push (auto)"
)

# Ensure we're in the repo root (script location)
Set-Location -Path $PSScriptRoot

# Basic guard that this is a git repo
git -c core.autocrlf=false status --porcelain | Out-Null
if ($LASTEXITCODE -ne 0) { throw "Git not initialized here (no .git folder)." }

# Stage, commit (if needed), pull --rebase, push
git add -A
if (-not (git diff --cached --quiet)) {
    $stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    git commit -m "$Message — $stamp"
    git pull --rebase
    git push origin HEAD
    Write-Host "Pushed changes."
} else {
    Write-Host "No changes to commit."
}
