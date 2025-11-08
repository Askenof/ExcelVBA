param(
  [string]$Workbook = "Sales - Fiscal Year.xlsm",   # your workbook name in this folder
  [string]$CommitMessage = "export: sync VBA",
  [switch]$SkipExcel                                # set to skip Excel COM (see Mode 2)
)

$ErrorActionPreference = "Stop"

# 0) Work in the repo root (the folder where this script lives)
Set-Location -Path $PSScriptRoot

# 1) (Optional) run the Excel macro ExportAllVba via COM
if (-not $SkipExcel) {
    # Launch Excel silently
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false

    try {
        $wbPath = Join-Path $PSScriptRoot $Workbook
        if (-not (Test-Path $wbPath)) { throw "Workbook not found: $wbPath" }

        $wb = $xl.Workbooks.Open($wbPath)
        # Run your macro that exports to Forms/Modules/Sheets
        $xl.Run("ExportAllVba")
        $wb.Save()
        $wb.Close($false)
    }
    finally {
        if ($xl -ne $null) { $xl.Quit() | Out-Null }
        [System.GC]::Collect() | Out-Null
        [System.GC]::WaitForPendingFinalizers()
    }
}

# 2) Stage, commit, push (respects your .gitignore)
$changes = git status --porcelain
if ([string]::IsNullOrWhiteSpace($changes)) {
    Write-Host "No changes to commit."
    exit 0
}

git add -A
$stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
git commit -m "$CommitMessage — $stamp"
git pull --rebase
git push origin main
Write-Host "Exported + pushed."
