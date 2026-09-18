
$scriptPath = 'C:\Power BI Backups\Final PS Script.txt'
$scriptContent = Get-Content -Path $scriptPath -Raw

$tokens = @()
$parseErrors = @()

$ast = [System.Management.Automation.Language.Parser]::ParseInput(
    $scriptContent,
    [ref]$tokens,
    [ref]$parseErrors
)

if ($parseErrors.Count -eq 0) {
    Write-Host "Parse OK - No syntax errors detected"
    exit 0
} else {
    Write-Host "Parse ERRORS found:"
    foreach ($parseError in $parseErrors) {
        Write-Host "Line $($parseError.Extent.StartLineNumber): $($parseError.Message)"
    }
    exit 1
}
