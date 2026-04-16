param(
    [Parameter(Mandatory = $true)] [string]$InputDocx,
    [Parameter(Mandatory = $true)] [string]$RulesJson,
    [Parameter(Mandatory = $true)] [string]$OutputDocx,
    [string]$JarPath = ""
)

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($JarPath)) {
    $JarPath = Join-Path $scriptDir "java/formatter.jar"
}

if (-not (Test-Path -LiteralPath $InputDocx)) {
    throw "InputDocx not found: $InputDocx"
}
if (-not (Test-Path -LiteralPath $RulesJson)) {
    throw "RulesJson not found: $RulesJson"
}
if (-not (Test-Path -LiteralPath $JarPath)) {
    Write-Warning "formatter.jar not found. Falling back to Python formatter."
    & python (Join-Path $scriptDir 'format_manuscript.py') --input $InputDocx --rules $RulesJson --output $OutputDocx
    if ($LASTEXITCODE -ne 0) {
        throw "Python formatter exited with code $LASTEXITCODE"
    }
    exit 0
}

$arguments = @(
    "-jar", $JarPath,
    "--input", (Resolve-Path $InputDocx),
    "--rules", (Resolve-Path $RulesJson),
    "--output", $OutputDocx
)

Write-Host "Running formatter launcher..."
& java @arguments
if ($LASTEXITCODE -ne 0) {
    throw "Formatter exited with code $LASTEXITCODE"
}
Write-Host "Formatted document written to $OutputDocx"
