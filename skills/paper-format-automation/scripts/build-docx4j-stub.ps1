param()

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$javaDir = Join-Path $scriptDir "java"
$srcDir = Join-Path $javaDir "src"
$outDir = Join-Path $javaDir "out"
$jarPath = Join-Path $javaDir "formatter.jar"

if (-not (Test-Path $srcDir)) {
    throw "Source directory not found: $srcDir"
}

New-Item -ItemType Directory -Force -Path $outDir | Out-Null
Get-ChildItem $outDir -Recurse -File -ErrorAction SilentlyContinue | Remove-Item -Force

$source = Join-Path $srcDir "PaperFormatAutomationFormatter.java"
& javac -d $outDir $source
if ($LASTEXITCODE -ne 0) {
    throw "javac failed with code $LASTEXITCODE"
}

if (Test-Path $jarPath) {
    Remove-Item -Force $jarPath
}

& jar cfe $jarPath PaperFormatAutomationFormatter -C $outDir .
if ($LASTEXITCODE -ne 0) {
    throw "jar creation failed with code $LASTEXITCODE"
}

Write-Host "Built stub formatter jar at $jarPath"
