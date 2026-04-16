param(
    [Parameter(Mandatory = $true)] [string]$InputTemplate,
    [string]$OutputDocx = ""
)

if (-not (Test-Path -LiteralPath $InputTemplate)) {
    throw "Input template not found: $InputTemplate"
}

$resolvedInput = (Resolve-Path $InputTemplate).Path
$extension = [IO.Path]::GetExtension($resolvedInput).ToLowerInvariant()
if ([string]::IsNullOrWhiteSpace($OutputDocx)) {
    $OutputDocx = [IO.Path]::ChangeExtension($resolvedInput, '.docx')
}
$outputDir = Split-Path -Parent $OutputDocx
if (-not [string]::IsNullOrWhiteSpace($outputDir) -and -not (Test-Path -LiteralPath $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

if ($extension -eq '.docx') {
    Copy-Item -LiteralPath $resolvedInput -Destination $OutputDocx -Force
    Write-Host "Template already in .docx format: $OutputDocx"
    exit 0
}

if ($extension -ne '.doc') {
    throw "Only .doc and .docx template inputs are supported."
}

$word = $null
$doc = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $doc = $word.Documents.Open($resolvedInput)
    $wdFormatXMLDocument = 12
    $doc.SaveAs([ref]$OutputDocx, [ref]$wdFormatXMLDocument)
    Write-Host "Converted template to .docx: $OutputDocx"
} catch {
    throw "Failed to convert .doc template to .docx. Ensure Microsoft Word is available. Original error: $($_.Exception.Message)"
} finally {
    if ($doc -ne $null) {
        try { $doc.Close() } catch {}
    }
    if ($word -ne $null) {
        try { $word.Quit() } catch {}
    }
}
