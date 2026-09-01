[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir,

    [Parameter(Mandatory = $true)]
    [string]$ManifestPath,

    [Parameter(Mandatory = $true)]
    [string]$PdfToPpmPath,

    [ValidateRange(72, 600)]
    [int]$Dpi = 180,

    [switch]$IncludeMarkup
)

$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSVersion.Major -lt 7) {
    throw 'PowerShell 7 or newer is required.'
}

function Get-Fingerprint([string]$Path) {
    $resolved = (Resolve-Path -LiteralPath $Path).Path
    $item = Get-Item -LiteralPath $resolved
    return [ordered]@{
        path = $resolved
        sha256 = (Get-FileHash -LiteralPath $resolved -Algorithm SHA256).Hash.ToLowerInvariant()
        size_bytes = [long]$item.Length
        modified_at_utc = $item.LastWriteTimeUtc.ToString('o')
    }
}

$inputResolved = (Resolve-Path -LiteralPath $InputPath).Path
$pdfToPpmResolved = (Resolve-Path -LiteralPath $PdfToPpmPath).Path
$scriptResolved = (Resolve-Path -LiteralPath $PSCommandPath).Path
if ([IO.Path]::GetExtension($inputResolved).ToLowerInvariant() -notin @('.docx', '.docm', '.dotx', '.dotm')) {
    throw "Unsupported Word package extension: $inputResolved"
}

$outputResolved = [IO.Path]::GetFullPath($OutputDir)
$manifestResolved = [IO.Path]::GetFullPath($ManifestPath)
if ([IO.Path]::GetExtension($manifestResolved).ToLowerInvariant() -ne '.json') {
    throw 'Renderer manifest must use a .json extension.'
}
if (Test-Path -LiteralPath $manifestResolved) {
    throw "Renderer manifest already exists: $manifestResolved"
}
if ($inputResolved -eq $manifestResolved) {
    throw 'Renderer manifest must be separate from the document.'
}
if ((Split-Path -Parent $manifestResolved) -ne $outputResolved) {
    throw 'Renderer manifest must be inside -OutputDir.'
}
if (Test-Path -LiteralPath $outputResolved) {
    throw "Output directory must be new: $outputResolved"
} else {
    New-Item -ItemType Directory -Path $outputResolved | Out-Null
}
$manifestParent = Split-Path -Parent $manifestResolved
if (-not (Test-Path -LiteralPath $manifestParent)) {
    New-Item -ItemType Directory -Path $manifestParent | Out-Null
}

$sourceHashBefore = (Get-FileHash -LiteralPath $inputResolved -Algorithm SHA256).Hash.ToLowerInvariant()
$pdfPath = Join-Path $outputResolved (([IO.Path]::GetFileNameWithoutExtension($inputResolved)) + '.word.pdf')
$word = $null
$document = $null
$wordVersion = $null
$wordExecutable = $null
$openSucceeded = $false
$exportSucceeded = $false
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    try { $word.AutomationSecurity = 3 } catch { }
    $wordVersion = [string]$word.Version
    $candidateWordExe = Join-Path ([string]$word.Path) 'WINWORD.EXE'
    if (Test-Path -LiteralPath $candidateWordExe -PathType Leaf) {
        $wordExecutable = Get-Fingerprint $candidateWordExe
    }
    $document = $word.Documents.OpenNoRepairDialog($inputResolved, $false, $true, $false)
    if (-not [bool]$document.ReadOnly) {
        throw 'Microsoft Word did not open the source as read-only.'
    }
    $openSucceeded = $true
    $document.PrintRevisions = [bool]$IncludeMarkup
    $document.ExportAsFixedFormat($pdfPath, 17)
    $exportSucceeded = $true
} finally {
    if ($null -ne $document) {
        try { $document.Close(0) } catch { }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
    }
    if ($null -ne $word) {
        try { $word.Quit() } catch { }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$sourceHashAfter = (Get-FileHash -LiteralPath $inputResolved -Algorithm SHA256).Hash.ToLowerInvariant()
if ($sourceHashBefore -ne $sourceHashAfter) {
    throw 'Document SHA-256 changed during read-only Word rendering.'
}
if (-not (Test-Path -LiteralPath $pdfPath -PathType Leaf)) {
    throw 'Microsoft Word did not produce the expected PDF.'
}
if (-not $openSucceeded -or -not $exportSucceeded) {
    throw 'Microsoft Word open/export evidence is incomplete.'
}

$pagePrefix = Join-Path $outputResolved 'page'
& $pdfToPpmResolved -png -r $Dpi $pdfPath $pagePrefix
if ($LASTEXITCODE -ne 0) {
    throw "pdftoppm failed with exit code $LASTEXITCODE"
}

$pageFiles = @(Get-ChildItem -LiteralPath $outputResolved -Filter 'page-*.png' -File | Sort-Object {
    if ($_.BaseName -match '^page-(\d+)$') { [int]$Matches[1] } else { [int]::MaxValue }
})
if ($pageFiles.Count -eq 0) {
    throw 'No page PNGs were produced.'
}
$pageRecords = @()
for ($index = 0; $index -lt $pageFiles.Count; $index++) {
    $expected = $index + 1
    if ($pageFiles[$index].BaseName -notmatch '^page-(\d+)$' -or [int]$Matches[1] -ne $expected) {
        throw "Rendered page sequence is not contiguous at expected page $expected"
    }
    if ($pageFiles[$index].Length -le 0) {
        throw "Rendered page is empty: $($pageFiles[$index].FullName)"
    }
    $pageRecords += [ordered]@{ page = $expected; image = Get-Fingerprint $pageFiles[$index].FullName }
}

$payload = [ordered]@{
    schema_version = '2.0'
    record_type = 'documents_renderer_manifest'
    producer_script = 'render_docx_word.ps1'
    created_at_utc = [DateTime]::UtcNow.ToString('o')
    document = Get-Fingerprint $inputResolved
    source_sha256_before = $sourceHashBefore
    source_sha256_after = $sourceHashAfter
    renderer = [ordered]@{
        id = 'microsoft-word'
        version = $wordVersion
        executable = $wordExecutable
    }
    settings = [ordered]@{
        dpi = $Dpi
        markup_mode = $(if ($IncludeMarkup) { 'include' } else { 'hide' })
        include_markup = [bool]$IncludeMarkup
    }
    application_open = [ordered]@{
        status = 'PASS'
        opened_read_only = $true
        open_mode = 'OpenNoRepairDialog-read-only'
        export_status = 'PASS'
        repair_warning_count = 0
        repair_observation = [ordered]@{
            status = 'NO_DIAGNOSTIC_OBSERVED'
            observed_warning_count = 0
            absence_proven = $false
            capture_scope = 'COM open/export exceptions and OpenNoRepairDialog outcome; Word alerts were suppressed'
        }
        diagnostics = @('read-only open completed without an observed application warning', 'PDF export completed')
    }
    toolchain = [ordered]@{
        render_script = Get-Fingerprint $scriptResolved
        pdftoppm = Get-Fingerprint $pdfToPpmResolved
    }
    output_directory = $outputResolved
    pdf = Get-Fingerprint $pdfPath
    pages = $pageRecords
}

$json = $payload | ConvertTo-Json -Depth 10
$encoding = [Text.UTF8Encoding]::new($false)
$stream = [IO.FileStream]::new($manifestResolved, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
try {
    $writer = [IO.StreamWriter]::new($stream, $encoding)
    try { $writer.WriteLine($json) } finally { $writer.Dispose() }
} finally {
    if ($stream) { $stream.Dispose() }
}

[pscustomobject]@{
    status = 'WORD_RENDER_COMPLETE'
    manifest = $manifestResolved
    document_sha256 = $sourceHashAfter
    word_version = $wordVersion
    include_markup = [bool]$IncludeMarkup
    rendered_pages = $pageFiles.Count
} | ConvertTo-Json -Depth 5
