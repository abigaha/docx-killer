param(
    [Parameter(Mandatory = $true)][string]$InputPath,
    [Parameter(Mandatory = $true)][string]$OutputDirectory
)

$ErrorActionPreference = 'Stop'
$source = (Resolve-Path -LiteralPath $InputPath).Path
$destination = [System.IO.Path]::GetFullPath($OutputDirectory)
[System.IO.Directory]::CreateDirectory($destination) | Out-Null
$pdf = Join-Path $destination ([System.IO.Path]::GetFileNameWithoutExtension($source) + '.pdf')
$application = $null
$document = $null
$cleanupWarning = $null
$exported = $false

try {
    $application = New-Object -ComObject Word.Application
    $application.Visible = $false
    $application.DisplayAlerts = 0
    $document = $application.Documents.Open($source, $false, $true)
    $document.ExportAsFixedFormat($pdf, 17)
    $exported = Test-Path -LiteralPath $pdf
    $renderer = $application.Name
    $version = $application.Version
} finally {
    if ($null -ne $document) {
        try { $document.Close($false) } catch { $cleanupWarning = $_.Exception.Message }
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($document) } catch {}
    }
    if ($null -ne $application) {
        try { $application.Quit() } catch {
            $message = $_.Exception.Message
            $cleanupWarning = if ($cleanupWarning) { "$cleanupWarning; $message" } else { $message }
        }
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($application) } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

if (-not $exported) { throw 'The Word-compatible application did not create a PDF' }
[ordered]@{
    status = 'rendered'
    renderer = "$renderer $version"
    pdf = $pdf
    cleanup_warning = $cleanupWarning
} | ConvertTo-Json -Compress
