param(
    [Parameter(Mandatory = $true)][string]$InputPath,
    [Parameter(Mandatory = $true)][string]$OutputDirectory,
    [string]$ApplicationProgId = 'Word.Application'
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$source = (Resolve-Path -LiteralPath $InputPath).Path
$destination = [System.IO.Path]::GetFullPath($OutputDirectory)
[System.IO.Directory]::CreateDirectory($destination) | Out-Null
$pdf = Join-Path $destination ([System.IO.Path]::GetFileNameWithoutExtension($source) + '.pdf')
if (Test-Path -LiteralPath $pdf) { throw "Output already exists: $pdf" }
$workDirectory = Join-Path $destination ('.docx-killer-' + [Guid]::NewGuid().ToString('N'))
[System.IO.Directory]::CreateDirectory($workDirectory) | Out-Null
$stagedPdf = Join-Path $workDirectory ([System.IO.Path]::GetFileName($pdf))
$application = $null
$document = $null
$cleanupWarning = $null
$exported = $false

try {
    try {
        $application = New-Object -ComObject $ApplicationProgId
        $application.Visible = $false
        $application.DisplayAlerts = 0
        try { $application.AutomationSecurity = 3 } catch {}
        try { $application.Options.UpdateLinksAtOpen = $false } catch {}
        $document = $application.Documents.Open($source, $false, $true, $false)
        $document.ExportAsFixedFormat($stagedPdf, 17)
        $exported = Test-Path -LiteralPath $stagedPdf -PathType Leaf
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
    [System.IO.File]::Move($stagedPdf, $pdf)
} finally {
    if ([System.IO.Directory]::GetParent($workDirectory).FullName -eq [System.IO.DirectoryInfo]::new($destination).FullName) {
        try { Remove-Item -LiteralPath $workDirectory -Recurse -Force } catch {
            $message = $_.Exception.Message
            $cleanupWarning = if ($cleanupWarning) { "$cleanupWarning; $message" } else { $message }
        }
    }
}

[ordered]@{
    status = 'rendered'
    application_progid = $ApplicationProgId
    renderer = "$ApplicationProgId / $renderer $version"
    pdf = $pdf
    cleanup_warning = $cleanupWarning
} | ConvertTo-Json -Compress
