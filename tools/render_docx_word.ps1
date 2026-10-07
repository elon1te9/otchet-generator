param(
    [Parameter(Mandatory = $true)][string]$InputDocx,
    [Parameter(Mandatory = $true)][string]$OutputPdf,
    [string]$OutputDocx,
    [switch]$UpdateFields,
    [switch]$PreserveToc,
    [string]$BodyStartHeading
)

$ErrorActionPreference = 'Stop'
$inputPath = (Resolve-Path -LiteralPath $InputDocx).Path
if ($PreserveToc -and -not $BodyStartHeading) { throw '-PreserveToc requires -BodyStartHeading from the current report plan.' }
$outputPath = [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $OutputPdf))
$outputDir = Split-Path -Parent $outputPath
New-Item -ItemType Directory -Force -Path $outputDir | Out-Null
$openPath = $inputPath
$readOnly = $true
if ($OutputDocx) {
    $docxPath = [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $OutputDocx))
    if ($docxPath -eq $inputPath) { throw 'OutputDocx must differ from InputDocx; preserve the original.' }
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $docxPath) | Out-Null
    Copy-Item -LiteralPath $inputPath -Destination $docxPath -Force
    $openPath = $docxPath
    $readOnly = $false
}

$word = $null
$document = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $document = $word.Documents.Open($openPath, $false, $readOnly, $false, '', '', $false, '', '', 0, 0, $false, $false, $false, $false)
    if ($UpdateFields) {
        if ($PreserveToc) {
            # Cached TOC fields can be materialized as paragraphs. Do not rely
            # solely on TablesOfContents.Range to exclude nested PAGEREF fields.
            $bodyStart = $null
            foreach ($paragraph in $document.Paragraphs) {
                if ($paragraph.Range.Text.Trim() -eq $BodyStartHeading.Trim() -and $paragraph.OutlineLevel -le 2) {
                    $bodyStart = $paragraph.Range.Start
                    break
                }
            }
            if ($null -eq $bodyStart) { throw 'Cannot identify protected title/TOC boundary.' }
            foreach ($field in $document.Fields) {
                if ($field.Code.Start -ge $bodyStart -and $field.Type -ne 13) { $field.Update() | Out-Null }
            }
        }
        else {
            $document.Fields.Update() | Out-Null
            foreach ($toc in $document.TablesOfContents) {
                $toc.Update() | Out-Null
                # Newly generated TOC entries can inherit Word's theme formatting.
                $toc.Range.Font.Name = 'Times New Roman'
                $toc.Range.Font.Size = 14
                $toc.Range.Font.Bold = 0
                $toc.Range.Font.Italic = 0
                $toc.Range.Font.Underline = 0
                $toc.Range.Font.Color = 0
                $toc.Range.ParagraphFormat.SpaceBefore = 0
                $toc.Range.ParagraphFormat.SpaceAfter = 0
                $toc.Range.ParagraphFormat.LineSpacingRule = 1
                $toc.Range.ParagraphFormat.FirstLineIndent = 0
            }
        }
        foreach ($section in $document.Sections) {
            foreach ($header in $section.Headers) { $header.Range.Fields.Update() | Out-Null }
            foreach ($footer in $section.Footers) { $footer.Range.Fields.Update() | Out-Null }
        }
    }
    $document.Repaginate()
    if ($UpdateFields -and -not $PreserveToc) {
        foreach ($toc in $document.TablesOfContents) { $toc.UpdatePageNumbers() | Out-Null }
        $document.Repaginate()
    }
    if ($OutputDocx) { $document.Save() }
    # wdExportFormatPDF = 17; optimized for print, include document structure tags.
    $document.ExportAsFixedFormat($outputPath, 17, $false, 0, 0, 1, 9999, 0, $true, $true, 1, $true, $true, $false)
    Write-Output $outputPath
}
finally {
    if ($document) {
        $document.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($document) | Out-Null
    }
    if ($word) {
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
