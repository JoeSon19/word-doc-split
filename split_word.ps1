<#
.SYNOPSIS
    Word 문서(.docx)를 지정된 페이지 수 단위로 분할합니다.
.EXAMPLE
    .\split_word.ps1 -InputFile "C:\docs\report.docx"
    .\split_word.ps1 -InputFile "report.docx" -PagesPerSplit 50
    .\split_word.ps1 -InputFile "report.docx" -OutputDir "C:\output"
#>

param(
    [Parameter(Mandatory=$true, Position=0)]
    [string]$InputFile,

    [Parameter(Mandatory=$false)]
    [int]$PagesPerSplit = 100,

    [Parameter(Mandatory=$false)]
    [string]$OutputDir
)

$ErrorActionPreference = "Stop"

# 입력 파일 절대 경로 변환
$InputFile = (Resolve-Path $InputFile).Path
if (-not (Test-Path $InputFile)) {
    Write-Error "파일을 찾을 수 없습니다: $InputFile"
    exit 1
}

$baseName = [System.IO.Path]::GetFileNameWithoutExtension($InputFile)
if (-not $OutputDir) {
    $OutputDir = [System.IO.Path]::GetDirectoryName($InputFile)
}
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

$word = $null
try {
    Write-Host "Word 애플리케이션을 시작합니다..."
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0  # wdAlertsNone

    Write-Host "문서를 열고 있습니다: $InputFile"
    $doc = $word.Documents.Open($InputFile, $false, $true)  # ReadOnly=$true

    # 레이아웃 계산을 위해 PrintView 전환 후 대기
    $word.ActiveWindow.View.Type = 3  # wdPrintView
    Start-Sleep -Seconds 2

    $totalPages = $doc.ComputeStatistics(2)  # wdStatisticPages
    $totalParts = [math]::Ceiling($totalPages / $PagesPerSplit)

    Write-Host "총 페이지 수: $totalPages"
    Write-Host "분할 단위: ${PagesPerSplit}페이지"
    Write-Host "생성될 파일 수: $totalParts"
    Write-Host ("-" * 40)

    for ($i = 0; $i -lt $totalParts; $i++) {
        $startPage = $i * $PagesPerSplit + 1
        $endPage = [math]::Min(($i + 1) * $PagesPerSplit, $totalPages)
        $partNum = "{0:D3}" -f ($i + 1)
        $outputPath = Join-Path $OutputDir "${baseName}_${partNum}.docx"

        # 시작 페이지로 이동
        $rngStart = $doc.GoTo(1, 1, $startPage)  # wdGoToPage, wdGoToAbsolute

        if ($endPage -lt $totalPages) {
            $rngEnd = $doc.GoTo(1, 1, ($endPage + 1))
            $rngStart.End = $rngEnd.Start - 1
        } else {
            $rngStart.End = $doc.Content.End
        }

        # 복사 후 새 문서에 붙여넣기
        $rngStart.Copy()

        $newDoc = $word.Documents.Add()
        $newDoc.Content.Delete()
        $newDoc.Content.Paste()

        # 페이지 설정 복사 (첫 번째 섹션)
        try {
            $src = $doc.Sections(1).PageSetup
            $dst = $newDoc.Sections(1).PageSetup
            $dst.TopMargin = $src.TopMargin
            $dst.BottomMargin = $src.BottomMargin
            $dst.LeftMargin = $src.LeftMargin
            $dst.RightMargin = $src.RightMargin
            $dst.PageWidth = $src.PageWidth
            $dst.PageHeight = $src.PageHeight
            $dst.Orientation = $src.Orientation
        } catch {
            # 페이지 설정 복사 실패 시 무시
        }

        $newDoc.SaveAs([string]$outputPath, [int]12)
        $newDoc.Close(0)

        $fileName = [System.IO.Path]::GetFileName($outputPath)
        Write-Host "파트 $($i + 1)/$totalParts 저장 완료: $fileName (p.${startPage}-${endPage})"
    }

    $doc.Close(0)
    Write-Host ("-" * 40)
    Write-Host "완료! ${totalParts}개 파일이 '$OutputDir'에 저장되었습니다."

} catch {
    Write-Error "오류 발생: $_"
    exit 1
} finally {
    if ($word) {
        try { $word.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
    [System.GC]::Collect()
}
