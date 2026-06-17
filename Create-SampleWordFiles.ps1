[CmdletBinding()]
param(
    [string]$OutputDirectory = "."
)

$ErrorActionPreference = "Stop"

$outPath = (Resolve-Path -LiteralPath $OutputDirectory).Path
$samples = @(
    @{ Name = "Northwind Quarterly Operations Brief"; Ext = "docx"; Format = 16; Text = "Northwind Inc.`r`nQuarterly Operations Brief`r`nHighlights: SLA improvements, pilot expansion, and backlog reduction." },
    @{ Name = "Fabrikam Vendor Evaluation Notes"; Ext = "doc"; Format = 0; Text = "Fabrikam Corp.`r`nVendor Evaluation Notes`r`nSummary: Three vendors shortlisted based on security and support scorecards." },
    @{ Name = "Alpine Expansion Kickoff Agenda"; Ext = "docx"; Format = 16; Text = "Alpine Systems`r`nExpansion Kickoff Agenda`r`nAgenda: Objectives, risks, owners, and timeline." },
    @{ Name = "Contoso Internal Policy Memo Sample"; Ext = "doc"; Format = 0; Text = "Contoso Ltd.`r`nInternal Policy Memo`r`nTopic: Hybrid work equipment reimbursement policy updates." },
    @{ Name = "Wingtip Product Launch Checklist"; Ext = "docx"; Format = 16; Text = "Wingtip Co.`r`nProduct Launch Checklist`r`nItems: QA signoff, legal review, and launch readiness." }
)

$word = $null
$created = 0
$failed = 0

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    foreach ($sample in $samples) {
        $doc = $null
        try {
            $doc = $word.Documents.Add()
            $doc.Content.Text = [string]$sample.Text

            $fileName = "{0}.{1}" -f $sample.Name, $sample.Ext
            $filePath = Join-Path -Path $outPath -ChildPath $fileName

            if (Test-Path -LiteralPath $filePath) {
                Remove-Item -LiteralPath $filePath -Force
            }

            $doc.SaveAs([string]$filePath, [int]$sample.Format)
            Write-Host "Created: $filePath"
            $created++
        }
        catch {
            Write-Warning "Failed: $($sample.Name).$($sample.Ext): $($_.Exception.Message)"
            $failed++
        }
        finally {
            if ($doc) {
                $doc.Close($false)
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
            }
        }
    }
}
finally {
    if ($word) {
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host ""
Write-Host "Created: $created"
Write-Host "Failed:  $failed"

if ($failed -gt 0) {
    exit 1
}
