# Tests the Claude Accept/Reject logic (ClaudeReview.bas) in a hidden Word instance.
# Safe to run while Word is open: it uses its own hidden copy of Word and never touches your documents.
# Usage: powershell -ExecutionPolicy Bypass -File .\tests\test-accept-reject.ps1
$ErrorActionPreference = 'Stop'
$bas = Join-Path (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)) 'ClaudeReview.bas'

# Importing a module needs "Trust access to the VBA project object model" - turned on only for the test
$secKey = 'HKCU:\Software\Microsoft\Office\16.0\Word\Security'
if (-not (Test-Path $secKey)) { New-Item $secKey -Force | Out-Null }
Set-ItemProperty $secKey -Name AccessVBOM -Value 1 -Type DWord
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)

$job = Start-Job -ArgumentList $bas {
    param($bas)
    $w = New-Object -ComObject Word.Application; $w.Visible = $false; $w.DisplayAlerts = 0
    $oldUser = $w.UserName; $oldInit = $w.UserInitials; $oldLocal = $w.Options.UseLocalUserInfo
    $fails = 0
    function Check($label, $got, $want) {
        if ("$got" -eq "$want") { "PASS  $label" } else { "FAIL  $label (got $got, expected $want)"; $script:fails++ }
    }
    try {
        $d = $w.Documents.Add()
        $null = $d.VBProject.VBComponents.Import($bas)
        $d.Content.Text = "one DEFENDANTS two statue three INC four compliant five 1.410 six Plaintif seven"
        $w.Options.UseLocalUserInfo = $true
        function Suggest($old, $new, $who) {
            $w.UserName = $who
            $r = $d.Content; $f = $r.Find; $f.MatchCase = $true
            if ($f.Execute($old)) {
                if ($who -eq 'Claude Review') { $null = $d.Comments.Add($r, "fix $old") }
                $d.TrackRevisions = $true; $r.Text = $new; $d.TrackRevisions = $false
            }
        }
        function Cursor($txt) { $r = $d.Content; $null = $r.Find.Execute($txt); $r.Collapse(1); $null = $r.Move(1, 2); $r.Select() }
        Suggest 'DEFENDANTS' 'DEFENDANT''S' 'Claude Review'
        Suggest 'statue' 'statute' 'Claude Review'
        Suggest 'INC' 'LLC' 'Claude Review'
        Suggest 'compliant' 'Complaint' 'Claude Review'
        Suggest '1.410' '1.140' 'Claude Review'
        Suggest 'Plaintif' 'Plaintiff' 'Someone Else'
        Check 'setup: 6 suggestions = 12 tracked changes, 5 comments' "$($d.Revisions.Count)/$($d.Comments.Count)" '12/5'

        $a = [object]$null
        Cursor 'LLC'; $w.Run('ClaudeReview.ClaudeAccept_Ribbon', [ref]$a)
        Check 'Accept button resolves only the suggestion at the cursor' "$($d.Revisions.Count)/$($d.Comments.Count)" '10/4'
        Check 'Accept button moves to the next suggestion' $w.Selection.Text 'compliant'

        Cursor '1.140'; $b = [object]$false; $w.Run('ClaudeReview.WordAccept_Repurposed', [ref]$a, [ref]$b)
        Check 'Right-click Accept resolves deletion + insertion together' "$($d.Revisions.Count)/$($d.Comments.Count)" '8/3'

        Cursor 'Plaintiff'; $b = [object]$false; $w.Run('ClaudeReview.WordAccept_Repurposed', [ref]$a, [ref]$b)
        Check "Right-click Accept leaves another author's change to Word" $d.Revisions.Count 8

        Cursor 'statute'; $w.Run('ClaudeReview.ClaudeReject_Ribbon', [ref]$a)
        Check 'Reject button removes the change pair and its comment' "$($d.Revisions.Count)/$($d.Comments.Count)" '6/2'

        Cursor 'DEFENDANT'; $b = [object]$false; $w.Run('ClaudeReview.WordReject_Repurposed', [ref]$a, [ref]$b)
        Check 'Right-click Reject resolves the pair and its comment' "$($d.Revisions.Count)/$($d.Comments.Count)" '4/1'
    } finally {
        $w.UserName = $oldUser; $w.UserInitials = $oldInit; $w.Options.UseLocalUserInfo = $oldLocal
        $no = [object]0; foreach ($x in @($w.Documents)) { $x.Close([ref]$no) }; $w.Quit([ref]$no)
    }
    if ($fails) { "$fails test(s) FAILED" } else { 'All tests passed.' }
}

try {
    if (Wait-Job $job -Timeout 120) { Receive-Job $job } else { 'TIMED OUT - likely a VBA compile error (try opening Word and running the macro to see it).' }
} finally {
    Remove-Job $job -Force
    Start-Sleep 2
    # Clean up the hidden Word if it was left behind (only ones started by this test)
    Get-CimInstance Win32_Process -Filter "Name='WINWORD.EXE'" |
        Where-Object { $before -notcontains $_.ProcessId -and $_.CommandLine -match '/Automation' } |
        ForEach-Object { Stop-Process -Id $_.ProcessId -Force }
    Remove-ItemProperty $secKey -Name AccessVBOM -ErrorAction SilentlyContinue
}
