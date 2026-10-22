# Builds ClaudeReview.dotm (macros + "Claude" button group on Word's Home tab) into Word's STARTUP folder.
# Word must be closed. Re-run after editing ClaudeReview.bas or moving this folder
# (not needed for review-prompt.md / claude-review.mjs edits).
# To uninstall: delete ClaudeReview.dotm from %APPDATA%\Microsoft\Word\STARTUP.
param(
    # Where to put the built add-in. Default: Word's STARTUP folder, so Word loads it automatically.
    [string]$Target = (Join-Path $env:APPDATA 'Microsoft\Word\STARTUP\ClaudeReview.dotm')
)
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$tmp  = Join-Path $env:TEMP 'ClaudeReview-build.dotm'

if (Get-Process WINWORD -ErrorAction SilentlyContinue) { throw 'Close Word first (all windows), then re-run.' }

# The macro calls claude-review.mjs in this folder with Node - fill in the real locations on this PC
$node = (Get-Command node -ErrorAction SilentlyContinue).Source
if (-not $node) { throw 'Node.js not found. Install Node.js 18 or newer (nodejs.org), then re-run.' }
if (-not (Get-Command claude -ErrorAction SilentlyContinue) -and -not (Test-Path "$env:USERPROFILE\.local\bin\claude.exe")) {
    Write-Warning 'Claude Code (claude) not found - the Review Document button needs it installed and signed in.'
}
$src = [IO.File]::ReadAllText((Join-Path $here 'ClaudeReview.bas'), [Text.Encoding]::Default)
$src = $src.Replace('{{ADDIN_DIR}}', $here).Replace('{{NODE_EXE}}', $node)
$bas = Join-Path $env:TEMP 'ClaudeReview-build.bas'
[IO.File]::WriteAllText($bas, $src, [Text.Encoding]::Default)

# Temporarily allow programmatic access to the VBA project
$secKey = 'HKCU:\Software\Microsoft\Office\16.0\Word\Security'
if (-not (Test-Path $secKey)) { New-Item $secKey -Force | Out-Null }
Set-ItemProperty $secKey -Name AccessVBOM -Value 1 -Type DWord

$word = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Add()
    $null = $doc.VBProject.VBComponents.Import($bas)
    if (Test-Path $tmp) { Remove-Item $tmp -Force }
    $fn = [string]$tmp.ToString(); $ff = [int]15; $doc.SaveAs2([ref]$fn, [ref]$ff)   # 15 = wdFormatXMLTemplateMacroEnabled
    $no = [object]0; $doc.Close([ref]$no)
} finally {
    if ($word) { $no = [object]0; $word.Quit([ref]$no); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
    # Always leave VBA project access OFF (Word's default) - it's only needed during the build.
    # Word re-saves this setting while shutting down, so wait for it to exit before removing it.
    for ($i = 0; $i -lt 30 -and (Get-Process WINWORD -ErrorAction SilentlyContinue); $i++) { Start-Sleep -Milliseconds 500 }
    Remove-ItemProperty $secKey -Name AccessVBOM -ErrorAction SilentlyContinue
    Remove-Item $bas -Force -ErrorAction SilentlyContinue
}

# Add the ribbon buttons (customUI part) to the template package
$ribbon = @'
<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
  <commands>
    <!-- Word's own Accept/Reject (right-click menu + Review tab): one click resolves a whole Claude suggestion -->
    <command idMso="ReviewAcceptChange" onAction="WordAccept_Repurposed" />
    <command idMso="ReviewRejectChange" onAction="WordReject_Repurposed" />
    <command idMso="ReviewAcceptChangeAndMoveToNext" onAction="WordAcceptNext_Repurposed" />
    <command idMso="ReviewRejectChangeAndMoveToNext" onAction="WordRejectNext_Repurposed" />
  </commands>
  <ribbon>
    <tabs>
      <tab idMso="TabHome">
        <group id="grpClaude" label="Claude" insertBeforeMso="GroupClipboard">
          <button id="btnClaudeReview" label="Review Document" size="large" imageMso="ReviewTrackChanges"
                  onAction="ClaudeReview_Ribbon"
                  screentip="Review with Claude"
                  supertip="Sends this document to Claude for a proofread and legal review. Suggestions come back as tracked changes and comments by 'Claude Review'." />
          <button id="btnClaudeAccept" label="Accept" imageMso="ReviewAcceptChange"
                  onAction="ClaudeAccept_Ribbon"
                  screentip="Accept Claude suggestion"
                  supertip="Accepts the whole Claude suggestion at the cursor in one click, removes its comment, and moves to the next one." />
          <button id="btnClaudeReject" label="Reject" imageMso="ReviewRejectChange"
                  onAction="ClaudeReject_Ribbon"
                  screentip="Reject Claude suggestion"
                  supertip="Rejects the whole Claude suggestion at the cursor in one click, removes its comment, and moves to the next one." />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
'@
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
$zip = [IO.Compression.ZipFile]::Open($tmp, 'Update')
try {
    $e = $zip.CreateEntry('customUI/customUI.xml')
    $w = New-Object IO.StreamWriter($e.Open(), (New-Object Text.UTF8Encoding($false))); $w.Write($ribbon); $w.Close()

    $relsEntry = $zip.GetEntry('_rels/.rels')
    $r = New-Object IO.StreamReader($relsEntry.Open()); $rels = $r.ReadToEnd(); $r.Close()
    if ($rels -notmatch 'customUI/customUI.xml') {
        $rels = $rels -replace '</Relationships>', '<Relationship Id="rIdClaudeUI" Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility" Target="customUI/customUI.xml"/></Relationships>'
    }
    $relsEntry.Delete()
    $e = $zip.CreateEntry('_rels/.rels')
    $w = New-Object IO.StreamWriter($e.Open(), (New-Object Text.UTF8Encoding($false))); $w.Write($rels); $w.Close()
} finally { $zip.Dispose() }

New-Item -ItemType Directory -Force (Split-Path -Parent $Target) | Out-Null
Copy-Item $tmp $Target -Force
Remove-Item $tmp -Force
"Installed: $Target"
