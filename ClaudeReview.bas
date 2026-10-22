Attribute VB_Name = "ClaudeReview"
' Review Document - sends the open document's text (body, headers, footers, text boxes, footnotes) to Claude (via claude-review.mjs,
' which runs Claude Code with no tools) and applies the suggestions as tracked changes + comments by "Claude Review".
' The review instructions are in review-prompt.md (edit freely; no rebuild needed).
Option Explicit

Private Const NODE_EXE As String = "{{NODE_EXE}}"          ' filled in by build-addin.ps1
Private Const REVIEW_JS As String = "{{ADDIN_DIR}}\claude-review.mjs"   ' filled in by build-addin.ps1
Private Const REVIEWER As String = "Claude Review"
Private Const TITLE As String = "Claude Review"

' Every reviewed paragraph gets an ID (1, 2, 3...). Its Range is kept here so the edit can find it again later;
' Word keeps stored Ranges up to date as the document changes.
Private mRanges As Collection, mLabels As Collection, mLines() As String, mCount As Long

' Ribbon button callbacks
Public Sub ClaudeReview_Ribbon(control As Object)
    ClaudeReview
End Sub

Public Sub ClaudeAccept_Ribbon(control As Object)
    ResolveAndAdvance True
End Sub

Public Sub ClaudeReject_Ribbon(control As Object)
    ResolveAndAdvance False
End Sub

' Word's own Accept / Reject commands (right-click menu and the Review tab) are routed here by the add-in's
' ribbon XML. On a Claude suggestion they resolve the deletion + insertion together, plus its comment;
' on anything else cancelDefault stays False and Word does its normal thing.
Public Sub WordAccept_Repurposed(control As Object, ByRef cancelDefault)
    cancelDefault = ResolveAtCursor(True, True)
End Sub

Public Sub WordReject_Repurposed(control As Object, ByRef cancelDefault)
    cancelDefault = ResolveAtCursor(False, True)
End Sub

Public Sub WordAcceptNext_Repurposed(control As Object, ByRef cancelDefault)
    cancelDefault = ResolveAtCursor(True, True)
    If cancelDefault Then MoveToNextChange
End Sub

Public Sub WordRejectNext_Repurposed(control As Object, ByRef cancelDefault)
    cancelDefault = ResolveAtCursor(False, True)
    If cancelDefault Then MoveToNextChange
End Sub

Private Sub MoveToNextChange()
    On Error Resume Next
    Selection.Collapse wdCollapseEnd
    Selection.NextRevision False
End Sub

Private Sub ResolveAndAdvance(ByVal accept As Boolean)
    Dim pos As Range
    If Documents.Count = 0 Then Exit Sub
    Set pos = Selection.Range.Duplicate
    If Not ResolveAtCursor(accept, False) Then
        Application.StatusBar = "No Claude suggestion at the cursor."
        Exit Sub
    End If
    GoToNextSuggestion ActiveDocument, pos
End Sub

' Accepts or rejects the whole Claude suggestion at the cursor (deletion + insertion together) and removes its
' comment. Only touches changes by "Claude Review". Returns False if there's no Claude suggestion there.
' needRevision: only act if the cursor is on one of Claude's tracked changes (used for Word's own Accept/Reject).
Private Function ResolveAtCursor(ByVal accept As Boolean, ByVal needRevision As Boolean) As Boolean
    Dim doc As Document, sel As Range, rng As Range, near As Range, c As Comment, found As Comment
    Dim rv As Revision, i As Long

    If Documents.Count = 0 Then Exit Function
    Set doc = ActiveDocument
    Set sel = Selection.Range

    If needRevision Then
        If FirstClaudeRevision(sel) Is Nothing Then Exit Function
    End If

    ' The suggestion's comment, if the cursor is inside one
    For Each c In doc.Comments
        If IsSuggestionComment(c) Then
            If Touches(c.Scope, sel) Then Set found = c: Exit For
        End If
    Next c
    If found Is Nothing Then Set rng = sel.Duplicate Else Set rng = found.Scope.Duplicate

    ' Grow the range to cover the Claude changes right next to it (the deletion and its matching insertion)
    For i = 1 To 2
        Set near = rng.Duplicate
        near.MoveStart wdCharacter, -1: near.MoveEnd wdCharacter, 1
        For Each rv In near.Revisions
            If rv.Author = REVIEWER Then
                If Touches(rv.Range, near) Then
                    If rv.Range.Start < rng.Start Then rng.Start = rv.Range.Start
                    If rv.Range.End > rng.End Then rng.End = rv.Range.End
                End If
            End If
        Next rv
    Next i

    ' The comment sits on the original text, so a cursor in the new text can miss it; look again with the full pair
    If found Is Nothing Then
        For Each c In doc.Comments
            If IsSuggestionComment(c) Then
                If Touches(c.Scope, rng) Then Set found = c: Exit For
            End If
        Next c
    End If

    If FirstClaudeRevision(rng) Is Nothing And found Is Nothing Then Exit Function

    Application.UndoRecord.StartCustomRecord IIf(accept, "Accept", "Reject") & " Claude suggestion"
    On Error GoTo Done
    ' Look the changes up again after each one: accepting one can invalidate Word's handle to the next.
    For i = 1 To 50
        Set rv = FirstClaudeRevision(rng)
        If rv Is Nothing Then Exit For
        If accept Then rv.Accept Else rv.Reject
    Next i
    On Error Resume Next
    If Not found Is Nothing Then found.Delete   ' may already be gone with the text it was on
    ResolveAtCursor = True
Done:
    Application.UndoRecord.EndCustomRecord
End Function

' Range.Revisions can also return changes OUTSIDE the range (from the top of the document), so every
' revision is checked against the range's real position.
Private Function FirstClaudeRevision(ByVal rng As Range) As Revision
    Dim rv As Revision
    For Each rv In rng.Revisions
        If rv.Author = REVIEWER Then
            If Touches(rv.Range, rng) Then Set FirstClaudeRevision = rv: Exit Function
        End If
    Next rv
End Function

' True if a overlaps b (same part of the document). A collapsed b (just a cursor) counts if it sits inside a.
Private Function Touches(ByVal a As Range, ByVal b As Range) As Boolean
    If a.StoryType <> b.StoryType Then Exit Function
    If b.Start = b.End Then
        Touches = (a.Start <= b.Start And a.End >= b.End)
    Else
        Touches = (a.Start < b.End And a.End > b.Start)
    End If
End Function

Private Function IsSuggestionComment(ByVal c As Comment) As Boolean
    If c.Author <> REVIEWER Then Exit Function
    IsSuggestionComment = (Left$(c.Range.Text, 21) <> "Claude Review summary")
End Function

' Selects the next Claude comment after pos (wrapping to the top), or reports that none are left.
Private Sub GoToNextSuggestion(ByVal doc As Document, ByVal pos As Range)
    Dim c As Comment, first As Comment, nxt As Comment
    For Each c In doc.Comments
        If IsSuggestionComment(c) Then
            If first Is Nothing Then Set first = c
            If c.Scope.StoryType = pos.StoryType And c.Scope.Start >= pos.Start Then Set nxt = c: Exit For
        End If
    Next c
    If nxt Is Nothing Then Set nxt = first
    If nxt Is Nothing Then
        Application.StatusBar = "No more Claude comments. Any remaining header/footer changes: Review > Next."
    Else
        nxt.Scope.Select
    End If
End Sub

Public Sub ClaudeReview()
    Dim doc As Document, inFile As String, outFile As String, stamp As String
    Dim n As Long, lines() As String, res As String

    If Documents.Count = 0 Then Exit Sub
    Set doc = ActiveDocument
    If doc.ProtectionType <> wdNoProtection Then
        MsgBox "This document is protected, so tracked changes can't be added. Unprotect it first.", vbExclamation, TITLE
        Exit Sub
    End If

    Dim msg As String
    msg = "Send """ & doc.Name & """ to Claude for review?" & vbCrLf & vbCrLf & _
          "Suggestions come back as tracked changes and comments by """ & REVIEWER & """. " & _
          "Word will be busy for a minute or more while Claude reads it."
    If doc.Revisions.Count > 0 Then
        msg = msg & vbCrLf & vbCrLf & "Note: this document already has " & doc.Revisions.Count & _
              " tracked change(s). Claude will see deleted text too. Accepting or rejecting them first gives a cleaner review."
    End If
    If MsgBox(msg, vbOKCancel + vbQuestion, TITLE) <> vbOK Then Exit Sub

    stamp = Format$(Now, "yyyymmddhhnnss")
    inFile = Environ$("TEMP") & "\claude-review-" & stamp & ".txt"
    outFile = Environ$("TEMP") & "\claude-review-" & stamp & ".tsv"

    Application.StatusBar = "Claude Review: reading document..."
    n = ExportParagraphs(doc, inFile)
    If n = 0 Then MsgBox "There's no text to review.", vbInformation, TITLE: GoTo Cleanup

    Application.StatusBar = "Claude Review: Claude is reviewing """ & doc.Name & """ (see the black window)..."
    CreateObject("WScript.Shell").Run """" & NODE_EXE & """ --no-warnings """ & REVIEW_JS & """ """ & _
        inFile & """ """ & outFile & """", 1, True   ' visible progress window, wait

    If Dir(outFile) = "" Then MsgBox "The review didn't produce a result.", vbExclamation, TITLE: GoTo Cleanup
    res = ReadUtf8(outFile)
    lines = Split(res, vbCrLf)
    If lines(0) <> "OK" Then
        MsgBox "The review failed:" & vbCrLf & vbCrLf & Mid$(lines(0), 5), vbExclamation, TITLE
        GoTo Cleanup
    End If

    ApplyReview doc, lines

Cleanup:
    On Error Resume Next
    If Dir(inFile) <> "" Then Kill inFile   ' don't leave client text lying around in TEMP
    If Dir(outFile) <> "" Then Kill outFile
    Application.StatusBar = False
End Sub

' Writes "[id] text" (body) or "[id|Header, section 2] text" (other parts) for every non-empty paragraph in the
' body, headers, footers, text boxes, footnotes and endnotes. Returns how many were written.
Private Function ExportParagraphs(ByVal doc As Document, ByVal file As String) As Long
    Dim sec As Section, hf As HeaderFooter, k As Long, shp As Shape

    Set mRanges = New Collection: Set mLabels = New Collection
    ReDim mLines(0 To 255): mCount = 0

    AddStory doc.Content, ""
    For Each sec In doc.Sections
        For k = 1 To 3   ' primary, first page, even pages
            Set hf = sec.Headers(k)
            If HeaderWanted(hf, sec) Then AddStory hf.Range, HFLabel("Header", k, sec, doc)
            Set hf = sec.Footers(k)
            If HeaderWanted(hf, sec) Then AddStory hf.Range, HFLabel("Footer", k, sec, doc)
        Next k
    Next sec
    For Each shp In doc.Shapes
        AddShape shp, "Text box"
    Next shp
    For Each shp In doc.Sections(1).Headers(wdHeaderFooterPrimary).Shapes   ' all header/footer shapes
        AddShape shp, "Header/footer text box"
    Next shp
    If doc.Footnotes.Count > 0 Then AddStory doc.StoryRanges(wdFootnotesStory), "Footnote"
    If doc.Endnotes.Count > 0 Then AddStory doc.StoryRanges(wdEndnotesStory), "Endnote"

    If mCount = 0 Then Exit Function
    ReDim Preserve mLines(0 To mCount - 1)
    WriteUtf8 file, Join(mLines, vbLf)
    ExportParagraphs = mCount
End Function

Private Function HeaderWanted(ByVal hf As HeaderFooter, ByVal sec As Section) As Boolean
    On Error Resume Next
    If Not hf.Exists Then Exit Function
    If sec.Index > 1 And hf.LinkToPrevious Then Exit Function   ' same text as the previous section's
    HeaderWanted = True
End Function

Private Function HFLabel(ByVal kind As String, ByVal k As Long, ByVal sec As Section, ByVal doc As Document) As String
    HFLabel = kind & Choose(k, "", " (first page)", " (even pages)")
    If doc.Sections.Count > 1 Then HFLabel = HFLabel & ", section " & sec.Index
End Function

Private Sub AddShape(ByVal shp As Shape, ByVal label As String)
    Dim i As Long, hasText As Boolean
    On Error Resume Next
    If shp.Type = 6 Then   ' msoGroup
        For i = 1 To shp.GroupItems.Count
            AddShape shp.GroupItems(i), label
        Next i
        Exit Sub
    End If
    hasText = shp.TextFrame.HasText
    If Err.Number <> 0 Then Exit Sub   ' pictures etc. have no text frame
    On Error GoTo 0
    If hasText Then AddStory shp.TextFrame.TextRange, label
End Sub

Private Sub AddStory(ByVal story As Range, ByVal label As String)
    Dim p As Paragraph, t As String
    For Each p In story.Paragraphs
        t = p.Range.Text
        Do While Len(t) > 0 And (Right$(t, 1) = vbCr Or Right$(t, 1) = Chr$(7))
            t = Left$(t, Len(t) - 1)
        Loop
        t = Replace(Replace(t, vbCr, " "), vbLf, " ")
        If Len(Trim$(Replace(Replace(t, vbTab, ""), Chr$(11), ""))) > 0 Then
            mCount = mCount + 1
            mRanges.Add p.Range, CStr(mCount)
            mLabels.Add label, CStr(mCount)
            If mCount > UBound(mLines) + 1 Then ReDim Preserve mLines(0 To UBound(mLines) * 2 + 1)
            mLines(mCount - 1) = "[" & mCount & IIf(label = "", "", "|" & label) & "] " & t
        End If
    Next p
End Sub

' Adds a comment; returns False where Word doesn't allow comments (headers and footers).
Private Function TryComment(ByVal doc As Document, ByVal rng As Range, ByVal text As String) As Boolean
    On Error Resume Next
    doc.Comments.Add rng, text
    TryComment = (Err.Number = 0)
End Function

Private Sub ApplyReview(ByVal doc As Document, ByRef lines() As String)
    Dim i As Long, f() As String, summary As String, missed As String
    Dim nEdits As Long, nComments As Long, ok As Boolean
    Dim oldTrack As Boolean, oldUser As String, oldInit As String, oldLocal As Boolean
    Dim pr As Range, rng As Range, part As Range
    Dim findTxt As String, oldMid As String, newMid As String, p As Long, where As String, noRoom As String

    oldTrack = doc.TrackRevisions
    oldUser = Application.UserName: oldInit = Application.UserInitials
    oldLocal = Options.UseLocalUserInfo

    On Error GoTo Failed
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "Claude Review"
    Options.UseLocalUserInfo = True             ' so the author shows as "Claude Review", not the Office account
    Application.UserName = REVIEWER: Application.UserInitials = "CR"
    doc.TrackRevisions = True

    For i = 1 To UBound(lines)
        f = Split(lines(i), vbTab)
        If UBound(f) >= 1 And f(0) = "SUMMARY" Then
            summary = f(1)
        ElseIf UBound(f) >= 6 Then
            p = CLng(f(0))
            findTxt = DecodeText(f(1)): oldMid = DecodeText(f(3)): newMid = DecodeText(f(4))
            Set rng = Nothing
            If p >= 1 And p <= mCount Then
                Set pr = mRanges(CStr(p))
                Set rng = FindInRange(pr, findTxt)
                where = mLabels(CStr(p)): If where = "" Then where = "Body"
            End If
            If rng Is Nothing Then
                missed = missed & " " & where & ": """ & findTxt & """" & _
                         IIf(f(6) = "E", " -> """ & newMid & """", "") & " (" & f(5) & ");"
            Else
                If f(5) <> "" Then
                    If Not TryComment(doc, rng, f(5)) Then
                        noRoom = noRoom & " " & where & ", """ & findTxt & """: " & f(5) & ";"
                    End If
                End If
                If f(6) = "E" Then
                    Set part = rng.Duplicate
                    part.MoveStart wdCharacter, CLng(f(2))
                    part.End = part.Start + Len(oldMid)
                    part.Text = newMid
                    nEdits = nEdits + 1
                Else
                    nComments = nComments + 1
                End If
            End If
        End If
    Next i

    If noRoom <> "" Then summary = summary & " [Notes Word couldn't attach as comments (headers/footers):" & noRoom & "]"
    If missed <> "" Then summary = summary & " [Couldn't place in the document:" & missed & "]"
    If summary <> "" Then doc.Comments.Add doc.Paragraphs(1).Range, "Claude Review summary: " & summary
    ok = True

Done:
    On Error Resume Next
    doc.TrackRevisions = oldTrack
    Application.UserName = oldUser: Application.UserInitials = oldInit
    Options.UseLocalUserInfo = oldLocal
    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
    If ok Then
        MsgBox "Review done: " & nEdits & " suggested edit(s) and " & nComments & " comment(s)." & vbCrLf & vbCrLf & _
               "Use Review > Accept / Reject to go through them. Ctrl+Z (once) removes the whole review.", _
               vbInformation, TITLE
    End If
    Exit Sub
Failed:
    MsgBox "Stopped while applying the review (" & nEdits & " edits applied so far):" & vbCrLf & Err.Description, _
           vbExclamation, TITLE
    Resume Done
End Sub

' Finds the first occurrence of txt inside pr (case-sensitive, literal). Returns Nothing if not found.
Private Function FindInRange(ByVal pr As Range, ByVal txt As String) As Range
    Dim rng As Range, head As String
    head = txt
    If Len(head) > 250 Then head = Left$(head, 250)   ' Word's Find is limited to 255 characters
    Set rng = pr.Duplicate
    With rng.Find
        .ClearFormatting
        .Text = FindCode(head)
        .Forward = True: .Wrap = wdFindStop: .Format = False
        .MatchCase = True: .MatchWholeWord = False: .MatchWildcards = False
        .MatchSoundsLike = False: .MatchAllWordForms = False
        If Not .Execute Then Exit Function
    End With
    If rng.Start < pr.Start Or rng.End > pr.End Then Exit Function
    If Len(txt) > Len(head) Then rng.MoveEnd wdCharacter, Len(txt) - Len(head)
    Set FindInRange = rng
End Function

' Turns literal text into Word Find syntax.
Private Function FindCode(ByVal s As String) As String
    s = Replace(s, "^", "^^")
    s = Replace(s, vbTab, "^t")
    s = Replace(s, Chr$(11), "^l")
    FindCode = s
End Function

' The helper marks tabs and manual line breaks as {TAB} / {BR}.
Private Function DecodeText(ByVal s As String) As String
    DecodeText = Replace(Replace(s, "{TAB}", vbTab), "{BR}", Chr$(11))
End Function

Private Sub WriteUtf8(ByVal file As String, ByVal text As String)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.WriteText text
    st.SaveToFile file, 2
    st.Close
End Sub

Private Function ReadUtf8(ByVal file As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.LoadFromFile file
    ReadUtf8 = st.ReadText
    st.Close
End Function
