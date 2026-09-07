Attribute VB_Name = "modClaimFigureTag"
' =================================================================
' 权利要求附图标记精准标注工具（修正版）
' 功能：从“附图标记说明如下”解析标号，在选中的权利要求区域为部件名补充（标号）。
' 保护：独权前序、发明名称、步骤前缀、已有标号、长短词嵌套。
' =================================================================

Option Explicit
Option Private Module

Private Const FIGURE_TAG_CONFLICT As String = "__FIGURE_TAG_CONFLICT__"

Sub 权利要求标号()
    On Error GoTo ErrorHandler

    Dim doc As Document
    Set doc = ActiveDocument

    Dim oldTrackRevisions As Boolean
    oldTrackRevisions = doc.TrackRevisions

    Dim targetRange As Range
    Set targetRange = Selection.Range

    If targetRange.Start = targetRange.End Then
        MsgBox "请先选中需要标号的权利要求区域。", vbExclamation
        Exit Sub
    End If

    Dim tagDict As Object
    Set tagDict = LoadFigureTagDictionary(doc)

    If tagDict Is Nothing Then
        MsgBox "未能从“附图标记说明如下”区域解析出有效标号。", vbCritical
        Exit Sub
    End If

    If tagDict.Count = 0 Then
        MsgBox "未能从“附图标记说明如下”区域解析出有效标号。", vbCritical
        Exit Sub
    End If

    Dim conflictReport As String
    conflictReport = FigureTagConflictReport(tagDict)
    If Len(conflictReport) > 0 Then
        MsgBox "附图标记说明中存在同名不同标号，已停止自动标注：" & vbCrLf & conflictReport, vbCritical
        Exit Sub
    End If

    Dim boundaryReport As String
    If SelectionBoundaryTouchesKnownName(targetRange, tagDict, boundaryReport) Then
        MsgBox "选区边界落在已知部件名称中，跨边界命中将跳过：" & vbCrLf & boundaryReport, vbExclamation
    End If

    Dim undoStarted As Boolean
    undoStarted = BeginCustomUndoRecord("权利要求标号")
    doc.TrackRevisions = True

    Dim protectedRanges As Collection
    Set protectedRanges = New Collection
    BuildClaimProtection targetRange, protectedRanges

    Dim keys As Variant
    keys = SortedKeysByLengthDesc(tagDict)

    Dim endPosition As Long
    endPosition = targetRange.End

    Dim i As Long
    Dim word As String
    Dim code As String
    Dim suffixBlacklist As Object
    Dim skippedRevisionCount As Long
    Dim skippedBoundaryCount As Long

    For i = LBound(keys) To UBound(keys)
        word = CStr(keys(i))
        code = CStr(tagDict.Item(word))
        Set suffixBlacklist = BuildSuffixBlacklist(word, tagDict)

        Dim currentFindRange As Range
        Set currentFindRange = doc.Range(targetRange.Start, endPosition)

        With currentFindRange.Find
        .ClearFormatting
        .Text = word
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
            .Forward = True
            .Wrap = wdFindStop
        End With

        Do While currentFindRange.Find.Execute
            If currentFindRange.End > endPosition Then
                skippedBoundaryCount = skippedBoundaryCount + 1
                Exit Do
            End If

            If RangeTouchesExistingRevision(currentFindRange) Then
                skippedRevisionCount = skippedRevisionCount + 1
                currentFindRange.Start = currentFindRange.End
                currentFindRange.End = endPosition
                GoTo ContinueCurrentMatch
            End If

            If IsInProtectedRange(currentFindRange.Start, protectedRanges) Then GoTo SkipMatch
            If IsStepPrefix(currentFindRange) Then GoTo SkipMatch
            If IsContainedInLongerTag(currentFindRange, word, tagDict) Then GoTo SkipMatch
            If IsPreviousTokenChar(currentFindRange) Then GoTo SkipMatch
            If IsAlreadyTagged(currentFindRange, code, True) Then GoTo SkipMatch

            Dim nextChar As String
            nextChar = NextCharAfterRange(currentFindRange)
            If IsClaimNextCharForbidden(nextChar, suffixBlacklist) Then GoTo SkipMatch

            Dim insertText As String
            insertText = "（" & code & "）"

            Dim insertRange As Range
            Set insertRange = doc.Range(currentFindRange.End, currentFindRange.End)
            insertRange.Text = insertText

            endPosition = endPosition + Len(insertText)
            UpdateProtectedRanges protectedRanges, insertRange.Start, Len(insertText)

            currentFindRange.Start = insertRange.End
            currentFindRange.End = endPosition
            GoTo ContinueLoop

SkipMatch:
            currentFindRange.Start = currentFindRange.End
            currentFindRange.End = endPosition

ContinueLoop:
ContinueCurrentMatch:
        Loop
    Next i

    EndCustomUndoRecord undoStarted
    doc.TrackRevisions = oldTrackRevisions
    If skippedRevisionCount > 0 Or skippedBoundaryCount > 0 Then
        MsgBox "执行完毕。已跳过既有修订命中 " & skippedRevisionCount & " 个、选区边界不完整命中 " & skippedBoundaryCount & " 个，请人工检查。", vbExclamation
    Else
        MsgBox "执行完毕：已在权利要求区域补充附图标记。", vbInformation
    End If
    Exit Sub

ErrorHandler:
    On Error Resume Next
    EndCustomUndoRecord undoStarted
    doc.TrackRevisions = oldTrackRevisions
    MsgBox "运行错误：" & Err.Description, vbCritical
End Sub

Private Function BeginCustomUndoRecord(ByVal recordName As String) As Boolean
    On Error Resume Next
    Err.Clear
    Application.UndoRecord.StartCustomRecord recordName
    BeginCustomUndoRecord = (Err.Number = 0)
    Err.Clear
End Function

Private Sub EndCustomUndoRecord(ByVal started As Boolean)
    On Error Resume Next
    If started Then Application.UndoRecord.EndCustomRecord
    Err.Clear
End Sub

Private Function RangeTouchesExistingRevision(ByVal rng As Range) As Boolean
    On Error GoTo ConservativeFallback

    If rng.Revisions.Count > 0 Then
        RangeTouchesExistingRevision = True
        Exit Function
    End If

    Dim probeStart As Long
    Dim probeEnd As Long
    probeStart = rng.Start - 1
    If probeStart < 0 Then probeStart = 0
    probeEnd = rng.End + 1
    If probeEnd > rng.Document.Content.End Then probeEnd = rng.Document.Content.End

    If rng.Document.Range(probeStart, probeEnd).Revisions.Count > 0 Then
        RangeTouchesExistingRevision = True
    End If
    Exit Function

ConservativeFallback:
    RangeTouchesExistingRevision = True
End Function

Private Function LoadFigureTagDictionary(ByVal doc As Document) As Object
    Dim searchRange As Range
    Set searchRange = doc.Content.Duplicate

    With searchRange.Find
        .ClearFormatting
        .Text = "附图标记说明如下"
        .Format = False
        .MatchWildcards = False
        .Forward = True
        .Wrap = wdFindStop
    End With

    If Not searchRange.Find.Execute Then
        Set LoadFigureTagDictionary = CreateStringMap()
        Exit Function
    End If

    Dim startPos As Long
    startPos = searchRange.End

    Dim endPos As Long
    endPos = FindTagAreaEnd(doc, startPos)

    Dim tagText As String
    tagText = doc.Range(startPos, endPos).Text
    Set LoadFigureTagDictionary = ParseFigureTagText(tagText)
End Function

Private Function FigureTagConflictReport(ByVal dict As Object) As String
    Dim key As Variant
    Dim value As String
    Dim report As String

    For Each key In dict.Keys
        value = CStr(dict.Item(key))
        If Left$(value, Len(FIGURE_TAG_CONFLICT)) = FIGURE_TAG_CONFLICT Then
            report = report & CStr(key) & "：" & Mid$(value, Len(FIGURE_TAG_CONFLICT) + 2) & vbCrLf
        End If
    Next key

    FigureTagConflictReport = report
End Function

Private Function SelectionBoundaryTouchesKnownName(ByVal targetRange As Range, ByVal tagDict As Object, ByRef detail As String) As Boolean
    Dim key As Variant
    Dim word As String

    For Each key In tagDict.Keys
        word = CStr(key)
        If NameCrossesBoundary(targetRange.Document, targetRange.Start, word) Then
            SelectionBoundaryTouchesKnownName = True
            detail = detail & "选区起点：" & word & vbCrLf
        End If
        If NameCrossesBoundary(targetRange.Document, targetRange.End, word) Then
            SelectionBoundaryTouchesKnownName = True
            detail = detail & "选区终点：" & word & vbCrLf
        End If
    Next key
End Function

Private Function NameCrossesBoundary(ByVal doc As Document, ByVal boundary As Long, ByVal word As String) As Boolean
    Dim startPos As Long
    Dim candidateStart As Long
    Dim candidateEnd As Long

    startPos = boundary - Len(word) + 1
    If startPos < 0 Then startPos = 0

    For candidateStart = startPos To boundary - 1
        candidateEnd = candidateStart + Len(word)
        If candidateEnd > boundary And candidateEnd <= doc.Content.End Then
            If doc.Range(candidateStart, candidateEnd).Text = word Then
                NameCrossesBoundary = True
                Exit Function
            End If
        End If
    Next candidateStart
End Function

Private Function NormalizeFigureCode(ByVal rawCode As String) As String
    rawCode = Trim(rawCode)

    Do While Len(rawCode) > 0
        Dim lastChar As String
        lastChar = Right(rawCode, 1)
        If lastChar = "-" Or lastChar = "－" Or lastChar = "—" Or lastChar = "." Or lastChar = "．" Then
            rawCode = Left(rawCode, Len(rawCode) - 1)
        Else
            Exit Do
        End If
    Loop

    NormalizeFigureCode = rawCode
End Function

Private Function FindTagAreaEnd(ByVal doc As Document, ByVal startPos As Long) As Long
    Dim maxEnd As Long
    maxEnd = startPos + 3000
    If maxEnd > doc.Content.End Then maxEnd = doc.Content.End

    Dim tagAreaRange As Range
    Set tagAreaRange = doc.Range(startPos, maxEnd)

    With tagAreaRange.Find
        .ClearFormatting
        .Text = "。"
        .Format = False
        .MatchWildcards = False
        .Forward = True
        .Wrap = wdFindStop
    End With

    If tagAreaRange.Find.Execute Then
        FindTagAreaEnd = tagAreaRange.End
    Else
        FindTagAreaEnd = maxEnd
    End If
End Function

Private Function SortedKeysByLengthDesc(ByVal dict As Object) As Variant
    Dim keys As Variant
    keys = dict.Keys

    Dim i As Long, j As Long
    Dim temp As Variant
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If Len(CStr(keys(i))) < Len(CStr(keys(j))) Then
                temp = keys(i)
                keys(i) = keys(j)
                keys(j) = temp
            End If
        Next j
    Next i

    SortedKeysByLengthDesc = keys
End Function

Private Sub BuildClaimProtection(ByVal targetRange As Range, ByRef protectedRanges As Collection)
    Dim invNames As Collection
    Set invNames = New Collection

    Dim para As Paragraph
    For Each para In targetRange.Paragraphs
        Dim pText As String
        pText = para.Range.Text

        If IsNumberedClaimLineCompat(pText) Then
            Dim featurePos As Long
            featurePos = InStr(1, pText, "其特征在于", vbTextCompare)

            If featurePos > 0 Then
                protectedRanges.Add Array(para.Range.Start, para.Range.Start + featurePos + Len("其特征在于") - 1)
                AddInventionNameFromPreamble pText, featurePos, invNames
            End If
        End If
    Next para

    Dim name As Variant
    For Each name In invNames
        AddAllOccurrencesToProtection targetRange, CStr(name), protectedRanges
    Next name
End Sub

Private Sub AddInventionNameFromPreamble(ByVal pText As String, ByVal featurePos As Long, ByRef invNames As Collection)
    Dim nameStart As Long
    nameStart = InStr(1, pText, "一种", vbTextCompare)
    If nameStart = 0 Then nameStart = InStr(1, pText, "一个", vbTextCompare)
    If nameStart = 0 Then nameStart = InStr(1, pText, "一项", vbTextCompare)

    If nameStart = 0 Or nameStart >= featurePos Then Exit Sub

    Dim rawName As String
    rawName = Mid(pText, nameStart + 2, featurePos - nameStart - 2)
    rawName = Split(Replace(rawName, "，", ","), ",")(0)
    rawName = Trim(rawName)

    If Len(rawName) > 1 Then AddUniqueString invNames, rawName
End Sub

Private Sub AddAllOccurrencesToProtection(ByVal targetRange As Range, ByVal textToProtect As String, ByRef protectedRanges As Collection)
    If Len(textToProtect) = 0 Then Exit Sub

    Dim scanRange As Range
    Set scanRange = targetRange.Duplicate

    With scanRange.Find
        .ClearFormatting
        .Text = textToProtect
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .Forward = True
        .Wrap = wdFindStop
    End With

    Do While scanRange.Find.Execute
        If scanRange.Start >= targetRange.End Then Exit Do
        protectedRanges.Add Array(scanRange.Start, scanRange.End)
        scanRange.Collapse wdCollapseEnd
    Loop
End Sub

Private Sub AddUniqueString(ByRef coll As Collection, ByVal value As String)
    Dim item As Variant
    For Each item In coll
        If CStr(item) = value Then Exit Sub
    Next item
    coll.Add value
End Sub

Private Function BuildSuffixBlacklist(ByVal word As String, ByVal tagDict As Object) As Object
    Dim result As Object
    Set result = CreateStringMap()

    Dim otherWord As Variant
    For Each otherWord In tagDict.Keys
        If Len(CStr(otherWord)) > Len(word) Then
            Dim pos As Long
            pos = InStr(1, CStr(otherWord), word, vbTextCompare)

            Do While pos > 0
                If pos + Len(word) <= Len(CStr(otherWord)) Then
                    Dim suffixChar As String
                    suffixChar = Mid(CStr(otherWord), pos + Len(word), 1)
                    If Not result.Exists(suffixChar) Then result.Add suffixChar, True
                End If
                pos = InStr(pos + 1, CStr(otherWord), word, vbTextCompare)
            Loop
        End If
    Next otherWord

    Set BuildSuffixBlacklist = result
End Function

Private Function IsClaimNextCharForbidden(ByVal nextChar As String, ByVal suffixBlacklist As Object) As Boolean
    If nextChar = "" Then Exit Function

    If IsAsciiLetterOrDigit(nextChar) Or nextChar = "." Or nextChar = "-" Or nextChar = "－" Then
        IsClaimNextCharForbidden = True
        Exit Function
    End If

    If suffixBlacklist.Exists(nextChar) Then IsClaimNextCharForbidden = True
End Function

Private Function IsContainedInLongerTag(ByVal foundRange As Range, ByVal word As String, ByVal tagDict As Object) As Boolean
    Dim otherWord As Variant
    Dim offset As Long
    Dim candidateStart As Long
    Dim candidateEnd As Long

    For Each otherWord In tagDict.Keys
        If Len(CStr(otherWord)) > Len(word) Then
            offset = InStr(1, CStr(otherWord), word, vbTextCompare)
            Do While offset > 0
                candidateStart = foundRange.Start - offset + 1
                candidateEnd = candidateStart + Len(CStr(otherWord))

                If candidateStart >= 0 And candidateEnd <= foundRange.Document.Content.End Then
                    If foundRange.Document.Range(candidateStart, candidateEnd).Text = CStr(otherWord) Then
                        IsContainedInLongerTag = True
                        Exit Function
                    End If
                End If

                offset = InStr(offset + 1, CStr(otherWord), word, vbTextCompare)
            Loop
        End If
    Next otherWord
End Function

Private Function IsPreviousTokenChar(ByVal rng As Range) As Boolean
    If rng.Start <= 0 Then Exit Function
    IsPreviousTokenChar = IsAsciiLetterOrDigit(rng.Document.Range(rng.Start - 1, rng.Start).Text)
End Function

Private Function IsAlreadyTagged(ByVal rng As Range, ByVal code As String, ByVal parenthesized As Boolean) As Boolean
    If Len(code) = 0 Then Exit Function

    Dim probeEnd As Long
    probeEnd = rng.End + Len(code) + 4
    If probeEnd > rng.Document.Content.End Then probeEnd = rng.Document.Content.End

    Dim tail As String
    tail = rng.Document.Range(rng.End, probeEnd).Text
    tail = LTrim$(Replace(Replace(tail, ChrW(160), " "), ChrW(12288), " "))

    If parenthesized Then
        IsAlreadyTagged = (Left$(tail, Len(code) + 2) = "（" & code & "）") Or _
                          (Left$(tail, Len(code) + 2) = "(" & code & ")")
    Else
        IsAlreadyTagged = (Left$(tail, Len(code)) = code)
    End If
End Function

Private Function IsStepPrefix(ByVal foundRange As Range) As Boolean
    Dim startPos As Long
    startPos = foundRange.Start - 8
    If startPos < 0 Then startPos = 0

    Dim txt As String
    txt = foundRange.Document.Range(startPos, foundRange.Start).Text

    Dim compact As String
    compact = UCase(Trim(Replace(Replace(txt, vbCr, ""), vbLf, "")))

    IsStepPrefix = (Right(compact, 1) = "S" Or Right(compact, 4) = "STEP" Or Right(txt, 2) = "步骤")
End Function

Private Function NextCharAfterRange(ByVal rng As Range) As String
    If rng.End >= rng.Document.Content.End Then
        NextCharAfterRange = ""
    Else
        NextCharAfterRange = rng.Document.Range(rng.End, rng.End + 1).Text
    End If
End Function

Private Function IsAsciiLetterOrDigit(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    Dim n As Long
    n = AscW(Left(ch, 1))
    IsAsciiLetterOrDigit = ((n >= 48 And n <= 57) Or (n >= 65 And n <= 90) Or (n >= 97 And n <= 122))
End Function

Private Function IsInProtectedRange(ByVal pos As Long, ByVal protectedRanges As Collection) As Boolean
    Dim item As Variant
    For Each item In protectedRanges
        If pos >= CLng(item(0)) And pos < CLng(item(1)) Then
            IsInProtectedRange = True
            Exit Function
        End If
    Next item
End Function

Private Sub UpdateProtectedRanges(ByRef protectedRanges As Collection, ByVal insertPos As Long, ByVal insertLen As Long)
    Dim newColl As Collection
    Set newColl = New Collection

    Dim i As Long
    Dim item As Variant
    For i = 1 To protectedRanges.Count
        item = protectedRanges(i)
        If insertPos < CLng(item(0)) Then
            item(0) = CLng(item(0)) + insertLen
            item(1) = CLng(item(1)) + insertLen
        ElseIf insertPos < CLng(item(1)) Then
            item(1) = CLng(item(1)) + insertLen
        End If
        newColl.Add item
    Next i

    Set protectedRanges = newColl
End Sub
