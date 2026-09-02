Attribute VB_Name = "modClaimFigureTag"
' =================================================================
' 权利要求附图标记精准标注工具（修正版）
' 功能：从“附图标记说明如下”解析标号，在选中的权利要求区域为部件名补充（标号）。
' 保护：独权前序、发明名称、步骤前缀、已有标号、长短词嵌套。
' =================================================================

Option Explicit
Option Private Module

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

    For i = LBound(keys) To UBound(keys)
        word = CStr(keys(i))
        code = CStr(tagDict(word))
        Set suffixBlacklist = BuildSuffixBlacklist(word, tagDict)

        Dim currentFindRange As Range
        Set currentFindRange = doc.Range(targetRange.Start, endPosition)

        With currentFindRange.Find
        .ClearFormatting
        .text = word
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
            .Forward = True
            .Wrap = wdFindStop
        End With

        Do While currentFindRange.Find.Execute
            If currentFindRange.End > endPosition Then Exit Do

            If IsInProtectedRange(currentFindRange.Start, protectedRanges) Then GoTo SkipMatch
            If IsStepPrefix(currentFindRange) Then GoTo SkipMatch

            Dim nextChar As String
            nextChar = NextCharAfterRange(currentFindRange)
            If IsClaimNextCharForbidden(nextChar, suffixBlacklist) Then GoTo SkipMatch

            Dim insertText As String
            insertText = "（" & code & "）"

            Dim insertRange As Range
            Set insertRange = doc.Range(currentFindRange.End, currentFindRange.End)
            insertRange.text = insertText

            endPosition = endPosition + Len(insertText)
            UpdateProtectedRanges protectedRanges, insertRange.Start, Len(insertText)

            currentFindRange.Start = insertRange.End
            currentFindRange.End = endPosition
            GoTo ContinueLoop

SkipMatch:
            currentFindRange.Start = currentFindRange.End
            currentFindRange.End = endPosition

ContinueLoop:
        Loop
    Next i

    doc.TrackRevisions = oldTrackRevisions
    MsgBox "执行完毕：已在权利要求区域补充附图标记。", vbInformation
    Exit Sub

ErrorHandler:
    On Error Resume Next
    doc.TrackRevisions = oldTrackRevisions
    MsgBox "运行错误：" & Err.Description, vbCritical
End Sub

Private Function LoadFigureTagDictionary(ByVal doc As Document) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    Dim searchRange As Range
    Set searchRange = doc.Content.Duplicate

    With searchRange.Find
        .ClearFormatting
        .text = "附图标记说明如下"
        .Format = False
        .MatchWildcards = False
        .Forward = True
        .Wrap = wdFindStop
    End With

    If Not searchRange.Find.Execute Then
        Set LoadFigureTagDictionary = result
        Exit Function
    End If

    Dim startPos As Long
    startPos = searchRange.End

    Dim endPos As Long
    endPos = FindTagAreaEnd(doc, startPos)

    Dim tagText As String
    tagText = doc.Range(startPos, endPos).text

    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.IgnoreCase = True
    regEx.pattern = "([0-9]+[0-9A-Za-z\.\-]*)\s*[-－—、,:：]?\s*([一-龥A-Za-z]+)"

    Dim matches As Object
    Set matches = regEx.Execute(tagText)

    Dim m As Object
    Dim code As String
    Dim word As String
    For Each m In matches
        code = NormalizeFigureCode(CStr(m.SubMatches(0)))
        word = Trim(CStr(m.SubMatches(1)))
        If Len(word) > 0 And word <> "图" Then
            If Not result.Exists(word) Then result.Add word, code
        End If
    Next m

    Set LoadFigureTagDictionary = result
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
        .text = "。"
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
    keys = dict.keys

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
    Dim claimStartRegEx As Object
    Set claimStartRegEx = CreateObject("VBScript.RegExp")
    claimStartRegEx.pattern = "^\s*\d+\s*[\.\．、]"
    claimStartRegEx.Global = False

    Dim invNames As Collection
    Set invNames = New Collection

    Dim para As Paragraph
    For Each para In targetRange.Paragraphs
        Dim pText As String
        pText = para.Range.text

        If claimStartRegEx.Test(Trim(pText)) Then
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
        .text = textToProtect
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
    Set result = CreateObject("Scripting.Dictionary")

    Dim otherWord As Variant
    For Each otherWord In tagDict.keys
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

    If nextChar = "（" Or nextChar = "(" Then
        IsClaimNextCharForbidden = True
        Exit Function
    End If

    If IsAsciiLetterOrDigit(nextChar) Or nextChar = "." Or nextChar = "-" Or nextChar = "－" Then
        IsClaimNextCharForbidden = True
        Exit Function
    End If

    If suffixBlacklist.Exists(nextChar) Then IsClaimNextCharForbidden = True
End Function

Private Function IsStepPrefix(ByVal foundRange As Range) As Boolean
    Dim startPos As Long
    startPos = foundRange.Start - 8
    If startPos < 0 Then startPos = 0

    Dim txt As String
    txt = foundRange.Document.Range(startPos, foundRange.Start).text

    Dim compact As String
    compact = UCase(Trim(Replace(Replace(txt, vbCr, ""), vbLf, "")))

    IsStepPrefix = (Right(compact, 1) = "S" Or Right(compact, 4) = "STEP" Or Right(txt, 2) = "步骤")
End Function

Private Function NextCharAfterRange(ByVal rng As Range) As String
    If rng.End >= rng.Document.Content.End Then
        NextCharAfterRange = ""
    Else
        NextCharAfterRange = rng.Document.Range(rng.End, rng.End + 1).text
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
