Attribute VB_Name = "modEmbodimentFigureTag"
' =================================================================
' 具体实施方式精准标号工具（修正版）
' 功能：从“附图标记说明如下”解析标号，在选中的具体实施方式区域为部件名补充纯标号。
' 例：供氧机组 -> 供氧机组2
' =================================================================

Option Explicit

Sub 具体实施方式标号()
    On Error GoTo ErrorHandler

    Dim doc As Document
    Set doc = ActiveDocument

    Dim oldTrackRevisions As Boolean
    oldTrackRevisions = doc.TrackRevisions

    Dim targetRange As Range
    Set targetRange = Selection.Range

    If targetRange.Start = targetRange.End Then
        MsgBox "请先选中需要标号的具体实施方式区域。", vbExclamation
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

            Dim nextChar As String
            nextChar = NextCharAfterRange(currentFindRange)

            If IsEmbodimentNextCharForbidden(nextChar, suffixBlacklist) Then
                currentFindRange.Start = currentFindRange.End
                currentFindRange.End = endPosition
            Else
                Dim insertRange As Range
                Set insertRange = doc.Range(currentFindRange.End, currentFindRange.End)
                insertRange.text = code

                endPosition = endPosition + Len(code)
                currentFindRange.Start = insertRange.End
                currentFindRange.End = endPosition
            End If
        Loop
    Next i

    doc.TrackRevisions = oldTrackRevisions
    MsgBox "标号执行完毕：已在具体实施方式区域补充纯数字/字母标号。", vbInformation
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

Private Function IsEmbodimentNextCharForbidden(ByVal nextChar As String, ByVal suffixBlacklist As Object) As Boolean
    If nextChar = "" Then Exit Function

    If nextChar = "（" Or nextChar = "(" Then
        IsEmbodimentNextCharForbidden = True
        Exit Function
    End If

    If IsAsciiLetterOrDigit(nextChar) Or nextChar = "." Or nextChar = "-" Or nextChar = "－" Then
        IsEmbodimentNextCharForbidden = True
        Exit Function
    End If

    If suffixBlacklist.Exists(nextChar) Then IsEmbodimentNextCharForbidden = True
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
