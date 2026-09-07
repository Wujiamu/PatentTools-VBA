Attribute VB_Name = "modPatentCompat"
' =================================================================
' 跨平台兼容层：不依赖 Windows Scripting Runtime 的基础规则。
' 目标：Word 2010 Windows、较新 Windows Word 和 Mac Word 均可使用。
' =================================================================

Option Explicit
Option Private Module

Public Function CreateStringMap() As Object
    Dim nativeMap As Object

    On Error Resume Next
    Set nativeMap = CreateObject("Scripting.Dictionary")
    On Error GoTo 0

    If nativeMap Is Nothing Then
        Set CreateStringMap = New clsStringMap
    Else
        On Error Resume Next
        nativeMap.CompareMode = vbTextCompare
        On Error GoTo 0
        Set CreateStringMap = nativeMap
    End If
End Function

Public Function ParseFigureTagText(ByVal tagText As String) As Object
    Dim result As Object
    Set result = CreateStringMap()

    Dim i As Long
    Dim n As Long
    Dim rawCode As String
    Dim code As String
    Dim figureName As String

    i = 1
    n = Len(tagText)

    Do While i <= n
        If IsFigureCodeStart(Mid$(tagText, i, 1)) And IsFigureCodeBoundary(tagText, i) Then
            rawCode = ReadFigureCode(tagText, i)
            code = NormalizeFigureCodeCompat(rawCode)

            SkipInlineSpaces tagText, i
            If i <= n Then
                If IsFigureCodeSeparator(Mid$(tagText, i, 1)) Then
                    i = i + 1
                    SkipInlineSpaces tagText, i
                End If
            End If

            figureName = ReadFigureName(tagText, i)
            If Len(figureName) > 0 And figureName <> "图" Then
                AddFigureTagEntry result, figureName, code
            Else
                i = i + 1
            End If
        Else
            i = i + 1
        End If
    Loop

    Set ParseFigureTagText = result
End Function

Private Sub AddFigureTagEntry(ByVal result As Object, ByVal word As String, ByVal code As String)
    Const conflictMarker As String = "__FIGURE_TAG_CONFLICT__"

    If Not result.Exists(word) Then
        result.Add word, code
    ElseIf CStr(result.Item(word)) <> code Then
        result.Item(word) = conflictMarker & "|" & CStr(result.Item(word)) & "|" & code
    End If
End Sub

Private Function ReadFigureCode(ByVal text As String, ByRef position As Long) As String
    Dim startPos As Long
    startPos = position

    Do While position <= Len(text) And IsFigureCodeChar(Mid$(text, position, 1))
        position = position + 1
    Loop

    ReadFigureCode = Mid$(text, startPos, position - startPos)
End Function

Private Function ReadFigureName(ByVal text As String, ByRef position As Long) As String
    Dim startPos As Long
    startPos = position

    If position > Len(text) Or Not IsFigureNameStart(Mid$(text, position, 1)) Then Exit Function

    Do While position <= Len(text) And IsFigureNameChar(Mid$(text, position, 1))
        position = position + 1
    Loop

    ReadFigureName = Trim$(Mid$(text, startPos, position - startPos))
End Function

Private Sub SkipInlineSpaces(ByVal text As String, ByRef position As Long)
    Do While position <= Len(text)
        If Mid$(text, position, 1) = " " Or Mid$(text, position, 1) = vbTab Or _
           Mid$(text, position, 1) = ChrW(160) Or Mid$(text, position, 1) = ChrW(12288) Then
            position = position + 1
        Else
            Exit Do
        End If
    Loop
End Sub

Private Function IsFigureCodeBoundary(ByVal text As String, ByVal position As Long) As Boolean
    If position <= 1 Then
        IsFigureCodeBoundary = True
        Exit Function
    End If

    Dim beforeChar As String
    beforeChar = Mid$(text, position - 1, 1)

    IsFigureCodeBoundary = (InStr(" " & vbTab & vbCr & vbLf & _
                                 "、，,；;：:。．！？!?－—-（([{【《", beforeChar) > 0) Or _
                           beforeChar = ChrW(160) Or beforeChar = ChrW(12288)
End Function

Private Function IsFigureCodeStart(ByVal ch As String) As Boolean
    IsFigureCodeStart = IsAsciiDigit(ch) Or IsFullWidthDigit(ch) Or _
                        IsAsciiLetter(ch) Or IsFullWidthLetter(ch)
End Function

Private Function IsFigureCodeChar(ByVal ch As String) As Boolean
    IsFigureCodeChar = IsFigureCodeStart(ch) Or ch = "." Or ch = "．" Or _
                       ch = "-" Or ch = "－" Or ch = "′" Or ch = "'" Or ch = "’"
End Function

Private Function IsFigureCodeSeparator(ByVal ch As String) As Boolean
    IsFigureCodeSeparator = (InStr("-－—、，,；;：:", ch) > 0)
End Function

Private Function IsFigureNameStart(ByVal ch As String) As Boolean
    If ch = "" Then Exit Function
    IsFigureNameStart = IsChineseCharCompat(ch) Or IsFigureCodeStart(ch)
End Function

Private Function IsFigureNameChar(ByVal ch As String) As Boolean
    If ch = "" Then Exit Function
    IsFigureNameChar = IsChineseCharCompat(ch) Or IsFigureCodeStart(ch) Or _
                       InStr("/_／-－′'’", ch) > 0
End Function

Private Function NormalizeFigureCodeCompat(ByVal rawCode As String) As String
    rawCode = Trim$(rawCode)
    Do While Len(rawCode) > 0
        If InStr("-－—.．", Right$(rawCode, 1)) > 0 Then
            rawCode = Left$(rawCode, Len(rawCode) - 1)
        Else
            Exit Do
        End If
    Loop
    NormalizeFigureCodeCompat = rawCode
End Function

Private Function IsAsciiDigit(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    IsAsciiDigit = (AscW(ch) >= 48 And AscW(ch) <= 57)
End Function

Private Function IsAsciiLetter(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    Dim value As Long
    value = AscW(ch)
    IsAsciiLetter = (value >= 65 And value <= 90) Or (value >= 97 And value <= 122)
End Function

Private Function IsFullWidthDigit(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    Dim value As Long
    value = AscW(ch)
    If value < 0 Then value = value + 65536
    IsFullWidthDigit = (value >= 65296 And value <= 65305)
End Function

Private Function IsFullWidthLetter(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    Dim value As Long
    value = AscW(ch)
    If value < 0 Then value = value + 65536
    IsFullWidthLetter = (value >= 65313 And value <= 65338) Or _
                        (value >= 65345 And value <= 65370)
End Function

Private Function IsChineseCharCompat(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    Dim value As Long
    value = AscW(ch)
    If value < 0 Then value = value + 65536
    IsChineseCharCompat = (value >= 13312 And value <= 40959)
End Function

Public Function IsNumberedClaimLineCompat(ByVal text As String) As Boolean
    Dim position As Long
    Dim startPos As Long

    position = 1
    SkipInlineSpaces text, position
    startPos = position

    Do While position <= Len(text) And IsClaimDigitCompat(Mid$(text, position, 1))
        position = position + 1
    Loop

    If position = startPos Then Exit Function
    SkipInlineSpaces text, position

    If position <= Len(text) Then
        IsNumberedClaimLineCompat = (InStr(".．、", Mid$(text, position, 1)) > 0)
    End If
End Function

Public Function LeadingClaimNumberCompat(ByVal text As String) As Long
    Dim position As Long
    Dim startPos As Long
    Dim value As Long
    Dim digitValue As Long

    position = 1
    SkipInlineSpaces text, position
    startPos = position

    Do While position <= Len(text) And IsClaimDigitCompat(Mid$(text, position, 1))
        digitValue = ClaimDigitValueCompat(Mid$(text, position, 1))
        If value > 214748364 Or (value = 214748364 And digitValue > 7) Then Exit Function
        value = value * 10 + digitValue
        position = position + 1
    Loop

    If position > startPos Then LeadingClaimNumberCompat = value
End Function

Public Function RemoveLeadingClaimNumberCompat(ByVal text As String) As String
    Dim position As Long
    Dim startPos As Long

    position = 1
    SkipInlineSpaces text, position
    startPos = position

    Do While position <= Len(text) And IsClaimDigitCompat(Mid$(text, position, 1))
        position = position + 1
    Loop

    If position = startPos Then
        RemoveLeadingClaimNumberCompat = text
        Exit Function
    End If

    SkipInlineSpaces text, position
    If position <= Len(text) And InStr(".．、", Mid$(text, position, 1)) > 0 Then
        position = position + 1
        SkipInlineSpaces text, position
    End If

    RemoveLeadingClaimNumberCompat = Mid$(text, position)
End Function

Public Function StripClaimReferencePrefixCompat(ByVal text As String) As String
    Dim position As Long
    Dim keywordPos As Long
    Dim endPos As Long
    Dim prefixLength As Long

    position = 1
    SkipInlineSpaces text, position

    If Mid$(text, position, 1) = "如" Then
        prefixLength = 1
    ElseIf Mid$(text, position, 2) = "根据" Then
        prefixLength = 2
    Else
        StripClaimReferencePrefixCompat = text
        Exit Function
    End If

    keywordPos = position + prefixLength
    SkipInlineSpaces text, keywordPos
    If Mid$(text, keywordPos, 4) <> "权利要求" Then
        StripClaimReferencePrefixCompat = text
        Exit Function
    End If

    endPos = FindClaimReferencePrefixEndCompat(text, keywordPos)
    If endPos = 0 Then
        StripClaimReferencePrefixCompat = text
    Else
        SkipInlineSpaces text, endPos
        If Mid$(text, endPos, 1) = "的" Then endPos = endPos + 1
        SkipInlineSpaces text, endPos
        StripClaimReferencePrefixCompat = Mid$(text, endPos)
    End If
End Function

Public Function ReplaceClaimReferencesCompat(ByVal text As String) As String
    Dim searchPos As Long
    Dim keywordPos As Long
    Dim endPos As Long
    Dim replaceStart As Long
    Dim replacement As String

    searchPos = 1
    Do
        keywordPos = InStr(searchPos, text, "权利要求", vbTextCompare)
        If keywordPos = 0 Then Exit Do

        replaceStart = ClaimReferenceStartCompat(text, keywordPos, replacement)
        endPos = FindClaimReferenceEndCompat(text, keywordPos)

        If endPos > 0 Then
            If replacement = "前述的" And Mid$(text, endPos, 1) = "的" Then endPos = endPos + 1
            text = Left$(text, replaceStart - 1) & replacement & Mid$(text, endPos)
            searchPos = replaceStart + Len(replacement)
        Else
            searchPos = keywordPos + Len("权利要求")
        End If
    Loop

    ReplaceClaimReferencesCompat = text
End Function

Public Function RemoveDrawingNumbersCompat(ByVal text As String) As String
    Dim position As Long
    Dim closePos As Long
    Dim candidate As String
    Dim openChar As String

    position = 1
    Do While position <= Len(text)
        openChar = Mid$(text, position, 1)
        If openChar = "(" Or openChar = "（" Then
            closePos = FindDrawingNumberCloseCompat(text, position + 1)
            If closePos > position + 1 Then
                candidate = Mid$(text, position + 1, closePos - position - 1)
                If IsAsciiAlphaNumericStringCompat(candidate) Then
                    text = Left$(text, position - 1) & Mid$(text, closePos + 1)
                    position = position
                Else
                    position = closePos + 1
                End If
            Else
                position = position + 1
            End If
        Else
            position = position + 1
        End If
    Loop

    RemoveDrawingNumbersCompat = text
End Function

Public Function ExtractCaseNoCompat(ByVal fileName As String) As String
    Dim position As Long
    Dim startPos As Long
    Dim candidate As String

    position = 1
    Do While position <= Len(fileName)
        If IsAsciiAlphaNumericCompat(Mid$(fileName, position, 1)) Then
            startPos = position
            Do While position <= Len(fileName) And IsCaseNoCharCompat(Mid$(fileName, position, 1))
                position = position + 1
            Loop

            candidate = Mid$(fileName, startPos, position - startPos)
            Do While Len(candidate) > 0 And Not IsAsciiAlphaNumericCompat(Right$(candidate, 1))
                candidate = Left$(candidate, Len(candidate) - 1)
            Loop
            If Len(candidate) >= 2 Then
                If IsAsciiAlphaNumericCompat(Left$(candidate, 1)) And _
                   IsAsciiAlphaNumericCompat(Right$(candidate, 1)) Then
                    ExtractCaseNoCompat = candidate
                    Exit Function
                End If
            End If
        Else
            position = position + 1
        End If
    Loop
End Function

Private Function IsClaimDigitCompat(ByVal ch As String) As Boolean
    IsClaimDigitCompat = IsAsciiDigit(ch) Or IsFullWidthDigit(ch)
End Function

Private Function ClaimDigitValueCompat(ByVal ch As String) As Long
    If IsAsciiDigit(ch) Then
        ClaimDigitValueCompat = AscW(ch) - 48
    ElseIf IsFullWidthDigit(ch) Then
        Dim value As Long
        value = AscW(ch)
        If value < 0 Then value = value + 65536
        ClaimDigitValueCompat = value - 65296
    End If
End Function

Private Function ClaimReferenceStartCompat(ByVal text As String, ByVal keywordPos As Long, ByRef replacement As String) As Long
    Dim position As Long
    position = keywordPos - 1
    SkipInlineSpacesBackCompat text, position

    If position > 0 And Mid$(text, position, 1) = "如" Then
        replacement = "前述"
        ClaimReferenceStartCompat = position
        Exit Function
    End If

    If position >= 2 And Mid$(text, position - 1, 2) = "根据" Then
        replacement = "前述"
        ClaimReferenceStartCompat = position - 1
        Exit Function
    End If

    replacement = "前述的"
    ClaimReferenceStartCompat = keywordPos
End Function

Private Function FindClaimReferenceEndCompat(ByVal text As String, ByVal keywordPos As Long) As Long
    Dim position As Long
    Dim startPos As Long

    position = keywordPos + Len("权利要求")
    SkipInlineSpaces text, position
    startPos = position
    Do While position <= Len(text) And IsClaimDigitCompat(Mid$(text, position, 1))
        position = position + 1
    Loop
    If position = startPos Then Exit Function

    SkipInlineSpaces text, position
    If position <= Len(text) And IsClaimRangeCharCompat(Mid$(text, position, 1)) Then
        position = position + 1
        SkipInlineSpaces text, position
        startPos = position
        Do While position <= Len(text) And IsClaimDigitCompat(Mid$(text, position, 1))
            position = position + 1
        Loop
        If position = startPos Then Exit Function
    End If

    SkipInlineSpaces text, position
    If Mid$(text, position, 4) = "中任一项" Then
        position = position + 4
    ElseIf Mid$(text, position, 3) = "任一项" Then
        position = position + 3
    End If

    Do While position <= Len(text)
        If Mid$(text, position, 2) = "所述" Then
            FindClaimReferenceEndCompat = position + 2
            Exit Function
        End If
        If InStr("，,。；;：:", Mid$(text, position, 1)) > 0 Or _
           Mid$(text, position, 1) = vbCr Or Mid$(text, position, 1) = vbLf Then
            Exit Function
        End If
        position = position + 1
    Loop
End Function

Private Function FindClaimReferencePrefixEndCompat(ByVal text As String, ByVal keywordPos As Long) As Long
    Dim position As Long
    Dim startPos As Long

    position = keywordPos + Len("权利要求")
    SkipInlineSpaces text, position
    startPos = position
    Do While position <= Len(text) And IsClaimDigitCompat(Mid$(text, position, 1))
        position = position + 1
    Loop
    If position = startPos Then Exit Function

    SkipInlineSpaces text, position
    If position <= Len(text) And IsClaimRangeCharCompat(Mid$(text, position, 1)) Then
        position = position + 1
        SkipInlineSpaces text, position
        startPos = position
        Do While position <= Len(text) And IsClaimDigitCompat(Mid$(text, position, 1))
            position = position + 1
        Loop
        If position = startPos Then Exit Function
    End If

    SkipInlineSpaces text, position
    If Mid$(text, position, 4) = "中任一项" Then
        position = position + 4
    ElseIf Mid$(text, position, 3) = "任一项" Then
        position = position + 3
    End If

    Do While position <= Len(text)
        If Mid$(text, position, 2) = "所述" Then
            FindClaimReferencePrefixEndCompat = position + 2
            Exit Function
        End If
        If InStr("，,。；;：:", Mid$(text, position, 1)) > 0 Or _
           Mid$(text, position, 1) = vbCr Or Mid$(text, position, 1) = vbLf Then
            FindClaimReferencePrefixEndCompat = position + 1
            Exit Function
        End If
        position = position + 1
    Loop
End Function

Private Function IsClaimRangeCharCompat(ByVal ch As String) As Boolean
    IsClaimRangeCharCompat = (InStr("至到-－—~～", ch) > 0)
End Function

Private Sub SkipInlineSpacesBackCompat(ByVal text As String, ByRef position As Long)
    Do While position > 0
        If Mid$(text, position, 1) = " " Or Mid$(text, position, 1) = vbTab Or _
           Mid$(text, position, 1) = ChrW(160) Or Mid$(text, position, 1) = ChrW(12288) Then
            position = position - 1
        Else
            Exit Do
        End If
    Loop
End Sub

Private Function FindDrawingNumberCloseCompat(ByVal text As String, ByVal position As Long) As Long
    Dim candidateStart As Long
    candidateStart = position

    Do While position <= Len(text) And IsAsciiAlphaNumericCompat(Mid$(text, position, 1))
        position = position + 1
    Loop

    If position > candidateStart Then
        If Mid$(text, position, 1) = ")" Or Mid$(text, position, 1) = "）" Then
            FindDrawingNumberCloseCompat = position
        End If
    End If
End Function

Private Function IsAsciiAlphaNumericStringCompat(ByVal text As String) As Boolean
    Dim position As Long
    If Len(text) = 0 Then Exit Function

    For position = 1 To Len(text)
        If Not IsAsciiAlphaNumericCompat(Mid$(text, position, 1)) Then Exit Function
    Next position
    IsAsciiAlphaNumericStringCompat = True
End Function

Private Function IsAsciiAlphaNumericCompat(ByVal ch As String) As Boolean
    IsAsciiAlphaNumericCompat = IsAsciiDigit(ch) Or IsAsciiLetter(ch)
End Function

Private Function IsCaseNoCharCompat(ByVal ch As String) As Boolean
    IsCaseNoCharCompat = IsAsciiAlphaNumericCompat(ch) Or ch = "_" Or ch = "-"
End Function
