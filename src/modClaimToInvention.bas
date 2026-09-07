Attribute VB_Name = "modClaimToInvention"
' 专利代理人专用：权利要求转发明内容（修正版）
' 目标：将选中的权利要求改写为“发明内容”式表述，并尽量保留原段落结构。

Option Explicit
Option Private Module

Sub 权利要求转发明内容()
    On Error GoTo ErrorHandler

    Dim doc As Document
    Set doc = ActiveDocument

    Dim oldTrackRevisions As Boolean
    oldTrackRevisions = doc.TrackRevisions

    Dim selRange As Range
    Set selRange = Selection.Range

    If selRange.Start = selRange.End Then
        MsgBox "错误：请先选中需要转换的权利要求段落。", vbExclamation
        Exit Sub
    End If

    Dim undoStarted As Boolean
    undoStarted = BeginCustomUndoRecord("权利要求转发明内容")
    doc.TrackRevisions = True

    Dim paraCount As Long
    paraCount = selRange.Paragraphs.Count

    Dim paraText() As String
    Dim paraRange() As Range
    ReDim paraText(1 To paraCount)
    ReDim paraRange(1 To paraCount)

    Dim i As Long
    For i = 1 To paraCount
        Set paraRange(i) = selRange.Paragraphs(i).Range.Duplicate
        paraText(i) = paraRange(i).Text
    Next i

    Dim claimIndexForPara() As Long
    ReDim claimIndexForPara(1 To paraCount)

    Dim claimStartPara() As Long
    Dim claimIsDependent() As Boolean
    Dim claimNumForClaim() As Long
    Dim claimCount As Long
    claimCount = 0

    Dim indepDict As Object
    Set indepDict = CreateStringMap()

    Dim currentClaimIdx As Long
    currentClaimIdx = 0

    Dim testText As String
    Dim claimNum As Long
    Dim tmp As String

    For i = 1 To paraCount
        testText = CleanParaText(paraText(i))

        If IsStopHeadingLine(testText) Then
            claimIndexForPara(i) = 0
            currentClaimIdx = 0

        ElseIf IsValidClaimLine(testText) Then
            currentClaimIdx = currentClaimIdx + 1
            claimCount = currentClaimIdx

            ReDim Preserve claimStartPara(1 To claimCount)
            ReDim Preserve claimIsDependent(1 To claimCount)
            ReDim Preserve claimNumForClaim(1 To claimCount)

            claimStartPara(claimCount) = i
            claimIndexForPara(i) = claimCount

            claimNum = ClaimNumberFromText(testText)
            claimNumForClaim(claimCount) = claimNum

            tmp = RemoveClaimNumber(testText)
            claimIsDependent(claimCount) = IsDependentClaimStart(tmp)

            If Not claimIsDependent(claimCount) Then
                If claimNum > 0 Then indepDict.Item(CStr(claimNum)) = True
            End If

        ElseIf IsMeaningfulLine(testText) Then
            If currentClaimIdx > 0 Then
                claimIndexForPara(i) = currentClaimIdx
            Else
                claimIndexForPara(i) = 0
            End If

        Else
            claimIndexForPara(i) = 0
        End If
    Next i

    Dim prefixMap As Object
    Set prefixMap = BuildIndependentPrefixMap(indepDict)

    Dim pRange As Range
    Dim newText As String
    Dim idx As Long
    Dim skippedRevisionCount As Long

    For i = paraCount To 1 Step -1
        idx = claimIndexForPara(i)
        If idx = 0 Then GoTo NextPara

        If RangeHasExistingRevision(paraRange(i)) Then
            skippedRevisionCount = skippedRevisionCount + 1
            GoTo NextPara
        End If

        newText = RewriteClaimParagraph( _
            paraText(i), _
            i = claimStartPara(idx), _
            claimIsDependent(idx), _
            claimNumForClaim(idx), _
            prefixMap)

        Set pRange = paraRange(i).Duplicate
        If pRange.End > pRange.Start Then pRange.End = pRange.End - 1
        pRange.Text = newText

NextPara:
    Next i

    EndCustomUndoRecord undoStarted
    doc.TrackRevisions = oldTrackRevisions
    If skippedRevisionCount > 0 Then
        MsgBox "完成。已跳过 " & skippedRevisionCount & " 个含既有修订的段落，请人工转换。", vbExclamation
    Else
        MsgBox "完成：权利要求已转换为发明内容表述。", vbInformation
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

Private Function RangeHasExistingRevision(ByVal rng As Range) As Boolean
    On Error GoTo ConservativeFallback
    RangeHasExistingRevision = (rng.Revisions.Count > 0)
    Exit Function

ConservativeFallback:
    RangeHasExistingRevision = True
End Function

Private Function RewriteClaimParagraph( _
    ByVal src As String, _
    ByVal isFirstPara As Boolean, _
    ByVal isDependent As Boolean, _
    ByVal claimNum As Long, _
    ByVal prefixMap As Object) As String

    Dim text As String
    text = CleanParaText(src)

    If isFirstPara Then
        text = RemoveClaimNumber(text)

        If isDependent Then
            text = TextAfterFeatureMarker(text)
        Else
            text = RewriteIndependentFeatureMarker(text)
        End If
    End If

    text = RemoveDrawingNumbers(text)
    text = ReplaceClaimReferences(text)
    text = Replace(text, "为前述的", "前述的")
    text = Replace(text, "所述", "")

    text = RemoveAllSpaces(text)
    text = TrimLeadingPunctuation(text)
    text = FixEndingPunctuation(text)

    If isFirstPara Then
        If isDependent Then
            text = "在一些实施例中，" & text
        ElseIf prefixMap.Exists(CStr(claimNum)) Then
            text = CStr(prefixMap.Item(CStr(claimNum))) & text
        Else
            text = "本发明提供" & text
        End If
    End If

    RewriteClaimParagraph = text
End Function

Private Function RewriteIndependentFeatureMarker(ByVal text As String) As String
    If InStr(text, "用于") > 0 Then
        RewriteIndependentFeatureMarker = text
        Exit Function
    End If

    text = Replace(text, "，其特征在于，包括", "，其包括")
    text = Replace(text, "，其特征在于：包括", "，其包括")
    text = Replace(text, "，其特征在于，", "，")
    text = Replace(text, "，其特征在于：", "，")
    text = Replace(text, "，其特征在于", "，")
    text = Replace(text, "其特征在于，", "")
    text = Replace(text, "其特征在于：", "")
    text = Replace(text, "其特征在于", "")

    RewriteIndependentFeatureMarker = text
End Function

Private Function TextAfterFeatureMarker(ByVal text As String) As String
    Dim marker As Variant
    Dim markers As Variant
    markers = Array("其特征在于，", "其特征在于：", "其特征在于")

    For Each marker In markers
        If InStr(text, CStr(marker)) > 0 Then
            TextAfterFeatureMarker = Mid(text, InStr(text, CStr(marker)) + Len(CStr(marker)))
            Exit Function
        End If
    Next marker

    TextAfterFeatureMarker = StripClaimReferencePrefixCompat(text)
End Function

Private Function ReplaceClaimReferences(ByVal text As String) As String
    ReplaceClaimReferences = ReplaceClaimReferencesCompat(text)
End Function

Private Function BuildIndependentPrefixMap(ByVal indepDict As Object) As Object
    Dim result As Object
    Set result = CreateStringMap()

    If indepDict.Count = 0 Then
        Set BuildIndependentPrefixMap = result
        Exit Function
    End If

    Dim keys As Variant
    keys = indepDict.Keys

    Dim i As Long, j As Long
    Dim t As Variant
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If CLng(keys(i)) > CLng(keys(j)) Then
                t = keys(i)
                keys(i) = keys(j)
                keys(j) = t
            End If
        Next j
    Next i

    For i = LBound(keys) To UBound(keys)
        Select Case i - LBound(keys) + 1
            Case 1
                result.Item(CStr(keys(i))) = "本发明提供"
            Case 2
                result.Item(CStr(keys(i))) = "本发明还提供"
            Case Else
                result.Item(CStr(keys(i))) = "本发明另提供"
        End Select
    Next i

    Set BuildIndependentPrefixMap = result
End Function

Private Function CleanParaText(ByVal text As String) As String
    text = Replace(text, vbCr, "")
    text = Replace(text, vbLf, "")
    text = Replace(text, ChrW(160), "")
    text = Replace(text, ChrW(12288), "")
    CleanParaText = Trim(text)
End Function

Private Function RemoveAllSpaces(ByVal text As String) As String
    text = Replace(text, " ", "")
    text = Replace(text, vbTab, "")
    text = Replace(text, ChrW(160), "")
    text = Replace(text, ChrW(12288), "")
    RemoveAllSpaces = text
End Function

Private Function IsValidClaimLine(ByVal text As String) As Boolean
    text = CleanParaText(text)
    If text = "" Then Exit Function
    IsValidClaimLine = IsNumberedClaimLineCompat(text)
End Function

Private Function ClaimNumberFromText(ByVal text As String) As Long
    ClaimNumberFromText = LeadingClaimNumberCompat(text)
End Function

Private Function RemoveClaimNumber(ByVal text As String) As String
    RemoveClaimNumber = RemoveLeadingClaimNumberCompat(text)
End Function

Private Function IsDependentClaimStart(ByVal text As String) As Boolean
    text = Trim(text)
    text = Replace(text, " ", "")
    text = Replace(text, vbTab, "")
    text = Replace(text, ChrW(160), "")
    text = Replace(text, ChrW(12288), "")
    IsDependentClaimStart = (Left$(text, 5) = "如权利要求") Or (Left$(text, 6) = "根据权利要求")
End Function

Private Function IsMeaningfulLine(ByVal text As String) As Boolean
    text = CleanParaText(text)
    If text = "" Then Exit Function

    text = Replace(text, "，", "")
    text = Replace(text, "。", "")
    text = Replace(text, "；", "")
    text = Replace(text, "：", "")
    text = Replace(text, ",", "")
    text = Replace(text, ".", "")
    text = Replace(text, ";", "")
    text = Replace(text, ":", "")

    IsMeaningfulLine = (Len(Trim(text)) > 0)
End Function

Private Function IsStopHeadingLine(ByVal text As String) As Boolean
    text = CleanParaText(text)
    IsStopHeadingLine = (InStr(text, "转换后的发明内容") > 0)
End Function

Private Function RemoveDrawingNumbers(ByVal text As String) As String
    RemoveDrawingNumbers = RemoveDrawingNumbersCompat(text)
End Function

Private Function TrimLeadingPunctuation(ByVal text As String) As String
    text = Trim(text)
    Do While Len(text) > 0
        If InStr("，,。；;：:", Left(text, 1)) > 0 Then
            text = Trim(Mid(text, 2))
        Else
            Exit Do
        End If
    Loop
    TrimLeadingPunctuation = text
End Function

Private Function FixEndingPunctuation(ByVal text As String) As String
    text = Trim(text)
    If text = "" Then
        FixEndingPunctuation = text
        Exit Function
    End If

    If InStr("。；;：:！？?!", Right(text, 1)) = 0 Then
        text = text & "。"
    End If

    FixEndingPunctuation = text
End Function
