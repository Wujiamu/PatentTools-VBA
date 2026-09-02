Attribute VB_Name = "modCleanFormatTrackLite"
' =================================================================
' 一键改格式：修订兼容轻量版
' 在修订模式下使用；要求先选中文本。
' 在选区内逐项处理，避免整段重写导致字号被统一。
' =================================================================

Option Explicit

Public Sub 一键改格式_选中部分修改_兼容修订()
    On Error GoTo ErrorHandler

    Dim rng As Range
    Set rng = Selection.Range

    If rng.Start = rng.End Then
        MsgBox "请先选中需要清理格式的文本。修订兼容轻量版不建议直接处理全文。", vbExclamation
        Exit Sub
    End If

    Dim oldScreenUpdating As Boolean
    oldScreenUpdating = Application.ScreenUpdating

    Application.ScreenUpdating = False
    Application.StatusBar = "正在清理选中文本..."

    Application.StatusBar = "1/8 清理隐形字符..."
    RemoveInvisibleCharacters rng

    Application.StatusBar = "2/8 反单引号标准化..."
    NormalizeBacktickPairs rng

    Application.StatusBar = "3/8 清理 Markdown 符号..."
    RemoveMarkdownMarksInRange rng

    Application.StatusBar = "4/8 标点标准化..."
    NormalizePunctuationInRange rng

    Application.StatusBar = "5/8 单引号标准化..."
    NormalizeSingleQuotesInRange rng

    Application.StatusBar = "6/8 双引号标准化..."
    NormalizeDoubleQuotesInRange rng

    Application.StatusBar = "7/8 中文句号标准化..."
    FixChinesePeriodInRange rng

    Application.StatusBar = "8/8 字体标准化..."
    StandardizeFonts rng

    Application.StatusBar = False
    Application.ScreenUpdating = oldScreenUpdating

    MsgBox "选中文本格式清理完成。", vbInformation
    Exit Sub

ErrorHandler:
    On Error Resume Next
    Application.StatusBar = False
    Application.ScreenUpdating = oldScreenUpdating
    MsgBox "运行错误：" & Err.Description, vbCritical
End Sub

Private Sub StandardizeFonts(ByVal rng As Range)
    With rng.Font
        .name = "Times New Roman"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .NameFarEast = "楷体"
    End With
End Sub

Private Sub RemoveInvisibleCharacters(ByVal baseRange As Range)
    Dim arr As Variant
    arr = Array(ChrW(160), ChrW(&H200B), ChrW(&H200C), ChrW(&H200D), ChrW(&HFEFF))
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        ReplaceInRange baseRange, CStr(arr(i)), ""
    Next i
End Sub

Private Sub NormalizeBacktickPairs(ByVal baseRange As Range)
    Dim total As Long
    total = CountTextInRange(baseRange, "`")
    If total < 2 Then Exit Sub

    Dim rng As Range
    Set rng = baseRange.Duplicate

    Dim counter As Long
    counter = 0

    Dim wasTracking As Boolean
    wasTracking = baseRange.Document.TrackRevisions
    baseRange.Document.TrackRevisions = False

    With rng.Find
        .ClearFormatting
        .text = "`"
        .Format = False
        .MatchWildcards = False
        .Forward = True
        .Wrap = wdFindStop
    End With

    Do While rng.Find.Execute
        If rng.Start < baseRange.Start Or rng.End > baseRange.End Then Exit Do
        counter = counter + 1

        If counter = total And total Mod 2 = 1 Then
            rng.Collapse wdCollapseEnd
        ElseIf counter Mod 2 = 1 Then
            rng.text = ChrW(&H2018)
            rng.Collapse wdCollapseEnd
        Else
            rng.text = ChrW(&H2019)
            rng.Collapse wdCollapseEnd
        End If
    Loop

    baseRange.Document.TrackRevisions = wasTracking
End Sub

Private Sub RemoveMarkdownMarksInRange(ByVal baseRange As Range)
    ReplaceInRange baseRange, "**", ""
    ReplaceInRange baseRange, "*", ""
    ReplaceInRange baseRange, "__", ""
    ReplaceInRange baseRange, "_", ""

    Dim para As Paragraph
    For Each para In baseRange.Paragraphs
        CleanMarkdownAtParagraphStart para, baseRange
    Next para
End Sub

Private Sub CleanMarkdownAtParagraphStart(ByVal para As Paragraph, ByVal limitRange As Range)
    Dim r As Range
    Set r = para.Range.Duplicate

    If r.Start < limitRange.Start Then r.Start = limitRange.Start
    If r.End > limitRange.End Then r.End = limitRange.End
    If r.End > r.Start Then r.End = r.End - 1
    If r.End <= r.Start Then Exit Sub

    Dim originalText As String
    originalText = r.text

    Dim cleaned As String
    cleaned = TrimLeadingMarkdown(originalText)

    If cleaned <> originalText Then r.text = cleaned
End Sub

Private Function TrimLeadingMarkdown(ByVal text As String) As String
    Dim t As String
    t = text

    Do While Len(t) > 0 And Left(t, 1) = "#"
        t = Mid(t, 2)
    Loop

    Do While Len(t) > 0 And (Left(t, 1) = " " Or Left(t, 1) = vbTab)
        t = Mid(t, 2)
    Loop

    If Left(t, 2) = "- " Or Left(t, 2) = "+ " Then
        t = Mid(t, 3)
    End If

    TrimLeadingMarkdown = t
End Function

Private Sub NormalizePunctuationInRange(ByVal baseRange As Range)
    Dim arrFind As Variant
    Dim arrRepl As Variant

    arrFind = Array(",", ";", ":", "?", "!", "(", ")", "<", ">", "[", "]", "/", "@", "#", "$", "%", "&", "+", "=", "|", "~", "^", "{", "}", "-")
    arrRepl = Array("，", "；", "：", "？", "！", "（", "）", "《", "》", "［", "］", "／", "＠", "＃", "＄", "％", "＆", "＋", "＝", "｜", "～", "＾", "｛", "｝", "－")

    Dim i As Long
    For i = LBound(arrFind) To UBound(arrFind)
        ReplaceInRange baseRange, CStr(arrFind(i)), CStr(arrRepl(i))
    Next i
End Sub

Private Sub NormalizeSingleQuotesInRange(ByVal baseRange As Range)
    Dim rng As Range
    Set rng = baseRange.Duplicate

    Dim counter As Long
    counter = 0

    Dim wasTracking As Boolean
    wasTracking = baseRange.Document.TrackRevisions
    baseRange.Document.TrackRevisions = False

    With rng.Find
        .ClearFormatting
        .text = "'"
        .Format = False
        .MatchWildcards = False
        .Forward = True
        .Wrap = wdFindStop
    End With

    Do While rng.Find.Execute
        If rng.Start < baseRange.Start Or rng.End > baseRange.End Then Exit Do
        counter = counter + 1
        If counter Mod 2 = 1 Then
            rng.text = ChrW(&H2018)
        Else
            rng.text = ChrW(&H2019)
        End If
        rng.Collapse wdCollapseEnd
    Loop

    baseRange.Document.TrackRevisions = wasTracking
End Sub

Private Sub NormalizeDoubleQuotesInRange(ByVal baseRange As Range)
    Dim rng As Range
    Set rng = baseRange.Duplicate

    Dim counter As Long
    counter = 0

    Dim wasTracking As Boolean
    wasTracking = baseRange.Document.TrackRevisions
    baseRange.Document.TrackRevisions = False

    With rng.Find
        .ClearFormatting
        .text = """"
        .Format = False
        .MatchWildcards = False
        .Forward = True
        .Wrap = wdFindStop
    End With

    Do While rng.Find.Execute
        If rng.Start < baseRange.Start Or rng.End > baseRange.End Then Exit Do
        counter = counter + 1
        If counter Mod 2 = 1 Then
            rng.text = ChrW(&H201C)
        Else
            rng.text = ChrW(&H201D)
        End If
        rng.Collapse wdCollapseEnd
    Loop

    baseRange.Document.TrackRevisions = wasTracking
End Sub

Private Sub FixChinesePeriodInRange(ByVal baseRange As Range)
    Dim rng As Range
    Set rng = baseRange.Duplicate

    With rng.Find
        .ClearFormatting
        .text = "."
        .Format = False
        .MatchWildcards = False
        .Forward = True
        .Wrap = wdFindStop
    End With

    Do While rng.Find.Execute
        If rng.Start < baseRange.Start Or rng.End > baseRange.End Then Exit Do
        If rng.Start > 0 Then
            Dim beforeCh As String
            beforeCh = rng.Document.Range(rng.Start - 1, rng.Start).text
            If IsChineseChar(beforeCh) Then rng.text = "。"
        End If
        rng.Collapse wdCollapseEnd
    Loop
End Sub

Private Sub ReplaceInRange(ByVal baseRange As Range, ByVal findText As String, ByVal replText As String)
    Dim rng As Range
    Set rng = baseRange.Duplicate

    With rng.Find
        .ClearFormatting
        .replacement.ClearFormatting
        .text = findText
        .replacement.text = replText
        .Format = False
        .MatchWildcards = False
        .Forward = True
        .Wrap = wdFindStop
        .Execute Replace:=wdReplaceAll
    End With
End Sub

Private Function CountTextInRange(ByVal baseRange As Range, ByVal findText As String) As Long
    Dim rng As Range
    Set rng = baseRange.Duplicate

    With rng.Find
        .ClearFormatting
        .text = findText
        .Format = False
        .MatchWildcards = False
        .Forward = True
        .Wrap = wdFindStop
    End With

    Do While rng.Find.Execute
        If rng.Start < baseRange.Start Or rng.End > baseRange.End Then Exit Do
        CountTextInRange = CountTextInRange + 1
        rng.Collapse wdCollapseEnd
    Loop
End Function

Private Function IsChineseChar(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function

    Dim n As Long
    n = AscW(Left(ch, 1))
    If n < 0 Then n = n + 65536

    IsChineseChar = (n >= &H4E00 And n <= &H9FFF)
End Function
