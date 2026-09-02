Attribute VB_Name = "modCleanFormatFull"
' =================================================================
' 一键改格式：全文无修订版
' 功能完整；运行时临时关闭修订模式，处理完成后恢复原状态。
' 适合定稿前对全文做格式清理。
' =================================================================

Option Explicit

Public Sub 一键改格式_全文修改_不兼容修订()
    On Error GoTo ErrorHandler

    Dim doc As Document
    Set doc = ActiveDocument

    Dim oldTrackRevisions As Boolean
    oldTrackRevisions = doc.TrackRevisions

    Dim oldScreenUpdating As Boolean
    oldScreenUpdating = Application.ScreenUpdating

    Application.ScreenUpdating = False
    Application.StatusBar = "正在进行全文格式清理..."

    doc.TrackRevisions = False

    Dim rng As Range
    Set rng = doc.Content

    Application.StatusBar = "1/7 字体标准化..."
    StandardizeFonts rng

    Application.StatusBar = "2/7 清理隐形字符..."
    RemoveInvisibleCharacters rng

    Application.StatusBar = "3/7 反单引号标准化..."
    NormalizeBacktickPairs doc

    Application.StatusBar = "4/7 清理 Markdown 符号..."
    RemoveMarkdownMarksInDocument doc

    Application.StatusBar = "5/7 标点标准化..."
    NormalizePunctuationInDocument doc

    Application.StatusBar = "6/7 双引号标准化..."
    NormalizeDoubleQuotes doc

    Application.StatusBar = "7/7 中文句号标准化..."
    FixChinesePeriodInDocument doc

    doc.TrackRevisions = oldTrackRevisions
    Application.StatusBar = False
    Application.ScreenUpdating = oldScreenUpdating

    MsgBox "全文格式清理完成。", vbInformation
    Exit Sub

ErrorHandler:
    On Error Resume Next
    doc.TrackRevisions = oldTrackRevisions
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

Private Sub RemoveInvisibleCharacters(ByVal rng As Range)
    Dim arr As Variant
    arr = Array(ChrW(160), ChrW(&H200B), ChrW(&H200C), ChrW(&H200D), ChrW(&HFEFF))

    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        ReplaceInRange rng, CStr(arr(i)), ""
    Next i
End Sub

Private Sub RemoveMarkdownMarksInDocument(ByVal doc As Document)
    ReplaceInRange doc.Content, "**", ""
    ReplaceInRange doc.Content, "*", ""
    ReplaceInRange doc.Content, "__", ""
    ReplaceInRange doc.Content, "_", ""

    Dim para As Paragraph
    For Each para In doc.Paragraphs
        CleanMarkdownAtParagraphStart para
    Next para
End Sub

Private Sub CleanMarkdownAtParagraphStart(ByVal para As Paragraph)
    Dim r As Range
    Set r = para.Range.Duplicate
    If r.End > r.Start Then r.End = r.End - 1

    Dim text As String
    text = r.text

    Dim cleaned As String
    cleaned = TrimLeadingMarkdown(text)

    If cleaned <> text Then r.text = cleaned
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

Private Sub NormalizePunctuationInDocument(ByVal doc As Document)
    Dim arrFind As Variant
    Dim arrRepl As Variant

    arrFind = Array(",", ";", ":", "?", "!", "'", "(", ")", "<", ">", "[", "]", "/", "@", "#", "$", "%", "&", "+", "=", "|", "~", "^", "{", "}", "-")
    arrRepl = Array("，", "；", "：", "？", "！", "’", "（", "）", "《", "》", "［", "］", "／", "＠", "＃", "＄", "％", "＆", "＋", "＝", "｜", "～", "＾", "｛", "｝", "－")

    Dim i As Long
    For i = LBound(arrFind) To UBound(arrFind)
        ReplaceInRange doc.Content, CStr(arrFind(i)), CStr(arrRepl(i))
    Next i
End Sub

Private Sub NormalizeBacktickPairs(ByVal doc As Document)
    Dim total As Long
    total = CountTextInRange(doc.Content, "`")
    If total < 2 Then Exit Sub

    Dim rng As Range
    Set rng = doc.Content.Duplicate

    Dim counter As Long
    counter = 0

    With rng.Find
        .ClearFormatting
        .text = "`"
        .Format = False
        .MatchWildcards = False
        .Forward = True
        .Wrap = wdFindStop
    End With

    Do While rng.Find.Execute
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
End Sub

Private Sub NormalizeDoubleQuotes(ByVal doc As Document)
    Dim rng As Range
    Set rng = doc.Content.Duplicate

    Dim counter As Long
    counter = 0

    With rng.Find
        .ClearFormatting
        .text = """"
        .Format = False
        .MatchWildcards = False
        .Forward = True
        .Wrap = wdFindStop
    End With

    Do While rng.Find.Execute
        counter = counter + 1
        If counter Mod 2 = 1 Then
            rng.text = ChrW(&H201C)
        Else
            rng.text = ChrW(&H201D)
        End If
        rng.Collapse wdCollapseEnd
    Loop
End Sub

Private Sub FixChinesePeriodInDocument(ByVal doc As Document)
    Dim rng As Range
    Set rng = doc.Content.Duplicate

    With rng.Find
        .ClearFormatting
        .text = "."
        .Format = False
        .MatchWildcards = False
        .Forward = True
        .Wrap = wdFindStop
    End With

    Do While rng.Find.Execute
        If rng.Start > 0 Then
            Dim beforeCh As String
            beforeCh = doc.Range(rng.Start - 1, rng.Start).text
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
