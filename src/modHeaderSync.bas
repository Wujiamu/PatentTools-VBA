Attribute VB_Name = "modHeaderSync"
' =================================================================
' 专利工具：页眉双模式同步（修正版）
' 智能模式：从文件名提取案号，替换各页眉第一行。
' 克隆模式：以首页页眉第一行为准，同步至全文分节。
' =================================================================

Option Explicit

Sub 页眉双模式同步()
    Dim doc As Document
    Set doc = ActiveDocument

    Dim oldTrackRevisions As Boolean
    oldTrackRevisions = doc.TrackRevisions

    On Error GoTo ErrorHandler

    Dim response As VbMsgBoxResult
    response = MsgBox("请选择页眉处理模式：" & vbCrLf & vbCrLf & _
                      "【是 (Yes)】 智能模式：从文件名提取案号，替换各页眉第一行。" & vbCrLf & _
                      "【否 (No)】  克隆模式：以首页页眉第一行为准，同步至全文分节。" & vbCrLf & _
                      "【取消】 退出。", _
                      vbYesNoCancel + vbQuestion, "专利页眉标准化工具")

    If response = vbCancel Then GoTo CleanExit

    Dim targetCaseNo As String
    If response = vbNo Then
        Dim baseHeaderRange As Range
        Set baseHeaderRange = GetBaseHeaderRange(doc)
        targetCaseNo = ExtractFirstLine(baseHeaderRange.text)

        If Len(Trim(targetCaseNo)) = 0 Then
            MsgBox "第一页页眉第一行（案号）为空，无法克隆。", vbExclamation
            GoTo CleanExit
        End If
    Else
        If doc.Path = "" Then
            MsgBox "文件尚未保存，无法从文件名提取案号。", vbCritical
            GoTo CleanExit
        End If

        Dim baseName As String
        If InStrRev(doc.name, ".") > 0 Then
            baseName = Left(doc.name, InStrRev(doc.name, ".") - 1)
        Else
            baseName = doc.name
        End If

        targetCaseNo = ExtractCaseNoFromFileName(baseName)
        If Len(Trim(targetCaseNo)) = 0 Then
            MsgBox "无法从文件名提取案号，请检查文件名格式。", vbExclamation
            GoTo CleanExit
        End If
    End If

    doc.TrackRevisions = True

    Dim sec As Section
    Dim hf As HeaderFooter
    Dim modifiedCount As Long
    modifiedCount = 0

    For Each sec In doc.Sections
        For Each hf In sec.Headers
            If hf.Exists Then
                If UpdateHeaderFirstLine(hf, targetCaseNo) Then modifiedCount = modifiedCount + 1
            End If
        Next hf
    Next sec

    MsgBox "处理完成！共更新 " & modifiedCount & " 处页眉。", vbInformation

CleanExit:
    doc.TrackRevisions = oldTrackRevisions
    Exit Sub

ErrorHandler:
    On Error Resume Next
    doc.TrackRevisions = oldTrackRevisions
    MsgBox "页眉处理失败：" & Err.Description, vbCritical
End Sub

Private Function GetBaseHeaderRange(ByVal doc As Document) As Range
    Dim firstHeader As HeaderFooter
    Dim primaryHeader As HeaderFooter

    Set firstHeader = doc.Sections(1).Headers(wdHeaderFooterFirstPage)
    Set primaryHeader = doc.Sections(1).Headers(wdHeaderFooterPrimary)

    If doc.Sections(1).PageSetup.DifferentFirstPageHeaderFooter Then
        If firstHeader.Exists Then
            If Not IsRangeEffectivelyEmpty(firstHeader.Range) Then
                Set GetBaseHeaderRange = firstHeader.Range
                Exit Function
            End If
        End If
    End If

    Set GetBaseHeaderRange = primaryHeader.Range
End Function

Private Function ExtractFirstLine(ByVal fullText As String) As String
    Dim result As String
    result = fullText
    result = Replace(result, Chr(7), "")
    result = Replace(result, vbLf, "")

    Dim pos As Long
    pos = InStr(result, vbCr)
    If pos > 0 Then result = Left(result, pos - 1)

    ExtractFirstLine = Trim(Replace(result, vbTab, ""))
End Function

Private Function ExtractCaseNoFromFileName(ByVal fileName As String) As String
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = False
    regEx.IgnoreCase = True
    regEx.pattern = "[A-Za-z0-9][A-Za-z0-9_\-]*[A-Za-z0-9]"

    If regEx.Test(fileName) Then
        ExtractCaseNoFromFileName = regEx.Execute(fileName)(0).value
    Else
        ExtractCaseNoFromFileName = ""
    End If
End Function

Private Function IsRangeEffectivelyEmpty(ByVal rng As Range) As Boolean
    Dim tempText As String
    tempText = rng.text
    tempText = Replace(tempText, vbCr, "")
    tempText = Replace(tempText, vbLf, "")
    tempText = Replace(tempText, vbTab, "")
    tempText = Replace(tempText, Chr(7), "")
    tempText = Replace(tempText, ChrW(160), "")

    IsRangeEffectivelyEmpty = (Len(Trim(tempText)) = 0)
End Function

Private Function UpdateHeaderFirstLine(ByVal hf As HeaderFooter, ByVal newCaseNo As String) As Boolean
    Dim hRange As Range
    Set hRange = hf.Range.Duplicate

    If hRange.Paragraphs.Count = 0 Then
        hRange.text = newCaseNo
        hRange.ParagraphFormat.Alignment = wdAlignParagraphRight
        UpdateHeaderFirstLine = True
        Exit Function
    End If

    Dim firstPara As Paragraph
    Set firstPara = hRange.Paragraphs(1)

    Dim paraRange As Range
    Set paraRange = firstPara.Range.Duplicate
    If paraRange.End > paraRange.Start Then paraRange.End = paraRange.End - 1

    paraRange.text = newCaseNo
    firstPara.Range.ParagraphFormat.Alignment = wdAlignParagraphRight

    UpdateHeaderFirstLine = True
End Function
