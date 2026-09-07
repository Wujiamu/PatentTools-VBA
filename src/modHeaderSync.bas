Attribute VB_Name = "modHeaderSync"
' =================================================================
' 专利工具：页眉双模式同步（修正版）
' 智能模式：从文件名提取案号，替换各页眉第一行。
' 克隆模式：以首页页眉第一行为准，同步至全文分节。
' =================================================================

Option Explicit
Option Private Module

Sub 页眉双模式同步()
    Dim doc As Document
    Set doc = ActiveDocument

    Dim oldTrackRevisions As Boolean
    oldTrackRevisions = doc.TrackRevisions
    Dim undoStarted As Boolean
    Dim skippedRevisionCount As Long
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
        targetCaseNo = ExtractFirstLine(baseHeaderRange.Text)

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
        If InStrRev(doc.Name, ".") > 0 Then
            baseName = Left(doc.Name, InStrRev(doc.Name, ".") - 1)
        Else
            baseName = doc.Name
        End If

        targetCaseNo = ExtractCaseNoFromFileName(baseName)
        If Len(Trim(targetCaseNo)) = 0 Then
            MsgBox "无法从文件名提取案号，请检查文件名格式。", vbExclamation
            GoTo CleanExit
        End If
    End If

    undoStarted = BeginCustomUndoRecord("页眉双模式同步")
    doc.TrackRevisions = True

    Dim sec As Section
    Dim hf As HeaderFooter
    Dim modifiedCount As Long
    modifiedCount = 0

    For Each sec In doc.Sections
        For Each hf In sec.Headers
            If hf.Exists Then
                If RangeHasExistingRevision(hf.Range) Then
                    skippedRevisionCount = skippedRevisionCount + 1
                ElseIf UpdateHeaderFirstLine(hf, targetCaseNo) Then
                    modifiedCount = modifiedCount + 1
                End If
            End If
        Next hf
    Next sec

    If skippedRevisionCount > 0 Then
        MsgBox "处理完成！共更新 " & modifiedCount & " 处页眉；跳过 " & skippedRevisionCount & " 处含既有修订的页眉。", vbExclamation
    Else
        MsgBox "处理完成！共更新 " & modifiedCount & " 处页眉。", vbInformation
    End If

CleanExit:
    EndCustomUndoRecord undoStarted
    doc.TrackRevisions = oldTrackRevisions
    Exit Sub

ErrorHandler:
    On Error Resume Next
    EndCustomUndoRecord undoStarted
    doc.TrackRevisions = oldTrackRevisions
    MsgBox "页眉处理失败：" & Err.Description, vbCritical
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
    ExtractCaseNoFromFileName = ExtractCaseNoCompat(fileName)
End Function

Private Function IsRangeEffectivelyEmpty(ByVal rng As Range) As Boolean
    Dim tempText As String
    tempText = rng.Text
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
        hRange.Text = newCaseNo
        hRange.ParagraphFormat.Alignment = wdAlignParagraphRight
        UpdateHeaderFirstLine = True
        Exit Function
    End If

    Dim firstPara As Paragraph
    Set firstPara = hRange.Paragraphs(1)

    Dim paraRange As Range
    Set paraRange = firstPara.Range.Duplicate
    If paraRange.End > paraRange.Start Then paraRange.End = paraRange.End - 1

    paraRange.Text = newCaseNo
    firstPara.Range.ParagraphFormat.Alignment = wdAlignParagraphRight

    UpdateHeaderFirstLine = True
End Function
