Attribute VB_Name = "modToolboxEntry"
Option Explicit

' Word 宏列表中的唯一公开入口。
Public Sub 专利撰写工具箱()
    On Error GoTo ErrorHandler

    frmPatentToolbox.Show vbModeless
    Exit Sub

ErrorHandler:
    MsgBox "无法打开专利撰写工具箱：" & Err.Description, vbCritical, "专利撰写工具箱"
End Sub
