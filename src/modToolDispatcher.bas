Attribute VB_Name = "modToolDispatcher"
Option Explicit
Option Private Module

' 面板使用稳定的内部编号，不直接依赖按钮名称。
Public Function RunPatentTool(ByVal toolId As String) As Boolean
    On Error GoTo ErrorHandler

    Select Case LCase$(Trim$(toolId))
        Case "claim-figure-tag"
            Call modClaimFigureTag.权利要求标号

        Case "claim-to-invention"
            Call modClaimToInvention.权利要求转发明内容

        Case "embodiment-figure-tag"
            Call modEmbodimentFigureTag.具体实施方式标号

        Case "header-sync"
            Call modHeaderSync.页眉双模式同步

        Case "clean-format-full"
            Call modCleanFormatFull.一键改格式_全文修改_不兼容修订

        Case "clean-format-selection"
            Call modCleanFormatTrackLite.一键改格式_选中部分修改_兼容修订

        Case Else
            Err.Raise vbObjectError + 5100, "modToolDispatcher", _
                      "未知的工具编号：" & toolId
    End Select

    RunPatentTool = True
    Exit Function

ErrorHandler:
    RunPatentTool = False
    MsgBox "工具执行失败：" & Err.Description, vbCritical, "专利撰写工具箱"
End Function
