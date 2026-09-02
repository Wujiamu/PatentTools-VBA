VERSION 5.00
Begin VB.UserForm frmPatentToolbox
   Caption         =   "专利撰写工具箱"
   ClientHeight    =   6900
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7800
   StartUpPosition =   1  'CenterOwner
   Begin MSForms.Label lblTitle
      Caption         =   "专利撰写工具箱"
      Height          =   360
      Left            =   240
      Top             =   180
      Width           =   7320
   End
   Begin MSForms.Label lblDocumentInfo
      Caption         =   "活动文档：未检测到 Word 文档    |    当前选区：不可用"
      Height          =   360
      Left            =   240
      Top             =   660
      Width           =   7320
   End
   Begin MSForms.Label lblHint
      Caption         =   "提示：需要选区的工具会在点击时读取当前 Word 选区；影响全文或修订的操作请先确认作用范围。"
      Height          =   480
      Left            =   240
      Top             =   960
      Width           =   7320
      WordWrap        =   True
   End
   Begin MSForms.Frame fraFigureTag
      Caption         =   "标号处理"
      Height          =   1260
      Left            =   180
      Top             =   1500
      Width           =   7440
      Begin MSForms.CommandButton cmdClaimFigureTag
         Caption         =   "权利要求标号"
         ControlTipText  =   "先选中需要处理的权利要求区域"
         Height          =   420
         Left            =   180
         TabIndex        =   0
         Top             =   300
         Width           =   3360
      End
      Begin MSForms.CommandButton cmdEmbodimentFigureTag
         Caption         =   "具体实施方式标号"
         ControlTipText  =   "先选中需要处理的具体实施方式区域"
         Height          =   420
         Left            =   3720
         TabIndex        =   1
         Top             =   300
         Width           =   3360
      End
   End
   Begin MSForms.Frame fraContent
      Caption         =   "内容转换"
      Height          =   900
      Left            =   180
      Top             =   2910
      Width           =   7440
      Begin MSForms.CommandButton cmdClaimToInvention
         Caption         =   "权利要求转发明内容"
         ControlTipText  =   "先选中需要转换的权利要求段落"
         Height          =   420
         Left            =   180
         TabIndex        =   0
         Top             =   300
         Width           =   7080
      End
   End
   Begin MSForms.Frame fraDocument
      Caption         =   "文档信息"
      Height          =   900
      Left            =   180
      Top             =   3930
      Width           =   7440
      Begin MSForms.CommandButton cmdHeaderSync
         Caption         =   "页眉双模式同步"
         ControlTipText  =   "根据文件名或首页页眉同步各节页眉"
         Height          =   420
         Left            =   180
         TabIndex        =   0
         Top             =   300
         Width           =   7080
      End
   End
   Begin MSForms.Frame fraFormat
      Caption         =   "格式清理"
      Height          =   1260
      Left            =   180
      Top             =   4950
      Width           =   7440
      Begin MSForms.CommandButton cmdCleanFormatFull
         Caption         =   "一键改格式：全文（不兼容修订）"
         ControlTipText  =   "处理全文并暂时关闭修订，执行前请确认"
         Height          =   420
         Left            =   180
         TabIndex        =   0
         Top             =   300
         Width           =   3360
      End
      Begin MSForms.CommandButton cmdCleanFormatSelection
         Caption         =   "一键改格式：选区（兼容修订）"
         ControlTipText  =   "仅处理当前选区并尽量兼容修订"
         Height          =   420
         Left            =   3720
         TabIndex        =   1
         Top             =   300
         Width           =   3360
      End
   End
   Begin MSForms.Label lblStatus
      Caption         =   "就绪。"
      Height          =   360
      Left            =   240
      Top             =   6390
      Width           =   4800
   End
   Begin MSForms.CommandButton cmdRefresh
      Caption         =   "刷新状态"
      Height          =   420
      Left            =   5220
      TabIndex        =   6
      Top             =   6330
      Width           =   1080
   End
   Begin MSForms.CommandButton cmdClose
      Cancel          =   True
      Caption         =   "关闭面板"
      Height          =   420
      Left            =   6420
      TabIndex        =   7
      Top             =   6330
      Width           =   1080
   End
End
Attribute VB_Name = "frmPatentToolbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mBusy As Boolean
Private mCloseRequested As Boolean

Private Sub UserForm_Initialize()
    Me.Caption = "专利撰写工具箱"
    lblStatus.Caption = "就绪。"
    UpdateDocumentInfo
End Sub

Private Sub UserForm_Activate()
    If Not mBusy Then UpdateDocumentInfo
End Sub

Private Sub cmdClaimFigureTag_Click()
    RunToolFromButton "claim-figure-tag"
End Sub

Private Sub cmdEmbodimentFigureTag_Click()
    RunToolFromButton "embodiment-figure-tag"
End Sub

Private Sub cmdClaimToInvention_Click()
    RunToolFromButton "claim-to-invention"
End Sub

Private Sub cmdHeaderSync_Click()
    RunToolFromButton "header-sync"
End Sub

Private Sub cmdCleanFormatFull_Click()
    RunToolFromButton "clean-format-full"
End Sub

Private Sub cmdCleanFormatSelection_Click()
    RunToolFromButton "clean-format-selection"
End Sub

Private Sub cmdRefresh_Click()
    UpdateDocumentInfo
    lblStatus.Caption = "状态已刷新。"
End Sub

Private Sub cmdClose_Click()
    mCloseRequested = True
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    mCloseRequested = True
End Sub

Private Sub RunToolFromButton(ByVal toolId As String)
    If mBusy Then Exit Sub

    mBusy = True
    On Error GoTo ErrorHandler

    UpdateDocumentInfo
    lblStatus.Caption = "正在执行，请稍候……"

    ' 隐藏面板，使 Word 文档和当前选区重新成为用户操作焦点。
    Me.Hide
    DoEvents

    Dim succeeded As Boolean
    succeeded = modToolDispatcher.RunPatentTool(toolId)

    If Not mCloseRequested Then
        UpdateDocumentInfo
        If succeeded Then
            lblStatus.Caption = "执行完成。"
        Else
            lblStatus.Caption = "执行失败，请根据提示检查文档。"
        End If
        Me.Show vbModeless
    End If

    mBusy = False
    Exit Sub

ErrorHandler:
    mBusy = False
    If Not mCloseRequested Then Me.Show vbModeless
    MsgBox "面板调用失败：" & Err.Description, vbCritical, "专利撰写工具箱"
End Sub

Private Sub UpdateDocumentInfo()
    On Error GoTo NoDocument

    If Documents.Count = 0 Then GoTo NoDocument

    Dim docName As String
    Dim selectionLength As Long

    docName = ActiveDocument.Name
    selectionLength = Selection.Range.End - Selection.Range.Start
    lblDocumentInfo.Caption = "活动文档：" & docName & "    |    当前选区：" & CStr(selectionLength) & " 个字符"
    Exit Sub

NoDocument:
    lblDocumentInfo.Caption = "活动文档：未检测到 Word 文档    |    当前选区：不可用"
End Sub
