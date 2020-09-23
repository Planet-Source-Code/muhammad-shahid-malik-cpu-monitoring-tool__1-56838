VERSION 5.00
Begin VB.Form FrmLittleBarGraph 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CPU Usage Bar & Graph"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6015
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picUsage 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   0
      Picture         =   "FrmLittleBarGraph.frx":0000
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   1
      Top             =   0
      Width           =   2415
      Begin VB.Label lblCpuUsage 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   1560
         TabIndex        =   2
         Top             =   0
         Width           =   930
      End
      Begin VB.Line Line1 
         X1              =   72
         X2              =   16
         Y1              =   72
         Y2              =   72
      End
   End
   Begin VB.PictureBox picGraph 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00400040&
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   2400
      ScaleHeight     =   126.23
      ScaleMode       =   0  'User
      ScaleWidth      =   100
      TabIndex        =   0
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "FrmLittleBarGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private QueryObject As Object

Private Sub Form_Activate()
   If FrmMain.Check1.Value = 1 Then
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
   End If
End Sub

Private Sub Form_Load()

If FrmMain.Option3.Value = True Then
FrmMain.tmrRefresh.Interval = 500
FrmGraph.Timer1.Interval = 500
FrmLittleBar.Timer1.Interval = 500
FrmLittleBarGraph.Timer1.Interval = 500
FrmNumber.Timer1.Interval = 500
FrmBigBar.Timer1.Interval = 500
End If

If FrmMain.Option4.Value = True Then
FrmMain.tmrRefresh.Interval = 1000
FrmGraph.Timer1.Interval = 1000
FrmNumber.Timer1.Interval = 1000
FrmBigBar.Timer1.Interval = 1000
End If

If FrmMain.Option5.Value = True Then
FrmMain.tmrRefresh.Interval = 2000
FrmGraph.Timer1.Interval = 2000
FrmLittleBar.Timer1.Interval = 2000
FrmLittleBarGraph.Timer1.Interval = 2000
FrmNumber.Timer1.Interval = 2000
FrmBigBar.Timer1.Interval = 2000
End If

If FrmMain.Option11.Value = True Then
Me.Top = Screen.Height / 2 - Me.Height / 2
Me.Left = Screen.Width / 2 - Me.Width / 2
End If

If FrmMain.Option12.Value = True Then
Me.Top = 0
Me.Left = Screen.Width - Me.Width
End If

If FrmMain.Option13.Value = True Then
Me.Top = 0
Me.Left = 0
End If

If FrmMain.Option14.Value = True Then
Me.Top = 0
Me.Left = Screen.Width / 2 - Me.Width / 2
End If

If FrmMain.Option15.Value = True Then
Me.Top = Screen.Height - Me.Height - 500
Me.Left = Screen.Width - Me.Width
End If

If FrmMain.Option16.Value = True Then
Me.Top = Screen.Height - Me.Height - 500
Me.Left = 0
End If

If FrmMain.Option17.Value = True Then
Me.Top = Screen.Height - Me.Height - 500
Me.Left = Screen.Width / 2 - Me.Width / 2
End If

    If IsWinNT Then
        Set QueryObject = New clsCPUUsageNT
    Else
        Set QueryObject = New clsCPUUsage
    End If
    QueryObject.Initialize

End Sub

Private Sub Timer1_Timer()
   Dim Ret As Long
    Dim Which As Long
    Ret = QueryObject.Query
    If Ret = -1 Then
        Timer1.Enabled = False
        lblCpuUsage.Caption = ":-("
        MsgBox "Error while retrieving CPU usage"
    Else
        DrawUsage Ret, line1, picGraph
        lblCpuUsage.Caption = CStr(Ret) & "%"
        DoEvents
    DoEvents
    End If
End Sub

