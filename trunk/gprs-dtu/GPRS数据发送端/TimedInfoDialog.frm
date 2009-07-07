VERSION 5.00
Begin VB.Form TimedInfoDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "信息"
   ClientHeight    =   1560
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer timeout_timer 
      Interval        =   1000
      Left            =   3840
      Top             =   1080
   End
   Begin VB.Label timeLabel 
      Alignment       =   2  'Center
      Caption         =   "30秒"
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label infoLabel 
      Alignment       =   2  'Center
      Caption         =   "显示信息"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3855
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "TimedInfoDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public Success As Boolean
Public Timeout As Integer

Private Sub Form_Load()
    Success = False
    Timeout = 30
    timeout_timer.Enabled = False
End Sub

Private Sub timeout_timer_Timer()
    If Timeout > 0 Then
        timeLabel.Caption = "操作剩余 " & str(Timeout) & " 秒"
        Timeout = Timeout - 1
    Else
        Success = False
        timeout_timer.Enabled = False
        Me.Hide
    End If
End Sub

Public Sub Start(ByRef infoText As String)
    infoLabel.Caption = infoText
    timeLabel.Caption = "操作剩余 " & str(Timeout) & " 秒"
    Success = False
    timeout_timer.Enabled = True
    Me.Show vbModal
End Sub

Public Sub Cancel()
    Success = True
    timeout_timer.Enabled = False
    Me.Hide
End Sub


