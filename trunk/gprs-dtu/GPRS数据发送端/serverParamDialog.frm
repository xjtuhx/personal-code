VERSION 5.00
Begin VB.Form serverParamDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "接收端参数配置"
   ClientHeight    =   2310
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3945
   Icon            =   "serverParamDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox portBox 
      Height          =   270
      Left            =   1680
      TabIndex        =   5
      Text            =   "56789"
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox ipBox 
      Height          =   270
      Left            =   1680
      TabIndex        =   4
      Text            =   "127.0.0.1"
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "取消"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "接收端口："
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "接收端地址："
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "serverParamDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public Cancelled As Boolean

Private Sub CancelButton_Click()
    Cancelled = True
    Me.Hide
End Sub

Private Sub Form_Load()
    Cancelled = True
    ipBox = GetProfileString(App.Path, "\Control.ini", CONNECT_INFO, CONNECT_IP)
    portBox = GetProfileString(App.Path, "\Control.ini", CONNECT_INFO, CONNECT_PORT)
End Sub

Private Sub OKButton_Click()
    If CheckIP(ipBox.Text) And CLng(portBox.Text) > 0 And CLng(portBox.Text) < 65535 Then
        WriteProFileString App.Path & "\Control.ini", CONNECT_INFO, CONNECT_IP, Trim(Me.ipBox.Text)
        WriteProFileString App.Path & "\Control.ini", CONNECT_INFO, CONNECT_PORT, Trim(Me.portBox.Text)
        Cancelled = False
        Me.Hide
    End If
End Sub
