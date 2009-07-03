VERSION 5.00
Begin VB.Form phoneDialFrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "拨号设置"
   ClientHeight    =   2505
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3630
   Icon            =   "phoneDialFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPhonePass 
      Height          =   270
      Left            =   1560
      TabIndex        =   7
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox txtPhoneUser 
      Height          =   270
      Left            =   1560
      TabIndex        =   6
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox txtPhoneNumber 
      Height          =   270
      Left            =   1560
      TabIndex        =   5
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "取消"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "连接！"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "密码："
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "用户名："
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "电话号码："
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "phoneDialFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public IfDialedUp As Boolean

Private Sub CancelButton_Click()
    IfDialedUp = False
    Me.Hide
End Sub

Private Sub Form_Load()
    txtPhoneNumber = GetProfileString(App.Path & "\Control.ini", MODEM_INFO, MODEM_NUMBER)
    txtPhoneUser = xorPWD(GetProfileString(App.Path & "\Control.ini", MODEM_INFO, MODEM_USER))
    txtPhonePass = xorPWD(GetProfileString(App.Path & "\Control.ini", MODEM_INFO, MODEM_PASS))
End Sub

Private Sub OKButton_Click()
    Dim temp As Long
    If txtPhoneNumber.Text = "" Or txtPhoneUser.Text = "" Or txtPhonePass.Text = "" Then
        temp = MsgBox("您没有输入有效的拨号网络参数。", vbOKOnly, "错误")
        Exit Sub
    End If
    temp = AddConnection("", txtPhoneNumber.Text, "", txtPhoneUser.Text, txtPhonePass.Text, "")
    Select Case temp
        Case ERROR_PORT_ALREADY_OPEN: temp = MsgBox("错误，端口已经打开！", vbOKOnly, "Error")
        Case ERROR_UNKNOWN: temp = MsgBox("未知的错误！", vbOKOnly, "Error")
        Case ERROR_REQUEST_TIMEOUT: temp = MsgBox("错误，请求超时！", vbOKOnly, "Error")
        Case ERROR_PASSWD_EXPIRED: temp = MsgBox("错误，您没有输入密码！", vbOKOnly, "Error")
        Case ERROR_NO_DIALIN_PERMISSION: temp = MsgBox("错误，没有拨号音！", vbOKOnly, "Error")
        Case ERROR_SERVER_NOT_RESPONDING: temp = MsgBox("错误，拨入的远程计算机没有响应！", vbOKOnly, "Error")
        Case ERROR_UNRECOGNIZED_RESPONSE: temp = MsgBox("错误，未知的响应！", vbOKOnly, "Error")
        Case ERROR_NO_RESPONSES: temp = MsgBox("错误，没有响应！", vbOKOnly, "Error")
        Case ERROR_DEVICE_NOT_READY: temp = MsgBox("错误，设备没有准备好！", vbOKOnly, "Error")
        Case ERROR_LINE_BUSY: temp = MsgBox("错误，占线！", vbOKOnly, "Error")
        Case ERROR_NO_ANSWER: temp = MsgBox("错误，没有应答信号！", vbOKOnly, "Error")
        Case ERROR_NO_CARRIER: temp = MsgBox("错误，没有载波信号！", vbOKOnly, "Error")
        Case ERROR_NO_DIALTONE: temp = MsgBox("错误，没有拨号音！", vbOKOnly, "Error")
        Case ERROR_AUTHENTICATION_FAILURE: temp = MsgBox("用户名密码出错！", vbOKOnly, "Error")
        Case ERROR_PPP_TIMEOUT: temp = MsgBox("PPP接入超时。", vbOKOnly, "Error")
        Case Else: IfDialedUp = True
    End Select
    WriteProFileString App.Path & "\Control.ini", MODEM_INFO, MODEM_NUMBER, Trim(txtPhoneNumber)
    WriteProFileString App.Path & "\Control.ini", MODEM_INFO, MODEM_USER, xorPWD(Trim(txtPhoneUser))
    WriteProFileString App.Path & "\Control.ini", MODEM_INFO, MODEM_PASS, xorPWD(Trim(txtPhonePass))
End Sub
