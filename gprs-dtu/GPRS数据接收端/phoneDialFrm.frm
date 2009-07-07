VERSION 5.00
Begin VB.Form phoneDialFrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "拨号设置"
   ClientHeight    =   2955
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3630
   Icon            =   "phoneDialFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbModem 
      Height          =   300
      Left            =   1560
      TabIndex        =   9
      Top             =   1800
      Width           =   1695
   End
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
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "拨号设备："
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1800
      Width           =   975
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
Public IfDialedUp As Boolean
Public Cancelled As Boolean

Private Sub CancelButton_Click()
    IfDialedUp = False
    Cancelled = True
    Me.Hide
End Sub

Private Sub Form_Load()
    txtPhoneNumber = GetProfileString(App.Path & "\Control.ini", MODEM_INFO, MODEM_NUMBER)
    txtPhoneUser = xorPWD(GetProfileString(App.Path & "\Control.ini", MODEM_INFO, MODEM_USER))
    txtPhonePass = xorPWD(GetProfileString(App.Path & "\Control.ini", MODEM_INFO, MODEM_PASS))
    Cancelled = True
    cmbModem.Clear
    
    On Error Resume Next

    strComputer = "."
    Set objWMIService = GetObject("winmgmts:" _
        & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

    Set colItems = objWMIService.ExecQuery("Select * from Win32_POTSModem")

    For Each objitem In colItems
        cmbModem.Text = objitem.Name
        cmbModem.AddItem objitem.Name
        Next
        
End Sub

Private Sub OKButton_Click()
    If txtPhoneNumber.Text = "" Or txtPhoneUser.Text = "" Or txtPhonePass.Text = "" Then
        temp = MsgBox("您没有输入有效的拨号网络参数。", vbOKOnly, "错误")
        Exit Sub
    End If
    WriteProFileString App.Path & "\Control.ini", MODEM_INFO, MODEM_NUMBER, Trim(txtPhoneNumber)
    WriteProFileString App.Path & "\Control.ini", MODEM_INFO, MODEM_USER, xorPWD(Trim(txtPhoneUser))
    WriteProFileString App.Path & "\Control.ini", MODEM_INFO, MODEM_PASS, xorPWD(Trim(txtPhonePass))
End Sub
