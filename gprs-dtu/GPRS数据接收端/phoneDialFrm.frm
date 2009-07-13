VERSION 5.00
Begin VB.Form phoneDialFrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "拨号设置"
   ClientHeight    =   3255
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3630
   Icon            =   "phoneDialFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.ComboBox cmbType 
      Height          =   300
      ItemData        =   "phoneDialFrm.frx":0CCA
      Left            =   1560
      List            =   "phoneDialFrm.frx":0CCC
      TabIndex        =   11
      Top             =   360
      Width           =   1695
   End
   Begin VB.ComboBox cmbModem 
      Height          =   300
      Left            =   1560
      TabIndex        =   9
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox txtPhonePass 
      Height          =   270
      Left            =   1560
      TabIndex        =   7
      Text            =   "163"
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox txtPhoneUser 
      Height          =   270
      Left            =   1560
      TabIndex        =   6
      Text            =   "163"
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox txtPhoneNumber 
      Height          =   270
      Left            =   1560
      TabIndex        =   5
      Text            =   "96136"
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "取消"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "拨号类型："
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "拨号设备："
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "密码："
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "用户名："
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "电话号码："
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   840
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

Private Sub cmbType_Click()
    Select Case cmbType.List(cmbType.ListIndex)
        Case NAME_VPN
            Label1.Caption = "拨号地址："
            cmbModem.Enabled = False
            txtPhoneNumber.Enabled = True
        Case NAME_MODEM
            Label1.Caption = "电话号码："
            cmbModem.Enabled = True
            txtPhoneNumber.Enabled = True
        Case NAME_ADSL
            Label1.Caption = "电话号码："
            cmbModem.Enabled = False
            txtPhoneNumber.Enabled = False
        Case NAME_CDMA1X
            Label1.Caption = "电话号码："
            cmbModem.Enabled = True
            txtPhoneNumber.Enabled = True
        End Select
End Sub

Private Sub Form_Load()
    txtPhoneNumber = GetProfileString(App.Path & "\Control.ini", MODEM_INFO, MODEM_NUMBER)
    txtPhoneUser = xorPWD(GetProfileString(App.Path & "\Control.ini", MODEM_INFO, MODEM_USER))
    txtPhonePass = xorPWD(GetProfileString(App.Path & "\Control.ini", MODEM_INFO, MODEM_PASS))
    Cancelled = True
    cmbModem.Clear
    
    On Error Resume Next
    
    cmbType.Clear
    cmbType.AddItem NAME_CDMA1X
    cmbType.AddItem NAME_VPN
    cmbType.AddItem NAME_MODEM
    cmbType.AddItem NAME_ADSL
    cmbType.ListIndex = 0

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
    On Error Resume Next
    If txtPhoneNumber.Text = "" Or txtPhoneUser.Text = "" Or txtPhonePass.Text = "" Then
        temp = MsgBox("您没有输入有效的拨号网络参数。", vbOKOnly, "错误")
        Cancelled = True
        Exit Sub
    End If
    WriteProFileString App.Path & "\Control.ini", MODEM_INFO, MODEM_TYPE, cmbType.List(cmbType.ListIndex)
    WriteProFileString App.Path & "\Control.ini", MODEM_INFO, MODEM_NUMBER, Trim(txtPhoneNumber)
    WriteProFileString App.Path & "\Control.ini", MODEM_INFO, MODEM_USER, xorPWD(Trim(txtPhoneUser))
    WriteProFileString App.Path & "\Control.ini", MODEM_INFO, MODEM_PASS, xorPWD(Trim(txtPhonePass))
    Cancelled = False
    Me.Hide
End Sub
