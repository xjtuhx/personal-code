VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "登录"
   ClientHeight    =   2220
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4230
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1311.649
   ScaleMode       =   0  'User
   ScaleWidth      =   3971.741
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSComDlg.CommonDialog openFileDialog 
      Left            =   3720
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "mdb"
      DialogTitle     =   "打开数据库"
      Filter          =   "*.mdb"
   End
   Begin VB.TextBox TxtServer 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1560
      TabIndex        =   6
      Top             =   240
      Width           =   2325
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1560
      TabIndex        =   1
      Top             =   720
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   390
      Left            =   735
      TabIndex        =   4
      Top             =   1620
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   390
      Left            =   2400
      TabIndex        =   5
      Top             =   1620
      Width           =   1140
   End
   Begin VB.TextBox txtPassWord 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "$"
      TabIndex        =   3
      Top             =   1125
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "服务器(&S):"
      Height          =   255
      Index           =   2
      Left            =   270
      TabIndex        =   7
      Top             =   240
      Width           =   1065
   End
   Begin VB.Label lblLabels 
      Caption         =   "用户名(&U):"
      Height          =   270
      Index           =   0
      Left            =   225
      TabIndex        =   0
      Top             =   750
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "密码(&P):"
      Height          =   270
      Index           =   1
      Left            =   225
      TabIndex        =   2
      Top             =   1140
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public IfConnDB As Boolean '__ 标记是否连接数据库成功

Private Sub cmdCancel_Click()
    '设置全局变量为 false
    '不提示失败的登录
    IfConnDB = False
    Me.Hide
    End
End Sub

Private Sub cmdOK_Click()
    '检查正确的密码
    
    'GLConnString = "Provider=SQLOLEDB.1;Password=" & txtPassWord & ";Persist Security Info=True;User ID=" & txtUserName & ";Initial Catalog=GPRSDATA;Data Source=" & TxtServer & ""
    GLConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Password=" & txtPassWord & _
                   ";Persist Security Info=False;User ID=" & txtUserName & _
                   ";Data Source=" & TxtServer & ""
    
    If GlConnOK = False Then
        MsgBox "连接失败！！", vbCritical, "提示信息"
        Exit Sub
    End If
    On Error GoTo ConnErr
        '保存输入正确的信息
        WriteProFileString App.Path & "\Control.ini", "服务器信息", "服务器", Trim(TxtServer)
        WriteProFileString App.Path & "\Control.ini", "服务器信息", "用户名", xorPWD(Trim(txtUserName))
        WriteProFileString App.Path & "\Control.ini", "服务器信息", "口令", xorPWD(Trim(txtPassWord))
    IfConnDB = True

    Me.Hide
Exit Sub

ConnErr:

    MsgBox Err.Description

    IfConnDB = False
    Me.Hide
End Sub

Private Sub Form_Load()
    '_________________ 读取用户及服务器信息
    
    TxtServer = GetProfileString(App.Path & "\Control.ini", "服务器信息", "服务器")
    txtUserName = xorPWD(GetProfileString(App.Path & "\Control.ini", "服务器信息", "用户名"))
    txtPassWord = xorPWD(GetProfileString(App.Path & "\Control.ini", "服务器信息", "口令"))
    
    
    IfConnDB = False
    
End Sub

Private Sub TxtServer_Click()
    'CancelError 为 True。
    On Error GoTo ErrHandler
    '设置过滤器。
    'openFileDialog.Filter = "所有文件 (*.*)|*.* | 数据库文件 (*.mdb)|*.mdb"
    '指定缺省过滤器。
    'openFileDialog.FilterIndex = 0
    '显示“打开”对话框。
    openFileDialog.ShowOpen
    '调用打开文件的过程。
    TxtServer.Text = openFileDialog.FileName
    Exit Sub

ErrHandler:
    '用户按“取消”按钮。
    Exit Sub
End Sub
