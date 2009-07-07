VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "数据库登录信息"
   ClientHeight    =   4485
   ClientLeft      =   7560
   ClientTop       =   5745
   ClientWidth     =   4950
   Icon            =   "frmDBLogon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSComDlg.CommonDialog openFileDialog 
      Left            =   4440
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   6588
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "result_table表"
      TabPicture(0)   =   "frmDBLogon.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label9"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtDB(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtUserName(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtPasswd(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtTableName(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtTimestamp(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "GPSData表"
      TabPicture(1)   =   "frmDBLogon.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtTimestamp(1)"
      Tab(1).Control(1)=   "txtTableName(1)"
      Tab(1).Control(2)=   "txtPasswd(1)"
      Tab(1).Control(3)=   "txtUserName(1)"
      Tab(1).Control(4)=   "txtDB(1)"
      Tab(1).Control(5)=   "Label10"
      Tab(1).Control(6)=   "Label8"
      Tab(1).Control(7)=   "Label6"
      Tab(1).Control(8)=   "Label5"
      Tab(1).Control(9)=   "Label4"
      Tab(1).ControlCount=   10
      Begin VB.TextBox txtTimestamp 
         Height          =   270
         Index           =   1
         Left            =   -73320
         TabIndex        =   22
         Text            =   "2009-5-15 11:11:34"
         ToolTipText     =   "请按照 yyyy-mm-dd hh:MM;ss的格式输入时间，例如 2009-01-01 00:00:00"
         Top             =   3060
         Width           =   2655
      End
      Begin VB.TextBox txtTimestamp 
         Height          =   270
         Index           =   0
         Left            =   1680
         TabIndex        =   20
         Text            =   "2009-5-15 11:11:34"
         ToolTipText     =   "请按照 yyyy-mm-dd hh:MM;ss的格式输入时间，例如 2009-01-01 00:00:00"
         Top             =   3060
         Width           =   2655
      End
      Begin VB.TextBox txtTableName 
         Height          =   270
         Index           =   1
         Left            =   -73320
         TabIndex        =   18
         Text            =   "GPSData"
         Top             =   2460
         Width           =   2655
      End
      Begin VB.TextBox txtTableName 
         Height          =   270
         Index           =   0
         Left            =   1680
         TabIndex        =   16
         Text            =   "result_table"
         Top             =   2460
         Width           =   2655
      End
      Begin VB.TextBox txtPasswd 
         Height          =   270
         Index           =   1
         Left            =   -73320
         TabIndex        =   14
         Top             =   1860
         Width           =   2655
      End
      Begin VB.TextBox txtUserName 
         Height          =   270
         Index           =   1
         Left            =   -73320
         TabIndex        =   13
         Top             =   1260
         Width           =   2655
      End
      Begin VB.TextBox txtDB 
         Height          =   270
         Index           =   1
         Left            =   -73320
         TabIndex        =   12
         Top             =   660
         Width           =   2655
      End
      Begin VB.TextBox txtPasswd 
         Height          =   270
         Index           =   0
         Left            =   1680
         TabIndex        =   8
         Top             =   1860
         Width           =   2655
      End
      Begin VB.TextBox txtUserName 
         Height          =   270
         Index           =   0
         Left            =   1680
         TabIndex        =   7
         Top             =   1260
         Width           =   2655
      End
      Begin VB.TextBox txtDB 
         Height          =   270
         Index           =   0
         Left            =   1680
         TabIndex        =   6
         Top             =   660
         Width           =   2655
      End
      Begin VB.Label Label10 
         Caption         =   "初始发送时间："
         Height          =   255
         Left            =   -74760
         TabIndex        =   21
         Top             =   3060
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "初始发送时间："
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   3060
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "数据表名："
         Height          =   255
         Left            =   -74760
         TabIndex        =   17
         Top             =   2460
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "数据表名："
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   2460
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "密码："
         Height          =   255
         Left            =   -74760
         TabIndex        =   11
         Top             =   1860
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "用户名："
         Height          =   255
         Left            =   -74760
         TabIndex        =   10
         Top             =   1260
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "数据源文件："
         Height          =   255
         Left            =   -74760
         TabIndex        =   9
         Top             =   660
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "密码："
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1860
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "用户名："
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1260
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "数据源文件："
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   660
         Width           =   1215
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "取消"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   3960
      Width           =   1215
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public IfConnDB As Boolean

Private Sub CancelButton_Click()
    IfConnDB = False
    Me.Hide
    End
End Sub

Private Sub Form_Load()
    txtDB(0) = GetProfileString(App.Path & "\Control.ini", SERVER_INFO, SERVER_A)
    txtUserName(0) = xorPWD(GetProfileString(App.Path & "\Control.ini", SERVER_INFO, SERVER_USER_A))
    txtPasswd(0) = xorPWD(GetProfileString(App.Path & "\Control.ini", SERVER_INFO, SERVER_PASS_A))
    txtTableName(0) = GetProfileString(App.Path & "\Control.ini", SERVER_INFO, SERVER_TABLE_A)
    txtTimestamp(0) = GetProfileString(App.Path & "\Control.ini", SERVER_INFO, SERVER_TIMESTAMP_A)
    txtDB(1) = GetProfileString(App.Path & "\Control.ini", SERVER_INFO, SERVER_B)
    txtUserName(1) = xorPWD(GetProfileString(App.Path & "\Control.ini", SERVER_INFO, SERVER_USER_B))
    txtPasswd(1) = xorPWD(GetProfileString(App.Path & "\Control.ini", SERVER_INFO, SERVER_PASS_B))
    txtTableName(1) = GetProfileString(App.Path & "\Control.ini", SERVER_INFO, SERVER_TABLE_B)
    txtTimestamp(1) = GetProfileString(App.Path & "\Control.ini", SERVER_INFO, SERVER_TIMESTAMP_B)
    
    IfConnDB = False
End Sub

Private Sub OKButton_Click()
    glConnStringA = "Provider=Microsoft.Jet.OLEDB.4.0;Jet OleDb:DataBase Password=" & txtPasswd(0) & _
                   ";User ID=" & txtUserName(0) & _
                   ";Data Source=" & txtDB(0) & ""
    glConnStringB = "Provider=Microsoft.Jet.OLEDB.4.0;Jet OleDb:DataBase Password=" & txtPasswd(1) & _
                   ";User ID=" & txtUserName(1) & _
                   ";Data Source=" & txtDB(1) & ""
                  
    
    Call GetConnection(glConnA, glConnStringA)
    Call GetConnection(glConnB, glConnStringB)
    If IsConnected(glConnA, "select * from result_table") = False Or _
       IsConnected(glConnB, "select * from GPSData") = False Then
        MsgBox "连接失败！！请检查参数配置是否正确！", vbCritical, "提示信息"
        Exit Sub
    End If
    On Error GoTo ConnErr
    '保存输入正确的信息
    WriteProFileString App.Path & "\Control.ini", SERVER_INFO, SERVER_A, Trim(txtDB(0))
    WriteProFileString App.Path & "\Control.ini", SERVER_INFO, SERVER_USER_A, xorPWD(Trim(txtUserName(0)))
    WriteProFileString App.Path & "\Control.ini", SERVER_INFO, SERVER_PASS_A, xorPWD(Trim(txtPasswd(0)))
    WriteProFileString App.Path & "\Control.ini", SERVER_INFO, SERVER_TABLE_A, Trim(txtTableName(0))
    WriteProFileString App.Path & "\Control.ini", SERVER_INFO, SERVER_TIMESTAMP_A, Trim(txtTimestamp(0))
    
    WriteProFileString App.Path & "\Control.ini", SERVER_INFO, SERVER_B, Trim(txtDB(1))
    WriteProFileString App.Path & "\Control.ini", SERVER_INFO, SERVER_USER_B, xorPWD(Trim(txtUserName(1)))
    WriteProFileString App.Path & "\Control.ini", SERVER_INFO, SERVER_PASS_B, xorPWD(Trim(txtPasswd(1)))
    WriteProFileString App.Path & "\Control.ini", SERVER_INFO, SERVER_TABLE_B, Trim(txtTableName(1))
    WriteProFileString App.Path & "\Control.ini", SERVER_INFO, SERVER_TIMESTAMP_B, Trim(txtTimestamp(1))
    IfConnDB = True
    Me.Hide
Exit Sub

ConnErr:

    MsgBox Err.Description

    IfConnDB = False
    Me.Hide
End Sub

Private Sub txtDB_Click(Index As Integer)
    'CancelError 为 True。
    On Error GoTo ErrHandler
    '设置过滤器。
    openFileDialog.Filter = "所有文件 (*.*)|*.* |数据库文件 (*.mdb)|*.mdb"
    '指定缺省过滤器。
    openFileDialog.FilterIndex = 2
    '显示“打开”对话框。
    openFileDialog.ShowOpen
    '调用打开文件的过程。
    txtDB(Index).Text = openFileDialog.FileName
    Exit Sub

ErrHandler:
    '用户按“取消”按钮。
    Exit Sub
End Sub
