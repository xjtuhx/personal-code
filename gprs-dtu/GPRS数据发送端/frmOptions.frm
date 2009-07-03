VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form optionsDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数配置"
   ClientHeight    =   4080
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   3375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   5953
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "网络参数配置"
      TabPicture(0)   =   "frmOptions.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "nameBox"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ipBox"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "portBox"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "数据库参数配置"
      TabPicture(1)   =   "frmOptions.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "dbPasswdBox"
      Tab(1).Control(1)=   "dbUsernameBox"
      Tab(1).Control(2)=   "dbFileBox"
      Tab(1).Control(3)=   "Label6"
      Tab(1).Control(4)=   "Label5"
      Tab(1).Control(5)=   "Label4"
      Tab(1).ControlCount=   6
      Begin VB.TextBox dbPasswdBox 
         Height          =   270
         Left            =   -72720
         TabIndex        =   14
         Text            =   "请输入数据库密码"
         Top             =   2040
         Width           =   3255
      End
      Begin VB.TextBox dbUsernameBox 
         Height          =   270
         Left            =   -72720
         TabIndex        =   13
         Text            =   "请输入数据库用户名"
         Top             =   1440
         Width           =   3255
      End
      Begin VB.TextBox dbFileBox 
         Height          =   270
         Left            =   -72720
         TabIndex        =   12
         Text            =   "点击打开数据库"
         Top             =   840
         Width           =   3255
      End
      Begin VB.TextBox portBox 
         Height          =   270
         Left            =   2280
         TabIndex        =   8
         Text            =   "56789"
         Top             =   2040
         Width           =   3255
      End
      Begin VB.TextBox ipBox 
         Height          =   270
         Left            =   2280
         TabIndex        =   7
         Text            =   "请输入接收端IP地址"
         Top             =   1440
         Width           =   3255
      End
      Begin VB.TextBox nameBox 
         Height          =   270
         Left            =   2280
         TabIndex        =   6
         Text            =   "终端1"
         ToolTipText     =   "用于在接收端区分不同发送端的名字，可以任意填写。"
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label Label6 
         Caption         =   "数据库密码："
         Height          =   255
         Left            =   -74160
         TabIndex        =   11
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "数据库用户名："
         Height          =   255
         Left            =   -74340
         TabIndex        =   10
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "数据库文件位置："
         Height          =   255
         Left            =   -74520
         TabIndex        =   9
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "数据接收端端口号："
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "数据接收端IP地址："
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "本机终端别名："
         Height          =   255
         Left            =   720
         TabIndex        =   3
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "取消"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   3600
      Width           =   1215
   End
End
Attribute VB_Name = "optionsDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim Cancelled As Boolean

Private Sub CancelButton_Click()
    Me.Hide
    Cancelled = True
    
End Sub

Private Sub Form_Load()
    Cancelled = True
    
End Sub

Private Sub OKButton_Click()
    If Validate_Input Then
        Me.Hide
        Cancelled = False
    Else
        MsgBox "参数配置有误，请检查参数配置", vbOKOnly, "错误"
        Cancelled = True
    End If
End Sub


Private Function Validate_Input() As Boolean
    Validate_Input = True
End Function
