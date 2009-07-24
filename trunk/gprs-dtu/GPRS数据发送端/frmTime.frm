VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTime 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "发送起讫时间"
   ClientHeight    =   2805
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4590
   Icon            =   "frmTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin TabDlg.SSTab SSTab1 
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   3625
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "result_table"
      TabPicture(0)   =   "frmTime.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtStartTime(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtEndTime(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "GPSData"
      TabPicture(1)   =   "frmTime.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtStartTime(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtEndTime(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.TextBox txtEndTime 
         Height          =   270
         Index           =   1
         Left            =   -73200
         TabIndex        =   10
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox txtStartTime 
         Height          =   270
         Index           =   1
         Left            =   -73200
         TabIndex        =   9
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtEndTime 
         Height          =   270
         Index           =   0
         Left            =   1800
         TabIndex        =   6
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox txtStartTime 
         Height          =   270
         Index           =   0
         Left            =   1800
         TabIndex        =   5
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "终止发送时间："
         Height          =   255
         Left            =   -74640
         TabIndex        =   8
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "起始发送时间："
         Height          =   255
         Left            =   -74640
         TabIndex        =   7
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "终止发送时间："
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "起始发送时间："
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
End
Attribute VB_Name = "frmTime"
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
    txtStartTime(0) = GetProfileString(App.Path & "\Control.ini", TIME_INFO, TIME_START_A, Now)
    txtStartTime(1) = GetProfileString(App.Path & "\Control.ini", TIME_INFO, TIME_START_B, Now)
    txtEndTime(0) = GetProfileString(App.Path & "\Control.ini", TIME_INFO, TIME_END_A, Now)
    txtEndTime(1) = GetProfileString(App.Path & "\Control.ini", TIME_INFO, TIME_END_B, Now)
    
    Cancelled = True
End Sub

Private Sub OKButton_Click()
    WriteProFileString App.Path & "\Control.ini", TIME_INFO, TIME_START_A, Trim(txtStartTime(0))
    WriteProFileString App.Path & "\Control.ini", TIME_INFO, TIME_START_B, Trim(txtStartTime(1))
    WriteProFileString App.Path & "\Control.ini", TIME_INFO, TIME_END_A, Trim(txtEndTime(0))
    WriteProFileString App.Path & "\Control.ini", TIME_INFO, TIME_END_B, Trim(txtEndTime(1))
    Cancelled = False
    Me.Hide
End Sub
