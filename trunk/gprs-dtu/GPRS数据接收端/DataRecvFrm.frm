VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form DataRecvFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GPRS���ݽ��ն�"
   ClientHeight    =   7080
   ClientLeft      =   1575
   ClientTop       =   2145
   ClientWidth     =   8460
   Icon            =   "DataRecvFrm.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   8460
   StartUpPosition =   2  '��Ļ����
   Begin MSWinsockLib.Winsock serverSock 
      Left            =   600
      Top             =   6000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imageBar 
      Left            =   1200
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataRecvFrm.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataRecvFrm.frx":0C5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataRecvFrm.frx":18B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataRecvFrm.frx":2502
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataRecvFrm.frx":3154
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataRecvFrm.frx":3E2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataRecvFrm.frx":4B08
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataRecvFrm.frx":57E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataRecvFrm.frx":6434
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataRecvFrm.frx":710E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataRecvFrm.frx":7DE8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1920
      Top             =   6000
   End
   Begin VB.Frame Frame2 
      Caption         =   "״̬��Ϣ"
      Height          =   5655
      Left            =   2760
      TabIndex        =   3
      Top             =   960
      Width           =   5655
      Begin RichTextLib.RichTextBox infoBox 
         Height          =   5295
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   9340
         _Version        =   393217
         ScrollBars      =   3
         TextRTF         =   $"DataRecvFrm.frx":8A3A
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "���Ͷ��б�"
      Height          =   5655
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2535
      Begin MSComctlLib.TreeView clientList 
         Height          =   5295
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   9340
         _Version        =   393217
         Style           =   7
         Appearance      =   1
      End
   End
   Begin MSComctlLib.StatusBar statusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   6705
      Width           =   8460
      _ExtentX        =   14923
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9737
            Text            =   "״̬��Ϣ"
            TextSave        =   "״̬��Ϣ"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "2009-7-3"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Text            =   "��ʾʱ��"
            TextSave        =   "��ʾʱ��"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar toolBar 
      Align           =   1  'Align Top
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8460
      _ExtentX        =   14923
      _ExtentY        =   1508
      ButtonWidth     =   1455
      ButtonHeight    =   1349
      Appearance      =   1
      ImageList       =   "imageBar"
      DisabledImageList=   "imageBar"
      HotImageList    =   "imageBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "��������"
            Key             =   "��������"
            Description     =   "��������"
            Object.ToolTipText     =   "ͨ��Modem�������ӵ�������"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�Ͽ�����"
            Key             =   "�Ͽ�����"
            Description     =   "�Ͽ�����"
            Object.ToolTipText     =   "�Ͽ���������"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "��ʼ����"
            Key             =   "��ʼ����"
            Description     =   "��ʼ����"
            Object.ToolTipText     =   "��ʼ���ܷ��Ͷ�����"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ֹͣ����"
            Key             =   "ֹͣ����"
            Description     =   "ֹͣ����"
            Object.ToolTipText     =   "ֹͣ�ӷ��Ͷ˽�������"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "��������"
            Key             =   "��������"
            Description     =   "��������"
            Object.ToolTipText     =   "��������ѡ��"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�˳�����"
            Key             =   "�˳�����"
            Description     =   "�˳�����"
            Object.ToolTipText     =   "�˳����ճ���"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "DataRecvFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    statusBar.Panels(3).Text = Time
End Sub

Private Sub Timer1_Timer()
    statusBar.Panels(3).Text = Time
End Sub

Private Sub toolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "��������"
            
        Case "�Ͽ�����"
            
        Case "��ʼ����"
            
        Case "ֹͣ����"
            
        Case "��������"
            
        Case "�˳�����"
            Unload Me
            End
    End Select

End Sub
