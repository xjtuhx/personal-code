VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form DataSendFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GPRS���ݷ��Ͷ�"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   6450
   StartUpPosition =   2  '��Ļ����
   Begin MSWinsockLib.Winsock sock 
      Left            =   5160
      Top             =   5880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5760
      Top             =   5760
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
            Picture         =   "DataSendFrm.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataSendFrm.frx":0C52
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataSendFrm.frx":18A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataSendFrm.frx":24F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataSendFrm.frx":3148
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataSendFrm.frx":3D9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataSendFrm.frx":49EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataSendFrm.frx":56C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataSendFrm.frx":63A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataSendFrm.frx":707A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataSendFrm.frx":7D54
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6360
      Top             =   5880
   End
   Begin MSComctlLib.StatusBar statusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   6375
      Width           =   6450
      _ExtentX        =   11377
      _ExtentY        =   661
      SimpleText      =   "״̬��"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6191
            Text            =   "״̬��"
            TextSave        =   "״̬��"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            Text            =   "��ʾ����"
            TextSave        =   "2009-7-4"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
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
      Width           =   6450
      _ExtentX        =   11377
      _ExtentY        =   1508
      ButtonWidth     =   1773
      ButtonHeight    =   1349
      Appearance      =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "���ӷ�����"
            Key             =   "���ӷ�����"
            Description     =   "���ӷ�����"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�Ͽ�������"
            Key             =   "�Ͽ�������"
            Description     =   "�Ͽ�������"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "��ʼ����"
            Key             =   "��ʼ����"
            Description     =   "��ʼ����"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ֹͣ����"
            Key             =   "ֹͣ����"
            Description     =   "ֹͣ����"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "��������"
            Key             =   "��������"
            Description     =   "��������"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�˳�����"
            Key             =   "�˳�����"
            Description     =   "�˳�����"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox infoBox 
      Height          =   5535
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   9763
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"DataSendFrm.frx":8A2E
   End
End
Attribute VB_Name = "DataSendFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    statusBar.Panels(3).Text = Time
End Sub

Private Sub Timer1_Timer()
    statusBar.Panels(3).Text = Time
End Sub

Private Sub toolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "���ӷ�����"
            'Զ��������
            If optionsDialog.ipBox = "" Then
                MsgBox "�����÷�����IP����", vbOKOnly, "ȱ�ٲ���"
            End If
            If optionsDialog.portBox = "" Then
                MsgBox "�����÷������˿ڲ���", vbOKOnly, "ȱ�ٲ���"
            End If
            
            sock.RemoteHost = optionsDialog.ipBox
            '����˿�
            sock.RemotePort = optionsDialog.portBox
            '������������
            sock.Connect
            
        Case "�Ͽ�������"
            sock.Close
        Case "��ʼ����"
            If Not sock.State = sckConnected Then
                MsgBox "�����ѶϿ������������ӷ�������", vbOKOnly, "������Ϣ"
            Else
                Dim i As Integer
                For i = 0 To 5
                    sock.SendData ("Hello! VB Winsock!")
                    Sleep (1000)
                Next i
                
            End If
        Case "ֹͣ����"
        Case "��������"
            optionsDialog.Show vbModal, DataSendFrm
        Case "�˳�����"
            Unload Me
            End
    End Select
End Sub

Private Sub sock_Close()
    MsgBox ("socket closed")
End Sub

Private Sub sock_Connect()
    MsgBox ("socket connected")
End Sub
