VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��¼"
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
   StartUpPosition =   2  '��Ļ����
   Begin MSComDlg.CommonDialog openFileDialog 
      Left            =   3720
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "mdb"
      DialogTitle     =   "�����ݿ�"
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
      Caption         =   "ȷ��"
      Default         =   -1  'True
      Height          =   390
      Left            =   735
      TabIndex        =   4
      Top             =   1620
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��"
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
      Caption         =   "������(&S):"
      Height          =   255
      Index           =   2
      Left            =   270
      TabIndex        =   7
      Top             =   240
      Width           =   1065
   End
   Begin VB.Label lblLabels 
      Caption         =   "�û���(&U):"
      Height          =   270
      Index           =   0
      Left            =   225
      TabIndex        =   0
      Top             =   750
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "����(&P):"
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


Public IfConnDB As Boolean '__ ����Ƿ��������ݿ�ɹ�

Private Sub cmdCancel_Click()
    '����ȫ�ֱ���Ϊ false
    '����ʾʧ�ܵĵ�¼
    IfConnDB = False
    Me.Hide
    End
End Sub

Private Sub cmdOK_Click()
    '�����ȷ������
    
    'GLConnString = "Provider=SQLOLEDB.1;Password=" & txtPassWord & ";Persist Security Info=True;User ID=" & txtUserName & ";Initial Catalog=GPRSDATA;Data Source=" & TxtServer & ""
    GLConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Password=" & txtPassWord & _
                   ";Persist Security Info=False;User ID=" & txtUserName & _
                   ";Data Source=" & TxtServer & ""
    
    If GlConnOK = False Then
        MsgBox "����ʧ�ܣ���", vbCritical, "��ʾ��Ϣ"
        Exit Sub
    End If
    On Error GoTo ConnErr
        '����������ȷ����Ϣ
        WriteProFileString App.Path & "\Control.ini", "��������Ϣ", "������", Trim(TxtServer)
        WriteProFileString App.Path & "\Control.ini", "��������Ϣ", "�û���", xorPWD(Trim(txtUserName))
        WriteProFileString App.Path & "\Control.ini", "��������Ϣ", "����", xorPWD(Trim(txtPassWord))
    IfConnDB = True

    Me.Hide
Exit Sub

ConnErr:

    MsgBox Err.Description

    IfConnDB = False
    Me.Hide
End Sub

Private Sub Form_Load()
    '_________________ ��ȡ�û�����������Ϣ
    
    TxtServer = GetProfileString(App.Path & "\Control.ini", "��������Ϣ", "������")
    txtUserName = xorPWD(GetProfileString(App.Path & "\Control.ini", "��������Ϣ", "�û���"))
    txtPassWord = xorPWD(GetProfileString(App.Path & "\Control.ini", "��������Ϣ", "����"))
    
    
    IfConnDB = False
    
End Sub

Private Sub TxtServer_Click()
    'CancelError Ϊ True��
    On Error GoTo ErrHandler
    '���ù�������
    'openFileDialog.Filter = "�����ļ� (*.*)|*.* | ���ݿ��ļ� (*.mdb)|*.mdb"
    'ָ��ȱʡ��������
    'openFileDialog.FilterIndex = 0
    '��ʾ���򿪡��Ի���
    openFileDialog.ShowOpen
    '���ô��ļ��Ĺ��̡�
    TxtServer.Text = openFileDialog.FileName
    Exit Sub

ErrHandler:
    '�û�����ȡ������ť��
    Exit Sub
End Sub
