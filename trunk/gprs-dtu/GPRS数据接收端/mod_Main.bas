Attribute VB_Name = "mod_Main"
Option Explicit

'=====================================================
'                       ȫ�ֱ�������
'=====================================================

Global glConnA As ADODB.Connection     'ȫ�����ݿ�����
Global glConnB As ADODB.Connection

Global glRAS As New RAS.RASEngine   'ȫ�ֲ��ſ�������

Global glServer As Boolean      'ȫ�ַ����Ƿ�����
Global glCenterDial As String   'ȫ���������Ĳ�������
Global glCenterIP As String     'ȫ���������� IP ��ַ

Global glLocalIP As String      '���ؾ����� IP ��ַ


Global glDBUSer As String       '���ݿ��û�������
Global glDBPass As String
Global glDBIP As String

Global glConnStringA As String 'ȫ�������ַ���
Global glConnStringB As String

Global glInfoTxtLen As Integer

'==================================================== ������
Sub Main()
    '________________________ ��ʼ��ȫ�ֱ���
    Set glConnA = New Connection
    Set glConnB = New Connection
    
    '________________________ ��ȡ����������Ϣ
    glCenterDial = GetProfileString(App.Path & "\control.ini", "����������Ϣ", "��������")
    glCenterIP = GetProfileString(App.Path & "\control.ini", "����������Ϣ", "��������IP")
    glLocalIP = GetProfileString(App.Path & "\control.ini", "��������Ϣ", "IP")
    'glWebUrl = GetProfileString(App.Path & "\Control.ini", "����������Ϣ", "��������URL")
    '________________________ ���ò���ʾ������
    Load DataRecvFrm
    SetFormNoClose DataRecvFrm
    DataRecvFrm.Enabled = False
    
    glInfoTxtLen = Len(DataRecvFrm.infoBox.Text)
    
    '________________________ ���õ�½���岢��¼���ݿ�
    Load frmLogin
    frmLogin.Show 1
    
    '________________________ �ж��Ƿ���ͬ���ݿⲢ����Ӧ�Ĵ���
    If Not frmLogin.IfConnDB Then
        MsgBox "�����������ݿ�ʧ�ܣ��޷���������", vbExclamation, "����"
        Unload frmLogin
        Unload DataRecvFrm
        End
    End If
    
    '________________________ ����������
    DataRecvFrm.Show
    DataRecvFrm.Enabled = True
    DataRecvFrm.SetFocus

    'ClearOnLine '����GPRS����״��
    'ClearCJZT

    '________________________ ��ȡ��ˢ���������е��豸�б�
    'FrmMain.FixTreeDrv
    'FrmMain.WebBrw.Navigate (glWebUrl)
    
    '________________________ ���Ʒ���״̬
    glServer = False
    'ChangeRoute '///�ı䱾��·�ɱ�
End Sub






