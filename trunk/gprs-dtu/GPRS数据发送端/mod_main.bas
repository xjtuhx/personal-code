Attribute VB_Name = "mod_Main"
Option Explicit

'=====================================================
'                       ȫ�ֱ�������
'=====================================================

Global GlConn As ADODB.Connection     'ȫ�����ݿ�����

Global glRAS As New RAS.RASEngine   'ȫ�ֲ��ſ�������

Global glServer As Boolean      'ȫ�ַ����Ƿ�����
Global glCenterDial As String   'ȫ���������Ĳ�������
Global glCenterIP As String     'ȫ���������� IP ��ַ

Global glLocalIP As String      '���ؾ����� IP ��ַ


Global glDBUSer As String       '���ݿ��û�������
Global glDBPass As String
Global glDBIP As String

Global GLConnString As String 'ȫ�������ַ���
Global glWebUrl As String '�豸��ʱ��Ϣ��ַ
Global ConnMSDB As Integer

Global glInfoTxtLen As Integer

'==================================================== ������
Sub Main()
    '________________________ ��ʼ��ȫ�ֱ���
    Set GlConn = New Connection
    '________________________ ��ȡ����������Ϣ
    glCenterDial = GetProfileString(App.Path & "\control.ini", "����������Ϣ", "��������")
    glCenterIP = GetProfileString(App.Path & "\control.ini", "����������Ϣ", "��������IP")
    'glLocalIP = GetProfileString(App.Path & "\control.ini", "��������Ϣ", "IP")
    'glWebUrl = GetProfileString(App.Path & "\Control.ini", "����������Ϣ", "��������URL")
    '________________________ ���ò���ʾ������
    Load DataSendFrm
    SetFormNoClose DataSendFrm
    DataSendFrm.Enabled = False
    
    glInfoTxtLen = Len(DataSendFrm.infoBox.Text)
    
    '________________________ ���õ�½���岢��¼���ݿ�
    Load frmLogin
    frmLogin.Show 1
    
    '________________________ �ж��Ƿ���ͬ���ݿⲢ����Ӧ�Ĵ���
    If Not frmLogin.IfConnDB Then
        MsgBox "�����������ݿ�ʧ�ܣ��޷���������", vbExclamation, "����"
        Unload frmLogin
        Unload DataSendFrm
        End
    End If
    
    '________________________ ����������
    DataSendFrm.Show
    DataSendFrm.Enabled = True
    DataSendFrm.SetFocus

    'ClearOnLine '����GPRS����״��
    'ClearCJZT

    '________________________ ��ȡ��ˢ���������е��豸�б�
    'FrmMain.FixTreeDrv
    'FrmMain.WebBrw.Navigate (glWebUrl)
    
    '________________________ ���Ʒ���״̬
    glServer = False
    'ChangeRoute '///�ı䱾��·�ɱ�
End Sub

