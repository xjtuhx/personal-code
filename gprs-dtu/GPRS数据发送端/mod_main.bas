Attribute VB_Name = "mod_Main"
Option Explicit

'=====================================================
'                       ȫ�ֱ�������
'=====================================================

Global glConnA As ADODB.Connection     'ȫ�����ݿ�����
Global glConnB As ADODB.Connection

Global glConnStringA As String 'ȫ�������ַ���
Global glConnStringB As String

Global glInfoTxtLen As Long

'==================================================== ������
Sub Main()
    '________________________ ��ʼ��ȫ�ֱ���
    Set glConnA = New Connection
    Set glConnB = New Connection
    
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

End Sub
