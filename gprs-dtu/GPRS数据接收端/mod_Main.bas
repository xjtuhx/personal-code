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

End Sub
