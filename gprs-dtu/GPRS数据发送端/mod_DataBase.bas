Attribute VB_Name = "mod_DataBase"
Option Explicit
'�ӳ�1��
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Function GlConnOK(ByRef conhdl As Connection, ByRef constring As String) As Boolean
'/////�ж����Ӻ���
    On Error Resume Next
    Dim rs As Recordset
    Set rs = conhdl.OpenSchema(adSchemaTables)
    If rs.RecordCount > 0 Then
        rs.Close
        Set rs = Nothing
        GlConnOK = True
        With DataSendFrm
            .statusBar.Panels(1).Text = "��������"
            .infoBox.SelStart = glInfoTxtLen
            .infoBox.SelText = CON_SUCCESS
        End With
        glInfoTxtLen = glInfoTxtLen + Len(CON_SUCCESS)
        Exit Function
    End If
    rs.Close
    Set rs = Nothing
    '///�ر����´����ݿ�����
    If conhdl.State = 1 Then conhdl.Close
    Set conhdl = Nothing
    Set conhdl = New Connection
    conhdl.Open constring
    If conhdl.State = 1 Then
        GlConnOK = True
        With DataSendFrm
            .statusBar.Panels(1).Text = "��������"
            .infoBox.SelStart = glInfoTxtLen
            .infoBox.SelText = RECON_SUCCESS
        End With
        glInfoTxtLen = glInfoTxtLen + Len(RECON_SUCCESS)
        Exit Function
    End If
    GlConnOK = False
    With DataSendFrm
        .statusBar.Panels(1).Text = "ʧȥ���ӣ�"
        .infoBox.SelStart = glInfoTxtLen
        .infoBox.SelText = CON_FAILURE
    End With
    glInfoTxtLen = glInfoTxtLen + Len(CON_FAILURE)
    Err.Clear
End Function




Public Sub ChangeRoute()
'=====================================================
'                   �ı��豸·����Ϣ
'=====================================================
On Err GoTo RouteERR
    Dim GIP As String
    Dim GGate As String
    
    Dim LIP As String
    Dim LGate As String
    
    Dim PreIP As String
    Dim Cmd As String
    Dim Loc As Integer, L As Integer, i As Integer
    
    GIP = glCenterIP
    GGate = GetProfileString(App.Path & "\control.ini", "����������Ϣ", "������������")
    LIP = GetProfileString(App.Path & "\control.ini", "��������Ϣ", "IP")
    LGate = GetProfileString(App.Path & "\control.ini", "��������Ϣ", "����")
    
    If LIP = "127.0.0.0" Then Exit Sub
    
    L = Len(GIP)
    Dim TmpI
    TmpI = 0
    For i = L To 1 Step -1
        If Mid(GIP, i, 1) = "." Then
            TmpI = TmpI + 1
            If TmpI = 2 Then
                Loc = i
                Exit For
            End If
        End If
    Next
    
    PreIP = Left(GIP, Loc - 1) & ".0.0"
    
    If Dir(App.Path & "\r.bat") <> "" Then
        Kill App.Path & "\r.bat"
    End If
    
    Open App.Path & "\r.bat" For Output As #1
    
    Cmd = "%windir%\system32\Route delete 0.0.0.0"
    Print #1, Cmd
    Cmd = "%windir%\system32\Route Add " & PreIP & " MASK 255.255.0.0 " & GGate
    Print #1, Cmd
    Cmd = "%windir%\system32\Route Add 0.0.0.0 MASK 0.0.0.0 " & LGate
    Print #1, Cmd
    
    Close #1
    
    Shell App.Path & "\r.bat"
Exit Sub
RouteERR:
    SaveERR "����·��ʱ��������!" & Err.Description
    Err.Clear
End Sub
Public Sub ClearCJCS()
    On Error GoTo ERR_ERR
    Dim tmpSQL As String
    Dim ClearCJCS As New Recordset
    tmpSQL = "select �豸��Ϣ.ID from �豸��Ϣ,GPRS��Ϣ��" & _
    " where (�豸��Ϣ.gprs_id =GPRS��Ϣ��.gprs_id and GPRS��Ϣ��.�Ƿ�ɼ� <> 0)  and �豸��Ϣ.�Ƿ�ɼ� <> 0 "
    ClearCJCS.Open tmpSQL, GlConn, adOpenDynamic, adLockReadOnly
    Do Until ClearCJCS.EOF
        tmpSQL = "update �豸��Ϣ set �ɼ�״̬  = '���ڲɼ�', �ɼ����� = 0 where id = " & ClearCJCS("ID")
        GlConn.Execute tmpSQL
        ClearCJCS.MoveNext
    Loop
    ClearCJCS.Close
    Set ClearCJCS = Nothing
Exit Sub
ERR_ERR:
    SaveERR "��ղɼ�����ʱ��������!" & Err.Description
    Err.Clear
End Sub


Public Function CreatLLZT() As Boolean
On Error GoTo ERR_ERR
    ClearLLZT '////��������豸
    Dim tmpSQL As String
    Dim LLZT_Rcd As New Recordset
    Dim Tmp_Update As New Recordset
    tmpSQL = "select * from �豸��Ϣ where �ɼ����� = '�����ɼ�'"
    LLZT_Rcd.Open tmpSQL, GlConn, adOpenDynamic, adLockReadOnly
    
    tmpSQL = "select * from ����״�� where id is null"
    Tmp_Update.Open tmpSQL, GlConn, adOpenDynamic, adLockPessimistic
    
    Do Until LLZT_Rcd.EOF
        Tmp_Update.AddNew
        Tmp_Update("DRV_ID") = LLZT_Rcd("ID")
        LLZT_Rcd.MoveNext
    Loop
    
    Tmp_Update.Update
    Tmp_Update.Close
        
    LLZT_Rcd.Close
    Set LLZT_Rcd = Nothing
    Set Tmp_Update = Nothing
    CreatLLZT = True
Exit Function
ERR_ERR:
    CreatLLZT = False
    SaveERR "��������״����ʱ��������!" & Err.Description
    Err.Clear
End Function
Public Sub ClearLLZT()
On Error GoTo Errset
    Dim tmpSQL As String
    '____ ��������豸
    tmpSQL = "DELETE FROM ����״��"
    GlConn.Execute tmpSQL
Exit Sub
Errset:
    SaveERR "��������豸ʱ��������!" & Err.Description
    Err.Clear
End Sub


Public Sub ClearCJZT()
On Error GoTo Errset
    Dim tmpSQL As String
    '____ �����豸��Ϣ
    tmpSQL = "update �豸��Ϣ set �ɼ�״̬  = 'ֹͣ�ɼ�', �ɼ����� = 0 where �Ƿ�ɼ� <> 0 "
    GlConn.Execute tmpSQL
Exit Sub
Errset:
    SaveERR "�����豸��Ϣʱ��������!" & Err.Description
    Err.Clear
End Sub

Public Sub ClearOnLine()
    Dim tmpSQL As String
    '____ ����豸
    tmpSQL = "update GPRS��Ϣ�� set �Ƿ����� = 0"
    On Error GoTo Errset
    GlConn.Execute tmpSQL
Exit Sub
Errset:
    SaveERR "����豸ʱ��������!" & Err.Description
    Err.Clear
End Sub



