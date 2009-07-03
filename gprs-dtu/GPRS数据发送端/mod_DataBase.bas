Attribute VB_Name = "mod_DataBase"
Option Explicit
'延迟1秒
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Function GlConnOK(ByRef conhdl As Connection, ByRef constring As String) As Boolean
'/////判断连接函数
    On Error Resume Next
    Dim rs As Recordset
    Set rs = conhdl.OpenSchema(adSchemaTables)
    If rs.RecordCount > 0 Then
        rs.Close
        Set rs = Nothing
        GlConnOK = True
        With DataSendFrm
            .statusBar.Panels(1).Text = "连接正常"
            .infoBox.SelStart = glInfoTxtLen
            .infoBox.SelText = CON_SUCCESS
        End With
        glInfoTxtLen = glInfoTxtLen + Len(CON_SUCCESS)
        Exit Function
    End If
    rs.Close
    Set rs = Nothing
    '///关闭重新打开数据库连接
    If conhdl.State = 1 Then conhdl.Close
    Set conhdl = Nothing
    Set conhdl = New Connection
    conhdl.Open constring
    If conhdl.State = 1 Then
        GlConnOK = True
        With DataSendFrm
            .statusBar.Panels(1).Text = "连接正常"
            .infoBox.SelStart = glInfoTxtLen
            .infoBox.SelText = RECON_SUCCESS
        End With
        glInfoTxtLen = glInfoTxtLen + Len(RECON_SUCCESS)
        Exit Function
    End If
    GlConnOK = False
    With DataSendFrm
        .statusBar.Panels(1).Text = "失去链接！"
        .infoBox.SelStart = glInfoTxtLen
        .infoBox.SelText = CON_FAILURE
    End With
    glInfoTxtLen = glInfoTxtLen + Len(CON_FAILURE)
    Err.Clear
End Function




Public Sub ChangeRoute()
'=====================================================
'                   改变设备路由信息
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
    GGate = GetProfileString(App.Path & "\control.ini", "数据中心信息", "数据中心网关")
    LIP = GetProfileString(App.Path & "\control.ini", "局域网信息", "IP")
    LGate = GetProfileString(App.Path & "\control.ini", "局域网信息", "网关")
    
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
    SaveERR "进行路由时发生错误!" & Err.Description
    Err.Clear
End Sub
Public Sub ClearCJCS()
    On Error GoTo ERR_ERR
    Dim tmpSQL As String
    Dim ClearCJCS As New Recordset
    tmpSQL = "select 设备信息.ID from 设备信息,GPRS信息表" & _
    " where (设备信息.gprs_id =GPRS信息表.gprs_id and GPRS信息表.是否采集 <> 0)  and 设备信息.是否采集 <> 0 "
    ClearCJCS.Open tmpSQL, GlConn, adOpenDynamic, adLockReadOnly
    Do Until ClearCJCS.EOF
        tmpSQL = "update 设备信息 set 采集状态  = '正在采集', 采集次数 = 0 where id = " & ClearCJCS("ID")
        GlConn.Execute tmpSQL
        ClearCJCS.MoveNext
    Loop
    ClearCJCS.Close
    Set ClearCJCS = Nothing
Exit Sub
ERR_ERR:
    SaveERR "清空采集次数时发生错误!" & Err.Description
    Err.Clear
End Sub


Public Function CreatLLZT() As Boolean
On Error GoTo ERR_ERR
    ClearLLZT '////清空流量设备
    Dim tmpSQL As String
    Dim LLZT_Rcd As New Recordset
    Dim Tmp_Update As New Recordset
    tmpSQL = "select * from 设备信息 where 采集类型 = '流量采集'"
    LLZT_Rcd.Open tmpSQL, GlConn, adOpenDynamic, adLockReadOnly
    
    tmpSQL = "select * from 流量状况 where id is null"
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
    SaveERR "生成流量状况表时发生错误!" & Err.Description
    Err.Clear
End Function
Public Sub ClearLLZT()
On Error GoTo Errset
    Dim tmpSQL As String
    '____ 清空流量设备
    tmpSQL = "DELETE FROM 流量状况"
    GlConn.Execute tmpSQL
Exit Sub
Errset:
    SaveERR "清空流量设备时发生错误!" & Err.Description
    Err.Clear
End Sub


Public Sub ClearCJZT()
On Error GoTo Errset
    Dim tmpSQL As String
    '____ 更新设备信息
    tmpSQL = "update 设备信息 set 采集状态  = '停止采集', 采集次数 = 0 where 是否采集 <> 0 "
    GlConn.Execute tmpSQL
Exit Sub
Errset:
    SaveERR "更新设备信息时发生错误!" & Err.Description
    Err.Clear
End Sub

Public Sub ClearOnLine()
    Dim tmpSQL As String
    '____ 清空设备
    tmpSQL = "update GPRS信息表 set 是否在线 = 0"
    On Error GoTo Errset
    GlConn.Execute tmpSQL
Exit Sub
Errset:
    SaveERR "清空设备时发生错误!" & Err.Description
    Err.Clear
End Sub



