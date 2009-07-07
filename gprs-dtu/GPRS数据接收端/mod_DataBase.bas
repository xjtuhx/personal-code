Attribute VB_Name = "mod_DataBase"
Option Explicit
'延迟1秒
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Sub GetConnection(ByRef con As Connection, ByRef str As String)
    On Error Resume Next
    If con Is Nothing Then
        Set con = New ADODB.Connection
    End If
    con.Open (str)
    If con.State = 1 Then
        With DataRecvFrm
            .statusBar.Panels(1).Text = "连接正常"
            .infoBox.SelStart = glInfoTxtLen
            .infoBox.SelText = RECON_SUCCESS
        End With
        glInfoTxtLen = glInfoTxtLen + Len(RECON_SUCCESS)
        Exit Sub
    End If
    With DataRecvFrm
        .statusBar.Panels(1).Text = "失去链接！"
        .infoBox.SelStart = glInfoTxtLen
        .infoBox.SelText = CON_FAILURE
    End With
    glInfoTxtLen = glInfoTxtLen + Len(CON_FAILURE)
    Err.Clear
End Sub

Public Function IsConnected(ByRef con As Connection, ByRef sql As String)
    On Error Resume Next
    If con Is Nothing Or con.State <> 1 Then
        IsConnected = False
        Exit Function
    End If
    
    Dim rs As Recordset
    Set rs = New Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, con, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount >= 0 Then
        IsConnected = True
    Else
        IsConnected = False
    End If
    
    If rs.State = 1 Then
        rs.Close
    End If
    
    Set rs = Nothing
    
End Function


Public Sub GetRecords(ByRef rs As Recordset, ByRef dbcon As ADODB.Connection, ByRef tableName As String, ByRef timestamp As String)
'=====================================================
'从数据表中获取从timestamp开始的信息
'=====================================================
    On Error Resume Next
    If rs Is Nothing Then
        Set rs = New Recordset
    End If
    If rs.State = 1 Then
        rs.Close
    End If
    If dbcon.State <> adStateOpen Then
        MsgBox "数据库连接失败！无法获取数据！", vbOKOnly, "出错提示"
        Exit Sub
    End If
    Dim sqlstring As String
    sqlstring = "select * from " & tableName & " where measuretime > #" & timestamp & "#"
    rs.CursorLocation = adUseClient
    rs.Open sqlstring, dbcon, adOpenDynamic, adLockReadOnly
    If Not rs.RecordCount >= 0 Then
        rs.Close
        Set rs = Nothing
    End If
        
End Sub


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
    Dim Loc As Integer, L As Integer, I As Integer
    
    GIP = glCenterIP
    GGate = GetProfileString(App.Path & "\control.ini", "数据中心信息", "数据中心网关")
    LIP = GetProfileString(App.Path & "\control.ini", "局域网信息", "IP")
    LGate = GetProfileString(App.Path & "\control.ini", "局域网信息", "网关")
    
    If LIP = "127.0.0.0" Then Exit Sub
    
    L = Len(GIP)
    Dim TmpI
    TmpI = 0
    For I = L To 1 Step -1
        If Mid(GIP, I, 1) = "." Then
            TmpI = TmpI + 1
            If TmpI = 2 Then
                Loc = I
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

