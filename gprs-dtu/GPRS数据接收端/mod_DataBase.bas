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
