Attribute VB_Name = "mod_DataBase"
Public Function GlConnOK() As Boolean
'/////判断连接函数
    On Error Resume Next
    Dim tmpSQL As String
    Dim TEST_Rcd As New Recordset
    TEST_Rcd.CursorLocation = adUseClient
    tmpSQL = "select * from result_table"
    TEST_Rcd.Open tmpSQL, GlConn, adOpenForwardOnly, adLockReadOnly
    If TEST_Rcd.State = 1 Then
        TEST_Rcd.Close
        Set TEST_Rcd = Nothing
        GlConnOK = True
        'FrmMain.TxtConnOK = "连接正常" & ConnMSDB
        With DataSendFrm
            .statusBar.Panels(1).Text = "连接正常" & ConnMSDB
            .infoBox.SelStart = glInfoTxtLen
            .infoBox.SelText = CON_SUCCESS
        End With
        glInfoTxtLen = glInfoTxtLen + Len(CON_SUCCESS)
        Exit Function
    End If
    TEST_Rcd.Close
    Set TEST_Rcd = Nothing
    '///关闭重新打开数据库连接
    If GlConn.State = 1 Then GlConn.Close
    Set GlConn = Nothing
    ConnMSDB = ConnMSDB + 1 '连接失败计数
    Set GlConn = New Connection
    GlConn.Open GLConnString
    If GlConn.State = 1 Then
        GlConnOK = True
        With DataSendFrm
            .statusBar.Panels(1).Text = "连接正常" & ConnMSDB
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

