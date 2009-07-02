Attribute VB_Name = "mod_Connection"
Option Explicit

Public Function GPRSconn() As Boolean
On Error GoTo GPRSConn_ERR
    Dim tmpConn As Boolean
    Dim tmpId As Integer
    
    '___ �ж����������Ƿ��Ѿ���
    
    If CheckConnID(glCenterDial) > -1 Then
        tmpConn = True
    Else
        tmpConn = False
    End If
    
    DoEvents
    
    '___ ���Դ���������
    
    If Not tmpConn Then
        tmpId = CheckDialID(glCenterDial)
        DoEvents
        If tmpId > -1 Then
            If Not DialUP(tmpId) Then
                tmpConn = False
            Else
                tmpConn = True
            End If
        Else
            tmpConn = False
        End If
    End If
    
    DoEvents
GPRSconn = tmpConn

Exit Function
GPRSConn_ERR:
    GPRSconn = False
    SaveERR "���ſ������ķ�������!" & Err.Description
    Err.Clear
End Function

'----------------------------------------------- ���Ҷ�Ӧ�Ĳ��� ID
Public Function CheckDialID(ByVal tmpName As String) As Integer

    Dim tmpPhoneTot As Integer
    Dim tmpPhoneID As Integer
    Dim I As Integer
    
    tmpPhoneID = -1
    tmpPhoneTot = glRAS.PhoneEntries.Count
    
    For I = 0 To tmpPhoneTot - 1
        If Trim(UCase(glRAS.PhoneEntries(I).EntryName)) = Trim(UCase(tmpName)) Then
            tmpPhoneID = I
            Exit For
        End If
    Next

    CheckDialID = I
    
End Function


'----------------------------------------------- ��������ͨ�Ĳ��� ID
Public Function CheckConnID(ByVal tmpName As String) As Integer

    Dim tmpConnTot As Integer
    Dim tmpConnID As Integer
    Dim I As Integer
    
    tmpConnID = -1
    
    tmpConnTot = glRAS.Connections.Count
    
    For I = 0 To tmpConnTot - 1
        If Trim(UCase(glRAS.Connections(I).EntryName)) = Trim(UCase(tmpName)) Then
            tmpConnID = I
            Exit For
        End If
    Next

    CheckConnID = tmpConnID
    
End Function


'----------------------------------------------- ���в���

Public Function DialUP(ByVal tmpId As Integer) As Boolean

    Dim tmpConn As RAS.RConnection
    
    On Error GoTo ErrDial
    Set tmpConn = glRAS.PhoneEntries(tmpId).DialEntry(False)
    On Error GoTo 0
    DialUP = True
    ChangeRoute
Exit Function

ErrDial:
    SaveERR "���в���ʱ��������!" & Err.Description
    DialUP = False
    Err.Clear

End Function



'------------------------------------------------ �Ͽ���������
Public Function HungUP(ByVal tmpId As Integer) As Boolean
    
    On Error GoTo ErrHung
    glRAS.Connections.RemoveConnection CInt(Str(tmpId))
    DoEvents
    On Error GoTo 0
    HungUP = True
    
Exit Function

ErrHung:

    HungUP = False
    Err.Clear

End Function

