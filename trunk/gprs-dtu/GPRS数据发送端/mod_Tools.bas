Attribute VB_Name = "mod_Tools"
Option Explicit

'========================= ������ǰ

Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10

'==========================


'==========================  Ŀ¼�Ի���
Private Type BrowseInfo
hWndOwner As Long
pIDLRoot As Long
pszDisplayName As Long
lpszTitle As Long
ulFlags As Long
lpfnCallback As Long
lParam As Long
iImage As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

'===========================


'=========================== ���ô���رհ�ť
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long

Const SC_CLOSE = &HF060

'----------------------------------------- ������ǰ
Public Sub SetFormOnTop(myForm As Object)
    SetWindowPos myForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub


'----------------------------------------- ȡ����ǰ
Public Sub SetFormNoTop(myForm As Object)
    SetWindowPos myForm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub


'----------------------------------------- ������ܡ�����
Public Function xorPWD(ByVal s As String) As String
    Dim temp As String
    Dim i
        temp = ""
        For i = 1 To Len(s)
            temp = temp + Chr(Asc(Mid(s, i, 1)) Xor 13)
        Next i
        xorPWD = temp
End Function


'===============================  ѡȡ�ļ��к���

Public Function BrowseForFolder(hWndOwner As Long, sPrompt As String) As String

    Dim iNull As Integer
    Dim lpIDList As Long
    Dim lResult As Long
    Dim sPath As String
    Dim udtBI As BrowseInfo
    
    With udtBI
    .hWndOwner = hWndOwner
    .lpszTitle = lstrcat(sPrompt, "")
    .ulFlags = BIF_RETURNONLYFSDIRS
    End With
    
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
    sPath = String$(MAX_PATH, 0)
    lResult = SHGetPathFromIDList(lpIDList, sPath)
    Call CoTaskMemFree(lpIDList)
    iNull = InStr(sPath, vbNullChar)
    If iNull Then
    sPath = Left$(sPath, iNull - 1)
    End If
    End If
    
    BrowseForFolder = sPath

End Function


'============================== ���ùرհ�ť
Public Sub SetFormNoClose(myForm As Object)
    
    Dim hMenu As Long
    
    hMenu = GetSystemMenu(myForm.hwnd, 0)
    RemoveMenu hMenu, &HF060, &H200&
    
End Sub

'========================= ��֤IP��
Public Function CheckIP(IPValue As String) As Boolean
    Dim IPReg As New RegExp
    IPReg.Pattern = "^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$"
    Dim posIP1 As Integer: Dim intIP1 As Integer
    Dim posIP2 As Integer: Dim intIP2 As Integer
    Dim posIP3 As Integer: Dim intIP3 As Integer
    Dim intIP4 As Integer
   
    If IPReg.Test(IPValue) = False Then
        CheckIP = False
        MsgBox "������� IP ��ַ��ʽ���ԣ����������룡"
        Exit Function
    End If
   
    posIP1 = InStr(1, IPValue, ".")
    intIP1 = CInt(Mid(IPValue, 1, posIP1 - 1))
    If intIP1 < 1 Or intIP1 > 224 Then
        CheckIP = False
        MsgBox "��һ��IP��ַ������ 1-223 ֮�䣡"
        Exit Function
    End If
   
    posIP2 = InStr(posIP1 + 1, IPValue, ".")
    intIP2 = CInt(Mid(IPValue, posIP1 + 1, posIP2 - posIP1 - 1))
    If intIP2 > 255 Then
        CheckIP = False
        MsgBox "�ڶ���IP��ַ������ 0-255 ֮�䣡"
        Exit Function
    End If
   
    posIP3 = InStr(posIP2 + 1, IPValue, ".")
    intIP3 = CInt(Mid(IPValue, posIP2 + 1, posIP3 - posIP2 - 1))
    If intIP3 > 255 Then
        CheckIP = False
        MsgBox "������IP��ַ������ 0-255 ֮�䣡"
        Exit Function
    End If
   
    intIP4 = CInt(Mid(IPValue, posIP3 + 1, Len(IPValue) - posIP3))
    If intIP4 > 255 Then
        CheckIP = False
        MsgBox "���Ķ�IP��ַ������ 0-255 ֮�䣡"
        Exit Function
    End If
   
    CheckIP = True
End Function

Public Function AppendLogInfo(infobox As RichTextBox, line As String)
    infobox.SelStart = glInfoTxtLen
    infobox.SelText = line & vbNewLine
    glInfoTxtLen = glInfoTxtLen + Len(line & vbNewLine)
End Function

