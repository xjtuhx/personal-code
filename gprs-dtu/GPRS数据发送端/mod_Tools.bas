Attribute VB_Name = "mod_Tools"
Option Explicit

'========================= 窗体最前

Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10

'==========================


'==========================  目录对话框
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


'=========================== 禁用窗体关闭按钮
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long

Const SC_CLOSE = &HF060

'----------------------------------------- 窗体最前
Public Sub SetFormOnTop(myForm As Object)
    SetWindowPos myForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub


'----------------------------------------- 取消最前
Public Sub SetFormNoTop(myForm As Object)
    SetWindowPos myForm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub


'----------------------------------------- 密码加密、解密
Public Function xorPWD(ByVal s As String) As String
    Dim temp As String
    Dim I
        temp = ""
        For I = 1 To Len(s)
            temp = temp + Chr(Asc(Mid(s, I, 1)) Xor 13)
        Next I
        xorPWD = temp
End Function


'===============================  选取文件夹函数

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


'============================== 禁用关闭按钮
Public Sub SetFormNoClose(myForm As Object)
    
    Dim hMenu As Long
    
    hMenu = GetSystemMenu(myForm.hwnd, 0)
    RemoveMenu hMenu, &HF060, &H200&
    
End Sub

