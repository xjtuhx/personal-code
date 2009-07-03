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

'=========================== 播放声音文件
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long


 Public Const SND_SYNC = &H0
 Public Const SND_ASYNC = &H1
 Public Const SND_NODEFAULT = &H2
 Public Const SND_MEMORY = &H4
 Public Const SND_ALIAS = &H10000
 Public Const SND_FILENAME = &H20000
 Public Const SND_RESOURCE = &H40004
 Public Const SND_ALIAS_ID = &H110000
 Public Const SND_ALIAS_START = 0
 Public Const SND_LOOP = &H8
 Public Const SND_NOSTOP = &H10
 Public Const SND_VALID = &H1F
 Public Const SND_NOWAIT = &H2000
 Public Const SND_VALIDFLAGS = &H17201F
 Public Const SND_RESERVED = &HFF000000
 Public Const SND_TYPE_MASK = &H170007

' SND_SYNC 播 放WAV 文 件， 播 放 完 毕 后 将 控 制 转 移 回 你 的 应 用 程 序 中。
' SND_ASYNC 播 放WAV 文 件， 然 后 将 控 制 立 即 转 移 回 你 的 应 用 程 序 中， 而 不 管 对WAV 文 件 的 播 放 是 否 结 束。
' SND_NODEFAULT 不 要 播 放 缺 省 的WAV 文 件， 以 免 发 生 某 些 意 外 的 错 误。
' SND_MEMORY 播 放 以 前 已 经 加 载 到 内 存 中 的WAV 文 件。
' SND_LOOP 循 环 播 放WAV 文 件。
' SND_NOSTOP 在 开 始 播 放 其 它 的WAV 文 件 之 前， 需 要 完 成 对 本WAV 文 件 的 播 放。
' 注 意：SND_LOOP 标 识 通 常 需 要 同SND_ASYNC 共 同 使 用， 也 即 在 两 个 标 识 之 间 添 加 与 播 放 符， 以 免 在 对WAV 文 件 进 行 播 放 的 时 候 将 系 统 挂 起。
'===========================



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



'============================= 播放一段声音
Public Sub PlaySound(ByVal F As String)

    Dim R As Long
    
    R = sndPlaySound(F, SND_SYNC)
    
End Sub
