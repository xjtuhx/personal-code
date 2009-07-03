Attribute VB_Name = "mod_Communication"
Option Explicit
Global hRasConn As Long '定义一个指向RAS调用的全局句柄
Public Const APINULL = 0&
Public Const UNLEN = 256
Public Const DNLEN = 15
Public Const PWLEN = 256
Public Const RAS95_MaxPhoneNumber = 128
Public Const RAS95_MaxEntryName = 256
Public Const RAS95_MaxCallbackNumber = RAS95_MaxPhoneNumber
Public Type RASDIALPARAMS95
    dwSize As Long
    szEntryName(RAS95_MaxEntryName) As Byte
    szPhoneNumber(RAS95_MaxPhoneNumber) As Byte
    szCallbackNumber(RAS95_MaxCallbackNumber) As Byte
    szUserName(UNLEN) As Byte
    szPassword(PWLEN) As Byte
    szDomain(DNLEN) As Byte
End Type
'**********************************
'* RAS调用错误代号 *
'**********************************
Public Const NOT_SUPPORTED = 120&
Public Const RASBASEERROR = 600&
Public Const SUCCESS = 0&
Public Const ERROR_PORT_ALREADY_OPEN = (RASBASEERROR + 2)
Public Const ERROR_UNKNOWN = (RASBASEERROR + 35)
Public Const ERROR_REQUEST_TIMEOUT = (RASBASEERROR + 38)
Public Const ERROR_PASSWD_EXPIRED = (RASBASEERROR + 48)
Public Const ERROR_NO_DIALIN_PERMISSION = (RASBASEERROR + 49)
Public Const ERROR_SERVER_NOT_RESPONDING = (RASBASEERROR + 50)
Public Const ERROR_UNRECOGNIZED_RESPONSE = (RASBASEERROR + 52)
Public Const ERROR_NO_RESPONSES = (RASBASEERROR + 60)
Public Const ERROR_DEVICE_NOT_READY = (RASBASEERROR + 66)
Public Const ERROR_LINE_BUSY = (RASBASEERROR + 76)
Public Const ERROR_NO_ANSWER = (RASBASEERROR + 78)
Public Const ERROR_NO_CARRIER = (RASBASEERROR + 79)
Public Const ERROR_NO_DIALTONE = (RASBASEERROR + 80)
Public Const ERROR_AUTHENTICATION_FAILURE = (RASBASEERROR + 91)
Public Const ERROR_PPP_TIMEOUT = (RASBASEERROR + 118)
'**********************************
'* RAS API 声明 *
'**********************************
Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (lpString1 As Any, ByVal lpString2 As String) As Long
Public Declare Function RasDial Lib "RasApi32.DLL" Alias "RasDialA" (lpRasDialExtensions As Any, ByVal lpszPhonebook As String, lprasdialparams As Any, ByVal dwNotifierType As Long, lpvNotifier As Long, lphRasConn As Long) As Long
Public Declare Function RasHangUp Lib "RasApi32.DLL" Alias "RasHangUpA" (ByVal hRasConn As Long) As Long

Public Function AddConnection(strNewEntryName As String, strNewPhoneNumber As String, strNewCallbackNumber As String, strNewUsername As String, strNewPassword As String, strNewDomain As String) As Integer

    Dim lngRetCode As Long
    Dim lngRetLstrcpy As Long
    Dim lngRetHangUp As Long
    Dim lprasdialparams As RASDIALPARAMS95
    lprasdialparams.dwSize = 1052 '在WINDOWS95/98中必须将dwSize设为1052
    '利用lstrcpy函数将字符串拷贝到BYTE数组
    lngRetLstrcpy = lstrcpy(lprasdialparams.szEntryName(0), strNewEntryName)
    lngRetLstrcpy = lstrcpy(lprasdialparams.szPhoneNumber(0), strNewPhoneNumber)
    lngRetLstrcpy = lstrcpy(lprasdialparams.szCallbackNumber(0), strNewCallbackNumber)
    lngRetLstrcpy = lstrcpy(lprasdialparams.szUserName(0), strNewUsername)
    lngRetLstrcpy = lstrcpy(lprasdialparams.szPassword(0), strNewPassword)
    lngRetLstrcpy = lstrcpy(lprasdialparams.szDomain(0), strNewDomain)
    '我们使用同步通信
    Screen.MousePointer = vbHourglass
    hRasConn = 0 '
    lngRetCode = RasDial(ByVal APINULL, vbNullString, lprasdialparams, APINULL, ByVal APINULL, hRasConn)
    Screen.MousePointer = vbDefault
    '测试有没有错误
    If lngRetCode Then
        lngRetHangUp = RasHangUp(hRasConn)
    End If
    AddConnection = lngRetCode
End Function

Public Sub RemoveConnection(H_RasConn As Long)
    Call RasHangUp(hRasConn)
End Sub

