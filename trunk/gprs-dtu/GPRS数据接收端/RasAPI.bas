Attribute VB_Name = "mod_RasAPI"
Option Explicit

Private Const RASDT_TYPE_Modem = "modem"
Private Const RASDT_TYPE_Isdn = "isdn"
Private Const RASDT_TYPE_X25 = "x25"
Private Const RASDT_TYPE_Vpn = "vpn"
Private Const RASDT_TYPE_Pad = "pad"
Private Const RASDT_TYPE_Generic = "GENERIC"
Private Const RASDT_TYPE_Serial = "SERIAL"
Private Const RASDT_TYPE_FrameRelay = "FRAMERELAY"
Private Const RASDT_TYPE_Atm = "ATM"
Private Const RASDT_TYPE_Sonet = "SONET"
Private Const RASDT_TYPE_SW56 = "SW56"
Private Const RASDT_TYPE_Irda = "IRDA"
Private Const RASDT_TYPE_Parallel = "PARALLEL"
Private Const RASDT_TYPE_PPPoE = "PPPoE"

Public Enum RASDT_TYPE
    RASDT_Modem = 0
    RASDT_Isdn = 1
    RASDT_X25 = 2
    RASDT_Vpn = 3
    RASDT_Pad = 4
    RASDT_Generic = 5
    RASDT_Serial = 6
    RASDT_FrameRelay = 7
    RASDT_Atm = 8
    RASDT_Sonet = 9
    RASDT_SW56 = 10
    RASDT_Irda = 11
    RASDT_Parallel = 12
    RASDT_PPPoE = 13
End Enum

Private Const IS_TEXT_UNICODE_ASCII16 = &H1
Private Const IS_TEXT_UNICODE_REVERSE_ASCII16 = &H10
Private Const CP_ACP = 0  '  default to ANSI code page
Private Const ERROR_SUCCESS = 0&
Private Const RASP_Amb = &H10000
Private Const RASP_PppNbf = &H803F&
Private Const RASP_PppIpx = &H802B&
Private Const RASP_PppIp = &H8021&
Private Const RASP_PppLcp = &HC021&
Private Const RASP_Slip = &H20000
Private Const RASIDS_Disabled = &HFFFFFFFF
Private Const RDEOPT_PausedStates = &H2
Private Const WM_RASDIALEVENT = &HCCCD&

'RASCONNSTATE enum
Public Const RASCS_PAUSED = &H1000&
Public Const RASCS_DONE = &H2000&
'begin enum
Public Const RASCS_OpenPort = 0&
Public Const RASCS_PortOpened = 1&
Public Const RASCS_ConnectDevice = 2&
Public Const RASCS_DeviceConnected = 3&
Public Const RASCS_AllDevicesConnected = 4&
Public Const RASCS_Authenticate = 5&
Public Const RASCS_AuthNotify = 6&
Public Const RASCS_AuthRetry = 7&
Public Const RASCS_AuthCallback = 8&
Public Const RASCS_AuthChangePassword = 9&
Public Const RASCS_AuthProject = 10&
Public Const RASCS_AuthLinkSpeed = 11&
Public Const RASCS_AuthAck = 12&
Public Const RASCS_ReAuthenticate = 13&
Public Const RASCS_Authenticated = 14&
Public Const RASCS_PrepareForCallback = 15&
Public Const RASCS_WaitForModemReset = 16&
Public Const RASCS_WaitForCallback = 17&
Public Const RASCS_Projected = 18&
Public Const RASCS_StartAuthentication = 19&     'Windows 95 only
Public Const RASCS_CallbackComplete = 20&        'Windows 95 only
Public Const RASCS_LogonNetwork = 21&            'Windows 95 only
Public Const RASCS_Interactive = RASCS_PAUSED
Public Const RASCS_RetryAuthentication = RASCS_PAUSED + 1&
Public Const RASCS_CallbackSetByCaller = RASCS_PAUSED + 2&
Public Const RASCS_PasswordExpired = RASCS_PAUSED + 3&
Public Const RASCS_Connected = RASCS_DONE
Public Const RASCS_Disconnected = RASCS_DONE + 1&
'end enum

'**********************************
'*     RAS Error Return Codes     *
'**********************************
Public Const NOT_SUPPORTED = 120&

Public Const RASBASE = 600&
Public Const SUCCESS = 0&

Public Const PENDING = (RASBASE + 0)
Public Const ERROR_INVALID_PORT_HANDLE = (RASBASE + 1)
Public Const ERROR_PORT_ALREADY_OPEN = (RASBASE + 2)
Public Const ERROR_BUFFER_TOO_SMALL = (RASBASE + 3)
Public Const ERROR_WRONG_INFO_SPECIFIED = (RASBASE + 4)
Public Const ERROR_CANNOT_SET_PORT_INFO = (RASBASE + 5)
Public Const ERROR_PORT_NOT_CONNECTED = (RASBASE + 6)
Public Const ERROR_EVENT_INVALID = (RASBASE + 7)
Public Const ERROR_DEVICE_DOES_NOT_EXIST = (RASBASE + 8)
Public Const ERROR_DEVICETYPE_DOES_NOT_EXIST = (RASBASE + 9)
Public Const ERROR_BUFFER_INVALID = (RASBASE + 10)
Public Const ERROR_ROUTE_NOT_AVAILABLE = (RASBASE + 11)
Public Const ERROR_ROUTE_NOT_ALLOCATED = (RASBASE + 12)
Public Const ERROR_INVALID_COMPRESSION_SPECIFIED = (RASBASE + 13)
Public Const ERROR_OUT_OF_BUFFERS = (RASBASE + 14)
Public Const ERROR_PORT_NOT_FOUND = (RASBASE + 15)
Public Const ERROR_ASYNC_REQUEST_PENDING = (RASBASE + 16)
Public Const ERROR_ALREADY_DISCONNECTING = (RASBASE + 17)
Public Const ERROR_PORT_NOT_OPEN = (RASBASE + 18)
Public Const ERROR_PORT_DISCONNECTED = (RASBASE + 19)
Public Const ERROR_NO_ENDPOINTS = (RASBASE + 20)
Public Const ERROR_CANNOT_OPEN_PHONEBOOK = (RASBASE + 21)
Public Const ERROR_CANNOT_LOAD_PHONEBOOK = (RASBASE + 22)
Public Const ERROR_CANNOT_FIND_PHONEBOOK_ENTRY = (RASBASE + 23)
Public Const ERROR_CANNOT_WRITE_PHONEBOOK = (RASBASE + 24)
Public Const ERROR_CORRUPT_PHONEBOOK = (RASBASE + 25)
Public Const ERROR_CANNOT_LOAD_STRING = (RASBASE + 26)
Public Const ERROR_KEY_NOT_FOUND = (RASBASE + 27)
Public Const ERROR_DISCONNECTION = (RASBASE + 28)
Public Const ERROR_REMOTE_DISCONNECTION = (RASBASE + 29)
Public Const ERROR_HARDWARE_FAILURE = (RASBASE + 30)
Public Const ERROR_USER_DISCONNECTION = (RASBASE + 31)
Public Const ERROR_INVALID_SIZE = (RASBASE + 32)
Public Const ERROR_PORT_NOT_AVAILABLE = (RASBASE + 33)
Public Const ERROR_CANNOT_PROJECT_CLIENT = (RASBASE + 34)
Public Const ERROR_UNKNOWN = (RASBASE + 35)
Public Const ERROR_WRONG_DEVICE_ATTACHED = (RASBASE + 36)
Public Const ERROR_BAD_STRING = (RASBASE + 37)
Public Const ERROR_REQUEST_TIMEOUT = (RASBASE + 38)
Public Const ERROR_CANNOT_GET_LANA = (RASBASE + 39)
Public Const ERROR_NETBIOS_ERROR = (RASBASE + 40)
Public Const ERROR_SERVER_OUT_OF_RESOURCES = (RASBASE + 41)
Public Const ERROR_NAME_EXISTS_ON_NET = (RASBASE + 42)
Public Const ERROR_SERVER_GENERAL_NET_FAILURE = (RASBASE + 43)
Public Const WARNING_MSG_ALIAS_NOT_ADDED = (RASBASE + 44)
Public Const ERROR_AUTH_INTERNAL = (RASBASE + 45)
Public Const ERROR_RESTRICTED_LOGON_HOURS = (RASBASE + 46)
Public Const ERROR_ACCT_DISABLED = (RASBASE + 47)
Public Const ERROR_PASSWD_EXPIRED = (RASBASE + 48)
Public Const ERROR_NO_DIALIN_PERMISSION = (RASBASE + 49)
Public Const ERROR_SERVER_NOT_RESPONDING = (RASBASE + 50)
Public Const ERROR_FROM_DEVICE = (RASBASE + 51)
Public Const ERROR_UNRECOGNIZED_RESPONSE = (RASBASE + 52)
Public Const ERROR_MACRO_NOT_FOUND = (RASBASE + 53)
Public Const ERROR_MACRO_NOT_DEFINED = (RASBASE + 54)
Public Const ERROR_MESSAGE_MACRO_NOT_FOUND = (RASBASE + 55)
Public Const ERROR_DEFAULTOFF_MACRO_NOT_FOUND = (RASBASE + 56)
Public Const ERROR_FILE_COULD_NOT_BE_OPENED = (RASBASE + 57)
Public Const ERROR_DEVICENAME_TOO_LONG = (RASBASE + 58)
Public Const ERROR_DEVICENAME_NOT_FOUND = (RASBASE + 59)
Public Const ERROR_NO_RESPONSES = (RASBASE + 60)
Public Const ERROR_NO_COMMAND_FOUND = (RASBASE + 61)
Public Const ERROR_WRONG_KEY_SPECIFIED = (RASBASE + 62)
Public Const ERROR_UNKNOWN_DEVICE_TYPE = (RASBASE + 63)
Public Const ERROR_ALLOCATING_MEMORY = (RASBASE + 64)
Public Const ERROR_PORT_NOT_CONFIGURED = (RASBASE + 65)
Public Const ERROR_DEVICE_NOT_READY = (RASBASE + 66)
Public Const ERROR_READING_INI_FILE = (RASBASE + 67)
Public Const ERROR_NO_CONNECTION = (RASBASE + 68)
Public Const ERROR_BAD_USAGE_IN_INI_FILE = (RASBASE + 69)
Public Const ERROR_READING_SECTIONNAME = (RASBASE + 70)
Public Const ERROR_READING_DEVICETYPE = (RASBASE + 71)
Public Const ERROR_READING_DEVICENAME = (RASBASE + 72)
Public Const ERROR_READING_USAGE = (RASBASE + 73)
Public Const ERROR_READING_MAXCONNECTBPS = (RASBASE + 74)
Public Const ERROR_READING_MAXCARRIERBPS = (RASBASE + 75)
Public Const ERROR_LINE_BUSY = (RASBASE + 76)
Public Const ERROR_VOICE_ANSWER = (RASBASE + 77)
Public Const ERROR_NO_ANSWER = (RASBASE + 78)
Public Const ERROR_NO_CARRIER = (RASBASE + 79)
Public Const ERROR_NO_DIALTONE = (RASBASE + 80)
Public Const ERROR_IN_COMMAND = (RASBASE + 81)
Public Const ERROR_WRITING_SECTIONNAME = (RASBASE + 82)
Public Const ERROR_WRITING_DEVICETYPE = (RASBASE + 83)
Public Const ERROR_WRITING_DEVICENAME = (RASBASE + 84)
Public Const ERROR_WRITING_MAXCONNECTBPS = (RASBASE + 85)
Public Const ERROR_WRITING_MAXCARRIERBPS = (RASBASE + 86)
Public Const ERROR_WRITING_USAGE = (RASBASE + 87)
Public Const ERROR_WRITING_DEFAULTOFF = (RASBASE + 88)
Public Const ERROR_READING_DEFAULTOFF = (RASBASE + 89)
Public Const ERROR_EMPTY_INI_FILE = (RASBASE + 90)
Public Const ERROR_AUTHENTICATION_FAILURE = (RASBASE + 91)
Public Const ERROR_PORT_OR_DEVICE = (RASBASE + 92)
Public Const ERROR_NOT_BINARY_MACRO = (RASBASE + 93)
Public Const ERROR_DCB_NOT_FOUND = (RASBASE + 94)
Public Const ERROR_STATE_MACHINES_NOT_STARTED = (RASBASE + 95)
Public Const ERROR_STATE_MACHINES_ALREADY_STARTED = (RASBASE + 96)
Public Const ERROR_PARTIAL_RESPONSE_LOOPING = (RASBASE + 97)
Public Const ERROR_UNKNOWN_RESPONSE_KEY = (RASBASE + 98)
Public Const ERROR_RECV_BUF_FULL = (RASBASE + 99)
Public Const ERROR_CMD_TOO_LONG = (RASBASE + 100)
Public Const ERROR_UNSUPPORTED_BPS = (RASBASE + 101)
Public Const ERROR_UNEXPECTED_RESPONSE = (RASBASE + 102)
Public Const ERROR_INTERACTIVE_MODE = (RASBASE + 103)
Public Const ERROR_BAD_CALLBACK_NUMBER = (RASBASE + 104)
Public Const ERROR_INVALID_AUTH_STATE = (RASBASE + 105)
Public Const ERROR_WRITING_INITBPS = (RASBASE + 106)
Public Const ERROR_X25_DIAGNOSTIC = (RASBASE + 107)
Public Const ERROR_ACCT_EXPIRED = (RASBASE + 108)
Public Const ERROR_CHANGING_PASSWORD = (RASBASE + 109)
Public Const ERROR_OVERRUN = (RASBASE + 110)
Public Const ERROR_RASMAN_CANNOT_INITIALIZE = (RASBASE + 111)
Public Const ERROR_BIPLEX_PORT_NOT_AVAILABLE = (RASBASE + 112)
Public Const ERROR_NO_ACTIVE_ISDN_LINES = (RASBASE + 113)
Public Const ERROR_NO_ISDN_CHANNELS_AVAILABLE = (RASBASE + 114)
Public Const ERROR_TOO_MANY_LINE_ERRORS = (RASBASE + 115)
Public Const ERROR_IP_CONFIGURATION = (RASBASE + 116)
Public Const ERROR_NO_IP_ADDRESSES = (RASBASE + 117)
Public Const ERROR_PPP_TIMEOUT = (RASBASE + 118)
Public Const ERROR_PPP_REMOTE_TERMINATED = (RASBASE + 119)
Public Const ERROR_PPP_NO_PROTOCOLS_CONFIGURED = (RASBASE + 120)
Public Const ERROR_PPP_NO_RESPONSE = (RASBASE + 121)
Public Const ERROR_PPP_INVALID_PACKET = (RASBASE + 122)
Public Const ERROR_PHONE_NUMBER_TOO_LONG = (RASBASE + 123)
Public Const ERROR_IPXCP_NO_DIALOUT_CONFIGURED = (RASBASE + 124)
Public Const ERROR_IPXCP_NO_DIALIN_CONFIGURED = (RASBASE + 125)
Public Const ERROR_IPXCP_DIALOUT_ALREADY_ACTIVE = (RASBASE + 126)
Public Const ERROR_ACCESSING_TCPCFGDLL = (RASBASE + 127)
Public Const ERROR_NO_IP_RAS_ADAPTER = (RASBASE + 128)
Public Const ERROR_SLIP_REQUIRES_IP = (RASBASE + 129)
Public Const ERROR_PROJECTION_NOT_COMPLETE = (RASBASE + 130)
Public Const ERROR_PROTOCOL_NOT_CONFIGURED = (RASBASE + 131)
Public Const ERROR_PPP_NOT_CONVERGING = (RASBASE + 132)
Public Const ERROR_PPP_CP_REJECTED = (RASBASE + 133)
Public Const ERROR_PPP_LCP_TERMINATED = (RASBASE + 134)
Public Const ERROR_PPP_REQUIRED_ADDRESS_REJECTED = (RASBASE + 135)
Public Const ERROR_PPP_NCP_TERMINATED = (RASBASE + 136)
Public Const ERROR_PPP_LOOPBACK_DETECTED = (RASBASE + 137)
Public Const ERROR_PPP_NO_ADDRESS_ASSIGNED = (RASBASE + 138)
Public Const ERROR_CANNOT_USE_LOGON_CREDENTIALS = (RASBASE + 139)
Public Const ERROR_TAPI_CONFIGURATION = (RASBASE + 140)
Public Const ERROR_NO_LOCAL_ENCRYPTION = (RASBASE + 141)
Public Const ERROR_NO_REMOTE_ENCRYPTION = (RASBASE + 142)
Public Const ERROR_REMOTE_REQUIRES_ENCRYPTION = (RASBASE + 143)
Public Const ERROR_IPXCP_NET_NUMBER_CONFLICT = (RASBASE + 144)
Public Const ERROR_INVALID_SMM = (RASBASE + 145)
Public Const ERROR_SMM_UNINITIALIZED = (RASBASE + 146)
Public Const ERROR_NO_MAC_FOR_PORT = (RASBASE + 147)
Public Const ERROR_SMM_TIMEOUT = (RASBASE + 148)
Public Const ERROR_BAD_PHONE_NUMBER = (RASBASE + 149)
Public Const ERROR_WRONG_MODULE = (RASBASE + 150)
Public Const RASBASEEND = (RASBASE + 150)

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Type RASIPADDR
    a As Byte
    B As Byte
    c As Byte
    d As Byte
End Type

Private Enum RasEntryOptions
    RASEO_UseCountryAndAreaCodes = &H1
    RASEO_SpecificIpAddr = &H2
    RASEO_SpecificNameServers = &H4
    RASEO_IpHeaderCompression = &H8
    RASEO_RemoteDefaultGateway = &H10
    RASEO_DisableLcpExtensions = &H20
    RASEO_TerminalBeforeDial = &H40
    RASEO_TerminalAfterDial = &H80
    RASEO_ModemLights = &H100
    RASEO_SwCompression = &H200
    RASEO_RequireEncryptedPw = &H400
    RASEO_RequireMsEncryptedPw = &H800
    RASEO_RequireDataEncryption = &H1000
    RASEO_NetworkLogon = &H2000
    RASEO_UseLogonCredentials = &H4000
    RASEO_PromoteAlternates = &H8000
    RASEO_SecureLocalFiles = &H10000
    RASEO_RequireEAP = &H20000
    RASEO_RequirePAP = &H40000
    RASEO_RequireSPAP = &H80000
    RASEO_Custom = &H100000
    RASEO_PreviewPhoneNumber = &H200000
    RASEO_SharedPhoneNumbers = &H800000
    RASEO_PreviewUserPw = &H1000000
    RASEO_PreviewDomain = &H2000000
    RASEO_ShowDialingProgress = &H4000000
    RASEO_RequireCHAP = &H8000000
    RASEO_RequireMsCHAP = &H10000000
    RASEO_RequireMsCHAP2 = &H20000000
    RASEO_RequireW95MSCHAP = &H40000000
    RASEO_CustomScript = &H80000000
End Enum

Private Enum RASNetProtocols
    RASNP_NetBEUI = &H1
    RASNP_Ipx = &H2
    RASNP_Ip = &H4
End Enum

Private Enum RasFramingProtocols
    RASFP_Ppp = &H1
    RASFP_Slip = &H2
    RASFP_Ras = &H4
End Enum

Public Enum RasType
    RASET_Phone = 1
    RASET_Vpn = 2
    RASET_Direct = 3
    RASET_Internet = 4
    RASET_Broadband = 5
End Enum

Public Enum VpnStrategy
    VS_Default = 0
    VS_PptpOnly = 1
    VS_PptpFirst = 2
    VS_L2tpOnly = 3
    VS_L2tpFirst = 4
End Enum

Private Enum RASCredMask
    RASCM_UserName = &H1&
    RASCM_Password = &H2&
    RASCM_Domain = &H4&
    RASCM_DefaultCreds = &H8&
    RASCM_PreSharedKey = &H10&
    RASCM_ServerPreSharedKey = &H20&
    RASCM_DDMPreSharedKey = &H40&
End Enum

Private Type RASENTRY
    dwSize As Long
    dwfOptions As RasEntryOptions
    dwCountryID As Long
    dwCountryCode As Long
    szAreaCode(10) As Byte
    szLocalPhoneNumber(128) As Byte
    dwAlternateOffset As Long
    ipaddr As RASIPADDR
    ipaddrDns As RASIPADDR
    ipaddrDnsAlt As RASIPADDR
    ipaddrWins As RASIPADDR
    ipaddrWinsAlt As RASIPADDR
    dwFrameSize As Long
    dwfNetProtocols As RASNetProtocols
    dwFramingProtocol As RasFramingProtocols
    szScript(259) As Byte
    szAutodialDll(259) As Byte
    szAutodialFunc(259) As Byte
    szDeviceType(16) As Byte
    szDeviceName(128) As Byte
    szX25PadType(32) As Byte
    szX25Address(200) As Byte
    szX25Facilities(200) As Byte
    szX25UserData(200) As Byte
    dwChannels As Long
    dwReserved1 As Long
    dwReserved2 As Long
    dwSubEntries As Long
    dwDialMode As Long
    dwDialExtraPercent As Long
    dwDialExtraSampleSeconds As Long
    dwHangUpExtraPercent As Long
    dwHangUpExtraSampleSeconds As Long
    dwIdleDisconnectSeconds As Long
    dwType As RasType
    dwEncryptionType As Long
    dwCustomAuthKey As Long
    guidId As GUID
    szCustomDialDll(259) As Byte
    dwVpnStrategy As Long
    dwfOptions2 As Long
    dwfOptions3 As Long
    szDnsSuffix(255) As Byte
    dwTcpWindowSize As Long
    szPrerequisitePbk(259) As Byte
    szPrerequisiteEntry(256) As Byte
    dwRedialCount As Long
    dwRedialPause As Long
End Type

Private Type RASCREDENTIALS
    dwSize As Long
    dwMask As RASCredMask
    szUserName(256) As Byte
    szPassword(256) As Byte
    szDomain(15) As Byte
End Type

Private Type RASDIALEXTENSIONS
    'set dwsize to 16
    dwSize As Long
    dwfOptions As Long
    hwndParent As Long
    Reserved As Long
End Type

Private Type RASDIALPARAMS
    'set dwsize to 736 unless winver >= 400 then set to 1052
    dwSize As Long
    szEntryName(20) As Byte
    szPhoneNumber(128) As Byte
    szCallbackNumber(128) As Byte
    szUserName(256) As Byte
    szPassword(256) As Byte
    szDomain(15) As Byte
End Type

Private Type RASCONN
    'set dwsize to 32
    dwSize As Long
    hRasConn As Long
    szEntryName(20) As Byte
End Type

Private Type RAS_STATS
    dwSize As Long
    dwBytesXmited As Long
    dwBytesRcved As Long
    dwFramesXmited As Long
    dwFramesRcved As Long
    dwCrcErr As Long
    dwTimeoutErr As Long
    dwAlignmentErr As Long
    dwHardwareOverrunErr As Long
    dwFramingErr As Long
    dwBufferOverrunErr As Long
    dwCompressionRatioIn As Long
    dwCompressionRatioOut As Long
    dwBps As Long
    dwConnectDuration As Long
End Type

Private Type RASCONNSTATUS
    'set dwsize to 64 unless winver >= 400 then set to 288
    dwSize As Long
    rasconnstate As Long                            'RASCONNSTATE Enumeration
    dwError As Long
    szDeviceType(16) As Byte
    szDeviceName(32) As Byte
End Type

Private Type RASPPPIP
    'set dwsize to 40
    dwSize As Long
    dwError As Long
    szIpAddress(15) As Byte
    szServerAddress(15) As Byte
End Type

Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (lpString1 As Any, ByVal lpString2 As String) As Long
Private Declare Function RasGetEntryProperties Lib "rasapi32" Alias "RasGetEntryPropertiesA" (ByVal lpszPhonebook As String, ByVal lpszEntry As String, lpRasEntry As RASENTRY, ByVal dwEntryInfoSize As Long, ByVal lpbDeviceInfo As Long, ByVal dwDeviceInfoSize As Long) As Long
Private Declare Function RasSetEntryProperties Lib "rasapi32" Alias "RasSetEntryPropertiesA" (ByVal lpszPhonebook As String, ByVal lpszEntry As String, lpRasEntry As RASENTRY, ByVal dwEntryInfoSize As Long, ByVal lpbDeviceInfo As Long, ByVal dwDeviceInfoSize As Long) As Long
Private Declare Function RasGetCredentials Lib "rasapi32" Alias "RasGetCredentialsA" (ByVal lpszPhonebook As String, ByVal lpszEntry As String, lpCredentials As RASCREDENTIALS) As Long
Private Declare Function RasSetCredentials Lib "rasapi32" Alias "RasSetCredentialsA" (ByVal lpszPhonebook As String, ByVal lpszEntry As String, lpCredentials As RASCREDENTIALS, ByVal fClearCredentials As Long) As Long
Private Declare Function RasDeleteEntry Lib "rasapi32" Alias "RasDeleteEntryA" (ByVal lpszPhonebook As String, ByVal lpszEntry As String) As Long
Private Declare Function RasGetEntryDialParams Lib "RasApi32.DLL" Alias "RasGetEntryDialParamsA" (ByVal lpszPhonebook As String, lprasdialparams As RASDIALPARAMS, lpfPassword As Long) As Long
Private Declare Function RasSetEntryDialParams Lib "RasApi32.DLL" Alias "RasSetEntryDialParamsA" (ByVal lpszPhonebook As String, lprasdialparams As RASDIALPARAMS, ByVal fRemovePassword As Long) As Long
Private Declare Function RasDial Lib "RasApi32.DLL" Alias "RasDialA" (lpRasDialExtensions As RASDIALEXTENSIONS, ByVal lpszPhonebook As String, lprasdialparams As RASDIALPARAMS, ByVal dwNotifierType As Long, ByVal lpvNotifier As Long, lphRasConn As Long) As Long
Private Declare Function RasHangUp Lib "RasApi32.DLL" Alias "RasHangUpA" (ByVal hRasConn As Long) As Long
Private Declare Function RasEnumConnections Lib "RasApi32.DLL" Alias "RasEnumConnectionsA" (lpRasConn As Any, lpcb As Long, lpcConnections As Long) As Long
Private Declare Function RasGetConnectionStatistics Lib "RasApi32.DLL" (ByVal hRasConn As Long, lpStatistics As RAS_STATS) As Long
Private Declare Function RasGetConnectStatus Lib "RasApi32.DLL" Alias "RasGetConnectStatusA" (ByVal hRasConn As Long, lprasconnstatus As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function IsTextUnicode Lib "advapi32" (ByVal lpBuffer As Long, ByVal cb As Long, ByVal lpi As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function RasGetProjectionInfo Lib "RasApi32.DLL" Alias "RasGetProjectionInfoA" (ByVal hRasConn As Long, ByVal rasprojection As Long, lpprojection As Any, lpcb As Long) As Long
Private Declare Function RasGetErrorString Lib "RasApi32.DLL" Alias "RasGetErrorStringA" (ByVal uErrorValue As Long, ByVal lpszErrorString As String, ByVal cBufSize As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Public hRasConn As Long
Public RASDialErrorCode As Long

Private Function ExchangeRASType(ByVal Index As Long) As String
    Dim tempStr As String
    
    Select Case Index
        Case RASDT_Modem
            tempStr = RASDT_TYPE_Modem
        Case RASDT_Isdn
            tempStr = RASDT_TYPE_Isdn
        Case RASDT_X25
            tempStr = RASDT_TYPE_X25
        Case RASDT_Vpn
            tempStr = RASDT_TYPE_Vpn
        Case RASDT_Pad
            tempStr = RASDT_TYPE_Pad
        Case RASDT_Generic
            tempStr = RASDT_TYPE_Generic
        Case RASDT_Serial
            tempStr = RASDT_TYPE_Serial
        Case RASDT_FrameRelay
            tempStr = RASDT_TYPE_FrameRelay
        Case RASDT_Atm
            tempStr = RASDT_TYPE_Atm
        Case RASDT_Sonet
            tempStr = RASDT_TYPE_Sonet
        Case RASDT_SW56
            tempStr = RASDT_TYPE_SW56
        Case RASDT_Irda
            tempStr = RASDT_TYPE_Irda
        Case RASDT_Parallel
            tempStr = RASDT_TYPE_Parallel
        Case RASDT_PPPoE
            tempStr = RASDT_TYPE_PPPoE
        Case Else
            tempStr = ""
    End Select
    
    ExchangeRASType = tempStr
End Function

Public Function Create_PPP_Connection(ByVal sEntryName As String, ByVal dwRasType As RasType, ByVal DialVpnStrategy As VpnStrategy, ByVal sPhoneNumber As String, ByVal sUsername As String, ByVal sPassword As String, ByVal sDeviceName As String, ByVal sDeviceType As RASDT_TYPE, ByVal sUseIP As Long, ByVal sIPAddr As String, ByVal sUseDNS As Long, ByVal sDNS1 As String, ByVal sDNS2 As String, ByVal sUseDialRules As Long, ByVal sCountryID As String, ByVal sAreaCode As String) As Boolean
    Dim rEntry As RASENTRY
    Dim rCredential As RASCREDENTIALS

    Create_PPP_Connection = False
    
    With rEntry
        .dwSize = LenB(rEntry)
        .dwfOptions = RASEO_IpHeaderCompression Or RASEO_RemoteDefaultGateway Or RASEO_ModemLights Or RASEO_SwCompression Or RASEO_NetworkLogon Or RASEO_PreviewUserPw Or RASEO_ShowDialingProgress
        
        If sPhoneNumber <> "" Then
            .dwfOptions = .dwfOptions Or RASEO_PreviewPhoneNumber
            lstrcpy .szLocalPhoneNumber(0), ByVal sPhoneNumber
        End If
        
        lstrcpy .szDeviceName(0), ByVal sDeviceName
        lstrcpy .szDeviceType(0), ByVal ExchangeRASType(sDeviceType)
        
        If sUseIP <> 0 Then
            .dwfOptions = .dwfOptions Or RASEO_SpecificIpAddr
            .ipaddr = TransferAddrFromStr(ByVal sIPAddr)
        End If
        
        If sUseDNS <> 0 Then
            .dwfOptions = .dwfOptions Or RASEO_SpecificNameServers
            .ipaddrDns = TransferAddrFromStr(ByVal sDNS1)
            .ipaddrDnsAlt = TransferAddrFromStr(ByVal sDNS2)
        End If
        
        If sUseDialRules <> 0 Then
            .dwfOptions = .dwfOptions Or RASEO_UseCountryAndAreaCodes
            .dwCountryCode = Val(sCountryID)
            .dwCountryID = Val(sCountryID)
            lstrcpy .szAreaCode(0), ByVal sAreaCode
        End If
        
        .dwfNetProtocols = RASNP_Ip
        .dwFramingProtocol = RASFP_Ppp
        .dwRedialCount = 0
        .dwRedialPause = 60
        .dwIdleDisconnectSeconds = RASIDS_Disabled
        .dwType = dwRasType
        .dwVpnStrategy = DialVpnStrategy
        .dwTcpWindowSize = 0
    End With
    
    With rCredential
        .dwSize = LenB(rCredential)
        .dwMask = RASCM_UserName Or RASCM_Password Or RASCM_DefaultCreds
        lstrcpy .szUserName(0), ByVal sUsername
        lstrcpy .szPassword(0), ByVal sPassword
    End With
    
    If RasSetEntryProperties(vbNullString, sEntryName, rEntry, LenB(rEntry), 0, 0) = 0 Then
        If RasSetCredentials(vbNullString, sEntryName, rCredential, 0) = 0 Then
            Create_PPP_Connection = True
        End If
    End If
End Function

Private Function TransferAddrFromStr(ByVal sIPAddr As String) As RASIPADDR
    Dim sNum As Long
    Dim tempa As String
    Dim tempb As String
    Dim tempc As String
    Dim tempd As String
    
    If sIPAddr = "" Then
        GoTo TransferError
    End If
    
    sNum = InStr(sIPAddr, ".")
    If sNum > 0 Then
        tempa = Left(sIPAddr, sNum - 1)
        sIPAddr = Right(sIPAddr, Len(sIPAddr) - sNum)
    Else
        GoTo TransferError
    End If
    
    sNum = InStr(sIPAddr, ".")
    If sNum > 0 Then
        tempb = Left(sIPAddr, sNum - 1)
        sIPAddr = Right(sIPAddr, Len(sIPAddr) - sNum)
    Else
        GoTo TransferError
    End If
    
    sNum = InStr(sIPAddr, ".")
    If sNum > 0 Then
        tempc = Left(sIPAddr, sNum - 1)
        tempd = Right(sIPAddr, Len(sIPAddr) - sNum)
    Else
        GoTo TransferError
    End If
    
    With TransferAddrFromStr
        .a = Val(tempa)
        .B = Val(tempb)
        .c = Val(tempc)
        .d = Val(tempd)
    End With
    Exit Function
    
TransferError:
    With TransferAddrFromStr
        .a = 0
        .B = 0
        .c = 0
        .d = 0
    End With
    Exit Function
End Function

Public Function Delete_PPP_Connection(ByVal lpszEntry As String) As Boolean
    Delete_PPP_Connection = False
    
    If RasDeleteEntry(vbNullString, ByVal lpszEntry) = ERROR_SUCCESS Then
        Delete_PPP_Connection = True
    End If
End Function

Public Function Exists_PPP_Connection(ByVal lpszEntryName As String)
    Dim lprasdialparams As RASDIALPARAMS
    Dim lpRasDialExtensions As RASDIALEXTENSIONS
    Dim lpfPassword As Long
    
    Exists_PPP_Connection = True
    
    If Is_PPP_Connecting(ByVal lpszEntryName) = False Then
        With lpRasDialExtensions
            .dwSize = 16
            .dwfOptions = RDEOPT_PausedStates
            .hwndParent = vbNull
        End With
        
        With lprasdialparams
            .dwSize = 1052
            lstrcpy .szEntryName(0), ByVal lpszEntryName
        End With
        
        If RasGetEntryDialParams(vbNullString, lprasdialparams, lpfPassword) <> 0 Then
            Exists_PPP_Connection = False
        End If
    End If
End Function

Public Function Dial_PPP_Connection(ByVal lpszEntryName As String) As Boolean
    Dim lprasdialparams As RASDIALPARAMS
    Dim lpRasDialExtensions As RASDIALEXTENSIONS
    Dim lpfPassword As Long
    
    Dial_PPP_Connection = False
    
    If Is_PPP_Connecting(ByVal lpszEntryName) = False Then
        With lpRasDialExtensions
            .dwSize = 16
            .dwfOptions = RDEOPT_PausedStates
            .hwndParent = vbNull
        End With
        
        With lprasdialparams
            .dwSize = 1052
            lstrcpy .szEntryName(0), ByVal lpszEntryName
        End With
        
        hRasConn = 0
        RASDialErrorCode = 0
        If RasGetEntryDialParams(vbNullString, lprasdialparams, lpfPassword) = 0 Then
            If RasDial(lpRasDialExtensions, vbNullString, lprasdialparams, 1&, AddressOf RasDialFunc, hRasConn) = 0 Then
                Dial_PPP_Connection = True
            End If
        End If
    End If
End Function

Public Function Disconnect_PPP_Connection(ByVal lpszEntryName As String) As Boolean
    Dim temp As Long
    
    Disconnect_PPP_Connection = False
    
    Is_PPP_Connecting ByVal lpszEntryName
    
    If hRasConn <> 0 Then
        If RasHangUp(hRasConn) = 0 Then
            temp = GetTickCount()
            Do Until GetTickCount - temp >= 2000
                DoEvents
            Loop
            hRasConn = 0
            Disconnect_PPP_Connection = True
        End If
    End If
End Function

Public Function Is_PPP_Connecting(ByVal lpszEntryName As String) As Boolean
    Dim lpRasConn(255) As RASCONN
    Dim lpcConnections As Long
    Dim temp As String
    Dim i As Long
    
    Is_PPP_Connecting = False
    
    lpRasConn(0).dwSize = 32
    If RasEnumConnections(lpRasConn(0), LenB(lpRasConn(0)) * 256, lpcConnections) = 0 Then
        For i = 0 To lpcConnections - 1
            temp = ReadStringFromMemory(ByVal VarPtr(lpRasConn(i).szEntryName(0)), 21)
            Do While Right(temp, 1) = Chr(0)
                temp = Left(temp, Len(temp) - 1)
            Loop
            If temp = lpszEntryName Then
                hRasConn = lpRasConn(i).hRasConn
                Is_PPP_Connecting = True
                Exit Function
            End If
        Next
    End If
End Function

Private Function ReadStringFromMemory(ByVal Memory_Address As Long, ByVal Menory_Length As Long, Optional ByRef Text_Ascii As Boolean) As String
    Dim tmpBuffer() As Byte
    Dim tmpB As Byte
    Dim actSize As Long
    Dim idx As Long
    Dim tempNum As Long
    Dim tempStr As String
    
    ReDim tmpBuffer(Menory_Length * 2 + 2) As Byte
    actSize = MultiByteToWideChar(CP_ACP, 0, ByVal Memory_Address, ByVal Menory_Length, ByVal VarPtr(tmpBuffer(0)), ByVal (Menory_Length * 2 + 2))
    
    For idx = 0 To actSize * 2 - 1 Step 2
        CopyMemory ByVal VarPtr(tempNum), ByVal VarPtr(tmpBuffer(idx)), 2
        tempStr = tempStr + ChrW(tempNum)
        
        tmpB = tmpBuffer(idx)
        tmpBuffer(idx) = tmpBuffer(idx + 1)
        tmpBuffer(idx + 1) = tmpB
    Next
    If Right(tempStr, 1) = Chr(0) Then
        tempStr = Left(tempStr, Len(tempStr) - 1)
    End If
    
    If Not (IsMissing(Text_Ascii)) Then
        tempNum = IS_TEXT_UNICODE_REVERSE_ASCII16
        IsTextUnicode ByVal VarPtr(tmpBuffer(0)), ByVal actSize * 2, ByVal VarPtr(tempNum)
        tempNum = (tempNum And IS_TEXT_UNICODE_REVERSE_ASCII16)
        If tempNum = IS_TEXT_UNICODE_REVERSE_ASCII16 Then
            Text_Ascii = True
        Else
            Text_Ascii = False
        End If
    End If
    
    ReadStringFromMemory = tempStr
End Function

Public Function Get_PPP_Duration(ByVal lpszEntryName As String) As Long
    Dim lpStatistics As RAS_STATS
    
    If Is_PPP_Connecting(ByVal lpszEntryName) = True Then
        lpStatistics.dwSize = Len(lpStatistics)
        If RasGetConnectionStatistics(ByVal hRasConn, lpStatistics) = 0 Then
            Get_PPP_Duration = lpStatistics.dwConnectDuration
        End If
    End If
End Function

Public Function Get_PPP_TXByte(ByVal lpszEntryName As String) As Long
    Dim lpStatistics As RAS_STATS
    
    If Is_PPP_Connecting(ByVal lpszEntryName) = True Then
        lpStatistics.dwSize = Len(lpStatistics)
        If RasGetConnectionStatistics(ByVal hRasConn, lpStatistics) = 0 Then
            Get_PPP_TXByte = lpStatistics.dwBytesXmited
        End If
    End If
End Function

Public Function Get_PPP_RXByte(ByVal lpszEntryName As String) As Long
    Dim lpStatistics As RAS_STATS
    
    If Is_PPP_Connecting(ByVal lpszEntryName) = True Then
        lpStatistics.dwSize = Len(lpStatistics)
        If RasGetConnectionStatistics(ByVal hRasConn, lpStatistics) = 0 Then
            Get_PPP_RXByte = lpStatistics.dwBytesRcved
        End If
    End If
End Function

Public Function Get_PPP_Status(ByVal lpszEntryName As String) As Long
    Dim lprasconnstatus As RASCONNSTATUS
    
    If Is_PPP_Connecting(ByVal lpszEntryName) = True Then
        lprasconnstatus.dwSize = LenB(lprasconnstatus)
        If RasGetConnectStatus(ByVal hRasConn, lprasconnstatus) = 0 Then
            Get_PPP_Status = lprasconnstatus.rasconnstate
        End If
    Else
        Get_PPP_Status = RASCS_Disconnected
    End If
End Function

Public Function Get_Client_PPP_IPAddress(ByVal lpszEntryName As String) As String
    Dim lpraspppip As RASPPPIP
    Dim rasprojection As Long
    Dim lpcb As Long
    Dim tempIP As String
    
    If Is_PPP_Connecting(ByVal lpszEntryName) = True Then
        rasprojection = RASP_PppIp
        lpraspppip.dwSize = 40
        lpcb = 40
        If RasGetProjectionInfo(hRasConn, rasprojection, lpraspppip, lpcb) = 0 Then
            tempIP = ReadStringFromMemory(ByVal VarPtr(lpraspppip.szIpAddress(0)), 16)
            Do While Right(tempIP, 1) = Chr(0)
                tempIP = Left(tempIP, Len(tempIP) - 1)
            Loop
            Get_Client_PPP_IPAddress = tempIP
        End If
    End If
End Function

Public Function Get_Server_PPP_IPAddress(ByVal lpszEntryName As String) As String
    Dim lpraspppip As RASPPPIP
    Dim rasprojection As Long
    Dim lpcb As Long
    Dim tempIP As String
    
    If Is_PPP_Connecting(ByVal lpszEntryName) = True Then
        rasprojection = RASP_PppIp
        lpraspppip.dwSize = 40
        lpcb = 40
        If RasGetProjectionInfo(hRasConn, rasprojection, lpraspppip, lpcb) = 0 Then
            tempIP = ReadStringFromMemory(ByVal VarPtr(lpraspppip.szServerAddress(0)), 16)
            Do While Right(tempIP, 1) = Chr(0)
                tempIP = Left(tempIP, Len(tempIP) - 1)
            Loop
            Get_Server_PPP_IPAddress = tempIP
        End If
    End If
End Function

Public Function Get_Dial_Error_String(ByVal dwError As Long) As String
    Dim strRASErrorString As String
    
    strRASErrorString = Space(256)
    If RasGetErrorString(ByVal dwError, strRASErrorString, ByVal 256&) = 0 Then
        Do While Right(strRASErrorString, 1) = Chr(0)
            strRASErrorString = Left(strRASErrorString, Len(strRASErrorString) - 1)
        Loop
        Get_Dial_Error_String = strRASErrorString
    End If
End Function

Private Sub RasDialFunc(ByVal lpRasConn As Long, ByVal unMsg As Long, ByVal RasConnectStatus As Long, ByVal dwError As Long, ByVal dwExtendedError As Long)
    hRasConn = lpRasConn
    RASDialErrorCode = dwError
    
    Select Case RasConnectStatus
        Case RASCS_OpenPort
             DataRecvFrm.AppendInfoLine ("正在尝试打开通信端口")
        Case RASCS_PortOpened
             DataRecvFrm.AppendInfoLine ("通信端口已打开")
        Case RASCS_ConnectDevice
             DataRecvFrm.AppendInfoLine ("正在尝试连接设备")
        Case RASCS_DeviceConnected
             DataRecvFrm.AppendInfoLine ("已成功连接到设备")
        Case RASCS_AllDevicesConnected
             DataRecvFrm.AppendInfoLine ("物理链路已建立")
        Case RASCS_Authenticate
             DataRecvFrm.AppendInfoLine ("准备验证用户名和密码")
        Case RASCS_AuthNotify
             DataRecvFrm.AppendInfoLine ("正在验证用户名和密码")
        Case RASCS_AuthRetry
             DataRecvFrm.AppendInfoLine ("服务器已请求尝试另一个身份验证")
        Case RASCS_AuthCallback
             DataRecvFrm.AppendInfoLine ("服务器已请求一个回拨号码")
        Case RASCS_AuthChangePassword
             DataRecvFrm.AppendInfoLine ("已请求改变账号上的密码")
        Case RASCS_AuthProject
             DataRecvFrm.AppendInfoLine ("正在注册网络")
        Case RASCS_AuthLinkSpeed
             DataRecvFrm.AppendInfoLine ("正在计算链路速率")
        Case RASCS_AuthAck
             DataRecvFrm.AppendInfoLine ("身份验证请求正在确认中")
        Case RASCS_ReAuthenticate
             DataRecvFrm.AppendInfoLine ("准备开始回拨之后的身份验证")
        Case RASCS_Authenticated
             DataRecvFrm.AppendInfoLine ("已通过身份验证")
        Case RASCS_PrepareForCallback
             DataRecvFrm.AppendInfoLine ("线路即将取消连接，准备回拨")
        Case RASCS_WaitForModemReset
             DataRecvFrm.AppendInfoLine ("准备回拨之前，客户端等待调制解调器的重新设置")
        Case RASCS_WaitForCallback
             DataRecvFrm.AppendInfoLine ("等待服务器发出的接入拨号")
        Case RASCS_Projected
             DataRecvFrm.AppendInfoLine ("网络注册成功")
        Case RASCS_StartAuthentication      'Windows 95 only
             DataRecvFrm.AppendInfoLine ("用户身份验证已开始或已完成")
        Case RASCS_CallbackComplete         'Windows 95 only
             DataRecvFrm.AppendInfoLine ("已经回拨客户端")
        Case RASCS_LogonNetwork             'Windows 95 only
             DataRecvFrm.AppendInfoLine ("正在登录远程网络")
        Case RASCS_Interactive
             DataRecvFrm.AppendInfoLine ("拨号程序正在等待超级终端")
        Case RASCS_RetryAuthentication
             DataRecvFrm.AppendInfoLine ("拨号程序正在等待新的用户凭据")
        Case RASCS_CallbackSetByCaller
             DataRecvFrm.AppendInfoLine ("拨号程序正在等待客户端的回拨号码")
        Case RASCS_PasswordExpired
             DataRecvFrm.AppendInfoLine ("拨号程序希望用户提供一个新密码")
        Case RASCS_Connected
             DataRecvFrm.AppendInfoLine ("连接成功")
             With DataRecvFrm
                .ppp_status_timer.Enabled = True
                .toolBar.Buttons(BTN_CONNECT).Enabled = False
                .toolBar.Buttons(BTN_DISCONN).Enabled = True
                .toolBar.Buttons(BTN_START).Enabled = True
                .toolBar.Buttons(BTN_STOP).Enabled = False
             End With
        Case RASCS_Disconnected
             DataRecvFrm.AppendInfoLine ("连接失败")
    End Select
    
    If dwError <> 0 Then
        DataRecvFrm.statusBar.Panels(1) = DataRecvFrm.statusBar.Panels(1) & "，错误代码：" & CStr(dwError)
    End If
End Sub
