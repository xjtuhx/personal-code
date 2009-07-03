Attribute VB_Name = "mod_ReadINI"
Option Explicit


'_______ 声明读取 INI 文件的 API 函数

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'_______ 读取 INI 文件字符串

Public Function GetProfileString(ByVal INIFile As String, ByVal SectionName As String, ByVal KeyName As String, Optional DefaultValues As String = "") As String

    Dim Values As String * 255
    Dim N As Integer
    
    N = GetPrivateProfileString(SectionName, KeyName, DefaultValues, Values, Len(Values), INIFile)
    
    GetProfileString = Left(Values, N)
    
End Function

'_______ 写入 INI 文件字符串
Public Function WriteProFileString(ByVal INIFile As String, ByVal SectionName As String, ByVal KeyName As String, ByVal Para As String)

    Dim N As Integer
    
    N = WritePrivateProfileString(SectionName, KeyName, Para, INIFile)

End Function


