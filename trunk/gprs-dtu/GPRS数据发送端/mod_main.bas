Attribute VB_Name = "mod_Main"
Option Explicit

'=====================================================
'                       全局变量定义
'=====================================================

Global glConnA As ADODB.Connection     '全局数据库连接
Global glConnB As ADODB.Connection

Global glCenterDial As String   '全局数据中心拨号网络
Global glCenterIP As String     '全局数据中心 IP 地址

Global glConnStringA As String '全局连接字符串
Global glConnStringB As String

Global glInfoTxtLen As Integer

'==================================================== 主程序
Sub Main()
    '________________________ 初始化全局变量
    Set glConnA = New Connection
    Set glConnB = New Connection
    
    '________________________ 读取数据中心信息
    'glCenterDial = GetProfileString(App.Path & "\control.ini", "数据中心信息", "拨号网络")
    'glCenterIP = GetProfileString(App.Path & "\control.ini", "数据中心信息", "数据中心IP")
    'glLocalIP = GetProfileString(App.Path & "\control.ini", "局域网信息", "IP")
    '________________________ 调用并显示主窗体
    Load DataSendFrm
    SetFormNoClose DataSendFrm
    DataSendFrm.Enabled = False
    
    glInfoTxtLen = Len(DataSendFrm.infoBox.Text)
    
    '________________________ 调用登陆窗体并登录数据库
    Load frmLogin
    frmLogin.Show 1
    
    '________________________ 判断是否连同数据库并作相应的处理
    If Not frmLogin.IfConnDB Then
        MsgBox "由于连接数据库失败，无法继续程序。", vbExclamation, "错误"
        Unload frmLogin
        Unload DataSendFrm
        End
    End If
    
    '________________________ 启用主窗体
    DataSendFrm.Show
    DataSendFrm.Enabled = True
    DataSendFrm.SetFocus


    '________________________ 控制服务状态

End Sub
