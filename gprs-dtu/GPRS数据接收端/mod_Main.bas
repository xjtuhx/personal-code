Attribute VB_Name = "mod_Main"
Option Explicit

'=====================================================
'                       全局变量定义
'=====================================================

Global glConnA As ADODB.Connection     '全局数据库连接
Global glConnB As ADODB.Connection

Global glConnStringA As String '全局连接字符串
Global glConnStringB As String

Global glInfoTxtLen As Long

'==================================================== 主程序
Sub Main()
    '________________________ 初始化全局变量
    Set glConnA = New Connection
    Set glConnB = New Connection
    
   '________________________ 调用并显示主窗体
    Load DataRecvFrm
    SetFormNoClose DataRecvFrm
    DataRecvFrm.Enabled = False
    
    glInfoTxtLen = Len(DataRecvFrm.infoBox.Text)
    
    '________________________ 调用登陆窗体并登录数据库
    Load frmLogin
    frmLogin.Show 1
    
    '________________________ 判断是否连同数据库并作相应的处理
    If Not frmLogin.IfConnDB Then
        MsgBox "由于连接数据库失败，无法继续程序。", vbExclamation, "错误"
        Unload frmLogin
        Unload DataRecvFrm
        End
    End If
    
    '________________________ 启用主窗体
    DataRecvFrm.Show
    DataRecvFrm.Enabled = True
    DataRecvFrm.SetFocus

End Sub
