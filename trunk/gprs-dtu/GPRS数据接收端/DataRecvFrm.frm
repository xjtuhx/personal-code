VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form DataRecvFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GPRS数据接收端"
   ClientHeight    =   7080
   ClientLeft      =   1575
   ClientTop       =   2145
   ClientWidth     =   9810
   Icon            =   "DataRecvFrm.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   9810
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer ppp_status_timer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3120
      Top             =   6000
   End
   Begin MSWinsockLib.Winsock Listener 
      Left            =   2520
      Top             =   6000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   56789
   End
   Begin MSWinsockLib.Winsock Sock 
      Index           =   0
      Left            =   600
      Top             =   6000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   56789
   End
   Begin MSComctlLib.ImageList imageBar 
      Left            =   1200
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataRecvFrm.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataRecvFrm.frx":191C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataRecvFrm.frx":256E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataRecvFrm.frx":31C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataRecvFrm.frx":3E12
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataRecvFrm.frx":4AEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataRecvFrm.frx":57C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataRecvFrm.frx":64A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataRecvFrm.frx":70F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataRecvFrm.frx":7DCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataRecvFrm.frx":8AA6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1920
      Top             =   6000
   End
   Begin VB.Frame Frame2 
      Caption         =   "状态信息"
      Height          =   5655
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   9615
      Begin RichTextLib.RichTextBox infoBox 
         Height          =   5295
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   9340
         _Version        =   393217
         ScrollBars      =   3
         TextRTF         =   $"DataRecvFrm.frx":96F8
      End
   End
   Begin MSComctlLib.StatusBar statusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   6705
      Width           =   9810
      _ExtentX        =   17304
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12118
            Text            =   "状态信息"
            TextSave        =   "状态信息"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "2009-7-20"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Text            =   "显示时间"
            TextSave        =   "显示时间"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar toolBar 
      Align           =   1  'Align Top
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9810
      _ExtentX        =   17304
      _ExtentY        =   1508
      ButtonWidth     =   2090
      ButtonHeight    =   1349
      Appearance      =   1
      ImageList       =   "imageBar"
      DisabledImageList=   "imageBar"
      HotImageList    =   "imageBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "连接有线网络"
            Key             =   "连接有线网络"
            Description     =   "连接有线网络"
            Object.ToolTipText     =   "通过Modem拨号连接到互联网"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "断开有线网络"
            Key             =   "断开有线网络"
            Description     =   "断开有线网络"
            Object.ToolTipText     =   "断开网络连接"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "开始接收"
            Key             =   "开始接收"
            Description     =   "开始接收"
            Object.ToolTipText     =   "开始接受发送端数据"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "停止接收"
            Key             =   "停止接收"
            Description     =   "停止接收"
            Object.ToolTipText     =   "停止从发送端接受数据"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "关于"
            Key             =   "关于"
            Description     =   "关于"
            Object.ToolTipText     =   "参数配置选项"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "退出程序"
            Key             =   "退出程序"
            Description     =   "退出程序"
            Object.ToolTipText     =   "退出接收程序"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "DataRecvFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const BUSY As Boolean = False
Const FREE As Boolean = True
Dim ConnectState() As Boolean
Dim conType As String
Dim recordcount As Integer
Dim tickcount As Long


Private Sub Form_Load()
    statusBar.Panels(3).Text = Time
    ReDim Preserve ConnectState(0 To 1)
    On Error Resume Next
    
    ppp_status_timer.Enabled = False
    
    recordcount = 0
    tickcount = 0
    
    ConnectState(0) = FREE
    ConnectState(1) = FREE
    toolBar.Buttons(BTN_CONNECT).Enabled = True
    toolBar.Buttons(BTN_DISCONN).Enabled = False
    toolBar.Buttons(BTN_START).Enabled = False
    toolBar.Buttons(BTN_STOP).Enabled = False
End Sub

Private Sub Listener_ConnectionRequest(ByVal requestID As Long)
    Dim SockIndex As Integer
    Dim SockNum As Integer
    Dim answer
    On Error Resume Next
    '查找连接的用户数
    SockNum = UBound(ConnectState)
    If SockNum > 14 Then
        'Form1.Print SockIndex & ""
        Exit Sub
    End If
    '查找空闲的sock
    SockIndex = FindFreeSocket()
    '如果已有的sock都忙，而且sock数不超过15个，动态添加sock
    If SockIndex > SockNum Then
        Load Sock(SockIndex)
    End If
    ConnectState(SockIndex) = BUSY
    Sock(SockIndex).Tag = SockIndex
    '接受请求
    Sock(SockIndex).Accept (requestID)
    Dim line As String
    line = "接收到来自" & Sock(SockIndex).RemoteHostIP & "的连接请求"
    AppendInfoLine (line)
    
    answer = MsgBox(line & "，是否接受？", vbQuestion + vbYesNo, "是否接受请求？")
    If answer = vbNo Then
        line = "拒绝了来自" & Sock(SockIndex).RemoteHostIP & "的连接请求"
        AppendInfoLine (line)
        Sock(SockIndex).SendData ("DENY")
        Exit Sub
    Else
        line = "接受了来自" & Sock(SockIndex).RemoteHostIP & "的连接请求"
        AppendInfoLine (line)
        Sock(SockIndex).SendData ("ACCEPT")
    End If
    'Form1.Print SockIndex & "接受请求"
End Sub

Private Sub ppp_status_timer_Timer()
    Dim line As String
    Dim ms As Long
    Dim rx As Long
    Dim tx As Long
    If Is_PPP_Connecting(conType) Then
        ms = Get_PPP_Duration(conType) / 1000
        tx = Get_PPP_TXByte(conType) / 1024
        rx = Get_PPP_RXByte(conType) / 1024
        line = "连接持续时间: " & CStr(ms) & " 秒，已传输 " & CStr(tx) & " KB字节，已接收 " & _
            CStr(rx) & " KB字节。"
    
        statusBar.Panels(1) = line
        
        If Now - tickcount > 60 * 1000 Then
            line = "截至" & Date & " " & Time & "收到" & str(recordcount) & "条记录"
            AppendInfoLine (line)
            tickcount = Now
            recordcount = 0
        End If
    Else
        ' Connection broken
        line = "连接已经断开。"
        AppendInfoLine (line)
        ppp_status_timer.Enabled = False
        toolBar.Buttons(BTN_CONNECT).Enabled = True
        toolBar.Buttons(BTN_DISCONN).Enabled = False
        toolBar.Buttons(BTN_START).Enabled = False
        toolBar.Buttons(BTN_STOP).Enabled = False
    End If
End Sub

Private Sub Sock_Close(Index As Integer)
    If Sock(Index).State <> sckClosed Then
        Sock(Index).Close
    End If
    ConnectState(Index) = FREE
    'Form1.Print Index & "close"
End Sub


Private Sub Sock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim dx As String
    Dim tmpstr() As String
    Dim arraylen As Long
    Dim sql As String
    Dim line As String
    Dim i As Integer
    Dim j As Integer
    
    Sock(Index).GetData dx, vbString
        
    recordcount = recordcount + 1
    
    tmpstr = Split(dx, ",")
    For i = LBound(tmpstr) To UBound(tmpstr)
        If tmpstr(i) = frmLogin.txtTableName(1) Then
            line = ""
            For j = i + 1 To UBound(tmpstr)
                If tmpstr(j) = frmLogin.txtTableName(0) Or tmpstr(j) = frmLogin.txtTableName(1) Then
                    sql = "insert into " & tmpstr(i) & "values (" & Right(line, Len(line) - 1) & ")"
                    'glConnB.Execute sql
                    AppendInfoLine (sql)
                    Exit For
                Else
                    line = line & "," & tmpstr(j)
                End If
            Next j
        End If
        If tmpstr(i) = frmLogin.txtTableName(0) Then
            line = ""
            For j = i + 1 To UBound(tmpstr)
                If tmpstr(j) = frmLogin.txtTableName(0) Or tmpstr(j) = frmLogin.txtTableName(1) Then
                    sql = "insert into " & tmpstr(i) & "values (" & Right(line, Len(line) - 1) & ")"
                    'glConnA.Execute sql
                    AppendInfoLine (sql)
                    Exit For
                Else
                    line = line & "," & tmpstr(j)
                End If
            Next j
        End If
    Next i
    
    
    
End Sub

Public Function FindFreeSocket()
    Dim SockCount, i As Integer
    SockCount = UBound(ConnectState)
    For i = 0 To SockCount
        If ConnectState(i) = FREE Then
            FindFreeSocket = i
            Exit Function
        End If
    Next i
    ReDim Preserve ConnectState(0 To SockCount + 1)
    FindFreeSocket = UBound(ConnectState)
End Function


Private Sub Timer1_Timer()
    statusBar.Panels(3).Text = Time
End Sub

Private Sub toolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case BTN_CONNECT
            Dim ret As Boolean
            Dim line As String
            
            phoneDialFrm.Show vbModal
            
            If phoneDialFrm.Cancelled Then
                line = "用户取消"
                AppendInfoLine (line)
                Exit Sub
            End If
            
            conType = phoneDialFrm.cmbType.List(phoneDialFrm.cmbType.ListIndex)
            
            ret = Exists_PPP_Connection(conType)
            
            If ret = False Then
                '创建一个新的拨号连接
                line = "拨号连接不存在，正在新建拨号连接..."
                AppendInfoLine (line)
                
                Select Case conType
                    Case NAME_CDMA1X
                        ret = Create_PPP_Connection(NAME_CDMA1X, RASET_Phone, VS_Default, _
                                phoneDialFrm.txtPhoneNumber.Text, phoneDialFrm.txtPhoneUser.Text, _
                                phoneDialFrm.txtPhonePass.Text, phoneDialFrm.cmbModem.List(phoneDialFrm.cmbModem.ListIndex), _
                                RASDT_Modem, False, vbNullString, False, vbNullString, vbNullString, False, _
                                "86", "021")
                    Case NAME_VPN
                        ret = Create_PPP_Connection(NAME_VPN, RASET_Vpn, VS_Default, _
                                phoneDialFrm.txtPhoneNumber.Text, phoneDialFrm.txtPhoneUser.Text, _
                                phoneDialFrm.txtPhonePass.Text, vbNullString, _
                                RASDT_Vpn, False, vbNullString, False, vbNullString, vbNullString, False, _
                                vbNullString, vbNullString)
                    Case NAME_ADSL
                        ret = Create_PPP_Connection(NAME_ADSL, RASET_Broadband, VS_Default, _
                                vbNullString, phoneDialFrm.txtPhoneUser.Text, _
                                phoneDialFrm.txtPhonePass.Text, vbNullString, _
                                RASDT_PPPoE, False, vbNullString, False, vbNullString, vbNullString, False, _
                                vbNullString, vbNullString)
                    Case NAME_MODEM
                        ret = Create_PPP_Connection(NAME_MODEM, RASET_Phone, VS_Default, _
                                phoneDialFrm.txtPhoneNumber.Text, phoneDialFrm.txtPhoneUser.Text, _
                                phoneDialFrm.txtPhonePass.Text, phoneDialFrm.cmbModem.List(phoneDialFrm.cmbModem.ListIndex), _
                                RASDT_Modem, False, vbNullString, False, vbNullString, vbNullString, False, _
                                "86", "021")
                    Case NAME_DUMMY
                        ret = True
                        tickcount = Now
                        recordcount = 0
                        With DataRecvFrm
                            .toolBar.Buttons(BTN_CONNECT).Enabled = False
                            .toolBar.Buttons(BTN_DISCONN).Enabled = True
                            .toolBar.Buttons(BTN_START).Enabled = True
                            .toolBar.Buttons(BTN_STOP).Enabled = False
                        End With
                        Exit Sub
                End Select
                
                If ret = True Then
                    line = "连接创建成功！"
                    AppendInfoLine (line)
                Else
                    line = "连接创建失败，请重试！"
                    AppendInfoLine (line)
                    Exit Sub
                End If
            End If
            
            ret = Is_PPP_Connecting(conType)
            
            If ret = False Then
                line = "正在尝试拨号..."
                AppendInfoLine (line)
                
                ret = Dial_PPP_Connection(conType)
            End If
            
            If ret = True Then
                tickcount = Now
                recordcount = 0
                
            Else
                line = "拨号失败，请重试！"
                AppendInfoLine (line)
            End If
            
        Case BTN_DISCONN
            ret = Disconnect_PPP_Connection(conType)
            If ret = False Then
                statusBar.Panels(1) = "断开连接失败！"
                Exit Sub
            Else
                statusBar.Panels(1) = "断开连接成功！"
            End If
            
            ppp_status_timer.Enabled = False
            toolBar.Buttons(BTN_CONNECT).Enabled = True
            toolBar.Buttons(BTN_DISCONN).Enabled = False
            toolBar.Buttons(BTN_START).Enabled = False
            toolBar.Buttons(BTN_STOP).Enabled = False
        Case BTN_START
            '设置本机连接端口的localport属性的内容
            '请注意！必须是整体值
            ReDim Preserve ConnectState(0 To 1)
            On Error Resume Next
            ConnectState(0) = FREE
            ConnectState(1) = FREE
            Dim portNum As String
LoopTag:
            portNum = InputBox("请输入监听端口号", "接受参数配置", "56789")
            If portNum = "" Then MsgBox "还没输入！": GoTo LoopTag
            If Not IsNumeric(portNum) Then MsgBox "请输入数字！":  GoTo LoopTag
            Listener.LocalPort = portNum
            '将本机连接端口设置为监听模式
            Listener.Listen
            If Listener.State = sckListening Then
                line = LISTEN_SUCCESS & "侦听地址：" & Get_Client_PPP_IPAddress(conType) & " 侦听端口：" & Listener.LocalPort
                AppendInfoLine (line)
                toolBar.Buttons(BTN_START).Enabled = False
                toolBar.Buttons(BTN_STOP).Enabled = True
            Else
                line = LISTEN_FAILURE
                AppendInfoLine (line)
                toolBar.Buttons(BTN_START).Enabled = True
                toolBar.Buttons(BTN_STOP).Enabled = False
            End If
        Case BTN_STOP
            Dim SockCount, i As Integer
            SockCount = Sock.UBound
            For i = 0 To SockCount
                If Sock(i).State <> sckClosed Then
                    Sock(i).Close
                End If
            Next i
            line = LISTEN_CLOSED
            AppendInfoLine (line)
            toolBar.Buttons(BTN_START).Enabled = True
            toolBar.Buttons(BTN_STOP).Enabled = False
        Case BTN_PREF
            'MsgBox "Not Implemented", vbOKOnly, "N/A"
            frmAbout.Show 1
        Case BTN_QUIT
            Unload Me
            End
    End Select

End Sub

Public Sub AppendInfoLine(line As String)
    With infoBox
        .SelStart = glInfoTxtLen
        .SelText = line & vbNewLine
    End With
    glInfoTxtLen = glInfoTxtLen + Len(line & vbNewLine)
End Sub
