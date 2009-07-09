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
   ClientWidth     =   7545
   Icon            =   "DataRecvFrm.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   7545
   StartUpPosition =   2  '屏幕中心
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
      Width           =   7335
      Begin RichTextLib.RichTextBox infoBox 
         Height          =   5295
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   7095
         _ExtentX        =   12515
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
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8123
            Text            =   "状态信息"
            TextSave        =   "状态信息"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "2009-7-9"
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
      Width           =   7545
      _ExtentX        =   13309
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

Private Sub Form_Load()
    statusBar.Panels(3).Text = Time
    ReDim Preserve ConnectState(0 To 1)
    On Error Resume Next
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
    'Form1.Print SockIndex & "接受请求"
    Dim line As String
    line = "接收到来自" & Sock(SockIndex).RemoteHostIP & "的连接请求"
    AppendInfoLine (line)
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
    Dim pos As Long
    Dim tableName As String
    Sock(Index).GetData dx, vbString
    Dim sql As String
    AppendInfoLine (dx & str(Len(dx)))
    pos = InStr(dx, ",")
    tableName = ""
    If pos > 0 Then
        tableName = Left(dx, pos - 1)
    End If
    If tableName = frmLogin.txtTableName(1) Then
        ' GPSData
        sql = "insert into " & frmLogin.txtTableName(1) & " values (" & Right(dx, Len(dx) - pos) & ")"
        glConnB.Execute sql
    Else
        ' result_table
        sql = "insert into " & frmLogin.txtTableName(0) & " values (" & Right(dx, Len(dx) - pos) & ")"
        glConnA.Execute sql
    End If
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
            
            '判断拨号连接是否存在
            ret = Exists_PPP_Connection(NAME_MODEM)
            
            If ret = False Then
                '创建一个新的拨号连接
                line = "拨号连接不存在，正在新建拨号连接..."
                infoBox.SelStart = glInfoTxtLen
                infoBox.SelText = line & vbNewLine
                glInfoTxtLen = glInfoTxtLen + Len(line & vbNewLine)
                
                phoneDialFrm.Show vbModal
                
                If phoneDialFrm.Cancelled Then
                    line = "用户取消。"
                    infoBox.SelStart = glInfoTxtLen
                    infoBox.SelText = line & vbNewLine
                    glInfoTxtLen = glInfoTxtLen + Len(line & vbNewLine)
                    Exit Sub
                End If
                
                ret = Create_PPP_Connection(NAME_MODEM, RASET_Phone, VS_Default, phoneDialFrm.txtPhoneNumber, _
                        phoneDialFrm.txtPhoneUser, phoneDialFrm.txtPhonePass, phoneDialFrm.cmbModem.SelText, _
                        RASDT_Modem, False, vbNullString, False, vbNullString, vbNullString, False, vbNullString, _
                        vbNullString)
                
                If ret = True Then
                    line = "连接创建成功！"
                    infoBox.SelStart = glInfoTxtLen
                    infoBox.SelText = line & vbNewLine
                    glInfoTxtLen = glInfoTxtLen + Len(line & vbNewLine)
                Else
                    line = "连接创建失败，请重试！"
                    infoBox.SelStart = glInfoTxtLen
                    infoBox.SelText = line & vbNewLine
                    glInfoTxtLen = glInfoTxtLen + Len(line & vbNewLine)
                    Exit Sub
                End If
            End If
            
            ret = Is_PPP_Connecting(NAME_MODEM)
            
            If ret = False Then
                line = "正在尝试拨号..."
                infoBox.SelStart = glInfoTxtLen
                infoBox.SelText = line & vbNewLine
                glInfoTxtLen = glInfoTxtLen + Len(line & vbNewLine)
                
                ret = Dial_PPP_Connection(NAME_MODEM)
            End If
            
            If ret = True Then
                line = "网络连接已经建立！"
                infoBox.SelStart = glInfoTxtLen
                infoBox.SelText = line & vbNewLine
                glInfoTxtLen = glInfoTxtLen + Len(line & vbNewLine)
                
                '显示服务器IP信息
                MsgBox "接收端IP地址：" & Get_Client_PPP_IPAddress(NAME_MODEM)
                
                toolBar.Buttons(BTN_CONNECT).Enabled = False
                toolBar.Buttons(BTN_DISCONN).Enabled = True
                toolBar.Buttons(BTN_START).Enabled = True
                toolBar.Buttons(BTN_STOP).Enabled = False
            Else
                line = "拨号失败，请重试！"
                infoBox.SelStart = glInfoTxtLen
                infoBox.SelText = line & vbNewLine
                glInfoTxtLen = glInfoTxtLen + Len(line & vbNewLine)
            End If
            
        Case BTN_DISCONN
            ret = Disconnect_PPP_Connection(NAME_MODEM)
            If ret = False Then
                statusBar.Panels(1) = "断开连接失败！"
                Exit Sub
            Else
                statusBar.Panels(1) = "断开连接成功！"
            End If
            
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
                line = LISTEN_SUCCESS & "侦听地址：" & Get_Client_PPP_IPAddress(NAME_MODEM) & " 侦听端口：" & Listener.LocalPort
                infoBox.SelStart = glInfoTxtLen
                infoBox.SelText = line & vbNewLine
                glInfoTxtLen = glInfoTxtLen + Len(line & vbNewLine)
                toolBar.Buttons(BTN_START).Enabled = False
                toolBar.Buttons(BTN_STOP).Enabled = True
            Else
                infoBox.SelStart = glInfoTxtLen
                infoBox.SelText = LISTEN_FAILURE + vbNewLine
                glInfoTxtLen = glInfoTxtLen + Len(LISTEN_FAILURE & vbNewLine)
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
            infoBox.SelStart = glInfoTxtLen
            infoBox.SelText = LISTEN_CLOSED & vbNewLine
            glInfoTxtLen = glInfoTxtLen + Len(LISTEN_CLOSED & vbNewLine)
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
