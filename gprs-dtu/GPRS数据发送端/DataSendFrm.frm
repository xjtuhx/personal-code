VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form DataSendFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GPRS数据发送端"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8715
   Icon            =   "DataSendFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   8715
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer ppp_status_timer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2760
      Top             =   5880
   End
   Begin VB.Timer GPSData_timer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3360
      Top             =   5880
   End
   Begin VB.Timer result_table_timer 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   3960
      Top             =   5880
   End
   Begin MSWinsockLib.Winsock sock 
      Index           =   0
      Left            =   5160
      Top             =   5880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5760
      Top             =   5760
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
            Picture         =   "DataSendFrm.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataSendFrm.frx":191C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataSendFrm.frx":256E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataSendFrm.frx":31C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataSendFrm.frx":3E12
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataSendFrm.frx":4A64
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataSendFrm.frx":56B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataSendFrm.frx":6390
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataSendFrm.frx":706A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataSendFrm.frx":7D44
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataSendFrm.frx":8A1E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4560
      Top             =   5880
   End
   Begin MSComctlLib.StatusBar statusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6375
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   661
      SimpleText      =   "状态栏"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10186
            Text            =   "状态栏"
            TextSave        =   "状态栏"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            Text            =   "显示日期"
            TextSave        =   "2009-7-20"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Text            =   "显示时间"
            TextSave        =   "显示时间"
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox infoBox 
      Height          =   5535
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   9763
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"DataSendFrm.frx":96F8
   End
   Begin MSComctlLib.Toolbar toolBar 
      Align           =   1  'Align Top
      Height          =   855
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   1508
      ButtonWidth     =   1773
      ButtonHeight    =   1349
      Appearance      =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "连接服务器"
            Key             =   "连接服务器"
            Description     =   "连接服务器"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "断开服务器"
            Key             =   "断开服务器"
            Description     =   "断开服务器"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "开始传输"
            Key             =   "开始传输"
            Description     =   "开始传输"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "停止传输"
            Key             =   "停止传输"
            Description     =   "停止传输"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "关于"
            Key             =   "关于"
            Description     =   "关于"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "退出程序"
            Key             =   "退出程序"
            Description     =   "退出程序"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "DataSendFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim result_table As Recordset
Dim GPSData As Recordset
Dim result_table_last As String
Dim gpsdata_last As String
Dim conType As String
Dim gpscount As Integer
Dim rtcount As Integer


Private Sub Form_Load()
    Set result_table = New Recordset
    Set GPSData = New Recordset
    result_table_last = ""
    gpsdata_last = ""
    statusBar.Panels(3).Text = Time
    
    result_table_timer.Enabled = False
    GPSData_timer.Enabled = False
    Timer1.Enabled = True
    
    ppp_status_timer.Enabled = False
    
    gpscount = 0
    rtcount = 0
    
    conType = ""
    
    toolBar.Buttons(BTN_CONNECT).Enabled = True
    toolBar.Buttons(BTN_DISCONN).Enabled = False
    toolBar.Buttons(BTN_START).Enabled = False
    toolBar.Buttons(BTN_STOP).Enabled = False
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
    Else
        ' Connection broken
        ppp_status_timer.Enabled = False
        result_table_timer.Enabled = False
        GPSData_timer.Enabled = False
        sock(0).Close
        Call sock_Close(0)
    End If
End Sub

Private Sub result_table_timer_Timer()
    If sock(0).State <> sckConnected Then
        Exit Sub
    End If
    If result_table Is Nothing Or result_table.State = 0 Then
        If result_table_last = "" Then
            Call GetRecords(result_table, glConnA, _
                    frmLogin.txtTableName(0), frmLogin.txtTimestamp(0))
        Else
            Call GetRecords(result_table, glConnA, _
                    frmLogin.txtTableName(0), result_table_last)
        End If
    End If
    If result_table.EOF Then
        If result_table_last = "" Then
            Call GetRecords(result_table, glConnA, _
                    frmLogin.txtTableName(0), frmLogin.txtTimestamp(0))
        Else
            Call GetRecords(result_table, glConnA, _
                    frmLogin.txtTableName(0), result_table_last)
        End If
    End If
    
    If result_table.EOF Then
        Exit Sub
    End If
    Dim clip As String
    result_table_last = result_table.Fields("measuretime")
    clip = Trim(result_table.GetString(adClipString, 1, "','"))
    clip = Left(clip, Len(clip) - 1)
    clip = frmLogin.txtTableName(0) & ",'" & clip & "',"
    PrintLog ("发送:" & clip)
    rtcount = rtcount + 1
    sock(0).SendData (clip)
    
    PrintLog ("发送:" & str(gpscount + rtcount) & "条记录。")
    gpscount = 0
    rtcount = 0
End Sub

Private Sub GPSData_timer_Timer()
    If sock(0).State <> sckConnected Then
        Exit Sub
    End If
    If GPSData Is Nothing Or GPSData.State = 0 Then
        If gpsdata_last = "" Then
            Call GetRecords(GPSData, glConnB, _
                    frmLogin.txtTableName(1), frmLogin.txtTimestamp(1))
        Else
            Call GetRecords(GPSData, glConnB, _
                    frmLogin.txtTableName(1), gpsdata_last)
        End If
    End If
    If GPSData.EOF Then
        If gpsdata_last = "" Then
            Call GetRecords(GPSData, glConnB, _
                    frmLogin.txtTableName(1), frmLogin.txtTimestamp(1))
        Else
            Call GetRecords(GPSData, glConnB, _
                    frmLogin.txtTableName(1), gpsdata_last)
        End If
    End If
    
    If GPSData.EOF Then
        Exit Sub
    End If
    
    Dim clip As String
    gpsdata_last = GPSData.Fields("measuretime")
    clip = Trim(GPSData.GetString(adClipString, 1, "','"))
    clip = Left(clip, Len(clip) - 1)
    clip = frmLogin.txtTableName(1) & ",'" & clip & "',"
    PrintLog "发送:" & clip
    gpscount = gpscount + 1
    sock(0).SendData (clip)
End Sub


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
                Exit Sub
            End If
            
            conType = phoneDialFrm.cmbType.List(phoneDialFrm.cmbType.ListIndex)
            
            ret = Exists_PPP_Connection(conType)
            If ret = False Then
                '新建一个拨号连接
                line = "拨号连接不存在，正在新建拨号连接..."
                PrintLog (line)
                
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
                End Select
                
                If ret = True Then
                    line = "连接创建成功！"
                    PrintLog (line)
                Else
                    line = "连接创建失败！请重试！"
                    PrintLog (line)
                    Exit Sub
                End If
            End If
            
            ret = Is_PPP_Connecting(conType)
            If ret = False Then
                line = "正在尝试拨号..."
                PrintLog (line)

                ret = Dial_PPP_Connection(conType)
            End If
            
            If ret = True Then
                'toolBar.Buttons(BTN_CONNECT).Enabled = False
                toolBar.Buttons(BTN_DISCONN).Enabled = True
                
                serverParamDialog.Show vbModal
                If serverParamDialog.Cancelled = False Then
                    'If sock(0).State = sckOpen Then
                        'sock(0).Close
                    'End If
                    sock(0).Close
                    sock(0).RemoteHost = serverParamDialog.ipBox
                    '网络端口
                    sock(0).RemotePort = serverParamDialog.portBox
                    '发出连接命令
                    sock(0).Connect
                    
                    TimedInfoDialog.timeout = 15
                    TimedInfoDialog.Start "正在连接服务器 " & serverParamDialog.ipBox & " ...", 30
                    
                    If TimedInfoDialog.SUCCESS = False Then
                        line = SOCK_FAILURE & "服务器地址：" & serverParamDialog.ipBox
                        PrintLog (line)
                        Exit Sub
                    End If
                End If
            Else
                
                ppp_status_timer.Enabled = False
                
                line = "拨号失败，请重试！"
                PrintLog (line)
            End If
            
        Case BTN_DISCONN
            result_table_timer.Enabled = False
            GPSData_timer.Enabled = False
            sock(0).Close
            line = SOCK_CLOSED & "服务器地址：" & serverParamDialog.ipBox
            PrintLog (line)
            Call sock_Close(0)
            
            ret = Disconnect_PPP_Connection(conType)
            If ret = True Then
                statusBar.Panels(1) = "断开连接成功！"
                ppp_status_timer.Enabled = False
            Else
                statusBar.Panels(1) = "断开连接失败！"
            End If
        Case BTN_START
            If Not sock(0).State = sckConnected Then
                MsgBox "连接已断开，请重新连接服务器！", vbOKOnly, "出错信息"
            Else
                toolBar.Buttons(BTN_STOP).Enabled = True
                toolBar.Buttons(BTN_START).Enabled = False
                result_table_timer.Enabled = True
                GPSData_timer.Enabled = True
                
                gpscount = 0
                rtcount = 0
            End If
        Case BTN_STOP
            toolBar.Buttons(BTN_STOP).Enabled = False
            toolBar.Buttons(BTN_START).Enabled = True
            result_table_timer.Enabled = False
            GPSData_timer.Enabled = False
        Case BTN_PREF
            'optionsDialog.Show vbModal, DataSendFrm
            frmAbout.Show 1
        Case BTN_QUIT
            Unload Me
            End
    End Select
End Sub

Private Sub sock_Close(Index As Integer)
    'MsgBox ("socket closed")
    sock(Index).Close
    toolBar.Buttons(BTN_CONNECT).Enabled = True
    toolBar.Buttons(BTN_DISCONN).Enabled = True
    toolBar.Buttons(BTN_START).Enabled = False
    toolBar.Buttons(BTN_STOP).Enabled = False
    GPSData_timer.Enabled = False
    result_table_timer.Enabled = False
End Sub

Private Sub sock_Connect(Index As Integer)
    'MsgBox ("socket connected")
    TimedInfoDialog.Cancel
    Dim line As String
    line = "连接请求已发送，等待服务器接收..."
    PrintLog (line)
End Sub


Private Sub sock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim data As String
    Dim line As String
    sock(Index).GetData data, vbString
    
    Select Case data
        Case "DENY"
            PrintLog ("连接请求被拒绝！")
            sock_Close (Index)
        Case "ACCEPT"
            PrintLog ("连接请求被接受！")
            line = SOCK_CONNECTED & "服务器地址：" & serverParamDialog.ipBox
            PrintLog (line)
            toolBar.Buttons(BTN_CONNECT).Enabled = False
            toolBar.Buttons(BTN_DISCONN).Enabled = True
            toolBar.Buttons(BTN_START).Enabled = True
            toolBar.Buttons(BTN_STOP).Enabled = False
    End Select
End Sub

Public Sub PrintLog(line As String)
    With infoBox
        .SelStart = glInfoTxtLen
        .SelText = line & vbNewLine
    End With
    glInfoTxtLen = glInfoTxtLen + Len(line & vbNewLine)
End Sub
