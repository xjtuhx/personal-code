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
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   6450
   StartUpPosition =   2  '屏幕中心
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
            Picture         =   "DataSendFrm.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataSendFrm.frx":0C52
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataSendFrm.frx":18A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataSendFrm.frx":24F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataSendFrm.frx":3148
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataSendFrm.frx":3D9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataSendFrm.frx":49EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataSendFrm.frx":56C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataSendFrm.frx":63A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataSendFrm.frx":707A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DataSendFrm.frx":7D54
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
      Width           =   6450
      _ExtentX        =   11377
      _ExtentY        =   661
      SimpleText      =   "状态栏"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6191
            Text            =   "状态栏"
            TextSave        =   "状态栏"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            Text            =   "显示日期"
            TextSave        =   "2009-7-5"
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
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   9763
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"DataSendFrm.frx":8A2E
   End
   Begin MSComctlLib.Toolbar toolBar 
      Align           =   1  'Align Top
      Height          =   855
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6450
      _ExtentX        =   11377
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
            Caption         =   "参数配置"
            Key             =   "参数配置"
            Description     =   "参数配置"
            ImageIndex      =   10
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


Private Sub Form_Load()
    Set result_table = New Recordset
    Set GPSData = New Recordset
    result_table_last = ""
    gpsdata_last = ""
    statusBar.Panels(3).Text = Time
    
    result_table_timer.Enabled = False
    GPSData_timer.Enabled = False
    Timer1.Enabled = True
    toolBar.Buttons(BTN_CONNECT).Enabled = True
    toolBar.Buttons(BTN_DISCONN).Enabled = False
    toolBar.Buttons(BTN_START).Enabled = False
    toolBar.Buttons(BTN_STOP).Enabled = False
End Sub

Private Sub result_table_timer_Timer()
    If sock(0).State <> sckConnected Then
        Exit Sub
    End If
    If result_table Is Nothing Or result_table.State <> 1 Then
EOFLOADA:
        If result_table_last = "" Then
            Call GetRecords(result_table, glConnA, _
                    frmLogin.txtTableName(0), frmLogin.txtTimestamp(0))
        Else
            Call GetRecords(result_table, glConnA, _
                    frmLogin.txtTableName(0), result_table_last)
        End If
    End If
    If result_table.EOF Then
        GoTo EOFLOADA
    End If
    Dim clip As String
    result_table_last = result_table.Fields("measuretime")
    clip = Trim(result_table.GetString(adClipString, 1, "','"))
    clip = Left(clip, Len(clip) - 1)
    clip = frmLogin.txtTableName(0) & ",'" & clip & "'"
    infoBox.SelStart = glInfoTxtLen
    infoBox.SelText = "发送:" & clip & vbNewLine
    glInfoTxtLen = glInfoTxtLen + Len("发送:" & clip & vbNewLine)
    sock(0).SendData (clip)
End Sub

Private Sub GPSData_timer_Timer()
    If sock(0).State <> sckConnected Then
        Exit Sub
    End If
    If GPSData Is Nothing Or GPSData.State <> 1 Then
EOFLOADB:
        If gpsdata_last = "" Then
            Call GetRecords(GPSData, glConnB, _
                    frmLogin.txtTableName(1), frmLogin.txtTimestamp(1))
        Else
            Call GetRecords(GPSData, glConnB, _
                    frmLogin.txtTableName(1), gpsdata_last)
        End If
    End If
    If GPSData.EOF Then
        GoTo EOFLOADB
    End If
    Dim clip As String
    gpsdata_last = GPSData.Fields("measuretime")
    clip = Trim(GPSData.GetString(adClipString, 1, "','"))
    clip = Left(clip, Len(clip) - 1)
    clip = frmLogin.txtTableName(1) & ",'" & clip & "'"
    infoBox.SelStart = glInfoTxtLen
    infoBox.SelText = "发送:" & clip & vbNewLine
    glInfoTxtLen = glInfoTxtLen + Len("发送:" & clip & vbNewLine)
    sock(0).SendData (clip)
End Sub

Private Sub Timer1_Timer()
    statusBar.Panels(3).Text = Time
End Sub

Private Sub toolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case BTN_CONNECT
            '远程主机名
            If optionsDialog.ipBox = "" Then
                MsgBox "请配置服务器IP参数", vbOKOnly, "缺少参数"
                Exit Sub
            End If
            If optionsDialog.portBox = "" Then
                MsgBox "请配置服务器端口参数", vbOKOnly, "缺少参数"
                Exit Sub
            End If
            If sock(0).State = sckOpen Then
                sock(0).Close
            End If
            sock(0).RemoteHost = optionsDialog.ipBox
            '网络端口
            sock(0).RemotePort = optionsDialog.portBox
            '发出连接命令
            sock(0).Connect
            
            TimedInfoDialog.Timeout = 15
            TimedInfoDialog.Start ("正在连接服务器...")
            
            If TimedInfoDialog.Success = False Then
                statusBar.Panels(1) = SOCK_FAILURE
                infoBox.SelStart = glInfoTxtLen
                infoBox.SelText = SOCK_FAILURE & vbNewLine
                glInfoTxtLen = glInfoTxtLen + Len(SOCK_FAILURE & vbNewLine)
            End If
            
        Case BTN_DISCONN
            result_table_timer.Enabled = False
            GPSData_timer.Enabled = False
            sock(0).Close
            Call sock_Close(0)
        Case BTN_START
            If Not sock(0).State = sckConnected Then
                MsgBox "连接已断开，请重新连接服务器！", vbOKOnly, "出错信息"
            Else
                toolBar.Buttons(BTN_STOP).Enabled = True
                toolBar.Buttons(BTN_START).Enabled = False
                result_table_timer.Enabled = True
                GPSData_timer.Enabled = True
            End If
        Case BTN_STOP
            toolBar.Buttons(BTN_STOP).Enabled = False
            toolBar.Buttons(BTN_START).Enabled = True
            result_table_timer.Enabled = False
            GPSData_timer.Enabled = False
        Case BTN_PREF
            optionsDialog.Show vbModal, DataSendFrm
        Case BTN_QUIT
            Unload Me
            End
    End Select
End Sub

Private Sub sock_Close(Index As Integer)
    'MsgBox ("socket closed")
    sock(Index).Close
    statusBar.Panels(1).Text = SOCK_CLOSED
    infoBox.SelStart = glInfoTxtLen
    infoBox.SelText = SOCK_CLOSED & vbNewLine
    glInfoTxtLen = glInfoTxtLen + Len(SOCK_CLOSED & vbNewLine)
    toolBar.Buttons(BTN_CONNECT).Enabled = True
    toolBar.Buttons(BTN_DISCONN).Enabled = False
    toolBar.Buttons(BTN_START).Enabled = False
    toolBar.Buttons(BTN_STOP).Enabled = False
    GPSData_timer.Enabled = False
    result_table_timer.Enabled = False
End Sub

Private Sub sock_Connect(Index As Integer)
    'MsgBox ("socket connected")
    TimedInfoDialog.Cancel
    statusBar.Panels(1).Text = SOCK_CONNECTED
    infoBox.SelStart = glInfoTxtLen
    infoBox.SelText = SOCK_CONNECTED & vbNewLine
    glInfoTxtLen = glInfoTxtLen + Len(SOCK_CONNECTED & vbNewLine)
    toolBar.Buttons(BTN_CONNECT).Enabled = False
    toolBar.Buttons(BTN_DISCONN).Enabled = True
    toolBar.Buttons(BTN_START).Enabled = True
    toolBar.Buttons(BTN_STOP).Enabled = False
End Sub
