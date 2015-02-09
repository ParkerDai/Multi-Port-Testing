VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FMain 
   Caption         =   "Multi-Port Testing V1.6r (2014/10/16)"
   ClientHeight    =   9720
   ClientLeft      =   4185
   ClientTop       =   1230
   ClientWidth     =   10665
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9720
   ScaleWidth      =   10665
   Begin VB.Timer Timer7 
      Interval        =   1000
      Left            =   120
      Top             =   3120
   End
   Begin VB.ListBox lstSetting 
      Height          =   420
      Left            =   6960
      TabIndex        =   116
      Top             =   4680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CheckBox chk_10049 
      Caption         =   "Prevent 10049 Err"
      Height          =   255
      Left            =   9120
      TabIndex        =   115
      ToolTipText     =   "Reconnect every 5 seconds"
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox txtTimeout 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7680
      TabIndex        =   114
      Text            =   "5"
      Top             =   600
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Height          =   480
      Left            =   4920
      TabIndex        =   110
      Top             =   3560
      Width           =   1485
      Begin VB.OptionButton opAscHex 
         Caption         =   "ASC"
         Height          =   220
         Index           =   1
         Left            =   740
         TabIndex        =   112
         Top             =   180
         Width           =   720
      End
      Begin VB.OptionButton opAscHex 
         Caption         =   "Hex"
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   111
         Top             =   130
         Value           =   -1  'True
         Width           =   615
      End
   End
   Begin VB.CommandButton txtSendOnce 
      Caption         =   "Send Once"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   109
      Top             =   4500
      Width           =   1215
   End
   Begin VB.TextBox txtLoopbackStr 
      Height          =   375
      Left            =   2880
      TabIndex        =   108
      ToolTipText     =   "Hex without \ or 0x"
      Top             =   4920
      Width           =   3975
   End
   Begin VB.TextBox txtSendStr 
      Height          =   375
      Left            =   2880
      TabIndex        =   107
      ToolTipText     =   "Hex without \ or 0x"
      Top             =   4080
      Width           =   3975
   End
   Begin VB.ComboBox cbLocalIP 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3840
      TabIndex        =   105
      Top             =   230
      Width           =   1695
   End
   Begin VB.OptionButton opDataBit 
      Caption         =   "5 bit"
      Height          =   255
      Index           =   3
      Left            =   8160
      TabIndex        =   104
      Top             =   240
      Width           =   735
   End
   Begin VB.OptionButton opDataBit 
      Caption         =   "6 bit"
      Height          =   255
      Index           =   2
      Left            =   7440
      TabIndex        =   103
      Top             =   240
      Width           =   735
   End
   Begin VB.OptionButton opDataBit 
      Caption         =   "7 bit"
      Height          =   255
      Index           =   1
      Left            =   6720
      TabIndex        =   102
      Top             =   240
      Width           =   735
   End
   Begin VB.OptionButton opDataBit 
      Caption         =   "8 bit"
      Height          =   255
      Index           =   0
      Left            =   6000
      TabIndex        =   101
      Top             =   240
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.CheckBox chk_LoopbackStr 
      Caption         =   "Custom Loopback String"
      Height          =   255
      Left            =   2880
      TabIndex        =   100
      ToolTipText     =   "Define your Return String"
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CheckBox chk_SendStr 
      Caption         =   "Custom Send String"
      Height          =   255
      Left            =   2880
      TabIndex        =   99
      ToolTipText     =   "Define your Send String"
      Top             =   3720
      Width           =   1815
   End
   Begin VB.TextBox txtDuration 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7680
      TabIndex        =   97
      Text            =   "99999"
      Top             =   1080
      Width           =   735
   End
   Begin VB.CheckBox chk_SkipOnError 
      Caption         =   "Skip Error"
      Height          =   255
      Left            =   7920
      TabIndex        =   96
      ToolTipText     =   "Skips error checking"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CheckBox chk_AutoReconnect 
      Caption         =   "Auto Reconnect"
      Height          =   255
      Left            =   9120
      TabIndex        =   95
      ToolTipText     =   "Reconnect every 5 seconds"
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox txUDPLocalPort 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   94
      Text            =   "4660"
      ToolTipText     =   "Local Port for UDP"
      Top             =   1150
      Width           =   735
   End
   Begin VB.CheckBox chk_StopOnError 
      Caption         =   "Stop On Error"
      Height          =   255
      Left            =   6600
      TabIndex        =   91
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton butSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   89
      ToolTipText     =   "Save Error list to txt file"
      Top             =   1000
      Width           =   1695
   End
   Begin VB.CommandButton butDebug 
      Caption         =   "Debug"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   88
      ToolTipText     =   "Enlarge error list"
      Top             =   570
      Width           =   1695
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Left            =   10200
      Top             =   840
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   0
      TabIndex        =   67
      Top             =   1080
      Width           =   4215
      Begin VB.OptionButton opUDPMode 
         Caption         =   "UDP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   93
         Top             =   120
         Width           =   855
      End
      Begin VB.OptionButton opClientMode 
         Caption         =   "Listen"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2400
         TabIndex        =   69
         Top             =   120
         Width           =   1095
      End
      Begin VB.OptionButton opServerMode 
         Caption         =   "Client to Remote IP"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   68
         Top             =   120
         Value           =   -1  'True
         Width           =   2175
      End
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   10200
      Top             =   1320
   End
   Begin VB.CommandButton butTCPLinkTest 
      Caption         =   "TCP Link Test"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   8880
      TabIndex        =   63
      Top             =   5040
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.ListBox lstErrorList 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3000
      Left            =   2880
      TabIndex        =   61
      Top             =   6600
      Width           =   7695
   End
   Begin VB.Timer Timer3 
      Interval        =   20
      Left            =   2160
      Top             =   5280
   End
   Begin VB.CommandButton Cmd_Clear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   51
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   120
      Top             =   2520
   End
   Begin VB.CommandButton Cmd_Stop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   44
      Top             =   3840
      Width           =   2175
   End
   Begin VB.CommandButton Cmd_Send 
      Caption         =   "Send"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   43
      Top             =   3120
      Width           =   2175
   End
   Begin VB.CheckBox chkCOM 
      Caption         =   "Com1"
      Height          =   252
      Index           =   8
      Left            =   9600
      TabIndex        =   42
      Top             =   2280
      Width           =   852
   End
   Begin VB.CommandButton Cmd_Connect 
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   41
      Top             =   2400
      Width           =   2175
   End
   Begin VB.TextBox txtSocketPort 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1440
      TabIndex        =   39
      Text            =   "4660"
      Top             =   1800
      Width           =   1212
   End
   Begin VB.CheckBox chkCOM 
      Caption         =   "Com1"
      Height          =   252
      Index           =   32
      Left            =   9600
      TabIndex        =   38
      Top             =   3360
      Width           =   852
   End
   Begin VB.CheckBox chkCOM 
      Caption         =   "Com1"
      Height          =   252
      Index           =   31
      Left            =   8640
      TabIndex        =   37
      Top             =   3360
      Width           =   852
   End
   Begin VB.CheckBox chkCOM 
      Caption         =   "Com1"
      Height          =   252
      Index           =   30
      Left            =   7680
      TabIndex        =   36
      Top             =   3360
      Width           =   852
   End
   Begin VB.CheckBox chkCOM 
      Caption         =   "Com1"
      Height          =   252
      Index           =   29
      Left            =   6720
      TabIndex        =   35
      Top             =   3360
      Width           =   852
   End
   Begin VB.CheckBox chkCOM 
      Caption         =   "Com1"
      Height          =   252
      Index           =   28
      Left            =   5760
      TabIndex        =   34
      Top             =   3360
      Width           =   852
   End
   Begin VB.CheckBox chkCOM 
      Caption         =   "Com1"
      Height          =   252
      Index           =   27
      Left            =   4800
      TabIndex        =   33
      Top             =   3360
      Width           =   852
   End
   Begin VB.CheckBox chkCOM 
      Caption         =   "Com1"
      Height          =   252
      Index           =   26
      Left            =   3840
      TabIndex        =   32
      Top             =   3360
      Width           =   852
   End
   Begin VB.CheckBox chkCOM 
      Caption         =   "Com1"
      Height          =   252
      Index           =   25
      Left            =   2880
      TabIndex        =   31
      Top             =   3360
      Width           =   852
   End
   Begin VB.CheckBox chkCOM 
      Caption         =   "Com1"
      Height          =   252
      Index           =   24
      Left            =   9600
      TabIndex        =   30
      Top             =   3000
      Width           =   852
   End
   Begin VB.CheckBox chkCOM 
      Caption         =   "Com1"
      Height          =   252
      Index           =   23
      Left            =   8640
      TabIndex        =   29
      Top             =   3000
      Width           =   852
   End
   Begin VB.CheckBox chkCOM 
      Caption         =   "Com1"
      Height          =   252
      Index           =   22
      Left            =   7680
      TabIndex        =   28
      Top             =   3000
      Width           =   852
   End
   Begin VB.CheckBox chkCOM 
      Caption         =   "Com1"
      Height          =   252
      Index           =   21
      Left            =   6720
      TabIndex        =   27
      Top             =   3000
      Width           =   852
   End
   Begin VB.CheckBox chkCOM 
      Caption         =   "Com1"
      Height          =   252
      Index           =   20
      Left            =   5760
      TabIndex        =   26
      Top             =   3000
      Width           =   852
   End
   Begin VB.CheckBox chkCOM 
      Caption         =   "Com1"
      Height          =   252
      Index           =   19
      Left            =   4800
      TabIndex        =   25
      Top             =   3000
      Width           =   852
   End
   Begin VB.CheckBox chkCOM 
      Caption         =   "Com1"
      Height          =   252
      Index           =   18
      Left            =   3840
      TabIndex        =   24
      Top             =   3000
      Width           =   852
   End
   Begin VB.CheckBox chkCOM 
      Caption         =   "Com1"
      Height          =   252
      Index           =   17
      Left            =   2880
      TabIndex        =   23
      Top             =   3000
      Width           =   852
   End
   Begin VB.CheckBox chkCOM 
      Caption         =   "Com1"
      Height          =   252
      Index           =   16
      Left            =   9600
      TabIndex        =   22
      Top             =   2640
      Width           =   852
   End
   Begin VB.CheckBox chkCOM 
      Caption         =   "Com1"
      Height          =   252
      Index           =   15
      Left            =   8640
      TabIndex        =   21
      Top             =   2640
      Width           =   852
   End
   Begin VB.CheckBox chkCOM 
      Caption         =   "Com1"
      Height          =   252
      Index           =   14
      Left            =   7680
      TabIndex        =   20
      Top             =   2640
      Width           =   852
   End
   Begin VB.CheckBox chkCOM 
      Caption         =   "Com1"
      Height          =   252
      Index           =   13
      Left            =   6720
      TabIndex        =   19
      Top             =   2640
      Width           =   852
   End
   Begin VB.CheckBox chkCOM 
      Caption         =   "Com1"
      Height          =   252
      Index           =   12
      Left            =   5760
      TabIndex        =   18
      Top             =   2640
      Width           =   852
   End
   Begin VB.CheckBox chkCOM 
      Caption         =   "Com1"
      Height          =   252
      Index           =   11
      Left            =   4800
      TabIndex        =   17
      Top             =   2640
      Width           =   852
   End
   Begin VB.CheckBox chkCOM 
      Caption         =   "Com1"
      Height          =   252
      Index           =   10
      Left            =   3840
      TabIndex        =   16
      Top             =   2640
      Width           =   852
   End
   Begin VB.CheckBox chkCOM 
      Caption         =   "Com1"
      Height          =   252
      Index           =   9
      Left            =   2880
      TabIndex        =   15
      Top             =   2640
      Width           =   852
   End
   Begin VB.CheckBox chkCOM 
      Caption         =   "Com1"
      Height          =   252
      Index           =   7
      Left            =   8640
      TabIndex        =   14
      Top             =   2280
      Width           =   852
   End
   Begin VB.CheckBox chkCOM 
      Caption         =   "Com1"
      Height          =   252
      Index           =   6
      Left            =   7680
      TabIndex        =   13
      Top             =   2280
      Width           =   852
   End
   Begin VB.CheckBox chkCOM 
      Caption         =   "Com1"
      Height          =   252
      Index           =   5
      Left            =   6720
      TabIndex        =   12
      Top             =   2280
      Width           =   852
   End
   Begin VB.CheckBox chkCOM 
      Caption         =   "Com1"
      Height          =   252
      Index           =   4
      Left            =   5760
      TabIndex        =   11
      Top             =   2280
      Width           =   852
   End
   Begin VB.CheckBox chkCOM 
      Caption         =   "Com1"
      Height          =   252
      Index           =   3
      Left            =   4800
      TabIndex        =   10
      Top             =   2280
      Width           =   852
   End
   Begin VB.CheckBox chkCOM 
      Caption         =   "Com1"
      Height          =   252
      Index           =   2
      Left            =   3840
      TabIndex        =   9
      Top             =   2280
      Width           =   852
   End
   Begin VB.CheckBox chkCOM 
      Caption         =   "Com1"
      Height          =   252
      Index           =   1
      Left            =   2880
      TabIndex        =   8
      Top             =   2280
      Width           =   852
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   10200
      Top             =   360
   End
   Begin VB.TextBox txtSendInterval 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5160
      TabIndex        =   7
      Text            =   "200"
      Top             =   680
      Width           =   735
   End
   Begin VB.TextBox txtSendLen 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   2160
      TabIndex        =   5
      Text            =   "100"
      Top             =   680
      Width           =   975
   End
   Begin VB.CommandButton butArp 
      Caption         =   "ARP -d"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   5640
      Top             =   240
   End
   Begin VB.TextBox txtIP 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1380
      TabIndex        =   0
      Text            =   "10.0.160.5"
      Top             =   240
      Width           =   1392
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   2
      Left            =   3240
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   1
      Left            =   2880
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   3
      Left            =   3600
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   4
      Left            =   3960
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   5
      Left            =   4320
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   6
      Left            =   4680
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   7
      Left            =   5040
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   8
      Left            =   5400
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   9
      Left            =   2880
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   10
      Left            =   3240
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   11
      Left            =   3600
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   12
      Left            =   3960
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   13
      Left            =   4320
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   14
      Left            =   4680
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   15
      Left            =   5040
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   16
      Left            =   5400
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   17
      Left            =   2880
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   18
      Left            =   3240
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   19
      Left            =   3600
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   20
      Left            =   3960
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   21
      Left            =   4320
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   22
      Left            =   4680
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   23
      Left            =   5040
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   24
      Left            =   5400
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   25
      Left            =   2880
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   26
      Left            =   3240
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   27
      Left            =   3600
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   28
      Left            =   3960
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   29
      Left            =   4320
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   30
      Left            =   4680
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   31
      Left            =   5040
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   32
      Left            =   5400
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CheckBox chkLoopBack 
      Caption         =   "Loop Back"
      Height          =   252
      Left            =   7920
      TabIndex        =   62
      ToolTipText     =   "Switch to loopback mode"
      Top             =   1800
      Width           =   1092
   End
   Begin VB.TextBox txtExceedSend 
      Height          =   375
      Left            =   8520
      TabIndex        =   71
      Text            =   "0"
      ToolTipText     =   "The latest occurrence of a latency that is larger than Send Interval"
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox txtLagestSend 
      Height          =   375
      Left            =   8520
      TabIndex        =   72
      Text            =   "0"
      ToolTipText     =   "The largest latency since last clear"
      Top             =   4440
      Width           =   1695
   End
   Begin VB.TextBox total_port 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3480
      TabIndex        =   75
      Text            =   "16"
      Top             =   1800
      Width           =   375
   End
   Begin VB.CheckBox chkSelectAll 
      Caption         =   "Select All"
      Height          =   255
      Left            =   4200
      TabIndex        =   76
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CheckBox chkSelectOdd 
      Caption         =   "Select Odd"
      Height          =   255
      Left            =   5400
      TabIndex        =   77
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CheckBox chkSelectEven 
      Caption         =   "Select Even"
      Height          =   255
      Left            =   6600
      TabIndex        =   78
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CheckBox chk_pingpong 
      Caption         =   "Ping Pong"
      Height          =   255
      Left            =   5400
      TabIndex        =   90
      ToolTipText     =   "Wait for loopback data before sending again, select to burst traffic"
      Top             =   1440
      Value           =   1  '核取
      Width           =   1215
   End
   Begin VB.CheckBox chk_OverSend 
      Caption         =   "Exceed Timeout"
      Height          =   495
      Left            =   7560
      TabIndex        =   92
      ToolTipText     =   "The latest occurrence of a latency that is larger than Send Interval. Check to record to the error list"
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "Timeout (cycle) :"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   113
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Local IP:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2880
      TabIndex        =   106
      Top             =   240
      Width           =   1035
   End
   Begin VB.Label lblDuration 
      Caption         =   "Duration (min) :"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   98
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label24 
      Caption         =   "TAvgLatency :"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   87
      ToolTipText     =   "Average latency of all active ports"
      Top             =   9120
      Width           =   1455
   End
   Begin VB.Label lTAvgLatency 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  '單線固定
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   86
      ToolTipText     =   "Average latency of all active ports"
      Top             =   9120
      Width           =   1215
   End
   Begin VB.Label Label26 
      Caption         =   "ms"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9720
      TabIndex        =   85
      Top             =   6120
      Width           =   615
   End
   Begin VB.Label Label25 
      Caption         =   "Avg Latency"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   84
      ToolTipText     =   "Average value of latency"
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Label lAvgLatency 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  '單線固定
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   83
      ToolTipText     =   "Average value of latency"
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label Label23 
      Caption         =   "ms"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9720
      TabIndex        =   82
      Top             =   5520
      Width           =   615
   End
   Begin VB.Label Label22 
      Caption         =   "Latency"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   81
      ToolTipText     =   "The time that data takes to come back"
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label lLatency 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  '單線固定
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   80
      ToolTipText     =   "The time that data takes to come back"
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label lblCOM 
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   70
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label lTerror 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  '單線固定
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   66
      Top             =   8520
      Width           =   1215
   End
   Begin VB.Label Label14 
      Caption         =   "Total Error :"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   65
      Top             =   8520
      Width           =   1095
   End
   Begin VB.Label lTloss 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  '單線固定
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   64
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "Byte/Sec"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   60
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "Sec"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   59
      Top             =   5520
      Width           =   615
   End
   Begin VB.Label Label11 
      Caption         =   "Average Rate :"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   58
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "Loop Time :"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   57
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label lPerformance 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  '單線固定
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   56
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Label lLoopTime 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  '單線固定
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   55
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label lError 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  '單線固定
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   54
      Top             =   7320
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "Error :"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   53
      Top             =   7320
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Total Loss :"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   52
      Top             =   7920
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Loss :"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   50
      Top             =   6720
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Tx :"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   49
      Top             =   5520
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Rx :"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   48
      Top             =   6120
      Width           =   495
   End
   Begin VB.Label lLoss 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  '單線固定
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   47
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Label lTx 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  '單線固定
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   46
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label lRx 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  '單線固定
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   45
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Port from :"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   40
      Top             =   1800
      Width           =   1212
   End
   Begin VB.Label lblSendInterval 
      Caption         =   "Send Interval (ms):"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblSendLen 
      Caption         =   "Send Length (byte) :"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   4
      Top             =   720
      Width           =   2115
   End
   Begin VB.Label lblTime 
      BorderStyle     =   1  '單線固定
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   2
      Top             =   5040
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Remote IP:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Top             =   240
      Width           =   1155
   End
   Begin VB.Label lPort_Tesed 
      Caption         =   "Ports:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   79
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label16 
      Caption         =   "Largest"
      Height          =   375
      Left            =   7920
      TabIndex        =   74
      ToolTipText     =   "The largest latency since last clear"
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label15 
      Caption         =   "mS"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10080
      TabIndex        =   73
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'this procedure saves debug list to a file
Private Sub SaveLoadListbox(plstLB As ListBox, pstrFileName As String, pstrSaveOrLoad As String)
On Error GoTo Err

Dim strListItems As String
Dim i As Long

Select Case pstrSaveOrLoad
   Case "save"
    Open pstrFileName For Output As #1
    For i = 0 To plstLB.ListCount - 1
        plstLB.Selected(i) = True
        Print #1, plstLB.List(plstLB.ListIndex)
    Next
    Close #1

   Case "load"
   plstLB.Clear
   If Dir(pstrFileName) <> "" Then
    Open pstrFileName For Input As #1
    While Not EOF(1)
      Line Input #1, strListItems
      plstLB.AddItem strListItems
    Wend
    End If
    
    Close #1
End Select

Err:
End Sub

Private Sub butSave_Click()
    Call SaveLoadListbox(lstErrorList, "debug.txt", "save")
End Sub

Private Sub chk_10049_Click()
    If chk_10049.Value Then
        cbLocalIP.Enabled = False
    Else
        cbLocalIP.Enabled = True
    End If
End Sub

Private Sub chk_LoopbackStr_Click()
    Dim outString As String
    Dim i As Integer
    
    If chk_LoopbackStr.Value Then
        txtLoopbackStr.Enabled = True
        'outString = InputBox("Enter Hex")
        outString = txtLoopbackStr.Text
        
        If Len(outString) > 0 Then
            If opAscHex(0).Value Then                       'Hex mode
                ReDim Custom_Return(1 To Len(outString) / 2) As Byte
                
                For i = 1 To Len(outString) / 2
                    Custom_Return(i) = CByte("&H" & Mid(outString, i * 2 - 1, 2))
                Next
                
            ElseIf opAscHex(1).Value Then                   'string mode
                Custom_Return = StrConv(outString, vbFromUnicode)
    
            End If
            
            chk_pingpong.Value = 1
            chk_pingpong.Enabled = False
            txtSendLen.Text = UBound(Custom_Return) + 1
            txtSendLen.Enabled = False
        Else                                            'if user press cancel, uncheck
            chk_LoopbackStr.Value = 0
            chk_pingpong.Enabled = True
            txtSendLen.Enabled = True
        End If
    Else
        chk_pingpong.Enabled = True
        txtSendLen.Enabled = True
    End If
End Sub

Private Sub chk_SendStr_Click()

    Dim inString As String
    Dim i As Integer
    
    If chk_SendStr.Value Then
        txtSendStr.Enabled = True
        'inString = InputBox("Enter Hex")
        inString = txtSendStr.Text
        
        If Len(inString) > 0 Then
            If opAscHex(0).Value Then                       'Hex mode
                ReDim Custom_Send(1 To Len(inString) / 2) As Byte
                
                For i = 1 To Len(inString) / 2
                    Custom_Send(i) = CByte("&H" & Mid(inString, i * 2 - 1, 2))
                Next
                
            ElseIf opAscHex(1).Value Then                   'string mode
                Custom_Send = StrConv(inString, vbFromUnicode)
    
            End If
            
            txtSendLen.Text = UBound(Custom_Send) + 1
            txtSendLen.Enabled = False
        Else                                            'if user press cancel, uncheck
            chk_SendStr.Value = 0
            txtSendLen.Enabled = True
        End If
    Else
        txtSendLen.Enabled = True
    End If
    
End Sub

'chkCOM is Check for all the COMs(32), this procedure handles check/uncheck
Private Sub chkCOM_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
  If Cmd_Connect.Caption = "DisConnect" Then
    If chkCOM(Index).Value = 1 Then
      chkCOM(Index).Value = 0
    Else
      chkCOM(Index).Value = 1
    End If
  End If
End Sub

'this procedure handles which port to show on the GUI when mouse moves onto the COM
Private Sub chkCOM_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If chkCOM(Index).Value Then
    lblCOM.Caption = "COM " & Index & " State"
    iShow_Port = Index
  End If
End Sub

'chkSelectAll selects all COMs according to the number specified by total port
Private Sub chkSelectAll_Click()
Dim i As Integer

    For i = 1 To total_port
        chkCOM(i).Value = chkSelectAll.Value
    Next i

End Sub

'Selects only odd ports
Private Sub chkSelectOdd_Click()
Dim i As Integer

    For i = 1 To total_port
        chkCOM(i).Value = chkSelectOdd.Value
        i = i + 1
    Next i
End Sub

'select only even ports
Private Sub chkSelectEven_Click()
Dim i As Integer

    For i = 1 To total_port
        i = i + 1
        chkCOM(i).Value = chkSelectEven.Value
    Next i
End Sub

'decides if the program runs in ping pong mode. Ping pong is to collect all data before sending again
Private Sub chk_pingpong_Click()
    Timer4.Interval = txtSendInterval.Text
    
    If chk_pingpong.Value = 0 And chkLoopBack.Value Then
        txtSendLen.Enabled = False
    Else
        txtSendLen.Enabled = True
    End If

End Sub

Private Sub chkLoopBack_Click()
  If chkLoopBack.Value = 1 Then
    txtSendInterval.Enabled = False
    lblSendLen.Caption = "Receive Length:"
  Else
    txtSendInterval.Enabled = True
    lblSendLen.Caption = "Send Length:"
  End If
End Sub

Private Sub Cmd_Clear_Click()
Dim i As Integer
  For i = 1 To MAX_COM_PORT
    iSend_Buf_index(i) = 0
    iReceive_Buf_index(i) = 0

    iLoop_Time(i) = 0
    bTCP_Receive(i) = False
    iRx(i) = 0
    iTx(i) = 0
    iLoss(i) = 0
    iError(i) = 0
    iDuration(i) = 0
    iFirstTick(i) = 0
    iLastTick(i) = 0
    iLatency(i) = 0
    iReceived_Count(i) = 0
    no_receive_count(i) = 0
  Next i
  While lstErrorList.ListCount > 0
    lstErrorList.RemoveItem 0
  Wend
  lTloss.Caption = 0
  lTerror.Caption = 0
  lTAvgLatency.Caption = 0
  lLatency.Caption = 0
  lAvgLatency.Caption = 0
  txtLagestSend.Text = 0
  txtExceedSend.Text = 0
  no_send_count = 0
  
  
End Sub

Private Sub Cmd_Connect_Click()
On Error GoTo Err:

    Dim i As Integer
    Dim a As Integer
    
    If Cmd_Connect.Caption = "Connect to Remote IP" Or Cmd_Connect.Caption = "Listen" Then
        Cmd_Clear_Click
    
        For i = 1 To MAX_COM_PORT
            If Port_Status(i) = DisConnect Then
                If chkCOM(i).Value = 1 Then                                     'connect different socket modes
                    If opServerMode.Value = True Then
                        Winsock1(i).Protocol = sckTCPProtocol
                        Winsock1(i).LocalPort = "0"
                        Winsock1(i).RemoteHost = txtIP.Text
                        Winsock1(i).RemotePort = txtSocketPort.Text + i - 1
                        'Winsock1(i).Bind "2222", cbLocalIP.Text
                        'Debug.Print Winsock1(i).State
                        Winsock1(i).Connect
                        Port_Status(i) = Waiting_Connect
                        Cmd_Connect.Caption = "Waiting..."
                        chkCOM(i).ForeColor = &HFF&                             'Red
                    ElseIf opClientMode.Value = True Then
                        
                        Winsock1(i).Protocol = sckTCPProtocol
                        'Winsock1(i).LocalPort = txtSocketPort.Text + i - 1
                        If chk_10049.Value Then
                            Winsock1(i).Bind txtSocketPort + i - 1
                        Else
                            Winsock1(i).Bind txtSocketPort + i - 1, cbLocalIP.Text
                        End If
                        
                        Winsock1(i).Listen
                        Port_Status(i) = Waiting_Connect
                        Cmd_Connect.Caption = "Waiting..."
                        chkCOM(i).ForeColor = &HFF&                             'Red
                    ElseIf opUDPMode.Value = True Then
                        Winsock1(i).Protocol = sckUDPProtocol
                        Winsock1(i).RemoteHost = txtIP.Text
                        'Winsock1(i).LocalPort = txUDPLocalPort.Text + i - 1
                        Winsock1(i).RemotePort = txtSocketPort.Text + i - 1
                        
                        If chk_10049.Value Then
                            Winsock1(i).Bind txUDPLocalPort.Text + i - 1
                        Else
                            Winsock1(i).Bind txUDPLocalPort.Text + i - 1, cbLocalIP.Text
                        End If
                        
                        If chkLoopBack.Value = False Then
                            Cmd_Send.Enabled = True
                        End If
                        
                        Port_Status(i) = Connecting
                        chkCOM(i).ForeColor = &HFF00&                           'Green
                    End If
                Else
                    chkCOM(i).ForeColor = &H80000012
            End If
            
            Timer1.Enabled = True                                               'start updating UI
            End If
        Next i
    End If
  
    If Cmd_Connect.Caption = "DisConnect" Then                                  'close connections
        For i = 1 To MAX_COM_PORT
            Winsock1(i).Close
            Port_Status(i) = DisConnect
            Send_Status(i) = Stop_Send
            chkCOM(i).ForeColor = &H80000012                                    'None
        Next i
        
        Cmd_Send.Enabled = False
        Cmd_Stop.Enabled = False
        Timer4.Enabled = False
    End If
    
    If Cmd_Connect.Caption = "DisConnect" Then
        If opClientMode.Value Then
            Cmd_Connect.Caption = "Listen"
        Else
            Cmd_Connect.Caption = "Connect to Remote IP"
            
        End If
        
        chkSelectAll.Enabled = True
        chkLoopBack.Enabled = True
    Else
        Cmd_Connect.Caption = "DisConnect"
        chkSelectAll.Enabled = False
        chkLoopBack.Enabled = False
    End If
  
    Cmd_Clear_Click
  Exit Sub

Err:
    If Err.Number = 10049 Then
        chk_10049.Value = 1
        MsgBox "Please try again"
    Else
        MsgBox Error & " in cmd_Click"
    End If
End Sub

Private Sub Cmd_Send_Click()
Dim Send_Buffer() As Byte
Dim i As Integer
Dim j As Integer
On Error Resume Next
ReDim Send_Buffer(1 To txtSendLen.Text) As Byte
    'sTemp = ""
    For i = 1 To MAX_COM_PORT
        If Port_Status(i) = Connecting Then
            
            If chk_SendStr.Value = 0 Then
                For j = 1 To CInt(txtSendLen.Text)
                    'ReDim Preserve Send_Buffer(1 To j) As Byte
                    
                    If opDataBit(0).Value Then                              'different databits
                        Send_Buffer(j) = iSend_Buf_index(i) Mod 256
                    ElseIf opDataBit(1).Value Then
                        Send_Buffer(j) = iSend_Buf_index(i) Mod 128
                    ElseIf opDataBit(2).Value Then
                        Send_Buffer(j) = iSend_Buf_index(i) Mod 64
                    Else
                        Send_Buffer(j) = iSend_Buf_index(i) Mod 32
                    End If
                    
                    iSend_Buf_index(i) = iSend_Buf_index(i) + 1
                Next j
            Else
                Send_Buffer = Custom_Send
                
                ''''''''''''''''''''''''''''''''''''''''''''''''''''
'                'ReDim Custom_Send(1 To 1460) As Byte
'                Dim Stringa(1 To 1460) As Byte
'                Dim Stringb(1 To 1460) As Byte
'                'Dim i As Integer
'                Dim k As Integer
'
'                For k = 1 To 1460
'                   Stringa(k) = 33
'                Next
'
'                For k = 1 To 1460
'                   Stringb(k) = 34
'                Next
'
'                Stringa(1460) = 13
'                Stringb(1) = 0
'
'                Winsock1(i).SendData Stringa
'                Winsock1(i).SendData Stringb
                
                '''''''''''''''''''''''''''''''''''''''''''''''''''
                
                txtSendLen.Text = UBound(Custom_Send) + 1
                
            End If
            
            QueryPerformanceCounter iFirstTick(i)                           'start latency count
            
            Winsock1(i).SendData Send_Buffer
            
            iReceived_Count(i) = 0
            iTx(i) = iTx(i) + CInt(txtSendLen.Text)
            Send_Status(i) = Repeat_Send
            iLoss(i) = iTx(i) - iRx(i) - iError(i)
            chkCOM(i).ForeColor = &H80FF&                                   'yellow
        End If
    Next i
    
    If Timer4.Enabled = False Then
        Timer4.Enabled = True
    End If
    
    Cmd_Send.Enabled = False
    Cmd_Stop.Enabled = True
Err:

End Sub


Private Sub Cmd_Stop_Click()
    Dim i As Integer
    Dim TLoss As Long
    
    Timer4.Enabled = False
    Timer6.Enabled = False
    TLoss = 0
    
    For i = 1 To MAX_COM_PORT
        If Port_Status(i) = Connecting Then
            Send_Status(i) = Stop_Send
            TLoss = TLoss + iTx(i) - iRx(i)
        End If
    Next i
    
    
    Cmd_Send.Enabled = True
    Cmd_Stop.Enabled = False
End Sub

Private Sub butArp_Click()
Dim RetVal As Variant
  On Error Resume Next
  Kill "arp_d.bat.bat"
  Open "arp_d.bat" For Output As #1
  Print #1, "arp -d " & txtIP.Text
  Close #1
  RetVal = Shell("arp_d.bat", 3)
End Sub

Private Sub butTCPLinkTest_Click()
  If Timer5.Enabled = False Then
    Timer5.Enabled = True
  Else
    Timer5.Enabled = False
  End If
End Sub

Private Sub butDebug_Click()
    If lstErrorList.Height = 3000 Then
        lstErrorList.Height = 8000
        lstErrorList.Top = 1680
    Else
        lstErrorList.Height = 3000
        lstErrorList.Top = 6600
    End If
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim a() As Byte

    For i = 1 To MAX_COM_PORT
        chkCOM(i).Caption = "Com" & i
        chkCOM(i).ForeColor = &H80000012
        iSend_Buf_index(i) = 0
        iReceive_Buf_index(i) = 0
        iRx(i) = 0
        iTx(i) = 0
        iLoss(i) = 0
        iError(i) = 0
        'str_buffer(i) = Empty
        iReceived_Count(i) = 0
        Loopback_DataG(i) = a
        no_receive_count(i) = 0
    Next i
    
    QueryPerformanceFrequency Freq '取得API的計算頻率
    no_send_count = 0
    
    Timer4.Interval = txtSendInterval.Text
    Timer6.Interval = txtSendInterval.Text
    lTloss.Caption = 0
    lTerror.Caption = 0
    iShow_Port = 0
    opServerMode_Click
    'Load fCom
    'ReDim Loopback_DataA(txtSendLen.Text - 1)
    Dim Buf(0 To 4095) As Byte
    Dim BufSize As Long: BufSize = UBound(Buf) + 1
    Dim rc As Long
    rc = GetIpAddrTable_API(Buf(0), BufSize, 1)
    If rc <> 0 Then Err.Raise vbObjectError, , "GetIpAddrTable failed with return value " & rc
    Dim NrOfEntries As Integer: NrOfEntries = Buf(1) * 256 + Buf(0)
    'If NrOfEntries = 0 Then GetIpAddrTable = Array(): Exit Function
    ReDim IpAddrs(0 To NrOfEntries - 1) As String
    
    For i = 0 To NrOfEntries - 1
        Dim j As Integer, s As String: s = ""
        For j = 0 To 3: s = s & IIf(j > 0, ".", "") & Buf(4 + i * 24 + j): Next
        cbLocalIP.AddItem (s)
    Next
    cbLocalIP.Text = Winsock1(1).LocalIP

    Call SaveLoadListbox(lstSetting, "setting.txt", "load")
    'If Dir("setting.txt") <> "" Then Kill ("setting.txt")
    
    
    If lstSetting.ListCount = 50 Then      'lstSettings must have same entries as recorded at eof, otherwise it's an old version
        txtIP.Text = lstSetting.List(0)
        txtSendLen.Text = lstSetting.List(1)
        txtSendInterval.Text = lstSetting.List(2)
        total_port.Text = lstSetting.List(3)
        txtSocketPort.Text = lstSetting.List(4)
        txUDPLocalPort.Text = lstSetting.List(5)
        txtTimeout.Text = lstSetting.List(6)
        chk_pingpong.Value = lstSetting.List(7)
        chk_StopOnError.Value = lstSetting.List(8)
        chk_SkipOnError.Value = lstSetting.List(9)
        chk_AutoReconnect.Value = lstSetting.List(10)
        chk_10049.Value = lstSetting.List(11)
        opServerMode.Value = lstSetting.List(12)
        opClientMode.Value = lstSetting.List(13)
        opUDPMode.Value = lstSetting.List(14)
        chk_OverSend.Value = lstSetting.List(15)
        chkLoopBack.Value = lstSetting.List(16)
        txtDuration.Text = lstSetting.List(17)
        If chkLoopBack.Value Then chkLoopBack_Click: DoEvents
        
        For i = 1 To MAX_COM_PORT
            chkCOM(i) = lstSetting.List(17 + i)
            chk_pingpong.Value = chkLoopBack.Value
        Next i
    End If
    
    If Dir("Auto_Test") <> "" Then
        If chkLoopBack.Value = 1 Then
            Call Cmd_Connect_Click
        Else
            Call Cmd_Connect_Click
            Call Cmd_Send_Click
        End If
    End If

    If Dir("debug.txt") <> "" Then
        Kill "debug.txt"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    
    lstSetting.Clear
    lstSetting.AddItem txtIP.Text
    lstSetting.AddItem txtSendLen.Text
    lstSetting.AddItem txtSendInterval.Text
    lstSetting.AddItem total_port.Text
    lstSetting.AddItem txtSocketPort.Text
    lstSetting.AddItem txUDPLocalPort.Text
    lstSetting.AddItem txtTimeout.Text
    lstSetting.AddItem chk_pingpong.Value
    lstSetting.AddItem chk_StopOnError.Value
    lstSetting.AddItem chk_SkipOnError.Value
    lstSetting.AddItem chk_AutoReconnect.Value
    lstSetting.AddItem chk_10049.Value
    lstSetting.AddItem opServerMode.Value
    lstSetting.AddItem opClientMode.Value
    lstSetting.AddItem opUDPMode.Value
    lstSetting.AddItem chk_OverSend.Value
    lstSetting.AddItem chkLoopBack.Value
    lstSetting.AddItem txtDuration.Text
    
    For i = 1 To MAX_COM_PORT
        lstSetting.AddItem chkCOM(i)
    Next i
    
    
    Call SaveLoadListbox(lstSetting, "setting.txt", "save")

End Sub
Private Sub opAscHex_Click(Index As Integer)
    chk_SendStr.Value = 0
    chk_LoopbackStr = 0
End Sub

Private Sub opClientMode_Click()
  Cmd_Connect.Caption = "Listen"
  txtIP.Enabled = False
  txUDPLocalPort.Enabled = False
  cbLocalIP.Enabled = True
End Sub



Private Sub opServerMode_Click()
  Cmd_Connect.Caption = "Connect to Remote IP"
  txtIP.Enabled = True
  txUDPLocalPort.Enabled = False
  cbLocalIP.Enabled = False
End Sub

Private Sub opUDPMode_Click()
    Cmd_Connect.Caption = "Connect to Remote IP"
    txtIP.Enabled = True
    txUDPLocalPort.Enabled = True
    chk_AutoReconnect.Enabled = False
    chk_AutoReconnect.Value = False
    cbLocalIP.Enabled = True
End Sub

Private Sub total_port_LostFocus()
    If total_port.Text > 32 Then
        total_port.Text = 32
    End If
End Sub

Private Sub txtSendInterval_Change()
On Error GoTo err1:
  Timer4.Interval = txtSendInterval.Text
  Exit Sub
err1:
  Timer4.Interval = 1000
End Sub

Private Sub Timer1_Timer()
Dim i As Integer

  If Cmd_Connect.Caption = "Waiting..." Then
     Cmd_Connect.Caption = "Connect"
     Cmd_Send.Enabled = False
  End If


'  For i = 1 To MAX_COM_PORT
'    If Port_Status(i) = Waiting_Connect Then
'      Winsock1(i).Close
'      Port_Status(i) = DisConnect
'      chkCOM(i).ForeColor = &H80000012
'    End If
'  Next i

  Timer1.Enabled = False
End Sub

'Timer2 loops every second
Private Sub Timer2_Timer()
    Dim i As Integer
    Dim TLoss As Long
    Dim activeport As Integer
    Dim TAvgLatency As Double
    Dim TCP_State As Integer
    Dim t0 As SYSTEMTIME
    
    GetLocalTime t0
    activeport = 0
    TLoss = 0

    'lblTime.Caption = Format(Now, "yyyy/mm/dd hh:mm:ss")
    
    If Timer4.Enabled = True Then
        For i = 1 To MAX_COM_PORT
            
            If Port_Status(i) = Connecting Then
                activeport = activeport + 1
                iLoop_Time(i) = iLoop_Time(i) + 1
                TLoss = TLoss + iTx(i) - iRx(i)
                
                If (iRx(i) > 0) Then
                    TAvgLatency = TAvgLatency + iDuration(i) / (iRx(i) / txtSendLen.Text)
                End If
                
                If chk_AutoReconnect.Value Then                 'if auto reconnect is enabled, check socket state and reconnect if necessary
                
                    TCP_State = Winsock1(i).State
                        
                    If TCP_State = 9 Or TCP_State = 8 Or TCP_State = 6 Then
                        
                        lstErrorList.AddItem t0.wHour & ":" & t0.wMinute & ":" & t0.wSecond & ":" & t0.wMilliseconds & " ---> Com " & i & " : Reconncetion"
                        Winsock1(i).Close
                        
                        If opServerMode.Value Then
                            Winsock1(i).Connect
                        ElseIf opClientMode.Value Then
                            Winsock1(i).Listen
                        End If
                    End If
                    
                    If TCP_State = 7 And Cmd_Send.Enabled = True Then
                        lstErrorList.AddItem t0.wHour & ":" & t0.wMinute & ":" & t0.wSecond & ":" & t0.wMilliseconds & " ---> Com " & i & " : Retransmission"
                        Port_Status(i) = Connecting
                        Cmd_Send_Click
                    End If
                End If
            End If
        Next i
    Else
        For i = 1 To MAX_COM_PORT
            If bTCP_Receive(i) = True Then
                iLoop_Time(i) = iLoop_Time(i) + 1
            End If
            
            bTCP_Receive(i) = False
        Next i
    End If
    
    If TLoss > lTloss.Caption Then
        lTloss.Caption = TLoss
    End If
    
    If activeport > 0 Then
        lTAvgLatency.Caption = Round(TAvgLatency / activeport, 2)
    End If
        
    If no_send_count > txtTimeout.Text And Cmd_Send.Enabled = False Then  'resend
        lstErrorList.AddItem t0.wHour & ":" & t0.wMinute & ":" & t0.wSecond & ":" & t0.wMilliseconds & " " & txtSendInterval.Text * txtTimeout.Text & "ms Timed Out Resend"
        no_send_count = 0
        
        For i = 1 To MAX_COM_PORT
            If iReceived_Count(i) < txtSendLen.Text Then
                iReceive_Buf_index(i) = iReceive_Buf_index(i) + txtSendLen.Text
            End If
        Next i
        
        Cmd_Send_Click
    End If
    
    If Cmd_Connect.Caption = "DisConnect" And chkLoopBack.Value And chk_AutoReconnect.Value Then
        
        For i = 1 To MAX_COM_PORT
            no_receive_count(i) = no_receive_count(i) + 1           'timeout count for receive data
        
            If no_receive_count(i) > 10 Then
                
                    If Port_Status(i) = Connecting Then
                        TCP_State = Winsock1(i).State
                        no_receive_count(i) = 0
                        lstErrorList.AddItem t0.wHour & ":" & t0.wMinute & ":" & t0.wSecond & ":" & t0.wMilliseconds & " ---> Com " & i & " : Reconncetion"
                        Winsock1(i).Close
                        
                        If opServerMode.Value Then
                            Winsock1(i).Connect
                        ElseIf opClientMode.Value Then
                            Winsock1(i).Listen
                        End If
                        Port_Status(i) = Connecting
                    End If
            End If
        Next i
    End If
End Sub

'Timer3 is the GUI timer, it updates GUI
Private Sub Timer3_Timer()
On Error Resume Next
    
    Dim TLoss, i As Integer
    TLoss = 0
    
    lRx.Caption = iRx(iShow_Port)
    lTx.Caption = iTx(iShow_Port)
    lLoss.Caption = iTx(iShow_Port) - iRx(iShow_Port)
    lError.Caption = iError(iShow_Port)
    lLoopTime.Caption = iLoop_Time(iShow_Port)

    
    For i = 1 To MAX_COM_PORT
        If Port_Status(i) = Connecting Then
            Send_Status(i) = Stop_Send
            TLoss = TLoss + iTx(i) - iRx(i)
        End If
    Next i
    lTloss.Caption = TLoss
        
    If lLoopTime.Caption >= txtDuration.Text * 60 Then
        Cmd_Stop_Click
    End If
    
    lLatency.Caption = iLatency(iShow_Port)
    
    If iRx(iShow_Port) > 0 Then
        lAvgLatency.Caption = Round(iDuration(iShow_Port) / (iRx(iShow_Port) / txtSendLen.Text), 2)
    End If
    
    If iLoop_Time(iShow_Port) = 0 Then
        lPerformance.Caption = 0
    Else
        lPerformance.Caption = iRx(iShow_Port) \ iLoop_Time(iShow_Port)
    End If
End Sub

Private Sub Timer4_Timer()

    If chk_pingpong.Value Then                      'wait in pingpong mode

        Dim i As Integer
        Dim all_received As Boolean
        
        all_received = True
        no_send_count = no_send_count + 1
        
        For i = 1 To MAX_COM_PORT                   'check if all bytes in all ports received
            'no_send_count = no_send_count + 1      '2014/7/16 additional no_send_count
            If chkCOM(i).Value = 1 Then
                If iReceived_Count(i) < txtSendLen.Text Then
                    all_received = False
                    Exit For
                End If
            End If
        Next i
            
        If all_received = True Then                 'contunue send if all ports received
            Cmd_Send_Click
            no_send_count = 0
            
            For i = 1 To MAX_COM_PORT
                iReceived_Count(i) = 0
            Next i
        End If
    Else
        iReceived_Count(i) = 0
        Cmd_Send_Click
    End If
End Sub

Private Sub Timer5_Timer()
  Cmd_Connect_Click
End Sub

Private Sub Timer6_Timer()
'On error GoTo Err:
'    Dim i As Integer
'
'    For i = 1 To MAX_COM_PORT
'        If chkCOM(i).Value = 1 Then             'if checked
'            'If iReceived_Count(i) >= txtSendLen.Text Then
'                iTx(i) = iTx(i) + Len(str_buffer(i))
'                iReceived_Count(i) = 0
'
'                Winsock1(i).SendData str_buffer(i)
'                Timer6.Enabled = False
'                str_buffer(i) = Empty
'            'End If
'        End If
'    Next i
'Err:
End Sub

Private Sub Timer7_Timer()
    Dim i As Integer
    If Dir("Auto_Test") = "" Then
        If chkLoopBack.Value = 0 Then
            If CInt(lTerror.Caption) > 0 Or CInt(lTloss.Caption) > 5 Then
                butSave_Click
            End If
            For i = 1 To total_port
                If chkCOM(i).ForeColor = &HFF& Then ' Red 表示網路連線有問題
                    butSave_Click
                    Exit For
                End If
            Next i
        End If
        Timer7.Enabled = False
    End If
    Static time As Long
    If time < CLng(txtDuration.Text) * 60 Then
        time = time + 1
    Else
        Cmd_Stop_Click
    End If
End Sub

Private Sub txtSendOnce_Click()
On Error Resume Next
    Dim inString As String
    Dim i As Integer
    Dim tByte() As Byte
 
    inString = txtSendStr.Text
    
    If Len(inString) > 0 Then
        If opAscHex(0).Value Then                       'Hex mode
            ReDim tByte(1 To Len(inString) / 2) As Byte
            
            For i = 1 To Len(inString) / 2
                tByte(i) = CByte("&H" & Mid(inString, i * 2 - 1, 2))
            Next
            
        ElseIf opAscHex(1).Value Then                   'string mode
            tByte = StrConv(inString, vbFromUnicode)

        End If
        
        For i = 1 To MAX_COM_PORT
            If Port_Status(i) = Connecting Then
                Winsock1(i).SendData tByte
            End If
        Next
    End If

    
End Sub

Private Sub txtTimeout_Change()

If IsNumeric(txtTimeout.Text) = False Then
    txtTimeout.Text = "5"
End If

End Sub

Private Sub Winsock1_Connect(Index As Integer)
  Cmd_Connect.Caption = "DisConnect"
  Port_Status(Index) = Connecting
  If chkLoopBack.Value = 0 Then
    Cmd_Send.Enabled = True
  End If
  chkCOM(Index).ForeColor = &HFF00&     'Green
  'iRx(Index) = 0
  'iTx(Index) = 0
  'iLoss(Index) = 0
'  lStartTime(Index).Caption = Format(Now, "yyyy/mm/dd hh:mm:ss")
End Sub

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
  If Winsock1(Index).State <> sckClosed Then Winsock1(Index).Close
  Winsock1(Index).Accept requestID
  
  Cmd_Connect.Caption = "DisConnect"
  Port_Status(Index) = Connecting
  If chkLoopBack.Value = 0 Then
    Cmd_Send.Enabled = True
  End If
  chkCOM(Index).ForeColor = &HFF00&     'Green
  'iRx(Index) = 0
  'iTx(Index) = 0
  'iLoss(Index) = 0
'  lStartTime(Index).Caption = Format(Now, "yyyy/mm/dd hh:mm:ss")
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error GoTo Err

    Dim t0 As SYSTEMTIME
    Dim i As Integer
    Dim j As Integer
    Dim sTemp() As Byte
    Dim debug_str As String
    Dim debug_str_pos As Integer
    Dim debug_output_len As Integer
    Dim have_error As Boolean
    Dim LoopBackDataL() As Byte
    Dim bstr As String
    Dim sTempLen As Integer
    Dim bExpected As Byte
    
    If bytesTotal = 0 Then
        lstErrorList.AddItem t0.wHour & ":" & t0.wMinute & ":" & t0.wSecond & ":" & t0.wMilliseconds & " ---> Com " & Index & " Empty Packet"
        Exit Sub
    End If
    
    LoopBackDataL = Loopback_DataG(Index)                                           'copy stored loopback data from global to local
    have_error = False
    GetLocalTime t0
    
    'ReDim sTemp(bytesTotal)
    
    Winsock1(Index).GetData sTemp, vbArray + vbByte, bytesTotal                     'get and store socket data in sTemp, type=byte array, length=bytesTotal
    
    sTempLen = UBound(sTemp) - LBound(sTemp) + 1                                    'find actual bytes in array (empty packet)
    
    If sTempLen <> bytesTotal Then
        bytesTotal = sTempLen
    End If
    
    chkCOM(Index).ForeColor = &HFF00&                                               'Green
       
    ReDim Preserve LoopBackDataL(iReceived_Count(Index) + bytesTotal - 1) As Byte 'resize loopback data array
    
    For i = 0 To bytesTotal - 1

        If chkLoopBack.Value And chk_pingpong.Value Then                            'if pingpong mode and in loopback
            LoopBackDataL(i + iReceived_Count(Index)) = sTemp(i)                    'store sTemp in loopback mode
            'Debug.Print UBound(LoopBackDataL) & " " & i & " " & LoopBackDataL(i + iReceived_Count(Index))
        End If
                                                                                    
        If opDataBit(0).Value Then                                                  'different databits
            bExpected = iReceive_Buf_index(Index) Mod 256
        ElseIf opDataBit(1).Value Then
            bExpected = iReceive_Buf_index(Index) Mod 128
        ElseIf opDataBit(2).Value Then
            bExpected = iReceive_Buf_index(Index) Mod 64
        Else
            bExpected = iReceive_Buf_index(Index) Mod 32
        End If
                                                                                    'if byte match or no error checking
        If sTemp(i) = bExpected Or chk_SkipOnError.Value = 1 Then
            iRx(Index) = iRx(Index) + 1
        Else                                                                        'bytes mismatch
            While lstErrorList.ListCount > 1000
                lstErrorList.RemoveItem 0
            Wend
        
            have_error = True
                                                                                    'add extra 0 to single hex character in output
            If sTemp(i) >= 16 And iReceive_Buf_index(Index) Mod &H100 >= 16 Then
                bstr = " : Rec = 0x" & Hex$(sTemp(i)) & " <> 0x"
            ElseIf sTemp(i) < 16 Then
                bstr = " : Rec = 0x0" & Hex$(sTemp(i)) & " <> 0x"
            ElseIf iReceive_Buf_index(Index) Mod &H100 < 16 Then
                bstr = " : Rec = 0x" & Hex$(sTemp(i)) & " <> 0x0"
            Else
                bstr = " : Rec = 0x0" & Hex$(sTemp(i)) & " <> 0x0"
            End If
                                                                                    'display error byte
            lstErrorList.AddItem t0.wMonth & "/" & t0.wDay & " " & t0.wHour & ":" & t0.wMinute & ":" & t0.wSecond & ":" & t0.wMilliseconds & " ---> Com " & Index & bstr & Hex$((iReceive_Buf_index(Index) Mod &H100)) & " " & i & " " & bytesTotal & " " & iLatency(Index)
            
            iError(Index) = iError(Index) + 1
            lTerror.Caption = lTerror.Caption + 1
            iReceive_Buf_index(Index) = sTemp(i)                                    'recenter receive index to error byte
        End If
        
        iReceive_Buf_index(Index) = iReceive_Buf_index(Index) + 1                   'increment receive index to next byte
    Next i
    
        '-------------------Debug Print------------------------'
    If have_error = True Then
        debug_str_pos = 0
        debug_output_len = 24                                                       'change line when length > 24
        
        While (debug_str_pos < bytesTotal)                                          'while not all bytes displayed
            
            If bytesTotal - debug_str_pos < debug_output_len Then                   'if left bytes if lesser than line length
                debug_output_len = bytesTotal - debug_str_pos                       'line length equal to left bytes
            End If
            
            For j = 0 To debug_output_len - 1                                       'append all bytes for one line
                If sTemp(debug_str_pos + j) < 16 Then                               'add extra 0 for Hex 0-F
                    debug_str = debug_str & " " & "0" & Hex$(sTemp(debug_str_pos + j))
                Else
                    debug_str = debug_str & " " & Hex$(sTemp(debug_str_pos + j))
                End If
            Next j
            
            debug_str_pos = debug_str_pos + debug_output_len                        'move index to bytes not displayed
            lstErrorList.AddItem debug_str                                          'display one line
            debug_str = ""
        Wend                                                                        'continue to next line
        
        If chk_StopOnError.Value Then
            Cmd_Stop_Click
        End If
    Else                                                                                'count latency when no error
        If chkLoopBack.Value = 0 And Cmd_Send.Enabled = False Then                      'don't do latency caluculation if in loopback mode or not sending
            If iReceived_Count(Index) + bytesTotal >= txtSendLen.Text Then              'if all data received, calculate
                QueryPerformanceCounter iLastTick(Index)
                iLatency(Index) = (iLastTick(Index) - iFirstTick(Index)) / Freq * 1000
                iDuration(Index) = iDuration(Index) + iLatency(Index)
                
                If (iLatency(Index) > txtSendInterval.Text * txtTimeout.Text) Then
                    Dim tlatency As String
                    tlatency = Round(iLatency(Index), 2)
                
                    If chk_OverSend.Value Then 'output only once
                        lstErrorList.AddItem Format(Now, "hh:mm:ss") & " ---> Com " & Index & " Exceed Send Interval: " & iLatency(Index)
                    End If
                
                    txtExceedSend.Text = tlatency
                
                    If tlatency > txtLagestSend.Text Then
                        txtLagestSend.Text = tlatency
                    End If
                End If
            End If
        End If
    End If
    '-------------------Debug Print------------------------'
    
    bTCP_Receive(Index) = True
    
    If chkLoopBack.Value Then                                                       'if on loopback side
        
        no_receive_count(Index) = 0                                                        'clear timeout count for data receive
        
        If chk_pingpong.Value Then
                                                                
            If iReceived_Count(Index) + bytesTotal >= txtSendLen.Text Then          'if all data received
                If chk_LoopbackStr.Value Then
                    iTx(Index) = iTx(Index) + UBound(Custom_Return) - LBound(Custom_Return) + 1
                    Winsock1(Index).SendData Custom_Return
                Else
                    iTx(Index) = iTx(Index) + UBound(LoopBackDataL) - LBound(LoopBackDataL) + 1
                    Winsock1(Index).SendData LoopBackDataL                          'loopback data
                End If
                
                
                iReceived_Count(Index) = 0                                          'clear buffer count
                'Loopback_DataG(Index) = Empty
            Else
                Loopback_DataG(Index) = LoopBackDataL
                iReceived_Count(Index) = iReceived_Count(Index) + bytesTotal
            End If
        Else
            iTx(Index) = iTx(Index) + UBound(sTemp) - LBound(sTemp) + 1
            Winsock1(Index).SendData sTemp
        End If
            
    '        If Timer6.Enabled = False Then
    '            Timer6.Enabled = True
    '        End If
            
    Else
        iReceived_Count(Index) = iReceived_Count(Index) + bytesTotal
        
    End If
    Exit Sub
    
Err:
 
    If chk_OverSend.Value = 1 Then
        lstErrorList.AddItem t0.wHour & ":" & t0.wMinute & ":" & t0.wSecond & ":" & t0.wMilliseconds & " ---> Com " & Index & " " & bytesTotal & " " & " " & Error
    End If
    
    Resume Next
End Sub
