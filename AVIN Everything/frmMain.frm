VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AVIN System Control Panel v1.0"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   201
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   394
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab tabMain 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5318
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   441
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "System Info"
      TabPicture(0)   =   "frmMain.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblPlatform"
      Tab(0).Control(1)=   "lblVer"
      Tab(0).Control(2)=   "Image6"
      Tab(0).Control(3)=   "lblCompName"
      Tab(0).Control(4)=   "lblUser"
      Tab(0).Control(5)=   "Frame4"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Phone Dialer"
      TabPicture(1)   =   "frmMain.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtDial"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "btnDial"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtCOM"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Memory Manager"
      TabPicture(2)   =   "frmMain.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "totRam"
      Tab(2).Control(1)=   "freeRam"
      Tab(2).Control(2)=   "totVM"
      Tab(2).Control(3)=   "freeVM"
      Tab(2).Control(4)=   "totPage"
      Tab(2).Control(5)=   "freePage"
      Tab(2).Control(6)=   "RamPerc"
      Tab(2).Control(7)=   "VMPerc"
      Tab(2).Control(8)=   "PagePerc"
      Tab(2).Control(9)=   "Image2"
      Tab(2).Control(10)=   "Image3"
      Tab(2).Control(11)=   "Image4"
      Tab(2).Control(12)=   "tmrMem"
      Tab(2).Control(13)=   "Frame2"
      Tab(2).ControlCount=   14
      TabCaption(3)   =   "Screen"
      TabPicture(3)   =   "frmMain.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblCurr"
      Tab(3).Control(1)=   "Image1"
      Tab(3).Control(2)=   "lstRes"
      Tab(3).Control(3)=   "Command1"
      Tab(3).Control(4)=   "Command2"
      Tab(3).Control(5)=   "Frame3"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "System Controls"
      TabPicture(4)   =   "frmMain.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "btnSysControl(0)"
      Tab(4).Control(1)=   "btnSysControl(1)"
      Tab(4).Control(2)=   "btnSysControl(2)"
      Tab(4).Control(3)=   "btnSysControl(3)"
      Tab(4).Control(4)=   "btnSysControl(4)"
      Tab(4).Control(5)=   "btnSysControl(5)"
      Tab(4).Control(6)=   "btnSysControl(6)"
      Tab(4).Control(7)=   "btnSysControl(7)"
      Tab(4).Control(8)=   "btnSysControl(8)"
      Tab(4).Control(9)=   "btnSysControl(9)"
      Tab(4).Control(10)=   "btnSysControl(10)"
      Tab(4).Control(11)=   "btnSysControl(11)"
      Tab(4).Control(12)=   "btnSysControl(12)"
      Tab(4).Control(13)=   "btnSysControl(13)"
      Tab(4).Control(14)=   "btnSysControl(14)"
      Tab(4).Control(15)=   "btnSysControl(15)"
      Tab(4).Control(16)=   "btnSysControl(16)"
      Tab(4).Control(17)=   "btnSysControl(17)"
      Tab(4).Control(18)=   "btnSysControl(18)"
      Tab(4).Control(19)=   "btnSysControl(19)"
      Tab(4).Control(20)=   "btnSysControl(20)"
      Tab(4).Control(21)=   "btnSysControl(21)"
      Tab(4).Control(22)=   "btnSysControl(22)"
      Tab(4).Control(23)=   "btnSysControl(23)"
      Tab(4).ControlCount=   24
      Begin VB.Frame Frame4 
         Caption         =   "Power Status"
         Height          =   975
         Left            =   -74880
         TabIndex        =   68
         Top             =   1680
         Width           =   3255
         Begin VB.Label lblLife 
            AutoSize        =   -1  'True
            Caption         =   "Battery Life:"
            Height          =   195
            Left            =   120
            TabIndex        =   71
            Top             =   720
            Width           =   840
         End
         Begin VB.Label lblBatt 
            AutoSize        =   -1  'True
            Caption         =   "Battery Status: N/A"
            Height          =   195
            Left            =   120
            TabIndex        =   70
            Top             =   480
            Width           =   1380
         End
         Begin VB.Label lblAC 
            AutoSize        =   -1  'True
            Caption         =   "AC Status: N/A"
            Height          =   195
            Left            =   120
            TabIndex        =   69
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "System Colors"
         Height          =   2535
         Left            =   -72240
         TabIndex        =   57
         Top             =   360
         Width           =   3015
         Begin VB.CommandButton Command3 
            Caption         =   "Reset"
            Enabled         =   0   'False
            Height          =   375
            Left            =   2040
            TabIndex        =   65
            Top             =   2040
            Width           =   615
         End
         Begin VB.TextBox txtBlue 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Enabled         =   0   'False
            Height          =   285
            Left            =   2040
            MaxLength       =   3
            TabIndex        =   64
            Top             =   1680
            Width           =   615
         End
         Begin VB.TextBox txtGreen 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Enabled         =   0   'False
            Height          =   285
            Left            =   2040
            MaxLength       =   3
            TabIndex        =   62
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox txtRed 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Enabled         =   0   'False
            Height          =   285
            Left            =   2040
            MaxLength       =   3
            TabIndex        =   60
            Top             =   480
            Width           =   615
         End
         Begin VB.ListBox lstSysCol 
            Height          =   2205
            ItemData        =   "frmMain.frx":04CE
            Left            =   120
            List            =   "frmMain.frx":050B
            TabIndex        =   58
            Top             =   240
            Width           =   1815
         End
         Begin VB.Shape shpFill 
            BackStyle       =   1  'Opaque
            Height          =   2175
            Left            =   2760
            Top             =   240
            Width           =   135
         End
         Begin VB.Label Label6 
            Caption         =   "Blue:"
            Height          =   255
            Left            =   2040
            TabIndex        =   63
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label Label5 
            Caption         =   "Green:"
            Height          =   255
            Left            =   2040
            TabIndex        =   61
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label4 
            Caption         =   "Red:"
            Height          =   255
            Left            =   2040
            TabIndex        =   59
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.CommandButton btnSysControl 
         Caption         =   "Enable Start"
         Height          =   255
         Index           =   23
         Left            =   -71040
         TabIndex        =   56
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton btnSysControl 
         Caption         =   "Disable Start"
         Height          =   255
         Index           =   22
         Left            =   -71040
         TabIndex        =   55
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CommandButton btnSysControl 
         Caption         =   "Get Snapshot"
         Enabled         =   0   'False
         Height          =   255
         Index           =   21
         Left            =   -72960
         TabIndex        =   54
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton btnSysControl 
         Caption         =   "Show Time"
         Height          =   255
         Index           =   20
         Left            =   -71040
         TabIndex        =   53
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton btnSysControl 
         Caption         =   "Flip Mouse Buttons"
         Height          =   255
         Index           =   19
         Left            =   -72960
         TabIndex        =   52
         Top             =   2280
         Width           =   1695
      End
      Begin VB.CommandButton btnSysControl 
         Caption         =   "Hide Time"
         Height          =   255
         Index           =   18
         Left            =   -71040
         TabIndex        =   51
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CommandButton btnSysControl 
         Caption         =   "Send E-Mail"
         Height          =   255
         Index           =   17
         Left            =   -71040
         TabIndex        =   50
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CommandButton btnSysControl 
         Caption         =   "Disconnect from the Internet"
         Height          =   495
         Index           =   16
         Left            =   -71040
         TabIndex        =   49
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton btnSysControl 
         Caption         =   "Connect to the Internet"
         Height          =   495
         Index           =   15
         Left            =   -71040
         TabIndex        =   48
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton btnSysControl 
         Caption         =   "Minimize All"
         Height          =   255
         Index           =   14
         Left            =   -72960
         TabIndex        =   47
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CommandButton btnSysControl 
         Caption         =   "Empty Recycle Bin"
         Height          =   255
         Index           =   13
         Left            =   -72960
         TabIndex        =   46
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CommandButton btnSysControl 
         Caption         =   "Close CD-ROM"
         Height          =   255
         Index           =   12
         Left            =   -72960
         TabIndex        =   45
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton btnSysControl 
         Caption         =   "Open CD-ROM"
         Height          =   255
         Index           =   11
         Left            =   -72960
         TabIndex        =   44
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton btnSysControl 
         Caption         =   "Enable Ctl+Alt+Del"
         Height          =   255
         Index           =   10
         Left            =   -72960
         TabIndex        =   43
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton btnSysControl 
         Caption         =   "Disable Ctl+Alt+Del"
         Height          =   255
         Index           =   9
         Left            =   -72960
         TabIndex        =   42
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton btnSysControl 
         Caption         =   "Show Desktop"
         Height          =   255
         Index           =   8
         Left            =   -74880
         TabIndex        =   41
         Top             =   2280
         Width           =   1695
      End
      Begin VB.CommandButton btnSysControl 
         Caption         =   "Hide Desktop"
         Height          =   255
         Index           =   7
         Left            =   -74880
         TabIndex        =   40
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton btnSysControl 
         Caption         =   "Run Screensaver"
         Height          =   255
         Index           =   6
         Left            =   -74880
         TabIndex        =   39
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton btnSysControl 
         Caption         =   "Show Taskbar"
         Height          =   255
         Index           =   5
         Left            =   -74880
         TabIndex        =   38
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton btnSysControl 
         Caption         =   "Hide Taskbar"
         Height          =   255
         Index           =   4
         Left            =   -74880
         TabIndex        =   37
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton btnSysControl 
         Caption         =   "Shutdown Menu"
         Height          =   255
         Index           =   3
         Left            =   -74880
         TabIndex        =   36
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton btnSysControl 
         Caption         =   "Force Logoff"
         Height          =   255
         Index           =   2
         Left            =   -74880
         TabIndex        =   35
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton btnSysControl 
         Caption         =   "Force Reboot"
         Height          =   255
         Index           =   1
         Left            =   -74880
         TabIndex        =   34
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton btnSysControl 
         Caption         =   "Force Shutdown"
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   33
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Set Mode"
         Height          =   495
         Left            =   -73200
         TabIndex        =   32
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Display Properties"
         Height          =   495
         Left            =   -73200
         TabIndex        =   31
         Top             =   2280
         Width           =   855
      End
      Begin VB.ListBox lstRes 
         Height          =   2205
         Left            =   -74880
         TabIndex        =   30
         Top             =   600
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         Caption         =   "Free System Resources"
         Height          =   855
         Left            =   -71160
         TabIndex        =   28
         Top             =   2040
         Width           =   1935
         Begin VB.Label lblFreeSys 
            Alignment       =   2  'Center
            Caption         =   "N/A%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Timer tmrMem 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   -72120
         Top             =   2160
      End
      Begin VB.TextBox txtCOM 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2280
         TabIndex        =   14
         Text            =   "3"
         ToolTipText     =   "The COM Port to dial to"
         Top             =   2640
         Width           =   255
      End
      Begin VB.Frame Frame1 
         Caption         =   "Speed Dial"
         Height          =   1815
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   5655
         Begin MSCommLib.MSComm ComPort 
            Left            =   3720
            Top             =   720
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DTREnable       =   -1  'True
         End
         Begin VB.CommandButton btnSpeed 
            Caption         =   "[Empty]"
            Height          =   375
            Index           =   7
            Left            =   3120
            TabIndex        =   10
            ToolTipText     =   "Right Click to Edit"
            Top             =   1320
            Width           =   2415
         End
         Begin VB.CommandButton btnSpeed 
            Caption         =   "[Empty]"
            Height          =   375
            Index           =   6
            Left            =   3360
            TabIndex        =   9
            ToolTipText     =   "Right Click to Edit"
            Top             =   960
            Width           =   2175
         End
         Begin VB.CommandButton btnSpeed 
            Caption         =   "[Empty]"
            Height          =   375
            Index           =   5
            Left            =   3360
            TabIndex        =   8
            ToolTipText     =   "Right Click to Edit"
            Top             =   600
            Width           =   2175
         End
         Begin VB.CommandButton btnSpeed 
            Caption         =   "[Empty]"
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   7
            ToolTipText     =   "Right Click to Edit"
            Top             =   1320
            Width           =   2415
         End
         Begin VB.CommandButton btnSpeed 
            Caption         =   "[Empty]"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   6
            ToolTipText     =   "Right Click to Edit"
            Top             =   960
            Width           =   2175
         End
         Begin VB.CommandButton btnSpeed 
            Caption         =   "[Empty]"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   5
            ToolTipText     =   "Right Click to Edit"
            Top             =   600
            Width           =   2175
         End
         Begin VB.CommandButton btnSpeed 
            Caption         =   "[Empty]"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   4
            ToolTipText     =   "Right Click to Edit"
            Top             =   240
            Width           =   2415
         End
         Begin VB.CommandButton btnSpeed 
            Caption         =   "[Empty]"
            Height          =   375
            Index           =   4
            Left            =   3120
            TabIndex        =   11
            ToolTipText     =   "Right Click to Edit"
            Top             =   240
            Width           =   2415
         End
         Begin VB.Image Image5 
            Height          =   480
            Left            =   2520
            Picture         =   "frmMain.frx":0613
            Stretch         =   -1  'True
            Top             =   720
            Width           =   600
         End
      End
      Begin VB.CommandButton btnDial 
         Caption         =   "Dial"
         Height          =   255
         Left            =   4800
         TabIndex        =   2
         ToolTipText     =   "Click to Dial the number/Cancel the call"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtDial 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Type your number here"
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label lblUser 
         AutoSize        =   -1  'True
         Caption         =   "User Name: N/A"
         Height          =   195
         Left            =   -74880
         TabIndex        =   67
         Top             =   1320
         Width           =   1185
      End
      Begin VB.Label lblCompName 
         AutoSize        =   -1  'True
         Caption         =   "Computer Name: N/A"
         Height          =   195
         Left            =   -74880
         TabIndex        =   66
         Top             =   1080
         Width           =   1530
      End
      Begin VB.Image Image1 
         Height          =   975
         Left            =   -73200
         Picture         =   "frmMain.frx":0A55
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   855
      End
      Begin VB.Image Image6 
         Height          =   1665
         Left            =   -71160
         Picture         =   "frmMain.frx":0E97
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   1860
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   -70680
         Picture         =   "frmMain.frx":3F19
         Top             =   960
         Width           =   480
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   -71160
         Picture         =   "frmMain.frx":435B
         Top             =   840
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   -70920
         Picture         =   "frmMain.frx":479D
         Top             =   600
         Width           =   480
      End
      Begin VB.Label PagePerc 
         AutoSize        =   -1  'True
         Caption         =   "N/A% Free"
         Height          =   195
         Left            =   -74880
         TabIndex        =   27
         Top             =   2640
         Width           =   780
      End
      Begin VB.Label VMPerc 
         AutoSize        =   -1  'True
         Caption         =   "N/A% Free"
         Height          =   195
         Left            =   -74880
         TabIndex        =   26
         Top             =   1800
         Width           =   780
      End
      Begin VB.Label RamPerc 
         AutoSize        =   -1  'True
         Caption         =   "N/A% Free"
         Height          =   195
         Left            =   -74880
         TabIndex        =   25
         Top             =   960
         Width           =   780
      End
      Begin VB.Label lblCurr 
         Caption         =   "Current Mode: N/AxN/A"
         Height          =   255
         Left            =   -74880
         TabIndex        =   24
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label lblVer 
         AutoSize        =   -1  'True
         Caption         =   "Version: N/A"
         Height          =   195
         Left            =   -74880
         TabIndex        =   23
         Top             =   720
         Width           =   915
      End
      Begin VB.Label lblPlatform 
         AutoSize        =   -1  'True
         Caption         =   "Platform: N/A"
         Height          =   195
         Left            =   -74880
         TabIndex        =   22
         Top             =   480
         Width           =   960
      End
      Begin VB.Label freePage 
         AutoSize        =   -1  'True
         Caption         =   "Free Page File: N/A"
         Height          =   195
         Left            =   -74880
         TabIndex        =   21
         Top             =   2400
         Width           =   1410
      End
      Begin VB.Label totPage 
         AutoSize        =   -1  'True
         Caption         =   "Total Page File: N/A"
         Height          =   195
         Left            =   -74880
         TabIndex        =   20
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label freeVM 
         AutoSize        =   -1  'True
         Caption         =   "Free Virtual Memory: N/A"
         Height          =   195
         Left            =   -74880
         TabIndex        =   19
         Top             =   1560
         Width           =   1785
      End
      Begin VB.Label totVM 
         AutoSize        =   -1  'True
         Caption         =   "Total Virtual Memory: N/A"
         Height          =   195
         Left            =   -74880
         TabIndex        =   18
         Top             =   1320
         Width           =   1830
      End
      Begin VB.Label freeRam 
         AutoSize        =   -1  'True
         Caption         =   "Free RAM: N/A"
         Height          =   195
         Left            =   -74880
         TabIndex        =   17
         Top             =   720
         Width           =   1110
      End
      Begin VB.Label totRam 
         AutoSize        =   -1  'True
         Caption         =   "Total RAM: N/A"
         Height          =   195
         Left            =   -74880
         TabIndex        =   16
         Top             =   480
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "(The mouse is usually COM1)"
         Height          =   255
         Left            =   2640
         TabIndex        =   15
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "COM Port:"
         Height          =   255
         Left            =   1440
         TabIndex        =   13
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Click Cancel To Hang Up!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   30
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   4335
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' DIALER - DIAL BUTTON
Private Sub btnDial_Click()
    If btnDial.Caption = "Dial" Then
        btnDial.Caption = "Cancel"
        Frame1.Visible = False
        txtDial.Enabled = False
        CancelFlag = False
        txtCOM.Enabled = False
        Dial txtDial.Text
    ElseIf btnDial.Caption = "Cancel" Then
        btnDial.Caption = "Dial"
        Frame1.Visible = True
        txtDial.Enabled = True
        CancelFlag = True
        txtCOM.Enabled = True
    End If
End Sub

' DIALER - SPEED DIAL BUTTONS
Private Sub btnSpeed_Click(Index As Integer)
    If Data.Dialer.spdNumber(Index) = "" Then Exit Sub
    txtDial.Text = Data.Dialer.spdNumber(Index)
    btnDial_Click
End Sub

' DIALER - SPEED DIAL BUTTONS (Right mouse click)
Private Sub btnSpeed_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbKeyRButton Then
        Data.Dialer.spdName(Index) = InputBox("Change name to:", "Change name...", btnSpeed(Index).Caption)
        Data.Dialer.spdNumber(Index) = InputBox("Change number to:", "Change number...", Data.Dialer.spdNumber(Index))
    End If
    btnSpeed(Index).Caption = Data.Dialer.spdName(Index)
    SaveSetting "AVINControler", "DialName", Str$(Index), Data.Dialer.spdName(Index)
    SaveSetting "AVINControler", "DialNum", Str$(Index), Data.Dialer.spdNumber(Index)
End Sub

' SCREEN - CLICK A BUTTON
Private Sub btnSysControl_Click(Index As Integer)
    Select Case Index
        Case 0: ShutDown
        Case 1: Restart
        Case 2: LogOff
        Case 3: ShutDown_DIALOG
        Case 4: TaskBarHide
        Case 5: TaskBarShow
        Case 6: runScreenSaver frmMain
        Case 7: DesktopIconsHide
        Case 8: DesktopIconsShow
        Case 9: ALT_CTRL_DEL_Disabled
        Case 10: ALT_CTRL_DEL_Enabled
        Case 11: OpenCDROM
        Case 12: CloseCDROM
        Case 13: EmptyRecycle
        Case 14: MinimizeAll
        Case 15: InternetConnect
        Case 16: InternetDisconnect
        Case 17: SendEmail
        Case 18: HideTime
        Case 19: FlipMouseButtons
        Case 20: ShowTime
        Case 21
        Case 22: disStart
        Case 23: enStart
    End Select
End Sub

' SCREEN - SHOW DISPLAY SETTINGS
Private Sub Command1_Click()
    Shell "rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,0", 5
End Sub

' SCREEN - CHANGE THE RESOLUTION
Private Sub Command2_Click()
    Dim L As Long, Flags As Long, X As Long
    X = lstRes.ListIndex
    Dev(X).dmFields = DM_BitsPerPel Or DM_PelsWidth Or DM_PelsHeight
    Flags = CDS_UpdateRegistry
    L = ChangeDisplaySettings(Dev(X), Flags)
    Select Case L
        Case Disp_Change_Restart
            L = MsgBox("This change will not take effect until you reboot the system.  Reboot now?", vbYesNo)
            If L = vbYes Then
                Flags = 0
                L = ExitWindowsEx(EWX_Reboot, Flags)
            End If
        Case Disp_Change_Successful
        Case Else
            MsgBox "Error changing resolution! Returned: " & L
    End Select
    Bits = GetDeviceCaps(hDC, BitsPixel)
    Wdth = Screen.Width / Screen.TwipsPerPixelX
    Hght = Screen.Height / Screen.TwipsPerPixelY
    lblCurr.Caption = "Current Mode: " & Wdth & "x" & Hght
End Sub

' SCREEN - RESET COLORS
Private Sub Command3_Click()
    SetSysColors 19, IndexArray(0), SavedColors(0)
End Sub

Private Sub Form_Load()
    SetWindowPos hWnd, -1, 0, 0, 0, 0, &H2 Or &H1
    ' DIALER
    ComPort.InputLen = 0
    For X = 0 To 7
        Data.Dialer.spdName(X) = GetSetting("AVINControler", "DialName", Str$(X), "")
        Data.Dialer.spdNumber(X) = GetSetting("AVINControler", "DialNum", Str$(X), "")
        If Data.Dialer.spdName(X) = "" Then Data.Dialer.spdName(X) = "[Empty]"
        btnSpeed(X).Caption = Data.Dialer.spdName(X)
    Next
    ' SYSTEM INFO
    OS.dwOSVersionInfoSize = Len(OS)
    GetVersionEx OS
    If OS.dwPlatformId = 0 Then lblPlatform.Caption = "Platform: Windows 3.1"
    If OS.dwPlatformId = 1 Then lblPlatform.Caption = "Platform: Windows 95/98/ME/2000"
    If OS.dwPlatformId = 2 Then lblPlatform.Caption = "Platform: Windows NT"
    lblVer.Caption = "Version: " & OS.dwMajorVersion & "." & OS.dwMinorVersion & "." & OS.dwBuildNumber
    lblCompName.Caption = "Computer Name: " & ComputerName
    lblUser.Caption = "User Name: " & UserName
    lblAC.Caption = "AC Status: " & ACStat
    lblBatt.Caption = "Battery Status: " & BattStat
    lblLife.Caption = "Battery Life: " & BattLife & " (" & BattPerc & "%)"
    ' SCREEN
    Dim L As Long
    Bits = GetDeviceCaps(hDC, BitsPixel)
    Wdth = Screen.Width / Screen.TwipsPerPixelX
    Hght = Screen.Height / Screen.TwipsPerPixelY
    MaxModes = 8
    ReDim Dev(0 To MaxModes) As DevMode
    NumModes = 0
    L = EnumDisplaySettings(ByVal 0, NumModes, Dev(NumModes))
    lblCurr.Caption = "Current Mode: " & Wdth & "x" & Hght
    Do While L
        lstRes.AddItem Dev(NumModes).dmPelsWidth & "x" & Dev(NumModes).dmPelsHeight & "x" & Dev(NumModes).dmBitsPerPel
        If Bits = Dev(NumModes).dmBitsPerPel And Wdth = Dev(NumModes).dmPelsWidth And Hght = Dev(NumModes).dmPelsHeight Then lstRes.ListIndex = lstRes.NewIndex
        NumModes = NumModes + 1
        If NumModes > MaxModes Then
            MaxModes = MaxModes + 8
            ReDim Preserve Dev(0 To MaxModes) As DevMode
        End If
        L = EnumDisplaySettings(ByVal 0, NumModes, Dev(NumModes))
    Loop
    NumModes = NumModes - 1
    For i = 0 To 18
        SavedColors(i) = GetSysColor(i)
    Next
End Sub

' SCREEN - COLOR LIST
Private Sub lstSysCol_Click()
    Dim Groups As Long, LeftOver As Long
    Dim Rb As String, Gb As String, Bb As String
    Rb = Left$(Hex$(SavedColors(lstSysCol.ListIndex)), 2)
    Gb = Mid$(Hex$(SavedColors(lstSysCol.ListIndex)), 3, 2)
    Bb = Right$(Hex$(SavedColors(lstSysCol.ListIndex)), 2)
    txtRed.Text = GetDec(Rb)
    txtGreen.Text = GetDec(Gb)
    txtBlue.Text = GetDec(Bb)
    shpFill.BackColor = RGB(txtRed.Text, txtGreen.Text, txtBlue.Text)
End Sub

' TABS - CHECK TURN ON MEMORY TIMER
Private Sub tabMain_Click(PreviousTab As Integer)
    If tabMain.Tab = 2 Then
        tmrMem.Enabled = True
    Else
        tmrMem.Enabled = False
    End If
End Sub

' MEMORY MANAGER - TIMER
Private Sub tmrMem_Timer()
    On Local Error Resume Next
    Memory.dwLength = Len(Memory)
    GlobalMemoryStatus Memory
    With Memory
        On Error Resume Next
        totRam.Caption = "Total RAM: " & Format$(.dwTotalPhys / 1024, "#,###") & " KB"
        freeRam.Caption = "Free RAM: " & Format$(.dwAvailPhys / 1024, "#,###") & " KB"
        RamPerc.Caption = Int((.dwAvailPhys / .dwTotalPhys) * 100) & "% Free"
        On Error Resume Next
        totVM.Caption = "Total Virtual Memory: " & Format$(.dwTotalVirtual / 1024, "#,###") & " KB"
        freeVM.Caption = "Free Virtual Memory: " & Format$(.dwAvailVirtual / 1024, "#,###") & " KB"
        VMPerc.Caption = Int((.dwAvailVirtual / .dwTotalVirtual) * 100) & "% Free"
        On Error Resume Next
        totPage.Caption = "Total Page File: " & Format$(.dwTotalPageFile / 1024, "#,###") & " KB"
        freePage.Caption = "Free Page File: " & Format$(.dwAvailPageFile / 1024, "#,###") & " KB"
        PagePerc.Caption = Int((.dwAvailPageFile / .dwTotalPageFile) * 100) & "% Free"
        On Error Resume Next
        a = (.dwAvailPageFile / 1024) + (.dwAvailPhys / 1024) + (.dwAvailVirtual / 1024)
        B = (.dwTotalPageFile / 1024) + (.dwTotalPhys / 1024) + (.dwTotalVirtual / 1024)
        lblFreeSys.Caption = Int((a / B) * 100) & "%"
    End With
End Sub

