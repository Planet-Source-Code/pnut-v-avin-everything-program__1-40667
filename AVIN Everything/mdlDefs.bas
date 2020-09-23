Attribute VB_Name = "mdlDefs"
' SYSTEM INFORMATION
Type OSVersionInfo
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128
End Type
Public Type SYSTEM_POWER_STATUS
        ACLineStatus As Byte
        BatteryFlag As Byte
        BatteryLifePercent As Byte
        Reserved1 As Byte
        BatteryLifeTime As Long
        BatteryFullLifeTime As Long
End Type
Public OS As OSVersionInfo
Public Power As SYSTEM_POWER_STATUS
Public Const AC_BackupPower = &H2
Public Const AC_Offline = &H0
Public Const AC_Online = &H1
Public Const AC_Unknown = &HFF
Public Const Battery_Charging = &H8
Public Const Battery_Critical = &H4
Public Const Battery_High = &H1
Public Const Battery_Low = &H2
Public Const Battery_NoBattery = &H80
Public Const Battery_Unknown = &HFF
Public Const Battery_LifeUnknown = &HFFFF
Public Const Battery_PercentageUnknown = &HFF

' PHONE DIALER
Type tDialInfo
    ComPort As Integer
End Type
Type tDialer
    spdName(7) As String
    spdNumber(7) As String
End Type

' MEMORY MANAGER
Type MemoryStatus
    dwLength        As Long
    dwMemoryLoad    As Long
    dwTotalPhys     As Long
    dwAvailPhys     As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual  As Long
    dwAvailVirtual  As Long
End Type
Public Memory As MemoryStatus

' SCREEN
Const CCHDeviceName = 32
Const CCHFormName = 32
Type DevMode
    dmDeviceName As String * CCHDeviceName
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFormName
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type
Const DM_BitsPerPel = &H40000
Const DM_PelsWidth = &H80000
Const DM_PelsHeight = &H100000
Const DM_DisplayFlags = &H200000
Const DM_DisplayFrequency = &H400000
Const BitsPixel = 12
Const CDS_UpdateRegistry = &H1
Const CDS_Test = &H2
Const CDS_FullScreen = &H4
Const CDS_Global = &H8
Const CDS_Set_Primary = &H10
Const CDS_Reset = &H40000000
Const CDS_SetRect = &H20000000
Const CDS_NoReset = &H10000000
Const Disp_Change_Successful = 0
Const Disp_Change_Restart = 1
Const Disp_Change_Failed = -1
Const Disp_Change_BadMode = -2
Const Disp_Change_NotUpdated = -3
Const Disp_Change_BadFlags = -4
Const Disp_Change_BadParam = -5
Const EWX_LogOff = 0
Const EWX_Shutdown = 1
Const EWX_Reboot = 2
Const EWX_Force = 4
Const COLOR_SCROLLBAR = 0
Const COLOR_BACKGROUND = 1
Const COLOR_ACTIVECAPTION = 2
Const COLOR_INACTIVECAPTION = 3
Const COLOR_MENU = 4
Const COLOR_WINDOW = 5
Const COLOR_WINDOWFRAME = 6
Const COLOR_MENUTEXT = 7
Const COLOR_WINDOWTEXT = 8
Const COLOR_CAPTIONTEXT = 9
Const COLOR_ACTIVEBORDER = 10
Const COLOR_INACTIVEBORDER = 11
Const COLOR_APPWORKSPACE = 12
Const COLOR_HIGHLIGHT = 13
Const COLOR_HIGHLIGHTTEXT = 14
Const COLOR_BTNFACE = 15
Const COLOR_BTNSHADOW = 16
Const COLOR_GRAYTEXT = 17
Const COLOR_BTNTEXT = 18
Public Dev() As DevMode, NumModes As Long, MaxModes As Long, Bits As Long, Wdth As Long, Hght As Long
Public SavedColors(18) As Long, IndexArray(18) As Long, NewColors(18) As Long

' SYSTEM CONTROLS
Type POINTAPI
    X As Long
    Y As Long
End Type
Type NOTIFYICONDATA
   cbSize As Long
   hWnd As Long
   uId As Long
   uFlags As Long
   uCallBackMessage As Long
   hIcon As Long
   szTip As String * 64
End Type

' ALL OF THE TABS
Type tTabs
    Dialer As tDialer
End Type
Public Data As tTabs
