Attribute VB_Name = "mdlSysInfo"
Public Function ComputerName() As String
    Dim a As String
    a = Space$(100)
    GetComputerName a, 100
    ComputerName = Trim(a)
End Function

Public Function UserName() As String
    Dim a As String
    a = Space$(100)
    GetUserName a, 100
    UserName = Trim(a)
End Function

Public Function ACStat() As String
    GetSystemPowerStatus Power
    Select Case Power.ACLineStatus
        Case AC_Online
            ACStat = "Online"
        Case AC_Offline
            ACStat = "Offline"
        Case AC_BackupPower
            ACStat = "Backup"
        Case AC_Unknown
            ACStat = "N/A"
    End Select
End Function


Public Function BattStat() As String
    GetSystemPowerStatus Power
    Select Case Power.BatteryFlag
        Case Battery_NoBattery
            BattStat = "No Battery"
        Case Battery_Unknown
            BattStat = "N/A"
    End Select
End Function

Public Function BattLife() As String
    GetSystemPowerStatus Power
    Select Case Power.BatteryLifeTime
        Case Battery_High
            BattLife = "High"
        Case Battery_Low
            BattLife = "Low"
        Case Battery_Critical
            BattLife = "CRITICALLY LOW"
        Case Battery_LifeUnknown
            BattLife = "N/A"
    End Select
End Function

Public Function BattPerc() As Integer
    GetSystemPowerStatus Power
    BattPerc = Power.BatteryLifePercent
    If Power.BatteryLifePercent = Battery_PercentageUnknown Then BattPerc = 0
End Function
