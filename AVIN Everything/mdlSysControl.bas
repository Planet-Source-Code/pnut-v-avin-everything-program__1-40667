Attribute VB_Name = "mdlSysControl"
Dim a As Long

Function ShutDown()
    Dim lngresult
    lngresult = ExitWindowsEx(1, 0&)
End Function

Function Restart()
    Dim lngresult
    lngresult = ExitWindowsEx(2, 0&)
End Function

Function LogOff()
    Dim lngresult
    lngresult = ExitWindowsEx(0, 0&)
End Function

Function TaskBarHide()
    Dim rtn
    rtn = FindWindow("Shell_traywnd", "")
    Call SetWindowPos(rtn, 0, 0, 0, 0, 0, &H80)
End Function

Function TaskBarShow()
    Dim rtn As Long
    rtn = FindWindow("Shell_traywnd", "")
    Call SetWindowPos(rtn, 0, 0, 0, 0, 0, &H40)
End Function

Function runScreenSaver(mainForm As Form)
    SendMessage mainForm.hWnd, &H112&, &HF140&, 0&
End Function

Function DesktopIconsShow()
    Dim hW As Long
    hW = FindWindowEx(0&, 0&, "Progman", vbNullString)
    ShowWindow hW, 5
End Function

Function DesktopIconsHide()
    Dim hW As Long
    hW = FindWindowEx(0&, 0&, "Progman", vbNullString)
    ShowWindow hW, 0
End Function

Function ALT_CTRL_DEL_Enabled()
    callMe (False)
End Function

Function ALT_CTRL_DEL_Disabled()
    callMe (True)
End Function

Private Sub callMe(huh As Boolean)
    Dim gd
    gd = SystemParametersInfo(97, huh, CStr(1), 0)
End Sub

Function OpenCDROM()
    Dim lngReturn As Long
    Dim strReturn As Long
    lngReturn = mciSendString("set CDAudio door open", strReturn, 127, 0)
End Function

Function CloseCDROM()
    Dim lngReturn As Long
    Dim strReturn As Long
    lngReturn = mciSendString("set CDAudio door closed", strReturn, 127, 0)
End Function

Function EmptyRecycle()
    Dim ReturnValue
    ReturnValue = SHEmptyRecycleBin(frmMain.hWnd, "", &H2)
End Function

Function MinimizeAll()
    Call keybd_event(&H5B, 0, 0, 0)
    Call keybd_event(77, 0, 0, 0)
    Call keybd_event(&H5B, 0, &H2, 0)
End Function

Function InternetConnect()
    Dim lResult As Long
    lResult = InternetAutodial(2, 0&)
End Function

Function InternetDisconnect()
    Dim lResult As Long
    lResult = InternetAutodialHangup(0&)
End Function

Function SendEmail()
    ShellExecute hWnd, "open", "mailto:", vbNullString, vbNullString, 5
End Function

Function FlipMouseButtons()
    If a = 0 Then a = 1: GoTo 1
    If a = 1 Then a = 0
1   ReturnValue = SwapMouseButton(a)
End Function

Function ShutDown_DIALOG()
    ShutDown_DIALOG = SHShutDownDialog(0)
End Function

Function HideTime()
    Dim P As Long, C As Long, a As Long
    P = FindWindow("Shell_TrayWnd", vbNullString)
    C = FindWindowEx(P&, 0&, "TrayNotifyWnd", vbNullString)
    a = FindWindowEx(C&, 0&, "TrayClockWClass", vbNullString)
    Call ShowWindow(a&, 0)
End Function

Public Sub ShowTime()
    Dim P As Long, C As Long, a As Long
    P = FindWindow("Shell_TrayWnd", vbNullString)
    C = FindWindowEx(P, 0&, "TrayNotifyWnd", vbNullString)
    a = FindWindowEx(C, 0&, "TrayClockWClass", vbNullString)
    Call ShowWindow(a, 5)
End Sub

Public Sub disStart()
    Dim Taskbar As Long, StartButton As Long
    Taskbar& = FindWindow("Shell_TrayWnd", vbNullString)
    StartButton& = FindWindowEx(Taskbar&, 0&, "Button", vbNullString)
    EnableWindow StartButton&, 0
End Sub

Public Sub enStart()
    Dim Taskbar As Long, StartButton As Long
    Taskbar& = FindWindow("Shell_TrayWnd", vbNullString)
    StartButton& = FindWindowEx(Taskbar&, 0&, "Button", vbNullString)
    EnableWindow StartButton&, 1
End Sub
