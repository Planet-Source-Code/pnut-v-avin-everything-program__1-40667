Attribute VB_Name = "mdlDialer"
Public CancelFlag As Boolean

Public Sub Dial(Number As String)
    Dim fromModem As String
    Dim toDial As String
    
    With frmMain.ComPort
        toDial = "ATDT" + Number + ";" + vbCr
        .CommPort = frmMain.txtCOM.Text
    
        On Error Resume Next
        .PortOpen = True
        If Err Then
            MsgBox "COM" & frmMain.txtCOM.Text & " Not Available, try another port!", vbCritical + vbOKOnly, "Error"
            GoTo 10
        End If
    
        .InBufferCount = 0
        .Output = toDial
    
        Do
            DoEvents
            If .InBufferCount Then
                fromModem = fromModem + .Input
                If InStr(fromModem, "OK") Then
                    Beep
                    MsgBox "Please pick up phone then press OK" & vbNewLine & vbNewLine & "Or press OK to hang up", vbOKOnly + vbInformation, "Number dialed!"
                    Exit Do
                End If
            End If
            If CancelFlag Then
                CancelFlag = False
                Exit Do
            End If
        Loop
        
        .Output = "ATH" + vbCr
        .PortOpen = False
    End With
10  With frmMain
        .btnDial.Caption = "Dial"
        .Frame1.Visible = True
        .txtDial.Enabled = True
        .txtCOM.Enabled = True
    End With
End Sub
