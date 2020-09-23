VERSION 5.00
Begin VB.Form frmDialCancel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dialing..."
   ClientHeight    =   1020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2685
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   2685
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCancel 
      Caption         =   "Can&cel"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblMain 
      AutoSize        =   -1  'True
      Caption         =   "Dialing 555-555-5555, please wait..."
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmDialCancel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub

Private Sub btnCancel_Click()
    frmMain.Enabled = True
    frmMain.Show
    Me.Hide
    Unload Me
End Sub

Private Sub Form_Load()
    lblMain.Caption = "Dialing " & frmMain.txtDial.Text & vbNewLine & "Please Wait..."
    lblMain.Left = 50
    Width = lblMain.Width + (lblMain.Width * 0.2)
    Height = lblMain.Height + btnCancel.Height + 200
    lblMain.Top = 50
    btnCancel.Top = lblMain.Height + 50
End Sub
