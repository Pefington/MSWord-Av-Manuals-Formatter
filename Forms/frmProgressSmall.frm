VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgressSmall
   ClientHeight    =   348
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10080
   OleObjectBlob   =   "frmProgressSmall.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmProgressSmall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkMaxi_Click()

If chkMaxi.Value = False Then
    frmProgress.chkMini.Value = False
    Me.hide
    frmProgress.Show
    If devMode Then Me.Left = Application.UsableWidth - (Me.Width + 100)
End If

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

If CloseMode = 0 Then
    If MsgBox("Are you sure you want to exit?" & vbNewLine & "I am only partly done.", _
        vbQuestion + vbYesNo, "Cannot wait?") = vbNo Then Cancel = True Else Call cfgEnd: End
End If

End Sub
