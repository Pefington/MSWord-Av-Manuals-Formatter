VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgress
   ClientHeight    =   6012
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8388.001
   OleObjectBlob   =   "frmProgress.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkDev_Click()

If chkDev.Value = True Then
    Call cfgDev
    Me.Show
    Me.Left = Application.UsableWidth - (Me.Width + 100)
ElseIf chkDev.Value = False Then
    Call cfgStart
    Me.Show
End If

End Sub

Private Sub chkMini_Click()

If chkMini.Value = True Then
    Me.hide
    frmProgressSmall.chkMaxi.Value = True
    frmProgressSmall.Show
End If

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

If CloseMode = 0 Then
    If MsgBox("Are you sure you want to exit?" & vbNewLine & "I am only partly done.", _
        vbQuestion + vbYesNo, "Cannot wait?") = vbNo Then Cancel = True Else Call cfgEnd: End
End If

End Sub
