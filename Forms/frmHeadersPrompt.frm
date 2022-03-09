VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHeadersPrompt
   Caption         =   "'Fake headers' are table headers that are not in the actual header, but in the main page of the document."
   ClientHeight    =   5700
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8616.001
   OleObjectBlob   =   "frmHeadersPrompt.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmHeadersPrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDone_Click()

Unload Me
Call FakeHeaders.CheckInput

End Sub

Private Sub cmdNoFake_Click()

skipFakes = True
frmProgress.stsFakeHeaders.Enabled = False
Unload Me

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

If CloseMode = 0 Then
    If MsgBox("Are you sure you want to exit?" & vbNewLine & "I have not even started yet!", _
        vbQuestion + vbYesNo, "Bruh...") = vbNo Then Cancel = True Else Documents(selDoc).Close (wdDoNotSaveChanges): Call cfgEnd: End
End If

End Sub
