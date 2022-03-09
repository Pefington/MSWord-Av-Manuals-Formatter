VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCoverSetup
   ClientHeight    =   6684
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8748.001
   OleObjectBlob   =   "frmCoverSetup.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmCoverSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCoverDone_Click()

With cmdCoverDone
    If allOK Then
        .Enabled = True
        Me.hide
    Else
        .Enabled = False
        lblCoverForm.Caption = "All the fields must be valid."
        lblCoverForm.ForeColor = wdColorRed
    End If
End With

End Sub

Private Sub cmdSkipCover_Click()

Me.hide
skipFakes = True
frmProgress.stsCover.Enabled = False

End Sub

Private Sub UserForm_Initialize()

cmdCoverDone.Enabled = False
refOK = True
authOK = True
issueOK = True
revOK = True
dateOK = True
allOK = False
fillDATE.Value = Format((DateSerial(Year(Now), Month(Now) + 1, 1)), "dd mmm yy")

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If CloseMode = 0 Then

        If MsgBox("Are you sure you want to exit?" & vbNewLine & "I have not even started yet!", _
            vbQuestion + vbYesNo, "Bruh...") = vbNo Then Cancel = True Else Documents(selDoc).Close (wdDoNotSaveChanges): Call cfgEnd: End

    End If

End Sub

Private Sub fillDEP_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)

With fillDEP
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub

Private Sub fillDEP_Change()

With fillDEP
    If .Value Like "[A-z][A-z][A-z]" Then
        .ForeColor = wdColorGreen
        depOK = True
    Else
        depOK = False
        .ForeColor = wdColorLightOrange
    End If
End With

End Sub

Private Sub fillDEP_Exit(ByVal Cancel As MSForms.ReturnBoolean)

With fillDEP
    If .Value Like "[A-z][A-z][A-z]" And .Value <> "DEP" Then
        .Value = UCase(.Value)
        .ForeColor = wdColorGreen
        lblCoverForm = "~ Please Fill ~"
        lblCoverForm.ForeColor = wdColorLightBlue
        Cancel = False
        depOK = True
    Else
        depOK = False
        Cancel = True
        .ForeColor = wdColorRed
        lblCoverForm.ForeColor = wdColorRed
        lblCoverForm = "Please enter a department."
        .SelStart = 0
        .SelLength = Len(.Value)
    End If
End With

End Sub

Private Sub fillREF_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)

With fillREF
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub

Private Sub fillREF_Change()

With fillREF
    If .Value Like "#" Or .Value Like "[0-5]#" Or .Value Like "0[0-5]#" Then
        .ForeColor = wdColorGreen
        refOK = True
    Else
        refOK = False
        .ForeColor = wdColorLightOrange
    End If
End With

End Sub

'Private Sub fillREF_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
'
'With fillREF
'    If .Value Like "#" Or .Value Like "[0-4]#" Or .Value Like "0[0-4]#" Then
'        Do Until .Value Like "0[0-4]#"
'            .Value = "0" & .Value
'            DoEvents
'        Loop
'        .ForeColor = wdColorGreen
'        lblCoverForm = "~ Please Fill ~"
'        lblCoverForm.ForeColor = wdColorLightBlue
'        Cancel = False
'        refOK = True
'    Else
'        .SelStart = 0
'        .SelLength = Len(.Value)
'        refOK = False
'        Cancel = True
'        .ForeColor = wdColorRed
'        lblCoverForm.ForeColor = wdColorRed
'        If Not IsNumeric(.Value) Then
'            lblCoverForm = "Reference number is.. well.. a number. You with me?"
'        ElseIf .Value > 49 Then
'            lblCoverForm = "Are you sure the department has that many manuals?"
'        End If
'    End If
'End With
'
'End Sub

Private Sub fillREF_Exit(ByVal Cancel As MSForms.ReturnBoolean)

With fillREF
    If .Value Like "#" Or .Value Like "[0-4]#" Or .Value Like "0[0-4]#" Then
        Do Until .Value Like "0[0-4]#"
            .Value = "0" & .Value
            DoEvents
        Loop
        .ForeColor = wdColorGreen
        lblCoverForm = "~ Please Fill ~"
        lblCoverForm.ForeColor = wdColorLightBlue
        Cancel = False
        refOK = True
    Else
        .SelStart = 0
        .SelLength = Len(.Value)
        refOK = False
        Cancel = True
        .ForeColor = wdColorRed
        lblCoverForm.ForeColor = wdColorRed
        If Not IsNumeric(.Value) Then
            lblCoverForm = "Reference number is.. well.. a number. You with me?"
        ElseIf .Value > 49 Then
            lblCoverForm = "Are you sure the department has that many manuals?"
        End If
    End If
End With

End Sub

Private Sub fillTITLE_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)

With fillTITLE
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub

Private Sub fillTITLE_Change()

With fillTITLE
    If Len(.Value) >= 8 Then
        .ForeColor = wdColorGreen
        titleOK = True
    Else
        titleOK = False
        .ForeColor = wdColorLightOrange
    End If
End With

End Sub

'Private Sub fillTITLE_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
'
'With fillTITLE
'    If Len(.Value) > 8 Then
'        .Value = UCase(.Value)
'        .ForeColor = wdColorGreen
'        lblCoverForm = "~ Please Fill ~"
'        lblCoverForm.ForeColor = wdColorLightBlue
'        Cancel = False
'        titleOK = True
'    Else
'        .SelStart = 0
'        .SelLength = Len(.Value)
'        titleOK = False
'        .ForeColor = wdColorRed
'        lblCoverForm.ForeColor = wdColorRed
'        Cancel = True
'        lblCoverForm = "Please enter a title."
'    End If
'End With
'
'End Sub

Private Sub fillTITLE_Exit(ByVal Cancel As MSForms.ReturnBoolean)

With fillTITLE
    If Len(.Value) > 8 Then
        .Value = UCase(.Value)
        .ForeColor = wdColorGreen
        lblCoverForm = "~ Please Fill ~"
        lblCoverForm.ForeColor = wdColorLightBlue
        Cancel = False
        titleOK = True
    Else
        .SelStart = 0
        .SelLength = Len(.Value)
        titleOK = False
        .ForeColor = wdColorRed
        lblCoverForm.ForeColor = wdColorRed
        Cancel = True
        lblCoverForm = "Please enter a title."
    End If
End With

End Sub

Private Sub fillSUBT_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)

With fillSUBT
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub

Private Sub fillSUBT_Change()

With fillSUBT
    If Len(.Value) >= 8 Then
        .ForeColor = wdColorGreen
        subtOK = True
    Else
        subtOK = False
        .ForeColor = wdColorLightOrange
    End If
End With

End Sub

Private Sub fillSUBT_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

With fillSUBT
    If Len(.Value) > 8 Then
        .Value = UCase(.Value)
        .ForeColor = wdColorGreen
        lblCoverForm = "~ Please Fill ~"
        lblCoverForm.ForeColor = wdColorLightBlue
        Cancel = False
        subtOK = True
    Else
        .SelStart = 0
        .SelLength = Len(.Value)
        subtOK = False
        .ForeColor = wdColorRed
        lblCoverForm.ForeColor = wdColorRed
        Cancel = True
        lblCoverForm = "Please enter a subtitle."
    End If
End With

End Sub

'Private Sub fillSUBT_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'
'With fillSUBT
'    If Len(.Value) > 8 Then
'        .Value = UCase(.Value)
'        .ForeColor = wdColorGreen
'        lblCoverForm = "~ Please Fill ~"
'        lblCoverForm.ForeColor = wdColorLightBlue
'        Cancel = False
'        subtOK = True
'    Else
'        .SelStart = 0
'        .SelLength = Len(.Value)
'        subtOK = False
'        .ForeColor = wdColorRed
'        lblCoverForm.ForeColor = wdColorRed
'        Cancel = True
'        lblCoverForm = "Please enter a subtitle."
'    End If
'End With
'
'End Sub

Private Sub fillAUTH_Change()

With fillAUTH
    If Len(.Value) >= 10 Then
        .ForeColor = wdColorGreen
        authOK = True
    Else
        .ForeColor = wdColorLightOrange
        authOK = False
    End If
End With

End Sub

Private Sub fillAUTH_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

With fillAUTH
    If Len(.Value) >= 10 Then
        .Value = StrConv(.Value, vbProperCase)
        .ForeColor = wdColorGreen
        lblCoverForm = "~ Please Fill ~"
        lblCoverForm.ForeColor = wdColorLightBlue
        Cancel = False
        authOK = True
    Else
        .SelStart = 0
        .SelLength = Len(.Value)
        authOK = False
        .ForeColor = wdColorRed
        lblCoverForm.ForeColor = wdColorRed
        Cancel = True
        lblCoverForm = "Please enter a publishing authority."
    End If
End With

End Sub

Private Sub fillISSUE_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)

With fillISSUE
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub

Private Sub fillISSUE_Change()

With fillISSUE
    If .Value Like "#" Or .Value Like "[0-1]#" Then
        .ForeColor = wdColorGreen
        issueOK = True
    Else
        issueOK = False
        .ForeColor = wdColorLightOrange
    End If
End With

End Sub

Private Sub fillISSUE_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

With fillISSUE
    If .Value Like "#" Or .Value Like "[0-1]#" Then
        If .Value Like "0#" Then .Value = Right(.Value, 1)
        .ForeColor = wdColorGreen
        lblCoverForm = "~ Please Fill ~"
        lblCoverForm.ForeColor = wdColorLightBlue
        Cancel = False
        issueOK = True
    Else
        .SelStart = 0
        .SelLength = Len(.Value)
        issueOK = False
        Cancel = True
        .ForeColor = wdColorRed
        lblCoverForm.ForeColor = wdColorRed
        If Not IsNumeric(.Value) And .Value <> vbNullString Then
            lblCoverForm = "Issue number is.. well.. a number. You with me?"
        ElseIf .Value > 19 And .Value <> vbNullString Then
            lblCoverForm = "Man you've got too many issues.."
        ElseIf .Value = vbNullString Then
            lblCoverForm = "Can't be empty."
        End If
    End If
End With

End Sub

'Private Sub fillISSUE_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'
'With fillISSUE
'    If .Value Like "#" Or .Value Like "[0-1]#" Then
'        If .Value Like "0#" Then .Value = Right(.Value, 1)
'        .ForeColor = wdColorGreen
'        lblCoverForm = "~ Please Fill ~"
'        lblCoverForm.ForeColor = wdColorLightBlue
'        Cancel = False
'        issueOK = True
'    Else
'        .SelStart = 0
'        .SelLength = Len(.Value)
'        issueOK = False
'        Cancel = True
'        .ForeColor = wdColorRed
'        lblCoverForm.ForeColor = wdColorRed
'        If Not IsNumeric(.Value) And .Value <> vbNullString Then
'            lblCoverForm = "Issue number is.. well.. a number. You with me?"
'        ElseIf .Value > 19 And .Value <> vbNullString Then
'            lblCoverForm = "Man you've got too many issues.."
'        ElseIf .Value = vbNullString Then
'            lblCoverForm = "Can't be empty."
'        End If
'    End If
'End With
'
'End Sub

Private Sub fillREV_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)

With fillREV
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub

Private Sub fillREV_Change()

With fillREV
    If .Value Like "#" Or .Value Like "[0-1]#" Then
        .ForeColor = wdColorGreen
        revOK = True
    Else
        revOK = False
        .ForeColor = wdColorLightOrange
    End If
End With

End Sub

Private Sub fillREV_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

With fillREV
    If .Value Like "#" Or .Value Like "[0]#" Then
        If .Value Like "0#" Then .Value = Right(.Value, 1)
        .ForeColor = wdColorGreen
        lblCoverForm = "~ Please Fill ~"
        lblCoverForm.ForeColor = wdColorLightBlue
        Cancel = False
        revOK = True
    Else
        .SelStart = 0
        .SelLength = Len(.Value)
        revOK = False
        .ForeColor = wdColorRed
        lblCoverForm.ForeColor = wdColorRed
        Cancel = True
        If Not IsNumeric(.Value) And .Value <> vbNullString Then
            lblCoverForm = "Revision number is.. well.. a number. You with me?"
        ElseIf .Value > 19 And .Value <> vbNullString Then
            lblCoverForm = "That looks like too many revisions. Might as well reissue, am I right?"
        ElseIf .Value = vbNullString Then
            lblCoverForm = "Can't be empty."
        End If
    End If
End With

End Sub

'Private Sub fillREV_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'
'With fillREV
'    If .Value Like "#" Or .Value Like "[0]#" Then
'        If .Value Like "0#" Then .Value = Right(.Value, 1)
'        .ForeColor = wdColorGreen
'        lblCoverForm = "~ Please Fill ~"
'        lblCoverForm.ForeColor = wdColorLightBlue
'        Cancel = False
'        revOK = True
'    Else
'        .SelStart = 0
'        .SelLength = Len(.Value)
'        revOK = False
'        .ForeColor = wdColorRed
'        lblCoverForm.ForeColor = wdColorRed
'        Cancel = True
'        If Not IsNumeric(.Value) And .Value <> vbNullString Then
'            lblCoverForm = "Revision number is.. well.. a number. You with me?"
'        ElseIf .Value > 19 And .Value <> vbNullString Then
'            lblCoverForm = "That looks like too many revisions. Might as well reissue, am I right?"
'        ElseIf .Value = vbNullString Then
'            lblCoverForm = "Can't be empty."
'        End If
'    End If
'End With
'
'End Sub

Private Sub fillDATE_Change()

With fillDATE
    If IsDate(.Value) And Format(Now, "yyyymmdd") < Format(.Value, "yyyymmdd") Then
        .ForeColor = wdColorGreen
        dateOK = True
    Else
        dateOK = False
        .ForeColor = wdColorLightOrange
    End If
End With

End Sub

Private Sub fillDATE_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

With fillDATE
    If IsDate(.Value) And Format(Now, "yyyymmdd") < Format(.Value, "yyyymmdd") Then
        .Value = Format(.Value, "dd mmm yy")
        .ForeColor = wdColorGreen
        lblCoverForm = "~ Please Fill ~"
        lblCoverForm.ForeColor = wdColorLightBlue
        Cancel = False
        dateOK = True
    Else
        .SelStart = 0
        .SelLength = Len(.Value)
        dateOK = False
        .ForeColor = wdColorRed
        lblCoverForm.ForeColor = wdColorRed
        Cancel = True
        If Not IsDate(.Value) Then
            lblCoverForm = "Please enter a valid date."
        ElseIf Format(Now, "yyyymmdd") > Format(.Value, "yyyymmdd") Then
            lblCoverForm = "I'm sorry. Are you from the past?"
        End If
    End If
End With

End Sub
