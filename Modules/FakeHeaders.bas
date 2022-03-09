Attribute VB_Name = "FakeHeaders"
Option Explicit

Sub Init()

    Call cfgDev
    frmHeadersPrompt.Show

End Sub

Sub CheckInput()

Selection.Expand wdParagraph: selP = Selection.Text

If selP = "" Or Len(selP) < 5 Then

    If MsgBox("Your cursor location doesn't seem right." & vbNewLine & vbNewLine & _
        "Try again?", vbRetryCancel, "Bruh...") = vbRetry Then

        Call Init: Exit Sub

        Else: End

    End If

End If

End Sub

Sub CleanTables()

stsUpdate "stsFakeHeaders", True
timeSubInit = Timer

Dim t As Table, tCount As Long, c As cell, i As Long, d As Long
tCount = Documents(selDoc).Tables.Count
d = 0
i = 0

For Each t In Documents(selDoc).Tables

    If i > tCount Then i = tCount

    For Each c In t.Range.Cells

        If devMode Then c.Range.Select

        If c.Range.Text = selP Then t.Delete: d = d + 1

    Next

    UpdateProgress (i / tCount)
    UpdateCaption (i & " of " & tCount & _
    " tables inspected for fake header, " & d & " tables deleted.")
    i = i + 1
Next

timeFakeHeaders = MinSec(Timer - timeSubInit)
stsUpdate "stsFakeHeaders", False

End Sub
