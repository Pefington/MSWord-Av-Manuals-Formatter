Attribute VB_Name = "Hyperlinks"
Option Explicit

Sub Clean()

stsUpdate "stsHyperlinks", True
timeSubInit = Timer

Dim i As Long, hCount As Long, str As String, pHcount As Long, tHcount As Long, _
    c As cell, cCount As Long, rg As Range, h As Hyperlink, link As Hyperlink
hCount = Documents(selDoc).Hyperlinks.Count
i = 0

With Documents(selDoc)

    For i = .TablesOfContents.Count To 1 Step -1

        .TablesOfContents(i).Range.Delete

    Next

    For i = hCount To 1 Step -1

        On Error GoTo ErrI
        With Documents(selDoc).Hyperlinks(i)
        On Error GoTo 0

UpdateProgress ((hCount - i) / hCount)
UpdateCaption (hCount - i & " of " & hCount & " hyperlinks processed.")

        str = .Range.Text

        If devMode Then .Range.Paragraphs(1).Range.Select

            If InStr(str, "@") <> 0 Or InStr(str, "www") <> 0 Then GoTo nextI

            If .Range.Tables.Count > 0 Then
                cCount = 0: tHcount = .Range.Tables(1).Range.Hyperlinks.Count
                For Each c In .Range.Tables(1).Range.Cells
                    If devMode Then c.Select
                    If c.Range.Characters.Count = 1 Then cCount = cCount + 1
                Next

                If tHcount >= (.Range.Tables(1).Range.Cells.Count - cCount) / 2 Then _
                    .Range.Tables(1).Delete: GoTo nextI
            End If

            For Each link In .Range.Paragraphs(1).Range.Hyperlinks
                link.Delete
            Next

            'On Error GoTo ErrorDelete
            .Range.Paragraphs(1).Range.Delete
            'On Error GoTo 0

            If pHcount > 1 Then i = i - pHcount + 1

        End With

nextI:
    Next

End With

timeHyperlinks = MinSec(Timer - timeSubInit)
stsUpdate "stsHyperlinks", False

Exit Sub
'**************************************************************
ErrI:
i = i - 1: Resume

errDelete:
h.Range.Paragraphs(1).Range.Cut: Resume Next
'**************************************************************
End Sub
