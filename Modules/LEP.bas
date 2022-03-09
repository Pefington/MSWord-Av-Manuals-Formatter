Attribute VB_Name = "LEP"
Option Explicit

Sub Generate()

'Call cfgDev

stsUpdate "stsLEP", True
timeSubInit = Timer

Dim s As Section, sCount As Long, rg As Range, i As Long, lStr As String, _
seqID As String, pCount As Long, p As Long, sTitle As String, fText As String, _
sContinue As Boolean, oldpCount As Long, rCount As Long, r As Row
sCount = Documents(selDoc).Sections.Count
ASTBlue = 12611584
i = 0

For Each s In Documents(selDoc).Sections

UpdateProgress (i / sCount)
UpdateCaption (i & " of " & sCount & " sections added to LEP.")

With s

If .Index = 1 Then GoTo nextS
sContinue = IIf(Not .Range.Paragraphs(1).Style Like "Heading 1*", True, False)

On Error Resume Next
If Not sContinue Then fText = _
.Footers(wdHeaderFooterPrimary).Range.Paragraphs(2).Range.Text
If Err.Number = 0 Then
    On Error GoTo 0
    fText = Replace(fText, vbCr, vbNullString)
    pCount = .Range.ComputeStatistics(wdStatisticPages) - 1
    DoEvents
    sTitle = Replace(.Range.Paragraphs(1).Range.Text, vbCr, vbNullString)
    lStr = Replace(.Range.Paragraphs(1).Range.ListFormat.ListString, vbCr, vbNullString)
    With Documents(selDoc).Sections(6).Range.Tables(1).Rows
        If Not sContinue Then
            .Add
            With .Last
                .Cells.Merge
                With .Range
                    If lStr Like "CHAPTER*" Then
                        .Text = lStr
                        seqID = Replace(lStr, " ", vbNullString)
                    Else
                        .Text = sTitle
                        seqID = Replace(sTitle, " ", vbNullString)
                    End If
                End With
            End With
        End If
        DoEvents
        .Add
        With .Last
            If .Cells.Count = 1 Then .Cells.Split NumColumns:=3
            With .Cells(1).Range
                .Text = fText
                Set rg = .Characters.Last.Previous
                If devMode Then rg.Select
            End With
            Documents(selDoc).Fields.Add rg, wdFieldSequence, seqID & "\* Arabic\n"
            .Cells(2).Range.Text = _
            Format(strDATE, "dd.mm.yyyy")
            .Cells(3).Range.Text = _
            Format(strDATE, "dd.mm.yyyy")
            .Range.Copy
            For p = 2 To pCount
                .Range.Paste
                If devMode Then .Select
                DoEvents
            Next
        End With
    End With
Else
    On Error GoTo -1
End If

End With

nextS:
On Error GoTo 0
i = i + 1
DoEvents
Next

UpdateCaption ("Recursively adding new LEP pages to LEP...")

Set rg = Documents(selDoc).Sections(6).Range.Tables(1).Range
With rg
.Rows.HeadingFormat = False
.InsertBreak wdSectionBreakContinuous
End With
With Documents(selDoc).Sections(7).Range.Tables(1)
    Set rg = .Range
    With rg.PageSetup.TextColumns
        .SetCount NumColumns:=2
        .EvenlySpaced = True
        .LineBetween = False
    End With
    .Rows(1).HeadingFormat = True
    pCount = .Range.Sections(1).Range.ComputeStatistics(wdStatisticPages) - 2
    .Rows(2).Delete
    With rg
        .Collapse wdCollapseStart
        Do Until .Rows.Last.Range.Text Like "LIST OF EFFECTIVE PAGES*"
            .MoveEnd wdRow, 1
        Loop
        .Collapse wdCollapseEnd
        .MoveEnd wdRow, 1
        .Copy
        oldpCount = 2
recount:
        For p = oldpCount To pCount
            .Paste
            .Collapse wdCollapseEnd
            If devMode Then .Select
            DoEvents
        Next
        oldpCount = pCount
        pCount = .Sections(1).Range.ComputeStatistics(wdStatisticPages) - 2
        If pCount > oldpCount Then GoTo recount
    End With
End With

i = 0
UpdateCaption ("Formatting LEP...")
rCount = Documents(selDoc).Sections(7).Range.Tables(1).Rows.Count
UpdateProgress (i / rCount)

For Each r In Documents(selDoc).Sections(7).Range.Tables(1).Rows
    With r
        If .Index = 1 Then .Borders(wdBorderBottom).color = wdColorWhite
        If .Cells.Count = 1 Then
            .Range.Style = "Table Header"
            .Shading.BackgroundPatternColor = ASTBlue
        End If
    End With
i = i + 1
DoEvents
Next

UpdateCaption ("Updating fields...")
DoEvents
Documents(selDoc).Sections(7).Range.Tables(1).Range.Fields.Update

timeLEP = MinSec(Timer - timeSubInit)
stsUpdate "stsLEP", False

End Sub
