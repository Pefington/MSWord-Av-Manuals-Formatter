Attribute VB_Name = "Sections"
Option Explicit

Sub Clean()

stsUpdate "stsSecRemoval", True
timeSubInit = Timer

Dim rg As Range, sCount As Long, i As Long, j As Long
sCount = Documents(selDoc).Sections.Count + 1
Set rg = Documents(selDoc).Range
i = 0: j = 0

    With rg.Find

        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^b"
        .Wrap = wdFindStop
        .Forward = True
        .MatchWildcards = False

        Do While .Execute

            If devMode Then rg.Select

            rg.InsertBefore (vbCr)
            rg.MoveStart (wdCharacter), 1
            If rg.Next(wdSection, 1).PageSetup.Orientation = _
            rg.Sections(1).PageSetup.Orientation Then rg.Delete: i = i + 1
            rg.Collapse (wdCollapseEnd)

UpdateProgress (j / sCount)
UpdateCaption (i & " of " & sCount & _
" sections breaks removed (preserving landscape sections).")
j = j + 1

        Loop

        .Text = "^l"
        .Wrap = wdFindContinue

        Do While .Execute

            If devMode Then rg.Select

            rg.InsertBefore (vbCr)
            rg.MoveStart (wdCharacter), 1
            rg.Delete
            rg.Collapse (wdCollapseEnd)

UpdateCaption ("Removing manual line breaks.")

        Loop

        .Text = "^n"

        Do While .Execute

            If devMode Then rg.Select

            rg.InsertBefore (vbCr)
            rg.MoveStart (wdCharacter), 1
            rg.Delete
            rg.Collapse (wdCollapseEnd)

UpdateCaption ("Removing column breaks.")

        Loop

        .Text = "^m"

        Do While .Execute

            If devMode Then rg.Select

            timeTotal = MinSec(Timer - timeInit)
            rg.InsertBefore (vbCr)
            rg.MoveStart (wdCharacter), 1
            rg.Delete
            rg.Collapse (wdCollapseEnd)

UpdateCaption ("Removing manual page breaks.")

        Loop

    End With

timeSectionsRemoval = MinSec(Timer - timeSubInit)
stsUpdate "stsSecRemoval", False

End Sub
