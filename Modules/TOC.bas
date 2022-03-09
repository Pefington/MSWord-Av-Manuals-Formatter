Attribute VB_Name = "TOC"
Option Explicit

Sub Clean()

stsUpdate "stsManTOC", True
timeSubInit = Timer

Dim s As Section, sCount As Long, rg As Range, i As Long, d As Long, _
oCount As Long, p As Long, pCount As Long
sCount = Documents(selDoc).Sections.Count
i = 0
d = 0

For Each s In Documents(selDoc).Sections

oCount = 0

UpdateProgress (i / sCount)
UpdateCaption (i & " of " & sCount & " sections analysed, " & d & " manual TOC sections removed.")


    With s.Range
        pCount = .Paragraphs.Count - 1
        Select Case pCount
        Case Is <= 10
            For p = 1 To pCount
                With .Paragraphs(p).Range
                    If devMode Then .Select
                    If .Style Like "Heading #" Or .Words.First = vbCr _
                    Then oCount = oCount + 1
                End With
            Next
            If oCount = pCount Then .Delete: d = d + 1
        Case Is > 10
            For p = 1 To 10
                With .Paragraphs(p).Range
                    If devMode Then .Select
                    If .Style Like "Heading #" Or .Words.First = vbCr _
                    Then oCount = oCount + 1
                End With
            Next
            If oCount = 10 Then .Delete: d = d + 1
        End Select
    End With

i = i + 1
Next

timeCleanTOC = MinSec(Timer - timeSubInit)
stsUpdate "stsManTOC", False

End Sub

Sub Generate()

stsUpdate "stsTOC", True
timeSubInit = Timer

Dim s As Section, sCount As Long, rg As Range, chpName As String, pCount As Long, _
nTOC As Long, i As Long, hasFWD As Boolean, hasROR As Boolean, hasHOR As Boolean, _
hasDL As Boolean, hasLEP As Boolean, sTitle As String, effDate As String
effDate = Format(strDATE, "dd.mm.yyyy")
sCount = Documents(selDoc).Sections.Count
nTOC = 1
i = 0

For Each s In Documents(selDoc).Sections

If i > sCount Then i = sCount
chpName = vbNullString

UpdateProgress (i / sCount)
UpdateCaption (i & " of " & sCount & " sections processed.")

With s
    If .Index = 1 Then GoTo nextS
    If .PageSetup.SectionStart = wdSectionContinuous Then
        .Headers(wdHeaderFooterPrimary).LinkToPrevious = True
        With .Footers(wdHeaderFooterPrimary)
            .LinkToPrevious = True
            .PageNumbers.RestartNumberingAtSection = False
        End With
        GoTo nextS
    End If
    With .Range
        With .Paragraphs(1).Range
            If devMode Then .Select
            sTitle = LCase(.Text)
        End With
        If sTitle Like "foreword*" Or sTitle Like "*preamble*" Then
            .Paragraphs(1).Style = "Heading 1 FWD"
            With s.Headers(wdHeaderFooterPrimary).Range.Tables(1).Range
                .Cells(4).Range.Text = "FWD"
                With .Cells(5).Range.Fields(1).Code
                    .Text = Replace(.Text, "HEADING 1", "HEADING 1 FWD")
                End With
            End With
            GoTo nextS
            hasFWD = True
        ElseIf sTitle Like "record*revision*" _
        Or sTitle Like "revision*record*" _
        Or sTitle Like "history*revision*" _
        Or sTitle Like "revision*history*" _
        Or sTitle Like "distribution*list*" _
        Or sTitle Like "list*distribution*" _
        Or sTitle Like "list*effective*page*" _
        Or sTitle Like "effective*page*list*" Then
            .Delete
            GoTo nextS
        ElseIf sTitle Like "appendix*" Then
            .Paragraphs(1).Style = "Heading 1 FWD"
            With s.Footers(wdHeaderFooterPrimary)
                .PageNumbers.IncludeChapterNumber = False
                Set rg = .Range.Paragraphs(2).Range
                With rg
                    .Collapse wdCollapseStart
                    .Text = "APP-"
                End With
            End With
            With s.Headers(wdHeaderFooterPrimary).Range.Tables(1).Range
                .Cells(4).Range.Text = "APP"
                With .Cells(5).Range.Fields(1).Code
                    .Text = Replace(.Text, "HEADING 1", "HEADING 1 FWD")
                End With
            End With
            GoTo nextS
        ElseIf sTitle Like "*definition*" Then
            .Paragraphs(1).Style = "Heading 1 FWD"
            With s.Footers(wdHeaderFooterPrimary)
                .PageNumbers.IncludeChapterNumber = False
                Set rg = .Range.Paragraphs(2).Range
                With rg
                    .Collapse wdCollapseStart
                    .Text = "DEF-"
                End With
            End With
            With s.Headers(wdHeaderFooterPrimary).Range.Tables(1).Range
                .Cells(4).Range.Text = "DEF"
                With .Cells(5).Range.Fields(1).Code
                    .Text = Replace(.Text, "HEADING 1", "HEADING 1 FWD")
                End With
            End With
            GoTo nextS
        ElseIf sTitle Like "*abbreviation*" Then
            .Paragraphs(1).Style = "Heading 1 FWD"
            With s.Footers(wdHeaderFooterPrimary)
                .PageNumbers.IncludeChapterNumber = False
                Set rg = .Range.Paragraphs(2).Range
                With rg
                    .Collapse wdCollapseStart
                    .Text = "ABB-"
                End With
            End With
            With s.Headers(wdHeaderFooterPrimary).Range.Tables(1).Range
                .Cells(4).Range.Text = "ABB"
                With .Cells(5).Range.Fields(1).Code
                    .Text = Replace(.Text, "HEADING 1", "HEADING 1 FWD")
                End With
            End With
            GoTo nextS
        Else
            pCount = .Paragraphs.Count
            With .Paragraphs(1).Range
                chpName = Replace(.ListFormat.ListString, " ", vbNullString)
                If chpName = vbNullString _
                Or Not chpName Like "CHAPTER#*" Then GoTo nextS
                .InsertAfter vbCr
            End With
            .Paragraphs(2).Style = "Body Text"
            Set rg = .Paragraphs(3).Range
            With rg
                .Collapse wdCollapseStart
                If devMode Then .Select
                .InsertBreak wdPageBreak
                .Start = s.Range.Paragraphs(4).Range.Start
                .End = s.Range.Paragraphs(pCount - 1).Range.End
                rg.bookmarks.Add chpName, rg
            End With
            .Paragraphs(2).Range.Fields.Add .Paragraphs(2).Range, wdFieldTOC, _
            "\b " & chpName & " \h \o " & Chr(34) & "1-3" & Chr(34), False
            nTOC = nTOC + 1
        End If

    End With
End With
nextS:
i = i + 1
quickNextS:
Next

With Documents(selDoc)

    If Not hasFWD Then
        With .Sections(2)
            If .Range.Paragraphs(1).Style = "Body Text" Then hasFWD = True
            UpdateCaption ("Generating FWD...")
            DoEvents
            If hasFWD Then
                Set rg = ThisDocument.Sections(4).Range
                With rg
                    .MoveEnd wdParagraph, -5
                    If devMode Then .Select
                    .Copy
                End With
            Else
                ThisDocument.Sections(4).Range.Copy
            End If
            Set rg = .Range
            With rg
                .Collapse wdCollapseStart
                If devMode Then .Select
                .Paste
            End With
        End With
        With .Sections(2)
            With .Headers(wdHeaderFooterPrimary).Range.Tables(1).Range
                .Cells(2).Range.Text = strSUBT
                .Cells(3).Range.Text = "REDACTED-" & strDEP & "-" & strREF
                .Cells(4).Range.Text = "FWD"
                With .Cells(5).Range.Fields(1).Code
                    .Text = Replace(.Text, "HEADING 1", "HEADING 1 FWD")
                End With
            End With
            With .Footers(wdHeaderFooterPrimary)
                .PageNumbers.IncludeChapterNumber = False
                Set rg = .Range.Paragraphs(2).Range
                With rg
                    .Collapse wdCollapseStart
                    .Text = "FWD-"
                End With
            End With
        End With
    End If

    With .Sections(3)
        UpdateCaption ("Generating ROR...")
        DoEvents
        Set rg = .Range
        With rg
            .Collapse wdCollapseStart
            If devMode Then .Select
            ThisDocument.Sections(5).Range.Copy
            .Paste
            With .Tables(1)
                .cell(2, 1).Range.Text = "Rev. " & strREV
                .cell(2, 2).Range.Text = Format(strDATE, "dd.mm.yyyy")
            End With
        End With
    End With
    With .Sections(3)
        With .Headers(wdHeaderFooterPrimary)
            .LinkToPrevious = False
            With .Range.Tables(1).Range
                .Cells(2).Range.Text = _
                strSUBT
                .Cells(3).Range.Text = _
                "REDACTED-" & strDEP & "-" & strREF
            End With
        End With
        With .Footers(wdHeaderFooterPrimary)
            .LinkToPrevious = False
            .PageNumbers.IncludeChapterNumber = False
        End With
    End With

    With .Sections(4)
        UpdateCaption ("Generating HOR...")
        DoEvents
        Set rg = .Range
        With rg
            .Collapse wdCollapseStart
            If devMode Then .Select
            ThisDocument.Sections(6).Range.Copy
            .Paste
            With .Tables(1)
                .cell(2, 1).Range.Text = "Rev. " & strREV
                .cell(2, 2).Range.Text = Format(strDATE, "dd.mm.yyyy")
            End With
        End With
    End With
    With .Sections(4)
        With .Headers(wdHeaderFooterPrimary)
            .LinkToPrevious = False
            With .Range.Tables(1).Range
                .Cells(2).Range.Text = _
                strSUBT
                .Cells(3).Range.Text = _
                "REDACTED-" & strDEP & "-" & strREF
            End With
        End With
        With .Footers(wdHeaderFooterPrimary)
            .LinkToPrevious = False
            .PageNumbers.IncludeChapterNumber = False
        End With
    End With

    With .Sections(5)
        UpdateCaption ("Generating DL...")
        Set rg = .Range
        With rg
            .Collapse wdCollapseStart
            If devMode Then .Select
            ThisDocument.Sections(7).Range.Copy
            .Paste
        End With
    End With
    With .Sections(5)
        With .Headers(wdHeaderFooterPrimary)
            .LinkToPrevious = False
            With .Range.Tables(1).Range
                .Cells(2).Range.Text = _
                strSUBT
                .Cells(3).Range.Text = _
                "REDACTED-" & strDEP & "-" & strREF
            End With
        End With
        With .Footers(wdHeaderFooterPrimary)
            .LinkToPrevious = False
            .PageNumbers.IncludeChapterNumber = False
        End With
    End With

    With .Sections(6)
        UpdateCaption ("Generating LEP...")
        Set rg = .Range
        With rg
            .Collapse wdCollapseStart
            If devMode Then .Select
            ThisDocument.Sections(8).Range.Copy
            .Paste
        End With
    End With
    With .Sections(6)
        With .Headers(wdHeaderFooterPrimary)
            .LinkToPrevious = False
            With .Range.Tables(1).Range
                .Cells(2).Range.Text = _
                strSUBT
                .Cells(3).Range.Text = _
                "REDACTED-" & strDEP & "-" & strREF
            End With
        End With
        With .Footers(wdHeaderFooterPrimary)
            .LinkToPrevious = False
            .PageNumbers.IncludeChapterNumber = False
        End With
    End With

End With

timeTOC = MinSec(Timer - timeSubInit)
stsUpdate "stsTOC", False

End Sub
