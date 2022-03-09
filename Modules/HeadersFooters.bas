Attribute VB_Name = "HeadersFooters"
Option Explicit

Sub Clean()

stsUpdate "stsHFRemoval", True
timeSubInit = Timer

Dim s As Section, sCount As Long, i As Long
sCount = Documents(selDoc).Sections.Count
i = 0

For Each s In Documents(selDoc).Sections

UpdateProgress (i / sCount)
UpdateCaption (i & " of " & sCount & " sections cleaned.")
i = i + 1

    If devMode Then s.Headers(wdHeaderFooterPrimary).Range.Select

    s.Headers(wdHeaderFooterPrimary).Range.Delete

    If devMode Then s.Footers(wdHeaderFooterPrimary).Range.Select

    s.Footers(wdHeaderFooterPrimary).Range.Delete

Next

timeHeadersFooters = MinSec(Timer - timeSubInit)
stsUpdate "stsHFRemoval", False

End Sub

Sub Setup()

stsUpdate "stsHF", True
timeSubInit = Timer

Dim s As Section, sCount As Long, i As Long, rg As Range, _
sContinue As Boolean
Set rg = ThisDocument.Paragraphs.Last.Range: i = 0
sCount = Documents(selDoc).Sections.Count: ASTBlue = 12611584

For Each s In Documents(selDoc).Sections

With s

    sContinue = IIf(Not .Range.Paragraphs(1).Style Like "Heading 1*", True, False)

    If .Index = 1 Then
        With .Borders(wdBorderLeft)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth225pt
            .color = ASTBlue
        End With
        With .Borders(wdBorderRight)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth225pt
            .color = ASTBlue
        End With
        With .Borders(wdBorderTop)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth225pt
            .color = ASTBlue
        End With
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth225pt
            .color = ASTBlue
        End With
        With .Borders
            .DistanceFrom = wdBorderDistanceFromPageEdge
            .AlwaysInFront = True
            .SurroundHeader = True
            .SurroundFooter = True
            .JoinBorders = False
            .DistanceFromTop = 24
            .DistanceFromLeft = 24
            .DistanceFromBottom = 24
            .DistanceFromRight = 24
            .Shadow = False
            .EnableFirstPageInSection = True
            .EnableOtherPagesInSection = True
        End With
    Else
        With .PageSetup
            If .Orientation = wdOrientPortrait Then
                .FooterDistance = CentimetersToPoints(0.6)
            Else
                .FooterDistance = CentimetersToPoints(0.4)
            End If
        End With

        Select Case .PageSetup.Orientation

        Case wdOrientPortrait
            ThisDocument.Tables(2).Range.Copy
            With .Headers(wdHeaderFooterPrimary)
                .LinkToPrevious = False
                With .Range
                    .Paste
                    With .Tables(1).Range
                        .Cells(2).Range.Text = _
                        strSUBT
                        .Cells(3).Range.Text = _
                        "REDACTED-" & strDEP & "-" & strREF
                    End With
                End With
            End With
            With ThisDocument
                rg.Start = .Paragraphs(42).Range.Start
                rg.End = .Paragraphs(43).Range.End
            End With
            rg.Copy
            With .Footers(wdHeaderFooterPrimary)
                .LinkToPrevious = False
                .Range.Paste
                With .PageNumbers
                    .NumberStyle = wdPageNumberStyleArabic
                    .IncludeChapterNumber = True
                    .HeadingLevelForChapter = 0
                    .ChapterPageSeparator = wdSeparatorHyphen
                    .RestartNumberingAtSection = IIf(sContinue, False, True)
                    .StartingNumber = 1
                End With
            End With

        Case wdOrientLandscape
            ThisDocument.Tables(3).Range.Copy
            With .Headers(wdHeaderFooterPrimary)
                .LinkToPrevious = False
                With .Range
                    .Paste
                    With .Tables(1).Range
                        .Cells(2).Range.Text = _
                        strSUBT
                        .Cells(3).Range.Text = _
                        "REDACTED-" & strDEP & "-" & strREF
                    End With
                End With
            End With
            With rg
                .Start = ThisDocument.Paragraphs(59).Range.Start
                .End = ThisDocument.Paragraphs(60).Range.End
                .Copy
            End With
            With .Footers(wdHeaderFooterPrimary)
                .LinkToPrevious = False
                .Range.Paste
                With .PageNumbers
                    .NumberStyle = wdPageNumberStyleArabic
                    .IncludeChapterNumber = True
                    .HeadingLevelForChapter = 0
                    .ChapterPageSeparator = wdSeparatorHyphen
                    .RestartNumberingAtSection = IIf(sContinue, False, True)
                    .StartingNumber = 1
                End With
            End With

        End Select

    End If

End With

UpdateProgress (i / sCount)
UpdateCaption ("Section " & i & " of " & sCount & " set up.")
i = i + 1

Next

timeHeadersSetup = MinSec(Timer - timeSubInit)
stsUpdate "stsHF", False

End Sub
