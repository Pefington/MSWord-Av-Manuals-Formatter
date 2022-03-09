Attribute VB_Name = "PageSetup"
Option Explicit

Sub LayoutSetup()

stsUpdate "stsLayout", True
timeSubInit = Timer

Dim ops As Long, i As Long, s As Section
ops = 13
i = 0

With Documents(selDoc).Range.PageSetup

    .TextColumns.SetCount (1)

UpdateProgress (i / ops)
UpdateCaption (ops & " parameters to set. Setting footer distance: 3mm.")
i = i + 1
    .HeaderDistance = CentimetersToPoints(0.3)

UpdateProgress (i / ops)
UpdateCaption (i & " of " & ops & " parameters set. Setting top margin: 1cm.")
i = i + 1
    .TopMargin = CentimetersToPoints(1)

UpdateProgress (i / ops)
UpdateCaption (i & " of " & ops & " parameters set. Setting bottom margin: 1cm.")
i = i + 1
    .BottomMargin = CentimetersToPoints(1)

UpdateProgress (i / ops)
UpdateCaption (i & " of " & ops & " parameters set. Setting left margin: 1cm.")
i = i + 1
    .LeftMargin = CentimetersToPoints(1)

UpdateProgress (i / ops)
UpdateCaption (i & " of " & ops & " parameters set. Setting right margin: 1cm.")
i = i + 1
    .RightMargin = CentimetersToPoints(1)

UpdateProgress (i / ops)
UpdateCaption (i & " of " & ops & " parameters set. Setting gutter: none.")
i = i + 1
    .Gutter = CentimetersToPoints(0)

UpdateProgress (i / ops)
UpdateCaption (i & " of " & ops & " parameters set. Setting NO odd-even pages headers.")
i = i + 1
    .OddAndEvenPagesHeaderFooter = False

UpdateProgress (i / ops)
UpdateCaption (i & " of " & ops & " parameters set. Setting NO different 1st page header.")
i = i + 1
    .DifferentFirstPageHeaderFooter = False

UpdateProgress (i / ops)
UpdateCaption (i & " of " & ops & " parameters set. Setting no mirror margins.")
i = i + 1
    .MirrorMargins = False

UpdateProgress (i / ops)
UpdateCaption (i & " of " & ops & " parameters set. Setting one page per sheet.")
    i = i + 1
    .TwoPagesOnOne = False

UpdateProgress (i / ops)
UpdateCaption (i & " of " & ops & " parameters set. Setting no booklet printing.")
    i = i + 1
    .BookFoldPrinting = False

UpdateProgress (i / ops)
UpdateCaption (i & " of " & ops & " parameters set.")

End With

For Each s In Documents(selDoc).Sections
    With s.PageSetup
        UpdateCaption ("Setting paper: A5, looking for landscape sections (" & i & " of scount checked).")
        If .Orientation = wdOrientPortrait Then
            .PaperSize = wdPaperA5
            .Orientation = wdOrientPortrait
        Else
            .PaperSize = wdPaperA5
            .Orientation = wdOrientLandscape
        End If
    End With
Next

timePageSetup = MinSec(Timer - timeSubInit)
stsUpdate "stsLayout", False

End Sub

Sub CoverPages()

'Call cfgdev

stsUpdate "stsCover", True
Dim r As Range, dateOrd As String, strDAY As String, strMONTH As String, strYEAR As String

strDAY = Day(strDATE)
strMONTH = Format(Month(strDATE), "mmmm")
strYEAR = Year(strDATE)

Select Case strDAY
    Case "1", "21", "31"
        dateOrd = "st"
    Case "2", "22"
        dateOrd = "nd"
    Case "3", "23"
        dateOrd = "rd"
    Case Else
        dateOrd = "th"
End Select

ThisDocument.Sections(1).Range.Copy

With Documents(selDoc)

    .Range(0, 0).Paste

    With .Tables(1)
        .Rows(1).Range.Text = "REDACTED-" & strDEP & "-" & strREF
        .Rows(1).Range.Style = "COVER Ref"
        .Rows(2).Range.Style = "COVER Logo"
        .Rows(3).Range.Text = strTITLE & vbCr & strSUBT
        .Rows(3).Range.Paragraphs(1).Style = "Cover Title"
        .Rows(3).Range.Paragraphs(2).Style = "Cover Subtitle"
        .Rows(6).Range.Text = "Issue-" & strISSUE & ", " & strMONTH & " " & strYEAR
        .Rows(6).Range.Style = "COVER Date"
        .Rows(7).Range.Text = "REDACTED-" & strDEP & "-" & strREF
        .Rows(7).Range.Style = "COVER Ref"
        .Rows(8).Range.Style = "COVER Logo"
        .Rows(9).Range.Text = strTITLE & vbCr & strSUBT
        .Rows(9).Range.Paragraphs(1).Style = "COVER Title"
        .Rows(9).Range.Paragraphs(2).Style = "COVER Subtitle"
        .Rows(10).Range.Style = "COVER Publisher"
        .Rows(11).Range.Text = strAUTH
        .Rows(11).Range.Style = "COVER Subtitle"
        .Rows(12).Range.Text = "Issue-" & strISSUE & ", Rev." & strREV _
        & ", " & strDAY & dateOrd & " " & strMONTH & " " & strYEAR
        .Rows(12).Range.Style = "COVER Date"
    End With

End With

stsUpdate "stsCover", False

End Sub
