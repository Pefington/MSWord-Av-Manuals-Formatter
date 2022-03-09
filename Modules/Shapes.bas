Attribute VB_Name = "Shapes"
Option Explicit

Sub CleanFloating()

'Call cfgdev

stsUpdate "stsShapes", True
timeSubInit = Timer

Dim sCount As Long, t As Long, i As Long, rg As Range, rg2 As Range, rgx As Range
Set rg2 = Documents(selDoc).Paragraphs(1).Range
sCount = Documents(selDoc).Shapes.Count

For i = sCount To 1 Step -1

    With Documents(selDoc).Shapes(i)

UpdateProgress ((sCount - i) / sCount)
UpdateCaption (sCount - i & " of " & sCount & " floating shapes processed.")

        If devMode Then .Anchor.Paragraphs(1).Range.Select: .Select: Debug.Print .Type

        If .Width <= Application.MillimetersToPoints(1) Or _
            .Height <= Application.MillimetersToPoints(1) Then .Delete: GoTo nextI

        Set rgx = .Anchor.Paragraphs(1).Range

        Select Case .Type

        Case msoTextBox
            If rgx.Start <> rg2.Start Then Set rg = rgx.Duplicate: Set rg2 = rgx.Duplicate
            If .TextFrame.TextRange.Characters.Count < 5 Then .Delete: GoTo nextI
            If .TextFrame.TextRange.Tables.Count > 0 Then
                For t = .TextFrame.TextRange.Tables.Count To 1 Step -1
                    On Error Resume Next
                    .TextFrame.TextRange.Tables(t).Range.Cut
                    On Error GoTo -1
                    On Error GoTo 0
                    With rg
                        .Collapse (wdCollapseEnd)
                        Do Until .Paragraphs(1).Range.Tables.Count = 0 _
                        And .StoryType = wdMainTextStory
                            .MoveEnd (wdParagraph), -1
                            If devMode Then .Paragraphs(1).Range.Select
                            DoEvents
                        Loop
                        .Paste
                    End With
                    .Delete
                    DoEvents
                Next
                GoTo nextI
            End If
            With rg
                .Collapse (wdCollapseEnd)
                If rgx.Tables.Count > 0 Then
                    Do Until .Paragraphs(1).Range.Tables.Count = 0 _
                    And .StoryType = wdMainTextStory
                        .MoveEnd (wdParagraph), -1
                        DoEvents
                    Loop
                    Documents(selDoc).Shapes(i).Anchor.Cut
                    .Paste
                End If
            End With
            If .Width = CentimetersToPoints(1) Or .Height < CentimetersToPoints(1) _
            Then .ShapeStyle = msoLineStylePreset17
            .TextFrame.AutoSize = True
            .WrapFormat.Type = wdWrapInline
            .Anchor.Style = "Image"
            rgx.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
            .LockAspectRatio = msoTrue
            If .Width > CentimetersToPoints(13) Then .Width = CentimetersToPoints(13)

        Case msoPicture
            .ConvertToInlineShape
            rgx.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
            .LockAspectRatio = msoTrue
            If .Width > CentimetersToPoints(13) Then .Width = CentimetersToPoints(13)
            .Anchor.Style = "Image"
            GoTo nextI

        Case msoFreeform
            .ShapeStyle = msoLineStylePreset17: GoTo nextI

        Case msoAutoShape
            .ShapeStyle = msoLineStylePreset17: GoTo nextI

        Case msoGroup
            rgx.ShapeRange(1).WrapFormat.Type = wdWrapInline
            If .Width = CentimetersToPoints(1) Or .Height < CentimetersToPoints(1) _
            Then .ShapeStyle = msoLineStylePreset17
            rgx.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
            .Anchor.Paragraphs(1).Style = "Image"
            .LockAspectRatio = msoTrue
            If .Width > CentimetersToPoints(13) Then .Width = CentimetersToPoints(13)
            If .Height > CentimetersToPoints(14) Then .Height = CentimetersToPoints(14)

        Case Else
            If .Width = CentimetersToPoints(1) Or .Height < CentimetersToPoints(1) _
            Then .ShapeStyle = msoLineStylePreset17
            rgx.InsertBefore (vbCr): rgx.InsertAfter (vbCr)
            .Anchor.Paragraphs(1).Style = "Image"
            rgx.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
            .WrapFormat.Type = wdWrapInline
            .LockAspectRatio = msoTrue
            If .Width > CentimetersToPoints(13) Then .Width = CentimetersToPoints(13)
            If .Height > CentimetersToPoints(14) Then .Height = CentimetersToPoints(14)
        End Select


    End With

nextI:
Next

timeShapes = (Timer - timeSubInit)
frmProgress.stsShapes.ForeColor = wdColorOrange
frmProgress.stsShapes.Font.Underline = False
DoEvents

End Sub

Sub CleanInline()

stsUpdate "stsShapes", True
timeSubInit = Timer

Dim sCount As Long, i As Long, s As InlineShape
sCount = Documents(selDoc).InlineShapes.Count
i = 0

For Each s In Documents(selDoc).InlineShapes

    With s

        If devMode Then .Select

UpdateProgress (i / sCount)
UpdateCaption (i & " of " & sCount & " inline shapes processed.")

        If .Width <= Application.MillimetersToPoints(1) Or _
            .Height <= Application.MillimetersToPoints(1) Then .Delete

        .LockAspectRatio = msoTrue
        If .Width > CentimetersToPoints(13) Then .Width = CentimetersToPoints(13)
        .Range.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
    End With
i = i + 1
Next

timeShapes = timeShapes + (Timer - timeSubInit): timeShapes = MinSec(timeShapes)
stsUpdate "stsShapes", False

End Sub
