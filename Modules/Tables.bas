Attribute VB_Name = "Tables"
Option Explicit

Sub Formatting()

stsUpdate "stsTables", True
timeSubInit = Timer

ASTBlue = 12611584

Dim t As Table, tCount As Long, tRows As Long, tCols As Long, iRow As Long, iCol As Long, _
cCount As Long, i As Long, HHead As Boolean, VHead As Boolean, rgH As Range, rgV As Range, _
selRow As Row, selCol As Column, c As cell, tableUniform As Boolean, rgHTmp As Range, _
rgVTmp As Range, tHeight As Long, tWidth As Single, pStart As Long, pEnd As Long, counter As Long
tCount = Documents(selDoc).Tables.Count: i = 0

For Each t In Documents(selDoc).Tables

UpdateProgress (i / tCount)
UpdateCaption (i & " of " & tCount & " tables formatted.")

    With t

        If devMode Then .Select: Selection.Collapse (wdCollapseStart)
        If .Borders.OutsideColor = wdColorRed _
         And .Borders.OutsideLineWidth = wdLineWidth300pt Then GoTo nextT

         .Borders.OutsideColor = ASTBlue

       tRows = .Rows.Count
       tCols = .Columns.Count
       cCount = .range.Cells.Count

       Set rgH = .cell(1, 1).range
       Set rgHTmp = rgH.Duplicate
       Set rgV = rgH.Duplicate
       Set rgVTmp = rgH.Duplicate
       rgH.Collapse wdCollapseStart
       rgV.Collapse wdCollapseStart
       HHead = False
       VHead = False
       iRow = 1
       iCol = 1
       tHeight = 0
       tWidth = 0

       .range.Font.Size = 9
       .AllowAutoFit = True
       .Spacing = 0
       .AllowPageBreaks = True
       .AutoFitBehavior (wdAutoFitContent)
       .TopPadding = MillimetersToPoints(1)
       .BottomPadding = MillimetersToPoints(1)
       .LeftPadding = MillimetersToPoints(2)
       .RightPadding = MillimetersToPoints(2)
       With .range
           .ParagraphFormat.Alignment = wdAlignParagraphCenter
           .Cells.VerticalAlignment = wdCellAlignVerticalCenter
       End With
       With .Rows
           .HeightRule = wdRowHeightAuto
           .Height = CentimetersToPoints(0)
           .LeftIndent = CentimetersToPoints(0)
       End With
       With .Columns
           .PreferredWidthType = wdPreferredWidthAuto
           .PreferredWidth = 0
       End With
       With .Borders
           On Error Resume Next
           .InsideLineStyle = wdLineStyleSingle
           .InsideLineWidth = wdLineWidth050pt
           .InsideColor = ASTBlue
           On Error GoTo 0
       End With
       .Spacing = 0
       With .Rows
           .Alignment = wdAlignRowCenter
           .LeftIndent = CentimetersToPoints(0)
       End With
       For Each c In .range.Cells
           With c.range
           If .Characters.Count = 1 Then .Bold = -1
           If devMode Then .Select
           End With
       Next

       If cCount = tRows * tCols Then
           Do Until .Rows(iRow).range.Bold <> -1
               With .Rows(iRow).range
                   rgH.End = .End
                   If devMode Then .Select
                   iRow = iRow + 1
                   DoEvents
               End With
               If iRow > tRows Then Exit Do
               DoEvents
           Loop
           If devMode Then rgH.Select
           .Columns(1).Select
           Do Until Selection.range.Bold <> -1
               .Columns(iCol).Select
               rgV.End = Selection.End
               iCol = iCol + 1
               If iCol > tCols Then Exit Do
               DoEvents
           Loop
           If devMode Then rgV.Select

       Else

           On Error Resume Next
           Do Until rgHTmp.Bold <> -1
               If iRow = 1 Then
                   rgH.End = .cell(iRow, iCol).range.End
                   If devMode Then rgH.Select
               Else
                   rgHTmp.End = .cell(iRow, iCol).range.End
                   If devMode Then rgHTmp.Select
               End If
               If iCol = tCols Then
                   iRow = iRow + 1
                   iCol = 1
                   If iRow > 2 And rgHTmp.Columns.Count = rgH.Columns.Count _
                   Then rgH.End = rgHTmp.Duplicate.End
                   If devMode Then rgH.Select: rgHTmp.Select
                   Set rgHTmp = .cell(iRow, iCol).range
               Else
                   iCol = iCol + 1
               End If
               If t.range.Bold = -1 Then Exit Do
               DoEvents
               If iCol >= tCols And iRow >= tRows Then Exit Do
           Loop
           On Error GoTo 0
           If devMode Then rgH.Select
           iRow = 1
           iCol = 1
           On Error Resume Next
           Do Until .cell(iRow, iCol).range.Bold <> -1
               If iCol = 1 Then
                   rgV.End = .cell(iRow, iCol).range.End
                   If devMode Then rgV.Select
               Else
                   rgVTmp.End = .cell(iRow, iCol).range.End
                   If devMode Then rgVTmp.Select
               End If
               If iRow = tRows Then
                   iCol = iCol + 1
                   iRow = 1
                   If iCol > 2 And rgVTmp.Rows.Count = rgV.Rows.Count _
                   Then rgV.End = rgVTmp.Duplicate.End
                   If devMode Then rgV.Select: rgVTmp.Select
                   Set rgVTmp = .cell(iRow, iCol).range
               Else
                   iRow = iRow + 1
               End If
               DoEvents
               If iCol >= tCols And iRow >= tRows Then Exit Do
           Loop
           On Error GoTo 0
           If devMode Then rgV.Select
       End If

       Select Case rgH.Cells.Count
           Case 0
               HHead = False
           Case 1
               On Error Resume Next
               t.cell(rgH.Cells(1).RowIndex, _
               rgH.Cells(1).ColumnIndex + 1).Select
               If Err.Number = 5941 And rgH.Cells(1).ColumnIndex _
               = tCols - 1 Then HHead = True
               On Error GoTo 0
           Case Is > 1
               HHead = True
       End Select
       Select Case rgV.Cells.Count
           Case 0
               VHead = False
           Case 1
               On Error Resume Next
               Err.Clear
               t.cell(rgH.Cells(1).RowIndex, _
               rgH.Cells(1).ColumnIndex + 1).Select
               If Err.Number = 5941 And rgV.Cells(1).RowIndex _
               = tRows - 1 Then VHead = True
               On Error GoTo 0
           Case Is > 1
               VHead = True
       End Select
       .range.Style = "Table Content"

       If HHead Then
           With rgH
               .Rows.HeadingFormat = True
               .Style = "Table Header"
               For Each c In .Cells
                   c.Borders.OutsideColor = wdColorWhite
                   With c.Shading
                       .Texture = wdTextureNone
                       Select Case .BackgroundPatternColorIndex
                       Case 0, 8, 15, 16
                           .BackgroundPatternColor = ASTBlue
                       Case Else
                           If t.Shading.BackgroundPatternColor <> -1 _
                           Then tableUniform = True
                           If tableUniform = True Then
                               .BackgroundPatternColor = ASTBlue
                           ElseIf .BackgroundPatternColor <> ASTBlue Then
                               c.range.Font.ColorIndex = wdAuto
                           End If
                       End Select
                   End With
               Next
           End With
       End If

       If VHead Then
           rgV.Select
           With Selection
               .Style = "Table Header"
               For Each c In .Cells
                   c.Borders.OutsideColor = wdColorWhite
                   With c.Shading
                       .Texture = wdTextureNone
                       Select Case .BackgroundPatternColorIndex
                       Case 0, 8, 15, 16
                           .BackgroundPatternColor = ASTBlue
                       Case Else
                           If t.Shading.BackgroundPatternColor <> -1 _
                           Then tableUniform = True
                           If tableUniform = True Then
                               .BackgroundPatternColor = ASTBlue
                           ElseIf .BackgroundPatternColor <> ASTBlue Then
                               c.range.Font.ColorIndex = wdAuto
                           End If
                       End Select
                   End With
               Next
           End With
       End If

       With .Borders
           .OutsideLineStyle = wdLineStyleSingle
           .OutsideLineWidth = wdLineWidth050pt
           .OutsideColor = ASTBlue
       End With

       For Each c In .range.Cells
           With c
           If .RowIndex = tRows Then tWidth = _
           tWidth + PointsToCentimeters(.Width)
           If devMode Then .Select
           End With
       Next
       Set rgH = .range
       With rgH
           .Collapse wdCollapseStart
           pStart = .Information(wdActiveEndPageNumber)
           .End = t.range.End
           .Collapse wdCollapseEnd
           pEnd = .Information(wdActiveEndPageNumber)
       End With
       tHeight = pEnd - pStart
       If tHeight >= 2 Then
           Select Case tWidth

           Case Is < 3
               With rgH
                   .Collapse wdCollapseStart
                   .InsertBreak wdSectionBreakContinuous
                   .End = t.range.End
                   .Collapse wdCollapseEnd
                   .InsertBreak wdSectionBreakContinuous
               End With
               t.Rows.HeadingFormat = False
               With t.range.Sections(1).PageSetup.TextColumns
                   .SetCount NumColumns:=3
                   .EvenlySpaced = True
                   .LineBetween = False
               End With

           Case 3 To 6
               With rgH
                   .Collapse wdCollapseStart
                   .InsertBreak wdSectionBreakContinuous
                   .End = t.range.End
                   .Collapse wdCollapseEnd
                   .InsertBreak wdSectionBreakContinuous
               End With
               t.Rows.HeadingFormat = False
               With t.range.Sections(1).PageSetup.TextColumns
                   .SetCount NumColumns:=2
                   .EvenlySpaced = True
                   .LineBetween = False
               End With

           End Select
       End If
    End With
nextT:
   HHead = False
   VHead = False
   tableUniform = False
   Selection.Collapse

i = i + 1
Next

timeTables = MinSec(Timer - timeSubInit)
stsUpdate "stsTables", False

End Sub
