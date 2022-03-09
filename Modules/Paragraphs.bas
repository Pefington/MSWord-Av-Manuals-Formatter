Attribute VB_Name = "Paragraphs"
Option Explicit

Dim p As paragraph, pCount As Long, ll(1) As Long, llPrev As Long, inPREAMBLE As Boolean, arrTmp As Variant, _
llInit As Long, inTable As Boolean, ltype(2) As WdListType, i As Long, VB As String, lTemplate As Variant, _
pStyle(1) As String, pLiSt(1) As String, pFont(1 To 3) As Long, bVB As String, lval(1 To 2) As Long, lStrFont As Font, _
pWord(9 To 19) As String, pChar(1 To 19) As String, pT As paragraph, tableError As Boolean, pFontName As String, _
rg As Range, inDEF As Boolean, inABB As Boolean, lStr As String, bodyIdt As Single, tFormatted As Boolean, pNumbered As Boolean, _
Idt(1) As Single, idtInit As Single, idtPrev As Single, TypeOfList As String, rgIdt As Single, pBulleted As Boolean
'for all arrays, double digit system with first digit as follows:(0 or nothing) is previous p
'(1) is current p, (2) is next p.
'for pWord and pChar the second digit is as follows: (1) first, ... , (9) last.

Sub Parse()

Call cfgDev

stsUpdate "stsParagraphs", True
timeSubInit = Timer

pCount = Documents(selDoc).Paragraphs.Count
bodyIdt = Documents(selDoc).Styles("Body Text").ParagraphFormat.LeftIndent
ASTBlue = 12611584
VB = "VBA1"
bVB = "bVBA1"
i = 0

For Each p In Documents(selDoc).Range.Paragraphs

If i > pCount Then i = pCount
UpdateProgress (i / pCount)
UpdateCaption (i & " of " & pCount & " paragraphs processed. ")

With p.Range

    If devMode Then
        .Select
        On Error Resume Next
        ActiveWindow.ScrollIntoView .Next.Paragraphs(1).Range, True
        On Error GoTo 0
        'Debug.Print i & " of " & pCount& & ": " & pStyle(1) & vbNewLine ' <= DEBUGGING ONLY
    End If

'___IMMEDIATE VARIABLES FOR QUICK TRIAGE____________________________________________________________________
    Erase pChar
    Erase pWord
    On Error Resume Next
    ll(1) = .ListFormat.ListLevelNumber
    lval(1) = .ListFormat.ListValue
    ltype(1) = .ListFormat.ListType
    pLiSt(1) = .ListStyle
    pStyle(1) = .Style
    pFontName = .Font.Name
    Idt(1) = p.LeftIndent
    pChar(11) = .Characters.First
    pWord(11) = LCase(Trim(.Words.First))
    lStr = .ListFormat.ListString
    lStrFont = .ListFormat.ListString
    lTemplate = .ListFormat.ListTemplate
    setpFont p.Range
    On Error GoTo 0
    inTable = IIf(.Tables.Count > 0, True, False)
    If Not inTable Then tableError = False: tFormatted = False

'___QUICK TRIAGE____________________________________________________________________________________________
    If .ShapeRange.Count > 0 Then .Style = "Image": GoTo nextP
    If pWord(11) = Chr(12) Then .Style = "Body Text": GoTo nextP
    If pChar(11) Like "[: ]" Or pChar(11) Like vbTab Then .Characters(1).Delete
    If .StoryType = wdTextFrameStory Then GoTo nextP
    On Error Resume Next
    If pWord(11) = vbCr Then .Delete: GoTo nextP
    On Error GoTo 0
    If pStyle(1) Like "TOC #" Then .Delete: GoTo nextP
    If pStyle(1) = VB Then TypeOfList = "Numbered"
    If pStyle(1) = bVB Then TypeOfList = "Bulleted"
    If pStyle(1) Like "*VBA*" Or pStyle(1) Like "Definition*" _
        Or pStyle(1) = "ABB" Then GoTo nextP
    If inTable And Not inABB Then
        If .Cells.Count = 0 Then
            GoTo nextP
        Else
            TableFormat p
            GoTo nextP
        End If
    End If
    If pLiSt(1) = "OM-Numbering" And lStr Like "#.#*" Then GoTo isOutline


'___VARIABLES FOR P-1 AND P+1 (doesn't exist error expected)________________________________________________
    On Error Resume Next
    ll(0) = .Previous.ListFormat.ListLevelNumber
    lval(2) = .Next.ListFormat.ListValue
    ltype(0) = .Previous.ListFormat.ListType
    ltype(2) = .Next.ListFormat.ListType
    pLiSt(0) = .Previous.ListStyle
    pStyle(0) = .Previous.Style
    Idt(0) = .Previous.Paragraphs(1).LeftIndent
    pChar(1) = LCase(.Previous(wdParagraph).Characters.First)
    pChar(12) = LCase(.Characters(2))
    pChar(9) = LCase(.Previous.Previous)
    pChar(19) = LCase(.Characters.Last.Previous)
    pWord(9) = LCase(Trim(.Previous(wdWord, 2)))
    pWord(12) = LCase(Trim(.Words(2)))
    pWord(13) = LCase(Trim(.Words(3)))
    pWord(19) = LCase(Trim(.Words(.Words.Count - 1)))
    On Error GoTo 0

'___INITIAL TRIAGE_________________________________________________________________________________________
    If Not inTable And pChar(12) = vbCr Then .Delete: GoTo nextP
    If LCase(pWord(11)) = "page" And IsNumeric(pWord(12)) And pWord(13) = vbCr _
        Then .Delete: GoTo nextP
    If pChar(11) Like "[a-z]" And pChar(9) <> Chr(32) And .Hyperlinks.Count = 0 _
        Then .Characters.First = UCase(.Characters.First)
    If Not inTable Then
        If lStr Like "*#.#*" Or lStr Like "#" Then GoTo isOutline
        If LCase(lStr) Like "chapter #" Then GoTo isOutline
        If CheckBullet(lStr) Then pBulleted = True: GoTo isBullet
        If checkNumbered(lStr) Then pNumbered = True: GoTo isNumbered
        If checkChapter(p) Then GoTo nextP
    End If

'___________________________________________________________________________________________________________
'___________________________________________________________________________________________________________
'___TRIAGE BY LIST TYPE BEGINS HERE_________________________________________________________________________
    Select Case ltype(1)

'___________________________________________________________________________________________________________
'___NO LIST_________________________________________________________________________________________________
    Case wdListNoNumbering
NoList:
        If .ParagraphFormat.OutlineLevel <> wdOutlineLevelBodyText _
            And lStr <> vbNullString Then GoTo isOutline
        If .Text Like "DEFINITION*" Or .Text Like "ABBREVIATION*" Then GoTo isOutline
        If CheckNote(p) Then GoTo nextP

'_______ListString CHECK FOR MANUAL FORMATTING______________________________________________________________
        If InStr(.Text, vbTab) > 0 And Not inABB Then
            arrTmp = Split(.Text, vbTab)
            arrTmp = arrTmp(LBound(arrTmp))
            If arrTmp Like "#*" And InStr(arrTmp, ".") > 0 Then
                ll(1) = 1 + Len(arrTmp) - Len(Replace(arrTmp, ".", vbNullString))
                GoTo isOutline
            End If
            GoTo isNumbered
        End If

'_______DEFINITIONS (inDEF)_________________________________________________________________________________
inDEF:
        If inDEF Then

'___________DEF WITH COLONS_________________________________________________________________________________
            If pStyle(0) = "Definition Bold" Then .Style = "Definition": GoTo nextP
            If (InStr(.Text, ":") > 0 And _
            (pFont(1) = 1 Or pFont(2) = 1 Or pFont(3) = 1)) _
            Or (pChar(19) = ":" And lStr Like "*)") Then
                If pStyle(0) = "Definition Bold" Then .Previous.Style = "Definition"
'                If pChar(19) = ":" Then
'                    If ltype(2) = wdListBullet Or ltype(2) = wdListSimpleNumbering Then
'                        .Style = "Definition"
'                        TypeOfList = "None"
'                        GoTo nextP
'                    Else
'                        .Style = "Definition Bold"
'                        .Characters.Last.Previous.Delete
'                        .Next.Style = "Definition"
'                        GoTo nextP
'                    End If
'                End If
                Set rg = .Duplicate
                With rg
                    .Collapse wdCollapseStart
                    .MoveEndUntil ":"
                    .InsertAfter vbCr
                    .Text = Replace(.Text, """", vbNullString)
                    .Style = "Definition Bold"
                    Do Until Not .Characters(.Characters.Count - 1) Like "[: ]"
                        .Characters.Last.Delete
                        If devMode Then rg.Select
                        DoEvents
                    Loop
                    .MoveStart wdParagraph, 1
                    .Style = "Definition"
                    TypeOfList = "None"
                End With
                CleanLeadIn rg
                GoTo nextP
'___________DEF WITHOUT COLONS______________________________________________________________________________
            ElseIf pFont(1) = 9 And pWord(11) <> "note" Then
                Set rg = .Duplicate
                With rg
                    .Collapse wdCollapseStart
                    Do Until .Next(wdCharacter, 1).Bold = False
                        .MoveEnd wdCharacter, 1
                        If devMode Then .Select
                        DoEvents
                    Loop
                    .Text = Replace(.Text, """", vbNullString)
                    .InsertAfter vbCr
                    .Style = "Definition Bold"
                    .MoveEnd wdCharacter, -1
                    Do Until Not .Characters.Last Like "[ .vbtab]"
                        .MoveEnd wdCharacter, -1
                        If devMode Then .Select
                        DoEvents
                    Loop
                    .InsertAfter vbCr
                    .Next(wdParagraph).Delete
                End With
                p.Style = "Definition"
                CleanLeadIn p.Range
                TypeOfList = "None"
                GoTo nextP
            End If

'___________DEF OTHER_______________________________________________________________________________________
            GoTo NoListOther
        End If

'_______ABBREVIATIONS (inABB)_______________________________________________________________________________
inABB:
        If inABB Then
'___________ABB IN TABLE____________________________________________________________________________________
            If inTable Then
                If tableError Then GoTo nextP
                On Error GoTo errTable
                .Tables(1).ConvertToText (vbTab)
                On Error GoTo 0
                inTable = False
                tableError = False
            End If
'___________ABB FOLLOW UP___________________________________________________________________________________
            If InStr(.Text, ": ") > 0 And pChar(19) <> ":" _
                Then .Text = Replace(.Text, ": ", vbTab)
            If pFont(1) = 9 And pWord(11) <> "note" And InStr(.Text, vbTab) = 0 Then
                Set rg = .Duplicate
                With rg
                    .Collapse wdCollapseStart
                    Do Until Not .Next(wdCharacter, 1).Bold
                        .MoveEnd wdCharacter, 1
                        If devMode Then .Select
                        DoEvents
                    Loop
                    .Text = Replace(.Text, """", vbNullString)
                    .InsertAfter vbTab
                End With
            End If
            If InStr(.Text, vbTab) = 0 Then
                .Style = "Body Text"
                pFontCheck p
                GoTo nextP
            End If
            .Style = "ABB"
            TypeOfList = "None"
            Set rg = .Duplicate
            With rg
                .Collapse wdCollapseStart
                .MoveEndUntil vbTab
                .Font.Bold = True
            End With

'___________NEW LETTER DELIMITATION_________________________________________________________________________
            If pChar(1) <> LCase(pChar(11)) And pStyle(0) = "ABB" Then
                If .Previous.ParagraphFormat.SpaceBefore = 12 Then
                With .Previous
                    .ParagraphFormat.SpaceAfter = 8
                    With .Borders(wdBorderBottom)
                        .LineStyle = wdLineStyleSingle
                        .LineWidth = wdLineWidth050pt
                        .color = ASTBlue
                    End With
                    .Borders.DistanceFromBottom = 12
                End With
'_______________NECESSARY (BUG) IF SINGLE ABB BEFORE LETTER CHANGE__________________________________________
                Else
                    .ParagraphFormat.SpaceBefore = 12
                    With .Borders(wdBorderTop)
                        .LineStyle = wdLineStyleSingle
                        .LineWidth = wdLineWidth050pt
                        .color = ASTBlue
                    End With
                    .Borders.DistanceFromTop = 6
                End If
            End If
            GoTo nextP
        End If
'_______STANDARD NO LIST____________________________________________________________________________________
NoListOther:
        .ParagraphFormat.Reset
        .Style = "Body Text"
        If inDEF And Not pStyle(0) Like "Heading #" Then
            .Paragraphs(1).LeftIndent = Idt(0)
            GoTo nextP
        End If
        If pChar(9) = ":" Then p.LeftIndent = Idt(0) Else TypeOfList = "None"
        pFontCheck p
        If (pWord(11) = "car" And pWord(12) = "section") _
        Or (pWord(12) = "car" And pWord(13) = "section") Then
        .Style = "Image"
        .Italic = True
        End If
        GoTo nextP
'___________________________________________________________________________________________________________
'___LIST BULLETED___________________________________________________________________________________________
    Case wdListBullet
isBullet:
        If CheckNote(p) Then GoTo nextP
        .Style = bVB

'_______NEW LIST (LIST INIT)________________________________________________________________________________
        Select Case pLiSt(0)
        Case "No List", "REDACTED NUMBERING"
            If TypeOfList = "Numbered" Then
                If llPrev < 9 Then .ListFormat.ListLevelNumber = llPrev + 1
                llPrev = ll(1)
            ElseIf TypeOfList = "Bulleted" Then
                SetIndent p
            Else
                idtInit = Idt(1)
                idtPrev = idtInit
                llInit = ll(1)
                llPrev = llInit
                TypeOfList = "Bulleted"
            End If
            If pChar(9) = ":" Then p.Previous.SpaceAfter = 6

'_______CONTINUE LIST AND SET NEW INDENT (BULLETED OR NUMBERED)_____________________________________________
        Case "bVBA", "VBA"
           SetIndent p
        End Select

        pFontCheck p
        CleanLeadIn p.Range
        GoTo nextP
'___________________________________________________________________________________________________________
'___LIST NUMBERED___________________________________________________________________________________________
    Case wdListSimpleNumbering
isNumbered:
        If CheckNote(p) Then GoTo nextP

'_______ABB/DEF TRIAGE______________________________________________________________________________________
        If inDEF Then
            If pWord(11) <> "note" And ll(1) = 1 And pChar(19) = ":" Then GoTo inDEF
            GoTo isBullet
        ElseIf inABB Then
            If (pFont(1) <> 0 Or pFont(3) <> 0) And pWord(11) <> "note" _
            Then GoTo inABB
        End If

'_______NEW LIST (LIST INIT)________________________________________________________________________________
        Select Case TypeOfList
        Case "Numbered", "None"
            .Style = VB
        Case "Bulleted"
            .Style = bVB
        End Select

        Select Case pLiSt(0)
        Case "No List", "REDACTED NUMBERING"

'_______CHECK IF MUST RESET NUMBERING_______________________________________________________________________
            If listContinue(p) = True Then SetIndent p Else _
                If .ListFormat.ListValue > 1 Then ResetNumbering p
            If pChar(12) = ")" Or pChar(12) = "." Then CleanLeadIn p.Range
            pFontCheck p
            CleanLeadIn p.Range
            GoTo nextP

'_______CONTINUE LIST AND SET NEW INDENT____________________________________________________________________
        Case "bVBA", "VBA"
            SetIndent p
        Case Else
            If ltype(1) = wdListBullet Then
                .Style = bVB
            ElseIf ltype(1) = wdListSimpleNumbering Then
                .Style = VB
            End If
        End Select
        pFontCheck p
        CleanLeadIn p.Range
        GoTo nextP
'___________________________________________________________________________________________________________
'___OUTLINE NUMBERING_______________________________________________________________________________________
    Case wdListOutlineNumbering
isOutline:
        If pChar(12) = vbCr Then .Delete: GoTo nextP
        If pLiSt(1) = "OM-Numbering" Then
            Select Case ll(1)
            Case 2, 3
                .Style = "Heading " & ll(1) - 1: GoTo OutlineTidy
            Case 6
                .Style = "Heading 3": GoTo OutlineTidy
            End Select
        End If

'_______PREAMBLE CHECK______________________________________________________________________________________
        If pWord(11) Like "foreword*" _
        Or (pWord(11) = "distribution" And pWord(12) Like "list*") _
        Or (pWord(11) Like "list*" And pWord(13) = "distribution") _
        Or (pWord(11) Like "record*" And pWord(13) Like "revision*") _
        Or (pWord(11) Like "revision*" And pWord(12) Like "record*") _
        Or (pWord(11) = "history" And pWord(13) Like "revision*") _
        Or (pWord(11) Like "revision*" And pWord(12) = "history") _
        Or (pWord(11) Like "list*" And pWord(13) Like "effective*") _
        Or (pWord(11) = "table" And pWord(13) Like "content*") Then
            .Style = "Heading 1 FWD"
            inPREAMBLE = True
'_______GENERAL OUTLINE_____________________________________________________________________________________
        Else
            If inPREAMBLE Then
                .Style = "Heading 1"
                inPREAMBLE = False
            Else
                .Style = "Heading " & ll(1)
            End If
            TypeOfList = "None"
        End If

'_______MANUAL NUMBERING CLEANING AND HEADING 1 SECTIONING__________________________________________________
OutlineTidy:
        checkSpecial
        CleanManFormat p
        If p.Style Like "Heading 1*" Then CheckH1 p
        GoTo nextP

'___HAVEN'T SEEN ONE YET, CALL 911 IF SIGHTED_______________________________________________________________
    Case wdListListNumOnly
        MsgBox ll(1) & " ListNum fields that can be used in the body of a paragraph." _
        & vbNewLine
    Case wdListMixedNumbering
        MsgBox ll(1) & " Mixed numeric list." & vbNewLine
    Case wdListPictureBullet
        MsgBox ll(1) & " Picture bulleted list." & vbNewLine
    End Select

'___FINAL CHECKS AND RESETS (and error handling reset)______________________________________________________
nextP:
    On Error GoTo 0
    If Not p.Style Like "Heading #*" And Not inTable Then .Font.Size = 11
    pNumbered = False
    pBulleted = False
    i = i + 1

End With

Next

timeParagraphs = MinSec(Timer - timeSubInit)
stsUpdate "stsParagraphs", False

Exit Sub

'***********************************************************************************************************
'ERROR HANDLING*********************************************************************************************

errRange:
pFontCheck p
Resume nextP

errTable:
p.Previous.Range.Characters.Last.InsertBefore vbCr
With p.Range.Tables(1).Borders
    .OutsideColor = wdColorRed
    .OutsideLineStyle = wdLineStyleSingle
    .OutsideLineWidth = wdLineWidth300pt
End With
p.Previous.Range.Delete
tableError = True
Resume nextP
'***********************************************************************************************************
'***********************************************************************************************************

End Sub

'###########################################################################################################
Private Function CheckBullet(ByRef lStr As Variant) As Boolean

CheckBullet = IIf((lStr <> vbNullString And Not lStr Like "[1-9]" _
And Not lStr Like "[A-z]") Or (lStr = vbNullString And pChar(11) Like "[-����]"), True, False)

End Function

'###########################################################################################################
Private Function checkNumbered(ByRef lStr As Variant) As Boolean

checkNumbered = IIf((InStr(lStr, ")") Or InStr(lStr, ".") _
    Or pWord(12) = ")") And Not pStyle(1) Like "Heading #*", True, False)

End Function

'###########################################################################################################
Private Function checkChapter(ByRef p As paragraph) As Boolean

If ((LCase(pWord(11)) Like "chapter" Or LCase(pWord(11)) Like "section") _
And IsNumeric(pWord(12))) Then
    With p.Range
        .Style = "Heading 1"
        .Words.First.Delete
        .Words.First.Delete
    End With
    CleanManFormat p
    CheckH1 p
    checkChapter = True
    checkSpecial
Else
    checkChapter = False
End If

End Function

'###########################################################################################################
Private Function checkSpecial()

    If pWord(11) Like "*abbreviation*" Or pWord(11) Like "*synonym*" Then
        inABB = True
        inDEF = False
    ElseIf pWord(11) Like "*definition*" Then
        inDEF = True
        inABB = False
    Else
        inABB = False
        inDEF = False
    End If

End Function

'###########################################################################################################
Private Function pFontCheck(ByRef p As paragraph)

With p.Range
    If inDEF Or inABB Then Exit Function
    If pFont(1) = 1 Then .Font.Bold = True
    If pFont(2) = 1 Then .Font.Italic = True
    If pFont(3) = 1 Then .Font.Underline = True
End With

End Function

'###########################################################################################################
Private Function setpFont(ByRef rg As Range)

Erase pFont
rg.MoveEnd wdCharacter, -2
With rg
    Select Case .Bold
        Case 0
            pFont(1) = 0
        Case -1, 1
            pFont(1) = 1
        Case Else
            pFont(1) = 9
    End Select
    Select Case .Italic
        Case 0
            pFont(2) = 0
        Case -1, 1
            pFont(2) = 1
        Case Else
            pFont(2) = 9
    End Select
        Select Case .Underline
        Case 0
            pFont(3) = 0
        Case -1, 1
            pFont(3) = 1
        Case Else
            pFont(3) = 9
    End Select
End With

End Function

'###########################################################################################################
Private Function CheckNote(ByRef p As paragraph) As Boolean

If pWord(11) = "note" And (pWord(12) = ":" Or pWord(12) = "-" Or _
pWord(12) = vbTab) Or IsNumeric(pWord(12)) Then
    Set rg = p.Range.Duplicate
    With rg
        If InStr(.Text, vbTab) > 0 Then .Text = Replace(.Text, vbTab, Chr(32))
        .Collapse wdCollapseStart
        .MoveEndUntil ":"
        If devMode Then .Select
        .ParagraphFormat.Reset
        .Style = "Body Text"
        If p.Previous.Range.Tables.Count = 0 Then .ParagraphFormat.LeftIndent = Idt(0)
        .Font.Bold = True
    End With
    CheckNote = True
Else
    CheckNote = False
    Exit Function
End If

End Function

'###########################################################################################################
Private Function listContinue(p As paragraph) As Boolean

If pStyle(0) = "Body Text" And Idt(0) <> bodyIdt Then
    listContinue = True
Else
    idtInit = Idt(1)
    idtPrev = idtInit
    llInit = ll(1)
    llPrev = llInit
    With p
        .Style = VB
    End With
    TypeOfList = "Numbered"
    listContinue = False
End If

End Function

'###########################################################################################################
Private Function ResetNumbering(ByRef p As paragraph)

Set rg = p.Range.Duplicate
With rg
    .Collapse wdCollapseStart
    .InsertBefore vbCr
    .Style = "VBA0"
End With

End Function

'###########################################################################################################
Private Function SetIndent(ByRef p As paragraph)

If pStyle(0) <> "Body Text" Then
    With p.Range
        If pWord(9) = "and" Or pWord(9) = "or" Then
            llPrev = ll(1)
            .ListFormat.ListLevelNumber = ll(0)
            Exit Function
        End If
        Select Case pChar(9)
        Case ";", ","
            llPrev = ll(1)
            .ListFormat.ListLevelNumber = ll(0)
        Case ":"
            llPrev = ll(1)
            .ListFormat.ListLevelNumber = ll(0) + 1
        Case Else
            SetIndentElse p
        End Select
    End With
Else
    SetIndentElse p
End If

End Function

'###########################################################################################################
Private Function SetIndentElse(ByRef p As paragraph)

With p.Range
'    If TypeOfList = "Bulleted" Then
'        .Style = bVB
'    ElseIf ltype(1) = wdListBullet Then
'        .Style = bVB
'    Else
'        .Style = VB
'    End If
'    If idtInit = 0 Then Exit Function
    Select Case ll(1)
    Case Is = llPrev
        lStr = llPrev - llInit
        Do Until lStr = 0
            .ListFormat.ListIndent
            lStr = lStr - 1
            If lStr < 0 Then Exit Do
            DoEvents
        Loop
    Case Is = llInit
        Exit Function
    Case Is > llPrev
        llPrev = ll(1)
        .ListFormat.ListIndent
    Case Is < llPrev
        lStr = llPrev - ll(1)
        llPrev = ll(1)
        Do Until lStr = 0 Or .Style Like "*1"
            .ListFormat.ListOutdent
            lStr = lStr - 1
            If lStr < 0 Then Exit Do
            DoEvents
        Loop
    End Select
End With

End Function

'###########################################################################################################
Private Function CleanLeadIn(ByRef pRange As Range)

With pRange.Characters
    Do Until .First Like "[A-z#]" And Not .First.Next Like "[).]"
        On Error GoTo errDelete
        .First.Delete
        On Error GoTo 0
        DoEvents
    Loop
    .First = UCase(.First)
End With

Exit Function
'___________________________________________________________________________________________________________
errDelete:
p.Range.Characters.First.Cut
Resume Next

End Function
'###########################################################################################################

Private Function CleanManFormat(ByRef p As paragraph)

If InStr(p.Range, vbTab) > 0 Then
    With p.Range.Characters
        Do Until .First Like "[A-z]"
            On Error GoTo errDelete
            .First.Delete
            On Error GoTo 0
            DoEvents
        Loop
        .First = UCase(.First)
    End With
End If

Exit Function
'___________________________________________________________________________________________________________
errDelete:
p.Range.Characters.First.Cut
Resume Next

End Function
'###########################################################################################################

Private Function CheckH1(ByRef p As paragraph)

Dim rgF As Range

If p.Range <> Documents(selDoc).Paragraphs(1).Range Then
    If InStr(p.Previous.Range.Text, Chr(12)) = 0 Then
        Set rgF = p.Range.Duplicate
        With rgF
            .Collapse wdCollapseStart
            .InsertBreak wdSectionBreakNextPage
            .MoveEnd wdParagraph, -1
            .Paragraphs(1).Range.Style = "Body Text"
        End With
    End If
End If

End Function
'###########################################################################################################

Private Function TableFormat(ByRef p As paragraph)

If Not tFormatted Then
    With p.Range.Tables(1).Borders
        On Error Resume Next
        .InsideLineStyle = wdLineStyleSingle
        .InsideLineWidth = wdLineWidth050pt
        .InsideColor = ASTBlue
        On Error GoTo 0
    End With
    tFormatted = True
End If
With p.Range.Cells(1)
    If .Range.Bold = -1 Then
        .Range.Style = "TABLE Header"
        .Borders.OutsideColor = wdColorWhite
        With .Shading
            .Texture = wdTextureNone
            .BackgroundPatternColor = ASTBlue
        End With
    Else
        p.Style = "TABLE Content"
    End If
End With

If pFontName Like "Wingdings*" Then p.Range.Font.Name = pFontName

If p.Next.Next.Range.Tables.Count = 0 Then
    With p.Range.Tables(1)
        .AllowAutoFit = True
        .Spacing = 0
        .AllowPageBreaks = True
        .AutoFitBehavior (wdAutoFitContent)
        .TopPadding = MillimetersToPoints(1)
        .BottomPadding = MillimetersToPoints(1)
        .LeftPadding = MillimetersToPoints(2)
        .RightPadding = MillimetersToPoints(2)
        With .Range
            .Font.Size = 9
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
            .OutsideLineStyle = wdLineStyleSingle
            .OutsideLineWidth = wdLineWidth050pt
            .OutsideColor = ASTBlue
        End With
        .Spacing = 0
        With .Rows
            .LeftIndent = CentimetersToPoints(0)
            .Alignment = wdAlignRowCenter
        End With
    End With
End If

End Function
