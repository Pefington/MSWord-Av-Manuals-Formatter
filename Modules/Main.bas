Attribute VB_Name = "Main"
Option Explicit

Public timeSubInit As Single, timeTotal As String, timeFakeHeaders As String, timeSectionsRemoval As String, _
skipFakes As Boolean, timePageSetup As String, timeParagraphs As String, timeSectionsRebuild As String, _
timeShapes As String, timeBookmarks As String, timeHeadersFooters As String, timeHyperlinks As String, _
timeSanitise As String, timeTables As String, timeEmptyParas As String, selP As String, devMode As Boolean, _
selDoc As String, strDEP As String, strREF As String, strTITLE As String, strSUBT As String, strAUTH As String, _
strISSUE As String, strREV As String, strDATE As Date, depOK As Boolean, refOK As Boolean, titleOK As Boolean, _
subtOK As Boolean, authOK As Boolean, issueOK As Boolean, revOK As Boolean, dateOK As Boolean, skipCover As Boolean, _
allOK As Boolean, timeInit As Single, ASTBlue As String, timeHeadersSetup As String, timeTOC As String, _
timeCleanTOC As String, timeLEP As String

Public Function MinSec(time) As String

MinSec = Format(time \ 60, "00") & ":" & Format(time Mod 60, "00")

End Function

Sub AvManFormatter()

Application.ScreenUpdating = True
With Windows(ThisDocument)
    .WindowState = wdWindowStateMaximize
    .View = wdPrintView
    With .View.Zoom
        .PageRows = 2
        .PageColumns = 4
    End With
End With
ActiveWindow.DocumentMap = False
'CommandBars("Styles").Visible = False

ASTBlue = 12611584

Dim mbReply As Long

    mbReply = MsgBox("Welcome to AvManFormatter Script." & vbNewLine & _
        vbNewLine & "Click OK to continue.", vbOKCancel, "AvManFormatter")

    If mbReply = 2 Then

        MsgBox "Goodbye~", vbInformation, "Bruh..."
        End

    End If

    MsgBox "We will now first select the document to process, then the REDACTED template.", vbOKOnly, "AvManFormatter"

Call OpenFile

timeInit = Timer

frmCoverSetup.Show

Do While frmCoverSetup.Visible

    If depOK And refOK And titleOK And subtOK And authOK And issueOK _
    And revOK And dateOK Then allOK = True Else allOK = False

    With frmCoverSetup.cmdCoverDone
        If allOK Then
            .ForeColor = wdColorGreen
            .Enabled = True
            .MousePointer = fmMousePointerDefault
        Else
            .Enabled = False
            .MousePointer = fmMousePointerNoDrop
        End If
    End With

    DoEvents

Loop

strDEP = frmCoverSetup.fillDEP
strREF = frmCoverSetup.fillREF
strTITLE = frmCoverSetup.fillTITLE
strSUBT = frmCoverSetup.fillSUBT
strAUTH = StrConv(frmCoverSetup.fillAUTH, vbProperCase)
strISSUE = frmCoverSetup.fillISSUE
strREV = frmCoverSetup.fillREV
strDATE = frmCoverSetup.fillDATE

Unload frmCoverSetup

Call FakeHeaders.Init

Do While frmHeadersPrompt.Visible
DoEvents
Loop

Call cfgStart

frmProgress.Show
DoEvents

Call bookmarks.Clean

Call Sections.Clean

Call Shapes.CleanFloating

If Not skipFakes Then Call FakeHeaders.CleanTables

Call Shapes.CleanInline

Call HeadersFooters.Clean

Call PageSetup.LayoutSetup

Call Hyperlinks.Clean

Call Sanitise.Initial

'MsgBox "Ready.": End

Call Paragraphs.Parse

Call TOC.Clean

Call Tables.Formatting

If Not skipCover Then Call PageSetup.CoverPages

Call HeadersFooters.Setup

Call TOC.Generate

Call LEP.Generate

Call OpsComplete

End Sub

Sub OpenFile()

Dim fd As FileDialog, FileSelected As Boolean, docPath As Variant, objFSO As Object

Set fd = Application.FileDialog(msoFileDialogOpen)

fd.Filters.Clear
fd.Filters.Add "Word Files", "*.doc*"
fd.AllowMultiSelect = False
fd.InitialFileName = Environ("UserProfile") & "\Downloads"
fd.Title = "Select the file to process."
fd.ButtonName = "Let's do this!"

FileSelected = fd.Show

If Not FileSelected Then
    Call cfgEnd
    MsgBox "Goodbye~", vbInformation, "Bruh..."
    End
End If

Windows(ThisDocument).WindowState = wdWindowStateMinimize

fd.Execute

docPath = Split(fd.SelectedItems(1), "\")
selDoc = docPath(UBound(docPath))

Call cfgDev

On Error Resume Next
Documents(selDoc).Convert
On Error GoTo 0

Call AttachTemplate

End Sub

Sub AttachTemplate()

Call cfgDev

Dim fd As FileDialog, FileSelected As Boolean

Set fd = Application.FileDialog(msoFileDialogFilePicker)

fd.Filters.Clear
fd.Filters.Add "REDACTED Template File", "*.dotx"
fd.AllowMultiSelect = False
fd.InitialFileName = Environ("UserProfile") & "\Downloads"
fd.Title = "Please locate up to date REDACTED.dotx"
fd.ButtonName = "That's the one!"

FileSelected = fd.Show

If Not FileSelected Then
    Documents(selDoc).Close
    Call cfgEnd
    MsgBox "I need a template.", vbInformation, "Bruh..."
    End
End If

With Documents(selDoc)
.UpdateStylesOnOpen = True
.AttachedTemplate = fd.SelectedItems(1)
.CopyStylesFromTemplate (fd.SelectedItems(1))
End With

End Sub

Sub OpsComplete()

timeTotal = MinSec(Timer - timeInit)

Documents(selDoc).UndoClear

Call cfgEnd
ActiveWindow.ScrollIntoView Documents(selDoc).Range.Paragraphs(1).Range, True

If skipFakes Then

MsgBox "All operations complete." & vbNewLine & vbNewLine & "Time taken (useful to Pef):" & _
    vbNewLine & vbNewLine & "Sorting shapes: " & timeShapes & vbNewLine & _
    "Removing fake headers: Skipped." & vbNewLine & _
    "Removing sections: " & timeSectionsRemoval & vbNewLine & _
    "Layout setup: " & timePageSetup & vbNewLine & _
    "Hyperlinks/TOCs cleanup: " & timeHyperlinks & vbNewLine & _
    "Sanitising: " & timeSanitise & vbNewLine & _
    "Paragraphs parsing: " & timeParagraphs & vbNewLine & _
    "Manual TOCs removal: " & timeCleanTOC & vbNewLine & _
    "Tables formatting: " & timeTables & vbNewLine & _
    "Headers and footers setup: " & timeHeadersSetup & vbNewLine & _
    "TOC generation: " & timeTOC & vbNewLine & _
    "LEP generation: " & timeLEP & vbNewLine & _
    "Total time since launch of program: " & timeTotal & vbNewLine & vbNewLine & _
    "When you click OK, I will end." & vbNewLine & vbNewLine & _
    "~~Goodbye~~", vbInformation, "AvManFormatter"
Else
MsgBox "All operations complete." & vbNewLine & vbNewLine & "Time taken (useful to Pef):" & _
    vbNewLine & vbNewLine & "Sorting shapes: " & timeShapes & vbNewLine & _
    "Removing fake headers: " & timeFakeHeaders & vbNewLine & _
    "Removing sections: " & timeSectionsRemoval & vbNewLine & _
    "Layout setup: " & timePageSetup & vbNewLine & _
    "Hyperlinks/TOCs cleanup: " & timeHyperlinks & vbNewLine & _
    "Sanitising: " & timeSanitise & vbNewLine & _
    "Paragraphs parsing: " & timeParagraphs & vbNewLine & _
    "Manual TOCs removal: " & timeCleanTOC & vbNewLine & _
    "Tables formatting: " & timeTables & vbNewLine & _
    "Headers and footers setup: " & timeHeadersSetup & vbNewLine & _
    "TOC generation: " & timeTOC & vbNewLine & _
    "LEP generation: " & timeLEP & vbNewLine & _
    "Total time since launch of program: " & timeTotal & vbNewLine & vbNewLine & _
    "When you click OK, I will end." & vbNewLine & vbNewLine & _
    "~~Goodbye~~", vbInformation, "AvManFormatter"
End If

End Sub
