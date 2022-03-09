Attribute VB_Name = "Sanitise"
Option Explicit

Sub Initial()

stsUpdate "stsSanitise1", True
timeSubInit = Timer
Dim i As Long, ops As Long: i = 0: ops = 11

With Documents(selDoc).Range.Find

UpdateCaption ("Searching for and removing multiple spaces.")
UpdateProgress (i / ops)
i = i + 1

    .ClearFormatting
    .Replacement.ClearFormatting
    .Text = " {2,}"
    .Replacement.Text = " "
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = False
    .MatchWholeWord = False
    .MatchAllWordForms = False
    .MatchSoundsLike = False
    .MatchWildcards = True
    .Execute Replace:=wdReplaceAll

UpdateCaption ("Searching for and removing multiple tab spaces.")
UpdateProgress (i / ops)
i = i + 1
    .Text = "^t{2,}"
    .Replacement.Text = vbTab
    .Execute Replace:=wdReplaceAll

UpdateCaption ("Sanitising colons.")
UpdateProgress (i / ops)
i = i + 1
    .Text = " :"
    .Replacement.Text = ":"
    .MatchWildcards = False
    .Execute Replace:=wdReplaceAll

UpdateCaption ("Sanitising commas.")
UpdateProgress (i / ops)
i = i + 1
    .Text = " ,"
    .Replacement.Text = ","
    .Execute Replace:=wdReplaceAll

UpdateCaption ("Searching for and removing spaces at paragraph start.")
UpdateProgress (i / ops)
i = i + 1
    .Text = "^p "
    .Replacement.Text = "^p"
    .Execute Replace:=wdReplaceAll

UpdateCaption ("Searching for and removing tabs at paragraph start.")
UpdateProgress (i / ops)
i = i + 1
    .Text = "^p^t"
    .Replacement.Text = "^p"
    .Execute Replace:=wdReplaceAll

UpdateCaption ("Searching for and removing tabs at paragraph end.")
UpdateProgress (i / ops)
i = i + 1
    .Text = "^t^p"
    .Replacement.Text = "^p"
    .Execute Replace:=wdReplaceAll

UpdateCaption ("Searching for and removing empty paragraphs (quick pass).")
UpdateProgress (i / ops)
i = i + 1
    .Text = "^13{2,}"
    .Replacement.Text = "^p"
    .MatchWildcards = True
    .Execute Replace:=wdReplaceAll

UpdateCaption ("Searching for and removing 'LEFT BLANK' type 1.")
UpdateProgress (i / ops)
i = i + 1
    .Text = "I?N?T?E?N?T?I?O?N?A?L?L?Y?L?E?F?T?B?L?A?N?K*^13"
    .Replacement.Text = vbNullString
    .Execute Replace:=wdReplaceAll

UpdateCaption ("Searching for and removing 'LEFT BLANK' type 2.")
UpdateProgress (i / ops)
i = i + 1
    .Text = "INTENTIONALLY?LEFT?BLANK*^13"
    .Replacement.Text = vbNullString
    .Execute Replace:=wdReplaceAll

UpdateCaption ("Searching for and removing 'LEFT BLANK' type 3.")
UpdateProgress (i / ops)
i = i + 1
    .Text = "INTENTIONALLY?BLANK*^13"
    .Replacement.Text = vbNullString
    .Execute Replace:=wdReplaceAll

UpdateCaption ("All done.")
UpdateProgress (i / ops)

End With

With Documents(selDoc)


    UpdateCaption ("Removing comments.")
    DoEvents

    For i = .Comments.Count To 1 Step -1
        .Comments(i).DeleteRecursively
    Next

End With

timeSanitise = MinSec(Timer - timeSubInit)
stsUpdate "stsSanitise1", False

End Sub
