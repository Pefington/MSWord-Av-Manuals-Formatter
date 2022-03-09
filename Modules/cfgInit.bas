Attribute VB_Name = "cfgInit"
Option Explicit

Public Sub cfgStart()

devMode = False

If selDoc = vbNullString Then Call cfgDev: Exit Sub
If selDoc = "AvManFormatter" Then MsgBox "I won't edit myself!", vbExclamation: End

Documents(selDoc).Activate
Windows(ThisDocument).WindowState = wdWindowStateMinimize
Application.ScreenUpdating = False
With Documents(selDoc)
    .GrammarChecked = True
    .Range.SpellingChecked = True
    .ShowGrammaticalErrors = False
    .ShowSpellingErrors = False
End With
With Options
    .CheckGrammarAsYouType = False
    .CheckSpellingAsYouType = False
    .AnimateScreenMovements = False
    .Pagination = False
End With
With Windows(selDoc)
    .View.Zoom.Percentage = 100
    .View = wdNormalView
    .WindowState = wdWindowStateMinimize
End With

End Sub

Public Sub cfgDev()

devMode = True
Application.ScreenUpdating = True
If selDoc = vbNullString Then selDoc = ActiveDocument
If selDoc = "AvManFormatter" Then MsgBox "I won't edit myself!", vbExclamation: End
With Documents(selDoc)
    .GrammarChecked = True
    .Range.SpellingChecked = True
    .ShowGrammaticalErrors = False
    .ShowSpellingErrors = False
End With
With Options
    .CheckGrammarAsYouType = False
    .CheckSpellingAsYouType = False
    .AnimateScreenMovements = True
    .Pagination = False
End With
With Windows(selDoc)
    .View = wdNormalView
    .View.Zoom.Percentage = 100
    .WindowState = wdWindowStateMaximize
End With

End Sub

Public Sub cfgEnd()

devMode = False
Windows(ThisDocument).WindowState = wdWindowStateMinimize
With Options
    .CheckGrammarAsYouType = True
    .CheckSpellingAsYouType = True
    .AnimateScreenMovements = True
    .Pagination = True
End With
With Documents(selDoc)
    .UndoClear
    .GrammarChecked = False
    .Range.SpellingChecked = False
    .ShowGrammaticalErrors = True
    .ShowSpellingErrors = True
    .Activate
End With
With Windows(selDoc)
    .View = wdPrintView
    .View.Zoom.Percentage = 100
    .WindowState = wdWindowStateMaximize
End With
Application.ScreenUpdating = True
Unload frmProgress
Unload frmProgressSmall

Documents(selDoc).Repaginate

End Sub
