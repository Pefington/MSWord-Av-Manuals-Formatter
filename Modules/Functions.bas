Attribute VB_Name = "Functions"
Option Explicit

'###########################################################################################################
Public Function UpdateProgress(ByRef Progress As Single)

frmProgress.infoProgress.Value = Progress * 100
frmProgressSmall.infoProgress.Value = Progress * 100
DoEvents

End Function

'###########################################################################################################
Public Function UpdateCaption(ByRef Caption As String)

timeTotal = MinSec(Timer - timeInit)
frmProgress.Caption = timeTotal & " - " & Caption
frmProgressSmall.Caption = timeTotal & " - " & Caption
DoEvents

End Function

'###########################################################################################################
Public Function stsUpdate(ByRef stsItem As String, Start As Boolean)

ASTBlue = 12611584

If Start Then
    frmProgress.Controls(stsItem).Font.Underline = True
    frmProgress.Controls(stsItem).ForeColor = ASTBlue
Else
    frmProgress.Controls(stsItem).Font.Underline = False
    frmProgress.Controls(stsItem).ForeColor = wdColorGreen
    frmProgress.Controls(stsItem).Value = True
End If

End Function
