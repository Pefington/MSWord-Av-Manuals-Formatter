Attribute VB_Name = "bookmarks"
Option Explicit

Sub Clean()

stsUpdate "stsBookmarks", True
DoEvents

Dim B As Bookmark

For Each B In Documents(selDoc).bookmarks

    B.Delete

Next

stsUpdate "stsBookmarks", False

End Sub
