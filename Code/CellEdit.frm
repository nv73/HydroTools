VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CellEdit 
   Caption         =   "Cell Editor"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6330
   OleObjectBlob   =   "CellEdit.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cellEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_initialize()

Dim currentSelection As String

currentSelection = Selection.Address

selectionArchive.Value = currentSelection

End Sub

Private Sub copySelection_Click()

Selection.Copy

End Sub

Private Sub deleteSelection_Click()

Selection.ClearContents

End Sub

Private Sub downButton_Click()

Selection.Offset(1, 0).Select

activeCellBox.Value = Selection.Address

End Sub

Private Sub editValue_Click()

Selection.Value = InputBox("Enter in desired new value: ", "Edit Cell")

End Sub

Private Sub leftButton_Click()

Selection.Offset(0, -1).Select

activeCellBox.Value = Selection.Address

End Sub

Private Sub pasteSelection_Click()

ActiveSheet.Paste

End Sub

Private Sub rightButton_Click()

Selection.Offset(0, 1).Select

activeCellBox.Value = Selection.Address

End Sub

Private Sub selectRange_Click()

Dim addressVal As String

addressVal = InputBox("Please enter desired range (ie. B6 or B6:C7): ", "Select Range")

Range(addressVal).Select

activeCellBox.Value = Selection.Address

End Sub

Private Sub upButton_Click()

Selection.Offset(-1, 0).Select

activeCellBox.Value = Selection.Address

End Sub

Private Sub UserForm_Terminate()

Range(selectionArchive.Value).Select

UserForm1.Show

AutoLog.Show

End Sub
