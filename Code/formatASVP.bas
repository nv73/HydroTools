Attribute VB_Name = "formatASVP"
Sub ASVP_Format()

'Used to format data from an SVP so it can be converted to an .sndvel file and edited in SV tool.
'Written by: Nick Viner
'Last updated: 08/21/2015
'Status: Working with stringent data structure requirements.

'Create necessary variables

Dim cellVal As Integer
Dim rowVal As Double
Dim cellString As String

'Set necessary base variable values as well as starting cell.

Cells(1, 1).Select
cellVal = 1
rowVal = 10

'Minor error catching

If IsNumeric(Cells(1, 1)) Then
    MsgBox ("ERROR: Data is of improper format. Ensure all metadata (ie. Now:, Battery Level:, RapidSVT:, etc.) has not been deleted by user")
    Exit Sub
End If
    
'Check with user to prevent accidental macro use
iRet = MsgBox("Do you wish to continue? If data is of improper format, it may be lost", vbYesNo)

    If iRet = vbNo Then
        Exit Sub
    End If


'Insert # into first 9 lines which always contain text with this particular format and need to be commented out.

While cellVal <= 9
    
    'Insert a # into the beginning of the active cell
    cellString = "#" & ActiveCell.Value
    ActiveCell = cellString
    
    'Move down to next cell
    cellVal = cellVal + 1
    Cells(cellVal, 1).Select
    
    Wend
    
    
'Delete middle column
    
Columns("B:B").Select
Selection.ClearContents


'Convert the contained depth values to negative units (value * -1)

Cells(rowVal, 1).Select

While IsEmpty(ActiveCell.Value) = False
    
    'Multiply depth by -1 to make it negative
    ActiveCell.Value = ActiveCell.Value * -1
    
    'Move the active cell.
    rowVal = rowVal + 1
    Cells(rowVal, 1).Select
    
    'Sometimes a < appears in the data. This will delete it and prevent it from screwing up the macro.
    If ActiveCell.Value = "<" Then
    
        ActiveCell.ClearContents
               
    End If
    
Wend
    


End Sub

