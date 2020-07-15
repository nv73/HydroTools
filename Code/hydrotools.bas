Attribute VB_Name = "hydrotools"
'-------------------------------------------------------'
'                                                       '
'-------------Hydrotools Function Library---------------'
'                                                       '
'-------------------------------------------------------'

Function getRange() As Variant

'Allow the user to select a range of cells. Returns a string in A1 format.

Dim rangeVal As String

rangeVal = Application.InputBox("Test", "Title", , , , , , 0)

rangeVal = hydrotools.deleteCharByChar(rangeVal, "=")

getRange = Application.ConvertFormula(rangeVal, xlR1C1, xlA1)

End Function

Function delCharByChar(inputstring As String, charDel As String) As String

'get length of input string
Dim stringLength As Integer

stringLength = Len(inputstring)

'Create a collection of characters from inputString

Dim stringCollection As New Collection

For p = 1 To stringLength
    
    stringCollection.Add (Mid(inputstring, p, 1))
   
Next p

'Take the character which the user wants to delete and remove it from the string

For p = stringLength To 1 Step -1

    If stringCollection(p) = charDel Then
    
        stringCollection.Remove (p)
        
    End If

Next p

'Take the new set of characters and reinsert them into a string, returning it to the user

Dim newstring As String

newstring = ""

For p = 1 To stringCollection.Count
    
    newstring = newstring & stringCollection(p)
        
Next p

delCharByChar = newstring

End Function

Function delCharByIndex(inputstring As String, indexDel As Integer) As String

'get length of input string
Dim stringLength As Integer

stringLength = Len(inputstring)

'Create a collection of characters from inputString

Dim stringCollection As New Collection

For p = 1 To stringLength
    
    stringCollection.Add (Mid(inputstring, p, 1))
   
Next p

stringCollection.Remove (indexDel)

Dim newstring As String

newstring = ""

For p = 1 To stringCollection.Count
    
    newstring = newstring & stringCollection(p)
        
Next p

delCharByIndex = newstring

End Function


Function getListLength(topOfListRange As String) As Integer
'Returns the length of the list


Dim currentRange As String
Dim tLength As Integer

Range(topOfListRange).Select

'Stores the current range so that it can be later selected
currentRange = ActiveCell.Address

'Runs through the rows until a blank value is encountered
While ActiveCell.Value <> ""

    ActiveCell.Offset(1, 0).Select
    tLength = tLength + 1

Wend

'A small, crude check to see if there is a break in the list.
If ActiveCell.Offset(1, 0).Value <> "" Then

    MsgBox ("Looks like you have yourself a break in the list. You should fix that or select a non broken column!")
    
End If

Range(currentRange).Select

getListLength = tLength

End Function

Function random(minVal As Integer, maxVal As Integer) As Integer

'A more user friendly version of VBs rand function

random = Int((maxVal - minVal + 1) * Rnd + minVal)

End Function

Function storeList(inputRange As String) As Collection

Dim currentRange As String

currentRange = inputRange

End Function

Function isActiveWorkBook() As Boolean

'Determines if the currently active workbook is Hydrotools.
If ActiveWorkbook.Name = "HydroTools_Active.xlsb" Then

    isActiveWorkBook = True
    
    Else
    
    isActiveWorkBook = False
    
End If
    
    
End Function

Function selectLeft()

'Selects the left most cell of the row containing the activecell.

Dim cellAddressLetter As String
Dim cellAddressNumber As Integer

'Parse the cell address into a letter and a number representing the column and row respectively.

cellAddressLetter = ActiveCell.Address

cellAddressLetter = delCharByChar(cellAddressLetter, "$")

cellAddressNumber = delCharByIndex(cellAddressLetter, 1)

cellAddressLetter = delCharByIndex(cellAddressLetter, 2)

'If the currently selected cell is not in column A, select the cell in the same row in column A.
If cellAddressLetter <> "A" Then

    Range("A" & cellAddressNumber).Select
    
End If

End Function

Function correctLogRange()

'First select the left most cell of the row containing the activecell.
hydrotools.selectLeft
ActiveCell.Offset(0, 1).Select

'if the cell above the activecell is blank, offset the activecell upwards until it borders a cell containing data.
If ActiveCell.Offset(-1, 0) = "" Then

    While ActiveCell.Offset(-1, 0) = ""

        ActiveCell.Offset(-1, 0).Select

    Wend
    
End If

'If the activecell already contains data, offset it down until an empty cell is reached.
If ActiveCell.Value <> "" Then

    While ActiveCell.Value <> ""
    
        ActiveCell.Offset(1, 0).Select
        
    Wend
    
End If

End Function


Function splitString(inputstring As String, delimiter As String) As Variant

Dim storestring() As String

storestring = Split(inputstring, delimiter)

splitString = storestring()

End Function
