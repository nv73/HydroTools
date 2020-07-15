VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MagCO 
   Caption         =   "Mag Cable Out"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7455
   OleObjectBlob   =   "MagCO.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MagCO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub loadFile_Click()

Dim xcoll As New Collection
Dim ycoll As New Collection
Dim linenumcoll As New Collection
Dim fixnum As New Collection
Dim cableoutcoll As New Collection
Dim adjustedCO As New Collection
Dim p As Integer
Dim rawFixes As New Collection
Dim expandedCO As New Collection
Dim expandedAdjustedCO As New Collection

On Error GoTo bigbaderror:

Application.ScreenUpdating = False

'Allows user to load and save the desired file which needs converting

Dim loadWindow
Dim filename As String

'Now to open up the cableout file from hydromap

loadWindow = Application.GetOpenFilename(Title:="LB FILE", FileFilter:="LB Cable Out Files *.lb* (*.lb*),")

Workbooks.Open (loadWindow)

'Delimit the imported data.------------------------------------------------------------------------
   
Columns("A:A").Select

Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
    Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
    :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1)), _
    TrailingMinusNumbers:=True
        
Columns("A:A").EntireColumn.AutoFit
'finish delimiting

'Populate collections with the imported data-------------------------------------------------------
Range("A1").Select

p = hydrotools.getListLength("A1")

For i = 1 To p
    
    fixnum.Add (ActiveCell.Value)
    cableoutcoll.Add (ActiveCell.Offset(0, 1).Value)
    'ActiveCell.Offset(0, 2).Value = ActiveCell.Offset(0, 1).Value * ActiveCell.Offset(0, 2).Value
    adjustedCO.Add (ActiveCell.Offset(0, 2).Value)
    ActiveCell.Offset(1, 0).Select
    Debug.Print (adjustedCO(i))
    
Next i

Application.DisplayAlerts = False

ActiveWorkbook.Close

Application.DisplayAlerts = True

'Open desired file raw mag file--------------------------------------------------------------------

loadWindow = Application.GetOpenFilename(Title:="Please select a mag file", FileFilter:="Mag Files *.mag* (*.mag*),")

Workbooks.Open (loadWindow)

'Delimit the imported data.------------------------------------------------------------------------
    
Columns("A:A").Select

Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
    Semicolon:=False, Comma:=True, Space:=True, Other:=False, FieldInfo _
    :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1)), _
    TrailingMinusNumbers:=True
        
Columns("A:A").EntireColumn.AutoFit

Columns("B:B").NumberFormat = "HH:MM:SS"
'Finish delimiting

Range("E1").Select

p = hydrotools.getListLength("E1")

'Add values to collections
For i = 1 To p
    
    ActiveCell.Value = ActiveCell.Value + xOff.Value
    ActiveCell.Offset(0, 1).Value = ActiveCell.Offset(0, 1).Value + yOff.Value
    xcoll.Add (ActiveCell.Value)
    ycoll.Add (ActiveCell.Offset(0, 1).Value)
    rawFixes.Add (ActiveCell.Offset(0, -4).Value)
    ActiveCell.Offset(1, 0).Select
    
Next i

Range("A1").Select

Dim n As Integer

n = 1

For i = 1 To p

    If rawFixes(i) = fixnum(n) Then
    
        expandedCO.Add (fixnum(n))
        expandedAdjustedCO.Add (adjustedCO(n))
        
    Else
    
        n = n + 1
        expandedCO.Add (fixnum(n))
        expandedAdjustedCO.Add (adjustedCO(n))
        
    End If
        
Next i

Range("F1").Select

For i = 1 To p

    ActiveCell.Value = ycoll(i) - expandedAdjustedCO(i)
    Debug.Print (ycoll(i) & " - " & expandedAdjustedCO(i))
    Debug.Print (ActiveCell.Value)
    ActiveCell.Offset(0, -1).Value = xcoll(i)
    ActiveCell.Offset(1, 0).Select
    
Next i

filename = Application.GetSaveAsFilename(ActiveWorkbook.Name, FileFilter:="Processed Mag File (*.proc), *.proc")

On Error Resume Next
            
ActiveWorkbook.SaveAs filename, xlCSV
            
ActiveWorkbook.Saved = True

Application.DisplayAlerts = False

'Close the active workbook and return to the primary hydrotools workspace

ActiveWorkbook.Close

Application.DisplayAlerts = True

Workbooks("HydroTools_Active.xlsb").Activate
Worksheets("Sheet1").Activate

Application.ScreenUpdating = True

bigbaderror:

    If Err.Number = 1004 Then
    
        MsgBox ("Whoops! Looks like you didn't enter a valid filename!")
        Application.ScreenUpdating = True
        
        If ActiveWorkbook.Name = "False.xlsx" Then
            ActiveWorkbook.Close
        End If
        
        Workbooks("HydroTools_Active.xlsb").Activate
        Worksheets("Sheet1").Activate
        
    End If

End Sub
