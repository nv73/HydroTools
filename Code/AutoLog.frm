VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AutoLog 
   Caption         =   "AutoLog 1.0"
   ClientHeight    =   10380
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5955
   OleObjectBlob   =   "AutoLog.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AutoLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_initialize()

Dim projectNum As String
Dim registryNumber As String
Dim vesselName As String
Dim sublocality As String
Dim currentDate As Date
Dim julianDay As Integer
Dim pageNum As Integer
Dim logTemplate As String
    
End Sub

Private Sub logGen_Click()

Dim sFileSaveName As Variant
Dim baseFilePath As String

baseFilePath = ActiveWorkbook.Path

logTemplate = baseFilePath & "\LogTemplates\projLogsExcel.xlsm"
currentDate = Now()

Range("L4").NumberFormat = "yyyy/mm/dd"

projectNum = InputBox("Please enter the project number", "Create Log Template")
registryNumber = InputBox("Please enter the registry number", "Create Log Template")
vesselName = InputBox("Please enter the vessel name", "Create Log Template")
sublocality = InputBox("Please describe breifly the project location", "Create Log Template")
julianDay = InputBox("Please enter the Julian date", "Create Log Template")

pageNum = (ActiveSheet.HPageBreaks.Count) + 1

Workbooks.Open (logTemplate)

Workbooks("projLogsExcel.xlsm").Activate

Range("G3").Value = projectNum
Range("G4").Value = registryNumber
Range("G5").Value = vesselName
Range("G6").Value = sublocality
Range("L4").Value = currentDate
Range("L5").Value = julianDay
Range("L3").Value = pageNum

Range("B10").Select

MultiPage1.Enabled = True

sFileSaveName = Application.GetSaveAsFilename(InitialFileName:=initialname, FileFilter:="Excel Files(*.xlsm), *.xlsm")

If sFileSaveName <> False Then
    ActiveWorkbook.SaveAs sFileSaveName
End If


End Sub

Private Sub newPage()

If ActiveCell.Value = "#" Then

    ActiveCell.Offset(-49, -1).Select
    
    Range(ActiveCell, ActiveCell.Offset(49, 14)).Select
    
    Selection.Copy
    
    ActiveCell.Offset(50, 0).Select
    
    ActiveSheet.Paste
        
    ActiveCell.Offset(9, 1).Select
    
    Range(ActiveCell, ActiveCell.Offset(39, 7)).Select
    
    Selection.RowHeight = 23
    
    Selection.ClearContents
    
    ActiveCell.Offset(0, 0).Select
            
        
End If


End Sub

Private Sub SOL_Click()

hydrotools.correctLogRange

'--------------------------------------------
If hydrotools.isActiveWorkBook = True Then

    MsgBox "Please ensure the log workbook is active before continuing"
    
    Exit Sub
    
End If
'--------------------------------------------

ActiveCell.Value = Now()

ActiveCell.NumberFormat = "hh:mm"

ActiveCell.Offset(0, 1) = InputBox("Input line name")

ActiveCell.Offset(0, 2) = InputBox("Input fix number")

ActiveCell.Offset(0, 3) = InputBox("Input heading")

ActiveCell.Offset(0, 4) = InputBox("Input Speed")

ActiveCell.Offset(0, 5) = InputBox("Input HDOP")

ActiveCell.Offset(0, 6) = InputBox("Input depth")

ActiveCell.Offset(0, 7) = "SOL" & " " & InputBox("Any extra comments?")

ActiveCell.Offset(1, 0).Select

'--------------------------------------------------------
If ActiveCell.Value = "#" Then

    ActiveCell.Offset(-49, -1).Select
    
    Range(ActiveCell, ActiveCell.Offset(49, 14)).Select
    
    Selection.Copy
    
    ActiveCell.Offset(50, 0).Select
    
    ActiveSheet.Paste
        
    ActiveCell.Offset(9, 1).Select
    
    Range(ActiveCell, ActiveCell.Offset(39, 7)).Select
    
    Selection.RowHeight = 23
    
    Selection.ClearContents
    
    ActiveCell.Offset(0, 0).Select
        
End If
'---------------------------------------------------------

End Sub

Private Sub EOL_Click()

hydrotools.correctLogRange

'--------------------------------------------
If hydrotools.isActiveWorkBook = True Then

    MsgBox "Please ensure the log workbook is active before continuing"
    
    Exit Sub
    
End If
'--------------------------------------------

ActiveCell.Value = Now()

ActiveCell.NumberFormat = "hh:mm"

ActiveCell.Offset(0, 1) = InputBox("Input line name")

ActiveCell.Offset(0, 2) = InputBox("Input fix number")

ActiveCell.Offset(0, 3) = InputBox("Input heading")

ActiveCell.Offset(0, 4) = InputBox("Input Speed")

ActiveCell.Offset(0, 5) = InputBox("Input HDOP")

ActiveCell.Offset(0, 6) = InputBox("Input depth")

ActiveCell.Offset(0, 7) = "EOL" & " " & InputBox("Any extra comments?")

ActiveCell.Offset(1, 0).Select

'--------------------------------------------------------
If ActiveCell.Value = "#" Then

    ActiveCell.Offset(-49, -1).Select
    
    Range(ActiveCell, ActiveCell.Offset(49, 14)).Select
    
    Selection.Copy
    
    ActiveCell.Offset(50, 0).Select
    
    ActiveSheet.Paste
        
    ActiveCell.Offset(9, 1).Select
    
    Range(ActiveCell, ActiveCell.Offset(39, 7)).Select
    
    Selection.RowHeight = 23
    
    Selection.ClearContents
    
    ActiveCell.Offset(0, 0).Select
        
End If
'---------------------------------------------------------

End Sub

Private Sub UserForm_Terminate()

UserForm1.Show

End Sub

Private Sub WX_Click()

hydrotools.correctLogRange

'--------------------------------------------
If hydrotools.isActiveWorkBook = True Then

    MsgBox "Please ensure the log workbook is active before continuing"
    
    Exit Sub
    
End If
'--------------------------------------------

Dim winds As Variant
Dim seas As Variant
Dim baroVal As Variant
Dim tempVal As Variant
Dim visVal As Variant

winds = windBox.Value
seas = seasBox.Value
baroVal = barometerBox.Value
tempVal = temperatureBox.Value
visVal = visBox.Value

ActiveCell.Value = Now()

ActiveCell.NumberFormat = "hh:mm"

ActiveCell.Offset(0, 7) = "Seas: " & seas & "ft" & "    winds: " & winds & "kts"
ActiveCell.Offset(1, 0).Select
ActiveCell.Offset(0, 7) = "Baro: " & baroVal & "mb" & "    temp: " & tempVal & "°F" & "    vis: " & visVal & "NM"

ActiveCell.Offset(1, 0).Select
'--------------------------------------------------------
If ActiveCell.Value = "#" Then

    ActiveCell.Offset(-49, -1).Select
    
    Range(ActiveCell, ActiveCell.Offset(49, 14)).Select
    
    Selection.Copy
    
    ActiveCell.Offset(50, 0).Select
    
    ActiveSheet.Paste
        
    ActiveCell.Offset(9, 1).Select
    
    Range(ActiveCell, ActiveCell.Offset(39, 7)).Select
    
    Selection.RowHeight = 23
    
    Selection.ClearContents
    
    ActiveCell.Offset(0, 0).Select
        
End If
'---------------------------------------------------------


End Sub

Private Sub comment_Click()

hydrotools.correctLogRange

'--------------------------------------------
If hydrotools.isActiveWorkBook = True Then

    MsgBox "Please ensure the log workbook is active before continuing"
    
    Exit Sub
    
End If
'--------------------------------------------

ActiveCell.Offset(0, 7) = InputBox("Please enter you comments", "Comment")

If ActiveCell.Offset(0, 7) = "" Then

    GoTo 222:
    
End If
    
ActiveCell.Value = Now()

ActiveCell.NumberFormat = "hh:mm"

ActiveCell.Offset(1, 0).Select

'--------------------------------------------------------
If ActiveCell.Value = "#" Then

    ActiveCell.Offset(-49, -1).Select
    
    Range(ActiveCell, ActiveCell.Offset(49, 14)).Select
    
    Selection.Copy
    
    ActiveCell.Offset(50, 0).Select
    
    ActiveSheet.Paste
        
    ActiveCell.Offset(9, 1).Select
    
    Range(ActiveCell, ActiveCell.Offset(39, 7)).Select
    
    Selection.RowHeight = 23
    
    Selection.ClearContents
    
    ActiveCell.Offset(0, 0).Select
        
End If
'---------------------------------------------------------

222:

End Sub

Private Sub ctdInWater_Click()

hydrotools.correctLogRange

'--------------------------------------------
If hydrotools.isActiveWorkBook = True Then

    MsgBox "Please ensure the log workbook is active before continuing"
    
    Exit Sub
    
End If
'--------------------------------------------

ActiveCell.Value = Now()

ActiveCell.NumberFormat = "hh:mm"

ActiveCell.Offset(0, 1) = "CTD"

ActiveCell.Offset(0, 7) = "CTD in water"

ActiveCell.Offset(1, 0).Select

End Sub

Private Sub CTD_Click()

hydrotools.correctLogRange

'--------------------------------------------
If hydrotools.isActiveWorkBook = True Then

    MsgBox "Please ensure the log workbook is active before continuing"
    
    Exit Sub
    
End If
'--------------------------------------------

Dim logctd As New Collection

ActiveCell.Value = Now()

ActiveCell.NumberFormat = "hh:mm"

ActiveCell.Offset(0, 1).Value = "CTD"

logctd.Add ctdFileName.Value & "    " & "WD = " & wDepth.Value & "m"

logctd.Add "X: " & xCoord.Value & "m" & "    " & "Y: " & yCoord.Value & "m"

logctd.Add "Lat: " & lat.Value & "°" & "    " & "Lon: " & lon.Value & "°"

logctd.Add "MB Depth: " & mbDepth.Value & "m" & "    " & "CTD Depth: " & ctdDepth.Value & "m"

logctd.Add "SB Depth: " & sbDepth.Value & "m"

logctd.Add "AML: " & amlSOS.Value & "m/s" & "    " & "HM: " & hmSOS.Value & "m/s"

logctd.Add "STBD Draft: " & sDraft.Value & "m" & "    " & "Port Draft: " & pDraft.Value & "m"

logctd.Add "Waterline/CRP: " & wlToCrp.Value & "m"

logctd.Add "Singlebeam Draft: " & sbDraft.Value & "m"

For p = 1 To 9

    If ActiveCell.Value = "#" Then
    
        ActiveCell.Offset(-49, -1).Select
    
        Range(ActiveCell, ActiveCell.Offset(49, 14)).Select
    
        Selection.Copy
    
        ActiveCell.Offset(50, 0).Select
    
        ActiveSheet.Paste
        
        ActiveCell.Offset(9, 1).Select
    
        Range(ActiveCell, ActiveCell.Offset(39, 7)).Select
    
        Selection.RowHeight = 23
    
        Selection.ClearContents
    
        ActiveCell.Offset(0, 0).Select
    
    End If
    
    ActiveCell.Offset(0, 7).Value = logctd(p)
    
    ActiveCell.Offset(1, 0).Select
    
Next p

End Sub


Private Sub LL_Click()

hydrotools.correctLogRange

'--------------------------------------------
If hydrotools.isActiveWorkBook = True Then

    MsgBox "Please ensure the log workbook is active before continuing"
    
    Exit Sub
    
End If
'--------------------------------------------

'---------------------------------------------------------
If ActiveCell.Value = "#" Then
    
    ActiveCell.Offset(-49, -1).Select
    
    Range(ActiveCell, ActiveCell.Offset(49, 14)).Select
    
    Selection.Copy
    
    ActiveCell.Offset(50, 0).Select
    
    ActiveSheet.Paste
        
    ActiveCell.Offset(9, 1).Select
  
    Range(ActiveCell, ActiveCell.Offset(39, 7)).Select
    
    Selection.RowHeight = 23
    
    Selection.ClearContents
    
    ActiveCell.Offset(0, 0).Select
    
End If
'---------------------------------------------------------

Dim logll As New Collection

ActiveCell.Value = Now()

ActiveCell.NumberFormat = "hh:mm"

ActiveCell.Offset(0, 1).Value = "Lead Line"

logll.Add llName.Value

logll.Add "MB Depth: " & llmbDepth.Value & "m"
 
logll.Add "SB Depth: " & llsbDepth.Value & "m"

logll.Add "LL Depth: " & llDepth.Value & "m"

For p = 1 To 4

    If ActiveCell.Value = "#" Then
    
        ActiveCell.Offset(-49, -1).Select
    
        Range(ActiveCell, ActiveCell.Offset(49, 14)).Select
    
        Selection.Copy
    
        ActiveCell.Offset(50, 0).Select
    
        ActiveSheet.Paste
        
        ActiveCell.Offset(9, 1).Select
    
        Range(ActiveCell, ActiveCell.Offset(39, 7)).Select
    
        Selection.RowHeight = 23
    
        Selection.ClearContents
    
        ActiveCell.Offset(0, 0).Select
    
    End If
    
    ActiveCell.Offset(0, 7).Value = logll(p)
    ActiveCell.Offset(1, 0).Select
    
Next p


End Sub

Private Sub save_Click()

'--------------------------------------------
If hydrotools.isActiveWorkBook = True Then

    MsgBox "Please ensure the log workbook is active before continuing"
    
    Exit Sub
    
End If
'--------------------------------------------

ActiveWorkbook.save

End Sub

Private Sub printSheet_Click()

'--------------------------------------------
If hydrotools.isActiveWorkBook = True Then

    MsgBox "Please ensure the log workbook is active before continuing"
    
    Exit Function
    
End If
'--------------------------------------------

ActiveWorkbook.PrintOut

End Sub

Private Sub loadSheet_Click()

Dim filename As String
    
filename = Application.GetOpenFilename(Title:="Please select a tide file", FileFilter:="Excel Files *.xlsm* (*.xlsm*),")

Workbooks.Open (filename)

MultiPage1.Enabled = True

While ActiveCell.Value <> "" And ActiveCell.Offset(0, 7).Value <> ""

    ActiveCell.Offset(1, 0).Select
    
Wend

End Sub

Private Sub closeLog_Click()

'--------------------------------------------
If hydrotools.isActiveWorkBook = True Then

    MsgBox "Please ensure the log workbook is active before continuing"
    
    Exit Sub
    
End If
'--------------------------------------------

ActiveWorkbook.Close

End Sub

Private Sub showCellEdit_Click()

cellEdit.Show

End Sub







