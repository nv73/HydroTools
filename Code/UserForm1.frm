VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "HydroTools"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10035
   OleObjectBlob   =   "UserForm1.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub UserForm_initialize()

'--------------------------------

'Written by: Nick Viner

'List of current bugs:

'If the saveAs application is cancelled, the program may crash crash (simple fix, but I am just lazy enough not to do it right now)
'It's gotten to the point where I cannot remember if I've fixed the above error but don't yet have the motivation to look into it.
'It doesn't seem like the error is present anymore, but I am going to leave this in the header just in case....

'Major Update log (as of 160912):

'150830: Forward/inverse now working as intended. Begun work on tide converter.
'150910: Tide converter now fully functional. Still working on error catching.
'150911: Added a simple distance/velocity/time calculator as well as a max ping calculator.
'        Fixed bug causing forward calculator to improperly calculate angles over 90 degrees.
'150912: Added an offset calculator for generating offsets for CARIS
'150913: Added a velocity converter.
'        Fixed rounding errors on the distance converter.
'151001: Fixed errors in the CARIS tide converter causing values not to properly paste and save.
'151003: Started working on a logging program extension.
'151004: Logging software work: SOL, EOL, Comment, New log, Weather, and CTD complete.
'151006: Finished leadline in autoLog. Coded a manual cell editor for autoLog.
'160327: Minor QOL fixes to CreateTides and sheet formatting. AutoLog bug fixes. Working on Mag data.
'160408: Improved on Mag outlier detection. Now uses changes in the slope to detect.
'160420: Added in a basic (and obscenely crude) calculator extension.
'160422: Added Miles to the unit conversion because I somehow missed that 9 months ago.
'160426: Begin creation of function library. Added a getListRange function and string manipulation. TideGen is fully operational & bug free!
'160426: Separated function library from userform to its own module "hydrotools". getRange, delcharbychar, delcharbyindex, getListLength, and random all working.
'160516: Fixed a number of small bugs within the Tide tool and Autologger. User manual is up to date.
'160528: Error catching in Autolog (includes two new functions).
'160908: Begun work on a more intensive mag cable out process.
'160912: Completed work on new cable out processing workflow. Searching for bugs.
'160919: Added in a tool to view and filter ASVP files (CTD).
'161122: Bug fixes in CTD cleaning.
'161221: Added ability to choose depth changes for filtering. Reset button. Extend button.

'--------------------------------
'The purpose of this userform is to allow the user to easily perform calculations and operations commonly used in the field of surveying.
'--------------------------------

'Initialize the global program buttons

'Sound speed variable to be used within all calculations of the application.

On Error GoTo sosErr:

Dim soundspeed As Double

speedOfSound.Value = 1500

soundspeed = speedOfSound.Value

'Create variables to be used to contain the values within the comboboxes. These will be arrays so a variant variable type is used.

Dim startUnit As Variant

Dim endUnit As Variant

Dim velUnitOne As Variant

Dim velUnitTwo As Variant

Dim monthArray As Variant

Dim dayArray As Variant

'Reset the array lengths to prevent blank spaces from appearing within the dropdown.

ReDim startUnit(8)

ReDim endUnit(8)

ReDim velUnitOne(4)

ReDim velUnitTwo(4)

ReDim monthArray(12)

ReDim dayArray(31)

'Assign values to the dropdown menu which the user can select.

startUnit = Array("Metres", "Centimetres", "Millimetres", "Survey Feet", "Feet", "Inches", "Nautical Miles", "Miles")

endUnit = Array("Metres", "Centimetres", "Millimetres", "Survey Feet", "Feet", "Inches", "Nautical Miles", "Miles")

velUnitOne = Array("Metres per Second", "Kilometres per Hour", "Miles per Hour", "Knots")

velUnitTwo = Array("Metres per Second", "Kilometres per Hour", "Miles per Hour", "Knots")

monthArray = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")

dayArray = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31)

'Set the number of columns within the drop-down menu.

startUnitsBox.ColumnCount = 1

finalUnitsBox.ColumnCount = 1

velOneBox.ColumnCount = 1

velTwoBox.ColumnCount = 1

monthBox.ColumnCount = 1

dayBox.ColumnCount = 1

'Assign easy to type variables to the two dropdowns for future reference.

startUnitsBox.List() = startUnit

finalUnitsBox.List() = endUnit

velOneBox.List() = velUnitOne

velTwoBox.List() = velUnitTwo

monthBox.List() = monthArray

dayBox.List() = dayArray

'------------------------------------------------------
'------------------------------------------------------
'Test Zone

'Testing delcharbyindex

Dim testIndex As Variant

ReDim testIndex(5)

testIndex = Array("1", "2", "3", "4", "5")

ComboBox1.ColumnCount = 1

ComboBox1.List() = testIndex

'------------------------------------------------------
'------------------------------------------------------
    
sosErr:
    
    If Err.Number = 13 Then
        MsgBox "Error: User entered a non-numeric or null value.", vbOKOnly, "Error"
        
    End If
    
End Sub

Private Sub helpButton_Click()

'Displays a small form showing the title of the program, name of programmer, and update time.
UserForm2.Show
    
End Sub

''''''''''''''''''''''''''''''''''''''''''
'
'Page 1: Unit Conversions and basic math
'
''''''''''''''''''''''''''''''''''''''''''

Private Sub cb1_Click()

'Variables to call the values in the dropdown (S1, s2), call the value the user entered (entry), ouput the final value (finalVal)

Dim S1 As String

Dim s2 As String

Dim entry As Variant

Dim finalVal As Double

Dim CF As Double

'-------- Unit conversions from metres -----------

Dim metrestoSurveyfeet As Double

Dim metrestoFeet As Double

Dim metrestoInches As Double

Dim metrestoNauticalmiles As Double

metrestoSurveyfeet = 3.28083333333

metrestoFeet = 3.28084

metrestoInches = 39.3701

metrestoNauticalmiles = 0.000539957

metrestoMiles = 0.000621371

'-------------------------------------------------

'Assign the active values within the comboboxes to variables S1 and s2

S1 = startUnitsBox.Value

s2 = finalUnitsBox.Value

entry = inputVal.Value

'Prevent the user from typing in any non-numeric value into the entry textbox.

On Error GoTo convertErr:
  
'Convert the initial value into metres

If S1 = "Metres" Then

    CF = 1
    Else
    If S1 = "Centimetres" Then
    
        CF = 0.01
        Else
        If S1 = "Millimetres" Then
        
            CF = 0.001
            Else
            If S1 = "Survey Feet" Then
            
                CF = 0.3048006
                Else
                If S1 = "Feet" Then
                
                    CF = 0.3048
                    Else
                    If S1 = "Inches" Then
                    
                        CF = 0.0254
                        Else
                        If S1 = "Nautical Miles" Then
                        
                            CF = 1852
                            Else
                            If S1 = "Miles" Then
                            
                                CF = 1609.34
                                
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End If

'Convert the metres value to the desired unit type

If s2 = "Metres" Then

    finalVal = (entry * CF)
    Else
    If s2 = "Centimetres" Then
    
        finalVal = (entry * CF) * 100
        Else
        If s2 = "Millimetres" Then
        
            finalVal = (entry * CF) * 1000
            Else
            If s2 = "Survey Feet" Then
            
                finalVal = (entry * CF) * metrestoSurveyfeet
                Else
                If s2 = "Feet" Then
                
                    finalVal = (entry * CF) * metrestoFeet
                    Else
                    If s2 = "Inches" Then
                    
                        finalVal = (entry * CF) * metrestoInches
                        Else
                        If s2 = "Nautical Miles" Then
                        
                            finalVal = (entry * CF) * metrestoNauticalmiles
                            Else
                            If s2 = "Miles" Then
                            
                                finalVal = (entry * CF) * metrestoMiles
                            
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End If
    
'Display the final results within a textbox so that it can be copy and pasted into a separate word document.

finalVal = Round(finalVal, 5)

testbox.Value = finalVal

Me.Repaint
        
convertErr:

    If Err.Number = 13 Then
    
        MsgBox "Error: User entered a non-numeric or null value.", vbOKOnly, "Error"
        
    End If
        
End Sub

Private Sub CommandButton2_Click()

Dim inputVal As Double
Dim resultVal As Double
Dim convertVal1 As String
Dim convertVal2 As String
Dim convertFactor As Double
Dim v1 As String
Dim v2 As String

On Error GoTo conversionErr:

inputVal = inputVel.Value
v1 = velOneBox.Value
v2 = velTwoBox.Value

Dim kmphToMps As Double
Dim mphToMps As Double
Dim knotToMps As Double

kmphToMps = 0.27778
mphToMps = 0.44704
knotToMps = 0.51444

If v1 = "Metres per Second" Then
    
    convertFactor = 1
    Else
    If v1 = "Kilometres per Hour" Then
        
        convertFactor = kmphToMps
        Else
        If v1 = "Miles per Hour" Then
            
            convertFactor = mphToMps
            Else
            If v1 = "Knots" Then
            
                convertFactor = knotToMps
            End If
        End If
    End If
End If

If v2 = "Metres per Second" Then
    
    resultVal = inputVal * convertFactor
    Else
    If v2 = "Kilometres per Hour" Then
        
        resultVal = (inputVal * convertFactor) * 3.6
        Else
        If v2 = "Miles per Hour" Then
            
            resultVal = (inputVal * convertFactor) * 2.23694
            Else
            If v2 = "Knots" Then
            
                resultVal = (inputVal * convertFactor) * 1.94386
                
            End If
        End If
    End If
End If

resultVal = Round(resultVal, 5)

outputVel.Value = resultVal

conversionErr:

    If Err.Number = 13 Then
    
        MsgBox "Error: User entered a non-numeric or null value.", vbOKOnly, "Error"
        
    End If

End Sub

Private Sub timeButton_Click()

'Allows the user to solve for time.

    timeBox.Enabled = False
    distBox.Enabled = True
    velBox.Enabled = True
    timeBox.BackColor = &H80000004
    distBox.BackColor = &H80000005
    velBox.BackColor = &H80000005

End Sub

Private Sub distButton_Click()

'Allows the user to solve for distance.

    timeBox.Enabled = True
    distBox.Enabled = False
    velBox.Enabled = True
    timeBox.BackColor = &H80000005
    distBox.BackColor = &H80000004
    velBox.BackColor = &H80000005

End Sub

Private Sub velButton_Click()

'Allows the user to solve for velocity.

    timeBox.Enabled = True
    distBox.Enabled = True
    velBox.Enabled = False
    timeBox.BackColor = &H80000005
    distBox.BackColor = &H80000005
    velBox.BackColor = &H80000004

End Sub

Private Sub solveVal_Click()

On Error GoTo dvtError: 'Error catching

If timeBox.Enabled = False Then 'Solves for time.

    timeBox.Value = distBox.Value / velBox.Value
    
    Else
    
    If distBox.Enabled = False Then 'Solves for distance.
    
        distBox.Value = velBox.Value * timeBox.Value
        
        Else
        
        If velBox.Enabled = False Then 'Solves for velocity.
        
            velBox.Value = distBox.Value / timeBox.Value
            
        End If
        
    End If
    
End If

dvtError:

    If Err.Number = 13 Then
    
        MsgBox "Error: User entered a non-numeric or null value.", vbOKOnly, "Error"
        
    End If
    
End Sub

Private Sub usePing_Click()

'Allows the user to choose whether or not to limit their max pings by ping-rate.

If usePing.Value = True Then
    
    pingRate.Enabled = True
    pingRate.BackColor = &H80000005
    depth.Enabled = False
    depth.BackColor = &H80000004
    
    Else
        
    pingRate.Enabled = False
    pingRate.BackColor = &H80000004
    depth.Enabled = True
    depth.BackColor = &H80000005
    
End If

End Sub


Private Sub solveButton_Click()

'Solves for either the maximum pings per metre given a depth and speed, or solves for pings per metre given a speed and ping rate.

Dim twtt As Double
Dim boatDist As Double
Dim boatMetres As Double

On Error GoTo solveErr:

soundspeed = speedOfSound.Value
    
    If usePing.Value = False Then
    
        twtt = (depth.Value / soundspeed) * 2    ' seconds taken for a ping to reach the seafloor and return.
        
        boatDist = twtt * (speed.Value * 0.514444)   ' Metres the boat travels during twtt.

        boatMetres = 1 / boatDist     ' How many pings can occur within a metre of distance travelled.
        
        resultPing.Caption = "Maximum Possible Pings per Metre"
        
        coverage.Value = boatMetres
    
    Else
    
        If usePing.Value = True Then
            
           resultPing.Caption = "Pings per Metre"
           
           coverage.Value = pingRate.Value / (speed.Value * 0.514444)
        
        End If
    
    End If
    
solveErr:

    If Err.Number = 13 Then
    
        MsgBox "Error: User entered a non-numeric or null value.", vbOKOnly, "Error"
        
    End If
    
End Sub

Private Sub CommandButton3_Click()

'Calculates the Julian date provided a normal YYYY/MM/DD date.

Dim monthSum As Integer
Dim dayFebruary As Integer

If yearBox.Value < 0 Then

    GoTo 909:
    
End If

If IsNumeric(yearBox.Value) = False Then
    
    GoTo 909:
    
End If

If monthBox.Value = "February" And dayBox.Value > 29 Then
    MsgBox "Please select a valid day for the month of February", , "Date Error"
    dayBox.Value = 1
    GoTo 909:
End If

If yearBox.Value Mod 4 = 0 Then
    dayFebruary = 29
    Else
    dayFebruary = 28
End If

If monthBox.Value = "January" Then
    monthSum = 0
    Else
    If monthBox.Value = "February" Then
        monthSum = 31
        Else
        If monthBox.Value = "March" Then
            monthSum = 31 + dayFebruary
            Else
            If monthBox.Value = "April" Then
                monthSum = 62 + dayFebruary
                Else
                If monthBox.Value = "May" Then
                    monthSum = 92 + dayFebruary
                    Else
                    If monthBox.Value = "June" Then
                        monthSum = 123 + dayFebruary
                        Else
                        If monthBox.Value = "July" Then
                            monthSum = 153 + dayFebruary
                            Else
                            If monthBox.Value = "August" Then
                                monthSum = 184 + dayFebruary
                                Else
                                If monthBox.Value = "September" Then
                                    monthSum = 215 + dayFebruary
                                    Else
                                    If monthBox.Value = "October" Then
                                        monthSum = 245 + dayFebruary
                                        Else
                                        If monthBox.Value = "November" Then
                                            monthSum = 271 + dayFebruary
                                            Else
                                            If monthBox.Value = "December" Then
                                                monthSum = 301 + dayFebruary
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

julianDate.Value = dayBox.Value + monthSum

909:

End Sub

'''''''''''''''''''''''''''''''''''
'
'Page 2: Basic vector calculations
'
'''''''''''''''''''''''''''''''''''

Private Sub inverseButton_Click()

    n1.Enabled = True
    e1.Enabled = True
    a2.Enabled = True
    d2.Enabled = True
    n2.Enabled = False
    e2.Enabled = False
    n2.BackColor = &H8000000F
    e2.BackColor = &H8000000F
    a2.BackColor = &H80000005
    d2.BackColor = &H80000005
    
End Sub

Private Sub fwdButton_Click()

    a2.Enabled = False
    d2.Enabled = False
    a2.BackColor = &H8000000F
    d2.BackColor = &H8000000F
    n2.BackColor = &H80000005
    e2.BackColor = &H80000005
    n1.Enabled = True
    e1.Enabled = True
    n2.Enabled = True
    e2.Enabled = True
    
End Sub

Private Sub cancelForm_Click()
    
    Workbooks("hydrotools_active.xlsb").Activate
    Unload Me
    'ThisWorkbook.Close savechanges:=False
    Worksheets("Sheet1").Activate
    
End Sub

Private Sub resetForm_Click()

    n1.Value = ""
    e1.Value = ""
    n2.Value = ""
    e2.Value = ""
    a2.Value = ""
    d2.Value = ""

End Sub

Private Sub calculateResult_Click()

'VARIABLE ASSIGNMENT
'------------------------------------------------------------
'Create variables to be calculated
Dim northing1 As Variant
Dim easting1 As Variant
Dim northing2 As Variant
Dim easting2 As Variant
Dim azimuth2 As Variant
Dim distance2 As Variant

'Create variables to output final results
Dim northResult As Double
Dim eastResult As Double
Dim distResult As Double
Dim azResult As Double

'Create variables which will help to calculate results
Dim deltaE As Double
Dim deltaN As Double
Dim angle As Double
Dim inverseEN As Double
Dim angleAd As Double
Dim toDeg As Double

'Assign the inputted values to the created variables
northing1 = n1.Value
easting1 = e1.Value
northing2 = n2.Value
easting2 = e2.Value
azimuth2 = a2.Value
distance2 = d2.Value
toDeg = (WorksheetFunction.Pi / 180)

'END VARIABLE ASSIGNMENT
'-------------------------------------------------------------

'Catch any errors which occur in the program
On Error GoTo errorCatch:

'Forward calculation
If a2.Enabled = True Then

    If azimuth2 <= 90 Then
        angle = azimuth2
        Else
        If azimuth2 <= 180 Then
            angle = azimuth2 - 90
            Else
            If azimuth2 <= 270 Then
              angle = azimuth2 - 180
              Else
              If azimuth2 <= 360 Then
                  angle = azimuth2 - 270
              End If
           End If
       End If
    End If

    deltaE = (distance2 * (Cos(angle * (WorksheetFunction.Pi / 180))))
    deltaN = (distance2 * (Sin(angle * (WorksheetFunction.Pi / 180))))

    northResult = northing1 + deltaN
    eastResult = easting1 + deltaE

    northFinal.Value = northResult
    eastFinal.Value = eastResult
    
End If

'Inverse calculation
If n2.Enabled = True Then
    
    deltaE = easting2 - easting1
    deltaN = northing2 - northing1
    
    inverseEN = Sqr(((deltaE) ^ 2) + ((deltaN) ^ 2))
    
    If northing2 > northing1 And easting2 > easting1 Then
        angleAd = 0
        Else
        If northing2 < northing1 And easting2 > easting1 Then
            angleAd = 90
            Else
            If northing2 > northing1 And easting2 < easting1 Then
                angleAd = 270
                Else
                If northing2 < northing1 And easting2 < easting1 Then
                    angleAd = 180
                End If
            End If
        End If
    End If

    angle = ((Atn((deltaN / deltaE))) * 57.2957795130823) + angleAd

    distFinal.Value = inverseEN
    azFinal.Value = angle
    
End If
    
'Error catching message
errorCatch:
    
    If Err.Number = 13 Then
        MsgBox "Error: User entered a non-numeric or null value.", vbOKOnly, "Error"
        
    End If
    
End Sub

'''''''''''''''''''''''''''''''''''''''
'
'Page 3: CARIS Tide file generator
'
'''''''''''''''''''''''''''''''''''''''

Private Sub loadFile_Click()

On Error GoTo openErr:
    
Dim filename As String
    
    filename = Application.GetOpenFilename(Title:="Please select a tide file", FileFilter:="Tide Files *.tid* (*.tid*),")
    
    filePathBox.Value = filename
    
    Workbooks.OpenText filename:=filename, DataType:=xlDelimited, Tab:=True, FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1))
    
    updateTide.Enabled = True
    
openErr:

    If Err.Number = 1004 Then
    
    filePathBox.Value = ""
    GoTo 9:
    
    End If
    
9:
    
End Sub

Private Sub createTide_Click()

Dim newBook As Workbook

    Set newBook = Workbooks.Add

    Application.DisplayAlerts = False

    Range("A1").Value = "-----------------------"

        Do
            fName = Application.GetSaveAsFilename(tide, FileFilter:="Tide Files *.tid* (*.tid*),")
    
            Loop Until fName <> False
    
    newBook.SaveAs filename:=fName + "tid", FileFormat:=xlTextWindows
       
    newBook.Close

End Sub

Private Sub CommandButton1_Click()
    
Dim stationVal As Variant
        
    stationVal = InputBox("Please enter the station number: ")
            
    ActiveWorkbook.FollowHyperlink Address:="http://tidesandcurrents.noaa.gov/waterlevels.html?id=" & stationVal

End Sub
Private Sub updateTide_Click()

Dim i As Integer
Dim c As Integer
Dim messageAnswer As Integer
Dim stationID As String
Dim newColCount As Integer

i = 2
c = 1

On Error GoTo 2:

messageAnswer = MsgBox("Have you copied the desired tide data from the NOAA website?", vbYesNo, "Update Tides")

    If messageAnswer = vbYes Then
    
        GoTo 420:
        
        Else
        
        If messageAnswer = vbNo Then
            
            MsgBox ("Once you have copied the desired values, please click 'Update Tides' again.")
            GoTo 3:
            
        End If
        
    End If
        
420:

Range("A2").Select
    
While ActiveCell.Value <> ""
    
   ActiveCell.Offset(1, 0).Select
   
   i = i + 1
   
Wend

Range("A" & i).Select

ActiveSheet.Paste

newColCount = i

Range("A" & i).Select

While ActiveCell.Value <> ""

    Cells(newColCount, 1).Select
    newColCount = newColCount + 1
    
Wend

Range("A" & i).Select

If prelim.Value = True Then

    Range(Cells(i, 5), Cells(newColCount, 5)).Select
    Selection.ClearContents
    Range(Cells(i, 3), Cells(newColCount, 3)).Select
    Selection.ClearContents
    Selection.Delete Shift:=xlToLeft
    Columns("A:A").Select
    Selection.NumberFormat = "yyyy/mm/dd"
    'Range(Cells(i, 2), Cells(newColCount, 2)).Select
    Columns("B:B").Select
    Selection.NumberFormat = "hh:mm"
    
    Else
    
    Range(Cells(i, 3), Cells(newColCount, 3)).Select
    Selection.ClearContents
    Selection.Delete Shift:=xlToLeft
    Range(Cells(i, 4), Cells(newColCount, 4)).Select
    Selection.ClearContents
    Selection.Delete Shift:=xlToLeft
    Columns("A:A").Select
    Selection.NumberFormat = "yyyy/mm/dd"
    'Range(Cells(i, 2), Cells(newColCount, 2)).Select
    Columns("B:B").Select
    Selection.NumberFormat = "hh:mm"
    
End If
        
ActiveWorkbook.SaveAs FileFormat:=xlTextWindows
    
ActiveWorkbook.Close

2:

   If Err.Number = 1004 Then
   
       GoTo 3:
   
   End If
    
3:

End Sub

'''''''''''''''''''''''''''''''''''''''
'
'Page 4: Offset calculator
'
'''''''''''''''''''''''''''''''''''''''

Private Sub calcOffsets_Click()
'A basic vessel offset calculator to be used when entering transducer offsets into CARIS.

On Error GoTo offsetErr:

mtx.Value = ctx.Value - cmx.Value
mty.Value = cty.Value - cmy.Value
mtz.Value = (ctz.Value - cmz.Value)

ntx.Value = ctx.Value - cnx.Value
nty.Value = cty.Value - cny.Value
ntz.Value = (ctz.Value - cnz.Value)

offsetErr:

    If Err.Number = 13 Then
    
        MsgBox "Error: User entered a non-numeric or null value.", vbOKOnly, "Error"
        
    End If
    
End Sub

Private Sub UserForm_Terminate()

    Application.Visible = True
    
    Worksheets("Sheet1").Activate
    
End Sub
Private Sub CommandButton4_Click()

Sheets("Sheet1").Activate

AutoLog.Show

End Sub

Private Sub startCalc_Click()

Calculator.Show

Worksheets("Sheet1").Activate

End Sub

'''''''''''''''''''''''''''''''''''''''''''''
'
'Page 5: G-880 Magnetometer Data processing.
'
'''''''''''''''''''''''''''''''''''''''''''''

Private Sub loadMag_Click()

On Error GoTo openErr2:

'------------------------------------------------------------------------------------------------------------------------
'Load the MAG file and format it properly to be examined

Dim filename As String
    
    filename = Application.GetOpenFilename(Title:="Please select a mag file", FileFilter:="Mag Files *.mag* (*.mag*),")
    
    Workbooks.OpenText (filename)
    
Columns("A:A").Select
Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
    Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo _
    :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), _
    Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1)), _
    TrailingMinusNumbers:=True
    
'get length/width of table-------------------

Application.ScreenUpdating = False

Columns("A:P").Select
Selection.NumberFormat = "0.000"
    
Dim cwidth As Integer
Dim clength As Integer

cwidth = 0
clength = 0

Range("A1").Select

While ActiveCell.Value <> ""

    ActiveCell.Offset(0, 1).Select
    
    cwidth = cwidth + 1
    
Wend

Range("A1").Select

While ActiveCell.Value <> ""

    ActiveCell.Offset(1, 0).Select
    clength = clength + 1
    
Wend

'End get length/width-----------------------

'generate collections---------------------

Dim fixnum As New Collection
Dim timeVal As New Collection
Dim dateVal As New Collection
Dim lat As New Collection
Dim lon As New Collection
Dim easting As New Collection
Dim northing As New Collection
Dim gamma As New Collection
Dim altitude As New Collection
Dim wd As New Collection
Dim signalStrength As New Collection

'End generating collections--------------

'Populate collections--------------------

Dim rowCount As Integer

'-----------------------------------
'Fixes

Range("A1").Select

For rowCount = 1 To clength

    fixnum.Add (ActiveCell.Value)
    ActiveCell.Offset(1, 0).Select
    
Next rowCount

Range("B1").Select
    
'-----------------------------------
'Times

For rowCount = 1 To clength

    timeVal.Add (ActiveCell.Value)
    ActiveCell.Offset(1, 0).Select
    
Next rowCount

Range("C1").Select

'-----------------------------------
'Dates

For rowCount = 1 To clength

    dateVal.Add (ActiveCell.Value)
    ActiveCell.Offset(1, 0).Select
    
Next rowCount

Range("D1").Select
    
'-----------------------------------
'Latitude

For rowCount = 1 To clength

    lat.Add (ActiveCell.Value)
    ActiveCell.Offset(1, 0).Select
    
Next rowCount

Range("E1").Select

'-----------------------------------
'Longitude

For rowCount = 1 To clength

    lon.Add (ActiveCell.Value)
    ActiveCell.Offset(1, 0).Select
    
Next rowCount

Range("F1").Select

'-----------------------------------
'Easting (X)

For rowCount = 1 To clength

    easting.Add (ActiveCell.Value)
    ActiveCell.Offset(1, 0).Select
    
Next rowCount

Range("G1").Select

'-----------------------------------
'Northing (Y)

For rowCount = 1 To clength

    northing.Add (ActiveCell.Value)
    ActiveCell.Offset(1, 0).Select
    
Next rowCount

Range("H1").Select

'-----------------------------------
'Gamma (nT)

For rowCount = 1 To clength

    gamma.Add (ActiveCell.Value)
    ActiveCell.Offset(1, 0).Select
    
Next rowCount

Range("I1").Select

'-----------------------------------
' Altitude

For rowCount = 1 To clength

    altitude.Add (ActiveCell.Value)
    ActiveCell.Offset(1, 0).Select
    
Next rowCount

Range("J1").Select

'-----------------------------------
'Water Depth

For rowCount = 1 To clength

    wd.Add (ActiveCell.Value)
    ActiveCell.Offset(1, 0).Select
    
Next rowCount

Range("K1").Select

'-----------------------------------
'Signal Strength

For rowCount = 1 To clength

    signalStrength.Add (ActiveCell.Value)
    ActiveCell.Offset(1, 0).Select
    
Next rowCount

Application.ScreenUpdating = True

'-----------------------------------

ActiveWorkbook.Saved = True

ActiveWorkbook.Close

'--------------------------------
'Total number of data points

pointNumBox.Value = gamma.Count

Application.ScreenUpdating = False

'----------------------------------------------------------------------------
'Extract the number of fixpoints and the start/end times from the Mag data.

Dim maxFix As Integer
Dim startTime As Double
Dim endTime As Double

Dim p As Integer

maxFix = 0

For p = 1 To clength
           
    If fixnum(p) >= maxFix Then
        
        maxFix = fixnum(p)
        
    End If
    
Next p

eventNumBox.Value = maxFix
startTimeBox.Value = Format(timeVal(1), "HH:MM:SS")
endTimeBox.Value = Format(timeVal(clength), "HH:MM:SS")

'--------------------------------------------------------------
'Calculate the mean gamma value

Dim meanGamma As Double
Dim n As Double

n = 0

For p = 1 To clength

    n = n + gamma(p)

Next p

meanGamma = n / (clength)

'--------------------------------------------------------------
'Find Standard Deviation of the dataset

Dim meanDiff As Double
Dim gammaStdDev As Double

n = 0

For p = 1 To clength

    n = n + ((gamma(p) - meanGamma) ^ 2)
            
Next p

meanDiff = n

meanDiff = meanDiff / (clength - 1)

gammaStdDev = meanDiff ^ 0.5

'------------------------------------------------------------------
'Utilize slope values to determine the presence of outliers in the dataset.

Dim setX As New Collection
Dim setSlope As New Collection
Dim setSum As Double
Dim setMean As Double

'Generate an x axis for which the gamma values can be referenced.

For x = 1 To clength

    setX.Add x
    
Next x

'Append slope values throughout the data set into a collection.

For f = 1 To (clength - 1)

    slope = Abs((gamma(f + 1) - gamma(f)) / (setX(f + 1) - setX(f)))
    
    setSlope.Add slope
    
Next f

'Determine the mean slope value

For getSum = 1 To (clength - 1)

    setSum = Abs(setSum) + Abs(setSlope(getSum))

Next getSum

setMean = setSum / clength - 1

'Determine the standard deviation (p) of the slopes

Dim meanDiffSD As Double
Dim stddev As Double
Dim sigmaV As Integer

n = 0

For p = 1 To clength - 1

    n = n + (setSlope(p) - setMean) ^ 2
    
Next p

meanDiffSD = n / clength - 1

stddev = meanDiff ^ 0.5

'Populate the GUI with information on the potential outliers

Dim sigmaCollGamma As New Collection
Dim sigmaCollTime As New Collection
Dim sigmaCollFix As New Collection
Dim sigmaCount As Integer

sigmaCount = 1

For Z = 1 To clength - 1

    If setSlope(Z) > setMean + (2 * stddev) Then
        
        sigmaCollGamma.Add gamma(Z)
        sigmaCollTime.Add timeVal(Z)
        sigmaCollFix.Add fixnum(Z)
        sigmaCount = sigmaCount + 1
        
    End If
    
Next Z

For Z = sigmaCollFix.Count To 2 Step -1

    If sigmaCollFix(Z) = sigmaCollFix(Z - 1) Then
        
        sigmaCollFix.Remove (Z)
        sigmaCollTime.Remove (Z)
        sigmaCollGamma.Remove (Z)
                            
    End If

Next Z

Dim fixRowLength As Integer

For Z = 1 To sigmaCollFix.Count

    With outlierBox
    
        outlierBox.ColumnCount = 3
        
        outlierBox.AddItem
        
        outlierBox.ColumnWidths = "65;30;30"
        
        outlierBox.List(Z - 1, 0) = Format(sigmaCollTime(Z), "HH:MM:SS")
        outlierBox.List(Z - 1, 1) = sigmaCollFix(Z)
        outlierBox.List(Z - 1, 2) = sigmaCollGamma(Z)
    
    End With

Next Z

SigmaBox.Value = sigmaCollFix.Count

'-------------------------------

Workbooks("HydroTools_Active.xlsb").Activate

Worksheets("gammaValues").Activate

For p = 1 To clength

    Cells(p, 1).Value = gamma(p)
    Cells(p, 2).Value = Format(fixnum(p), 0)
    
Next p

Chart13.Activate

Application.ScreenUpdating = True

saveChart.Enabled = True

'-------------------------------

openErr2:

    If Err.Number = 1004 Then
    
    filePathBox.Value = ""
    MsgBox "OH NO! AN ERROR OCCURED! Luckily this messagebox was here to catch it!"
    GoTo 10:
    
    End If
    
9:

    MsgBox "Mag data successfully loaded"

10:

End Sub

Private Sub saveChart_Click()

Dim filePath As String

Chart13.Activate

ActiveSheet.Copy

filePath = Application.GetSaveAsFilename(, FileFilter:="XLSX Files (*.xlsx), *.xlsx")

ActiveWorkbook.SaveAs filePath

ActiveWorkbook.Saved = True

ActiveWorkbook.Close

Workbooks("HydroTools_Active.xlsb").Activate

End Sub

Private Sub CommandButton5_Click()

On Error GoTo errCatch:

'Allows user to load the desired file which needs converting
Dim loadWindow

'Name of currently opened file. Not Necessary but I'll keep it for now
Dim currentFile As String

'Name of new file.
Dim newFile As String

'Creates the .lb filetype
Dim newFileType As String

'Default filename which appears in the save window.
Dim newFileName As String

'Variable used to store the active workbook.
Dim actBook As Workbook

'Variable used to store the active worksheet.
Dim actSheet As Worksheet

'Open desired file.--------------------------------------------------------------------------------

loadWindow = Application.GetOpenFilename(Title:="MAG CO FILE", FileFilter:="Mag Cable Out Files *.co* (*.co*),")

Workbooks.Open (loadWindow)

'Delimit the imported data.------------------------------------------------------------------------
    
Application.ScreenUpdating = False
    
Columns("A:A").Select

Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
    Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
    :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1)), _
    TrailingMinusNumbers:=True
        
Columns("A:A").EntireColumn.AutoFit

'Delete necessary columns.------------------------------------------------------------------------

Range("A:A").Delete Shift:=xlToLeft
Range("A:A").Delete Shift:=xlToLeft

Columns("C").NumberFormat = "0.000"
Columns("D").NumberFormat = "0.000"

'Change depth column to scale factor.--------------------------------------------------------------

Cells(1, 4).Select

Dim scaleFactor As Double

If sfCheckBox = True Then

    scaleFactor = InputBox("Enter the scale factor: ", "Scale Factor")
    
    While ActiveCell.Value <> ""

        ActiveCell.Value = ActiveCell.Offset(0, -1).Value * scaleFactor
        ActiveCell.Offset(1, 0).Select
    
    Wend
    
End If

'Select and Export by line.------------------------------------------------------------------------

Dim rowCount As Integer
Dim dayCount As Integer
Dim filename As String
Dim bookName As String
Dim sheetName As String
Dim dayVal As Variant

dayCount = 1
rowCount = 1

Range("A1").Select

sheetName = ActiveSheet.Name

While ActiveCell.Value <> ""

    If ActiveCell.Offset(1, 0).Value = ActiveCell.Value Then
    
        ActiveCell.Offset(1, 0).Select
        
        rowCount = rowCount + 1
        
        Else
            
            If sfCheckBox = False Then
            
                Range("D1").Select
            
                scaleFactor = InputBox("Input Scale Factor for line " & ActiveCell.Offset(0, -3).Value, "Scale Factor", "1.0")
            
                For j = 1 To rowCount
            
                    ActiveCell.Value = ActiveCell.Offset(0, -1).Value * scaleFactor
                    ActiveCell.Offset(1, 0).Select
                
                Next j
                
            End If
            
            Range("A1" & ":" & "F" & rowCount).Copy
                      
            Sheets.Add.Name = "mag_data" & dayCount
            
            Sheets("mag_data" & dayCount).Activate
            dayCount = dayCount + 1
            Range("A1").Select
            ActiveSheet.Paste
            Application.CutCopyMode = False
            
            bookName = ActiveWorkbook.Name
            
            ActiveSheet.Copy
            dayVal = Range("A1").Value
            Range("A:A").Delete Shift:=xlToLeft
            
            filename = Application.GetSaveAsFilename(dayVal, FileFilter:="Text Files (*.lb), *.lb")
            
            On Error Resume Next
            
            ActiveWorkbook.SaveAs filename, xlTextWindows
            
            ActiveWorkbook.Saved = True
            
            ActiveWorkbook.Close
            
            Workbooks(bookName).Activate
                        
            Sheets(sheetName).Activate
            
            Rows("1:" & rowCount).Select
            Rows("1:" & rowCount).Delete Shift:=xlUp
            
            Range("A1").Select
            rowCount = 1
            
            ActiveWorkbook.Saved = True
                                               
    End If

Wend

ActiveWorkbook.Close

Workbooks("HydroTools_Active.xlsb").Activate

'Restore screen refresh.-------------------------------------------------------------------------
Application.ScreenUpdating = True

errCatch:
    
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

Private Sub CommandButton7_Click()

Sheets("Sheet1").Activate

MultiPage1.Enabled = True

End Sub

Private Sub CommandButton8_Click()

MagCO.Show

End Sub

'''''''''''''''''''''''''''''''''''''''''''''
'
'Page 6: CTD Data Processing.
'
'''''''''''''''''''''''''''''''''''''''''''''

Private Sub cnvtoasvp_Click()

Dim myfile As String

Dim storeInfo As Variant

Dim textline As String

Dim linenum As Long

Dim depth As New Collection

Dim soundspeed As New Collection

linenum = 0

myfile = Application.GetOpenFilename()

Open myfile For Input As #1

Do Until EOF(1)

    linenum = linenum + 1
    Line Input #1, textline
    storeInfo = hydrotools.splitString(textline, "  ")
    
    If linenum > 149 Then
        
        depth.Add (storeInfo(1))
        soundspeed.Add (storeInfo(5))
        
    End If
    
Loop

Close #1

mynewfile = Application.GetSaveAsFilename & ".asvp"

Open mynewfile For Output As #1

For x = 1 To depth.Count - 1

    Print #1, depth(x) & vbTab & soundspeed(x)

Next x

Close #1

End Sub


Private Sub loadAsvp_Click()

Dim myfile As String

Dim storeInfo As Variant

Dim textline As String

Dim linenum As Long

Dim depth As New Collection

Dim soundspeed As New Collection

Dim i As Integer

i = 0
linenum = 0

myfile = Application.GetOpenFilename()

On Error GoTo loadError:

Open myfile For Input As #1

Do Until EOF(1)

    linenum = linenum + 1
    Line Input #1, textline
    storeInfo = hydrotools.splitString(textline, vbTab)
    
    If linenum > 1 Then
                
        depth.Add (storeInfo(0))
        soundspeed.Add (storeInfo(1))
        
    End If
    
Loop

Close #1

For Z = 1 To depth.Count

    With ctdInfo
    
        ctdInfo.ColumnCount = 2
        
        ctdInfo.AddItem
        
        ctdInfo.ColumnWidths = "60;30"
        
        ctdInfo.List(Z - 1, 0) = depth(Z)
        ctdInfo.List(Z - 1, 1) = soundspeed(Z)
    
    End With

Next Z

Dim harmonicMean As Double
Dim a As Double

For n = 1 To soundspeed.Count

    a = a + 1 / soundspeed(n)

Next n

harmonicMean = soundspeed.Count / a
hmean.Value = harmonicMean

'Determine the maximum CTD depth
Dim c As Integer
c = 0

While depth(depth.Count - c) Mod 100 = 0

    c = c + 1
    
Wend

ctdDepth.Value = depth(depth.Count - c)


resetList.Enabled = True
extendasvp.Enabled = True
saveAsvp.Enabled = True
loadAsvp.Enabled = False

loadError:

If Err.Number = 53 Then

    MsgBox ("Invalid File")
    
End If

End Sub

Private Sub filterAsvp_Click()

Dim depth As New Collection
Dim soundspeed As New Collection
Dim ctdArray As Variant
Dim depthFilter As Double

On Error GoTo continue:

depthFilter = InputBox("Enter in a delta-depth value for filtering: ", "Depth filter", 0.1)

For x = (ctdInfo.ListCount - 2) To 1 Step -1

    If Abs(ctdInfo.List(x, 0) - ctdInfo.List(x + 1, 0)) < depthFilter Then
    
        ctdInfo.RemoveItem (x + 1)
            
    End If
        
Next x

continue:

End Sub

Private Sub extendasvp_Click()

Dim newDepth As Double

filterAsvp.Enabled = True
newDepth = 123

While newDepth Mod 100 <> 0

    newDepth = InputBox("Depth to extend to (NOTE: value must be a multiple of 100) ", "Extend", 12000)

Wend

ctdInfo.List(ctdInfo.ListCount - 1, 0) = newDepth
ctdInfo.List(ctdInfo.ListCount - 1, 1) = ctdInfo.List(ctdInfo.ListCount - 2, 1)

End Sub


Private Sub saveAsvp_Click()

Dim myfile As String

Dim storevalue As Variant
Dim i As Integer
Dim j As Integer

On Error GoTo saveError:

myfile = Application.GetSaveAsFilename & ".asvp"

Open myfile For Output As #1

For x = 0 To ctdInfo.ListCount - 1

    Print #1, ctdInfo.List(x, 0) & vbTab & ctdInfo.List(x, 1)

Next x

Close #1

saveError:

    If Err.Number = 53 Then
    
        MsgBox ("No file name entered.")
        
    End If

End Sub

Private Sub resetList_Click()

ctdInfo.Clear
ctdDepth.Value = ""
hmean.Value = ""

extendasvp.Enabled = False
filterAsvp.Enabled = False
saveAsvp.Enabled = False
loadAsvp.Enabled = True

End Sub


'----------------------------------------------------------------------
'
'Test Subs
'
'----------------------------------------------------------------------

Private Sub CommandButton6_Click()

'Testing sub for delcharbyindex

Dim testString As String

testString = InputBox("Enter string: ")

TextBox1.Value = deleteCharacterByIndex(testString, ComboBox1.Value)

End Sub

