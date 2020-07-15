VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LinePlanner 
   Caption         =   "UserForm3"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3870
   OleObjectBlob   =   "LinePlanner.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LinePlanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub radioleft_Click()

radioright.Value = False

End Sub

Private Sub radioright_Click()

radioleft.Value = False

End Sub





Sub linePlan()

'User defined variables
Dim xCoord As Double
Dim yCoord As Double
Dim heading As Double
Dim spacing As Double
Dim lnum As Integer
Dim llength As Double

'Sub defined variables
Dim spcount As Integer
Dim x1 As Double
Dim y1 As Double
Dim x2 As Double
Dim y2 As Double
Dim slope As Double

'Loop variables
Dim tempx As Double
Dim tempy As Double
Dim deltax As Double
Dim deltay As Double

'Retrieve user defined values and assign them to variables
xCoord = coordx.Value
yCoord = coordy.Value
heading = lineheading.Value
spacing = linespacing.Value
lnum = linenum.Value
spspacing = shotpointspacing.Value
llength = linelength.Value
tempx = xCoord
tempy = yCoord

'Determine first lines characteristics

'Adjust the line length to maintain even shot point spacing
llength = llength - (llegth Mod spacing) + spacing
spcount = llength / spacing

deltay = hTrig.getSide("Sin", heading, 0, 0, spacing)
deltax = hTrig.getSide("Cos", heading, deltay, 0, spacing)


'Create an array which will be used as temporary storage for coordinate values
Dim coords As Variant
ReDim coords(2)
coords = Array(x, y)

'BEGIN ADDING THE COORDINATES YOU FOOL!!!!!
Dim linex As New Collection
Dim liney As New Collection

linex.Add xCoord
liney.Add yCoord

For Z = 1 To spcount
    
    tempx = tempx + deltax
    tempy = tempy + deltay
            
    linex.Add (tempx)
    liney.Add (tempy)
            
            
Next Z

End Sub

'TEST--------------------------------------
Private Sub CommandButton9_Click()

'User defined variables
Dim xCoord As Double
Dim yCoord As Double
Dim heading As Double
Dim spacing As Double
Dim lnum As Integer
Dim llength As Double

'Sub defined variables
Dim spcount As Integer
Dim x As Double
Dim y As Double

xCoord = coordx.Value
yCoord = coordy.Value
heading = lineheading.Value
spacing = linespacing.Value
lnum = linenum.Value
spspacing = shotpointspacing.Value
llength = linelength.Value

x = 1
y = 1

Dim endlinex As Double
endlinex = hTrig.getSide("Cos", heading, 559.1929, 0, llength)

Dim coords As Variant
ReDim coords(2)
coords = Array(x, y)

Debug.Print (endlinex)

End Sub
