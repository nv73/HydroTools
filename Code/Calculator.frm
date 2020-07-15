VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Calculator 
   Caption         =   "Calculator"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9705.001
   OleObjectBlob   =   "Calculator.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Calculator_initialize()

Sheet1.Activate

End Sub

Private Sub buttonAdd_Click()

calcOutput.Value = calcOutput.Value & " + "

End Sub

Private Sub buttonCos_Click()

calcOutput.Value = calcOutput.Value & "COS("

End Sub

Private Sub buttonDivide_Click()

calcOutput.Value = calcOutput.Value & " / "

End Sub

Private Sub buttonEight_Click()

Dim eight As Integer

eight = 8

calcOutput.Value = calcOutput.Value & 8
    
End Sub

Private Sub buttonEquals_Click()

On Error GoTo calcError:

Range("A1").Value = "=" & calcOutput.Value

result.Value = Range("A1").Value

calcError:

If Err.Number = 1004 Then

    MsgBox ("Invalid Expression")
    
End If

End Sub

Private Sub buttonExponent_Click()

Dim y As Variant

y = InputBox("Enter exponent value: ")

calcOutput.Value = calcOutput.Value & "^" & y

End Sub

Private Sub buttonFive_Click()

Dim five As Integer
five = 5
calcOutput.Value = calcOutput.Value & five


End Sub

Private Sub buttonFour_Click()

Dim four As Integer
four = 4
calcOutput.Value = calcOutput.Value & four


End Sub

Private Sub buttonMultiply_Click()

calcOutput.Value = calcOutput.Value & " * "

End Sub

Private Sub buttonNine_Click()

Dim nine As Integer
nine = 9
calcOutput.Value = calcOutput.Value & nine


End Sub

Private Sub buttonOne_Click()

Dim one As Integer
one = 1
calcOutput.Value = calcOutput.Value & one


End Sub

Private Sub buttonSeven_Click()

Dim seven As Integer
seven = 7
calcOutput.Value = calcOutput.Value & seven


End Sub

Private Sub buttonSin_Click()

calcOutput.Value = calcOutput.Value & "SIN("

End Sub

Private Sub buttonSix_Click()

Dim six As Integer
six = 6
calcOutput.Value = calcOutput.Value & six


End Sub

Private Sub buttonSquare_Click()

calcOutput.Value = calcOutput.Value & "^2"

End Sub

Private Sub buttonSquareRoot_Click()

calcOutput.Value = calcOutput.Value & "sqrt("

End Sub

Private Sub buttonSubtract_Click()

calcOutput.Value = calcOutput.Value & " - "

End Sub

Private Sub buttonTan_Click()

calcOutput.Value = calcOutput.Value & "TAN("

End Sub

Private Sub buttonThree_Click()

Dim three As Integer
three = 3
calcOutput.Value = calcOutput.Value & three


End Sub

Private Sub buttonTwo_Click()

Dim two As Integer
two = 2
calcOutput.Value = calcOutput.Value & two


End Sub

Private Sub buttonZero_Click()

Dim zero As Integer
zero = 0
calcOutput.Value = calcOutput.Value & zero


End Sub

Private Sub clearButton_Click()

calcOutput.Value = ""
result.Value = ""
Range("A1").Value = ""

End Sub

Private Sub leftBracket_Click()

calcOutput.Value = calcOutput.Value & "("

End Sub

Private Sub rightBracket_Click()

calcOutput.Value = calcOutput.Value & ")"

End Sub

Private Sub UserForm_Terminate()

UserForm1.Show

End Sub
