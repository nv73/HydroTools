Attribute VB_Name = "hTrig"
Function getSide(operator As String, angle As Double, opposite As Double, adjacent As Double, hypotenuse As Double) As Double

    operator = UCase(operator)
        
    If operator = "SIN" Then
        
        If opposite = 0 Then
            
            'opposite = angle * hypotenuse
                           
            getSide = Sin(angle * (WorksheetFunction.Pi / 180)) * hypotenuse
            
        End If
        
        If hypotenuse = 0 Then
            
            'hypotenuse = opposite / angle
                        
            getSide = opposite / Sin(angle * (WorksheetFunction.Pi / 180))
        
        End If
            
    Else
    
    If operator = "COS" Then
    
        If adjacent = 0 Then
            
            'adjacent = angle * hypotenuse
                        
            getSide = Cos(angle * (WorksheetFunction.Pi / 180)) * hypotenuse
            
        End If
            
        If hypotenuse = 0 Then
        
            'hypotenuse = adjacent / angle
                        
            getSide = adjacent / Cos(angle * (WorksheetFunction.Pi / 180))
            
        End If
            
    Else
    
    If operator = "TAN" Then
        
        If adjacent = 0 Then
            
            'adjacent = opposite / angle
                        
            getSide = opposite / Tan(angle * (WorksheetFunction.Pi / 180))
            
        End If
        
        If opposite = 0 Then
        
            'opposite = angle * adjacent
                        
            getSide = Tan(angle * (WorksheetFunction.Pi / 180)) * adjacent
            
        End If
        
    End If
    End If
    End If
    
End Function

Sub testTrig()

Dim test As Double

test = hTrig.getSide("Sin", 40, 0, 40, 60)

Debug.Print (test)
'Debug.Print (Sin(40 * (WorksheetFunction.Pi / 180)))

End Sub
