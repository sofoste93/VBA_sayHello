' ---							---
' -							      -	
'- 			Hello Sir!			   -	
' -							      -	
' ---							---


' - Written using macro recorder

Sub sayHello()
'
' sayHello Macro
' Greeting macro - my first macro
'

'
    ActiveCell.FormulaR1C1 = "Hello Sir!"
    Range("C3").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Columns("C:C").EntireColumn.AutoFit
    ActiveWorkbook.Save
End Sub
