Attribute VB_Name = "Module1"
Function CalculateClosurePeriod() As Variant
    If Range("ComboBox6").Value = "Closed" Then
        CalculateClosurePeriod = Application.WorksheetFunction.RoundUp(DateDiff("m", Range("TextBox5").Value, Sheets("Formula Sheet").Range("$E$26").Value), 0)
    Else
        CalculateClosurePeriod = ""
    End If
End Function

Function YourVBAFunction() As Variant
    Dim closurePeriod As Variant
    closurePeriod = CalculateClosurePeriod
    
    If closurePeriod = "" Then
        YourVBAFunction = ""
    Else
        YourVBAFunction = Application.WorksheetFunction.VLookup(closurePeriod, Sheets("Formula Sheet").Range("A:B"), 2, 0)
    End If
End Function

If Range("AE58").Value = "" Then
    ' Do nothing or handle the case when AE58 is empty
Else
    ' VLOOKUP function
    Dim result As Variant
    result = Application.WorksheetFunction.VLookup(Range("AE58").Value, Sheets("Formula Sheet").Range("$E$3:$F$19"), 2, 0)
    
    ' You can use the 'result' variable as needed
    ' For example, if you want to put the result in another cell:
    ' Range("YourTargetCell").Value = result
End If




