VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   13245
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   25275
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()

End Sub

Private Sub ComboBox9_Change()

End Sub

Private Sub CommandButton1_Click()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Claims Data")

    Dim emptyRow As Long
    emptyRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    
    Dim value1 As String
    Dim value2 As String
    Dim value3 As String
    Dim value4 As String
    Dim value5 As String
    Dim value6 As Date
    Dim value7 As String
    Dim value8 As String
    Dim value9 As String
    Dim value10 As String
    Dim value11 As Date
    Dim value12 As Integer
    Dim value13 As String
    Dim value14 As String
    Dim value15 As Date
    Dim value16 As Long
    Dim value17 As Date
    Dim value18 As Integer
    Dim value19 As Long
    Dim value20 As String
    Dim value21 As String
    Dim value22 As String
    Dim value23 As String
    Dim value24 As String
    Dim value25 As Date
    Dim value26 As Date
    Dim value27 As Date
    Dim value28 As Long
    Dim value29 As Integer
    Dim value30 As Integer
    Dim value31 As String
    Dim value32 As String
    Dim value33 As String
    Dim value34 As Currency
    Dim value35 As Currency
    Dim value36 As Currency
    Dim value37 As Currency
    Dim value38 As Currency
    Dim value39 As Currency
    Dim value40 As Currency
    Dim value41 As Date
    Dim value42 As Long
    Dim value43 As String
    Dim value44 As Date
    
    Dim yearsDD As Long
    Dim monthsDD As Long
    Dim daysDD As Long

    value1 = TextBox1.Value
    value2 = TextBox17.Value
    value3 = ComboBox1.Value
    value4 = ComboBox5.Value
    value5 = TextBox13.Value
    value6 = TextBox12.Value
    value7 = TextBox14.Value
    value8 = TextBox2.Value
    value9 = ComboBox4.Value
    value11 = TextBox5.Value
    value13 = ComboBox8.Value
    value14 = TextBox4.Value
    value15 = TextBox8.Value
    value17 = TextBox10.Value
    value20 = ComboBox3.Value
    value21 = CheckBox1.Value
    value22 = ComboBox2.Value
    value23 = ComboBox6.Value
    value24 = ComboBox7.Value
    value25 = TextBox7.Value
    value26 = TextBox15.Value
    value27 = TextBox3.Value
    value33 = ComboBox9.Value
    value34 = TextBox11.Value
    value38 = TextBox18.Value
    value39 = TextBox9.Value
    value40 = TextBox16.Value
    value41 = TextBox19.Value
    value43 = TextBox20.Value
    value44 = TextBox21.Value
    
    Dim startDate As Date
    Dim endDate As Date
    Dim datecc As Date
    Dim dcc As Date
    Dim ati As Date
    Dim notdate As Date
    Dim factor As Double
    
    startDate = DateValue(value6)
    endDate = DateValue(value17)
    notdate = DateValue(value11)
    datecc = DateValue(value26)
    dcc = DateValue(value25)
    ati = DateValue(value15)
    sheetdate = Sheets("Formula Sheet").Range("$E$26").Value
    factor = 1 - Sheets("Formula Sheet").Range("$E$23").Value
  
    yearsDD = DateDiff("yyyy", startDate, endDate)
    monthsDD = DateDiff("m", startDate, endDate) Mod 12
    daysDD = DateDiff("d", DateAdd("m", monthsDD, DateAdd("yyyy", yearsDD, startDate)), endDate)

    value10 = yearsDD & " Years " & monthsDD & " Months " & daysDD & " Days"
    value12 = Year(notdate)
    value16 = Application.WorksheetFunction.NetworkDays(ati, datecc)
    value18 = Year(endDate)
    value19 = Application.WorksheetFunction.NetworkDays(endDate, notdate)
    value28 = Application.WorksheetFunction.NetworkDays(dcc, datecc)
    value29 = Year(datecc)
    value42 = Application.WorksheetFunction.NetworkDays(notdate, datecc)
    
    If CheckBox1.Value = True Then
    
    value21 = "Yes"
    
    Else
    
    value21 = "No"
    
    End If
    
    If ComboBox6.Value = "Closed" Then
    
        value30 = Application.WorksheetFunction.RoundUp(DateDiff("m", notdate, sheetdate), 0)
        
    End If
    
    If value30 <> 0 Then
    
       value31 = Application.VLookup(value30, Sheets("Formula Sheet").Range("$A:$B"), 2, 0)
       
    End If
  
    If value31 <> "" Then
        
      value32 = Application.VLookup(value31, Sheets("Formula Sheet").Range("$E$3:$F$19"), 2, 0)
       
    End If
  
    If value32 <> "" Then
        
        value35 = value34 * (1 - value32)
       
    End If

    If ComboBox2 = "Pending" Then
    
       value36 = value34 * factor
        
    End If
       
    If value35 = 0 And value36 = 0 Then
        
        value37 = value34
        
    Else
    
        value37 = value35 + value36
        
    End If
   
    ws.Cells(emptyRow, 1).Value = value1
    ws.Cells(emptyRow, 2).Value = value2
    ws.Cells(emptyRow, 3).Value = value3
    ws.Cells(emptyRow, 4).Value = value4
    ws.Cells(emptyRow, 5).Value = value5
    ws.Cells(emptyRow, 6).Value = value6
    ws.Cells(emptyRow, 7).Value = value7
    ws.Cells(emptyRow, 8).Value = value8
    ws.Cells(emptyRow, 9).Value = value9
    ws.Cells(emptyRow, 10).Value = value10
    ws.Cells(emptyRow, 11).Value = value11
    ws.Cells(emptyRow, 12).Value = value12
    ws.Cells(emptyRow, 13).Value = value13
    ws.Cells(emptyRow, 14).Value = value14
    ws.Cells(emptyRow, 15).Value = value15
    ws.Cells(emptyRow, 16).Value = value16
    ws.Cells(emptyRow, 17).Value = value17
    ws.Cells(emptyRow, 18).Value = value18
    ws.Cells(emptyRow, 19).Value = value19
    ws.Cells(emptyRow, 20).Value = value20
    ws.Cells(emptyRow, 21).Value = value21
    ws.Cells(emptyRow, 22).Value = value22
    ws.Cells(emptyRow, 23).Value = value23
    ws.Cells(emptyRow, 24).Value = value24
    ws.Cells(emptyRow, 25).Value = value25
    ws.Cells(emptyRow, 26).Value = value26
    ws.Cells(emptyRow, 27).Value = value27
    ws.Cells(emptyRow, 28).Value = value28
    ws.Cells(emptyRow, 29).Value = value29
    ws.Cells(emptyRow, 30).Value = value30
    ws.Cells(emptyRow, 31).Value = value31
    ws.Cells(emptyRow, 32).Value = value32
    ws.Cells(emptyRow, 33).Value = value33
    ws.Cells(emptyRow, 34).Value = value34
    ws.Cells(emptyRow, 35).Value = value35
    ws.Cells(emptyRow, 36).Value = value36
    ws.Cells(emptyRow, 37).Value = value37
    ws.Cells(emptyRow, 38).Value = value38
    ws.Cells(emptyRow, 39).Value = value39
    ws.Cells(emptyRow, 40).Value = value40
    ws.Cells(emptyRow, 41).Value = value41
    ws.Cells(emptyRow, 42).Value = value42
    ws.Cells(emptyRow, 43).Value = value43
    ws.Cells(emptyRow, 44).Value = value44
  
    Unload Me
    CommandButton2_Click
    MsgBox ("Claim successfully added")
    
End Sub

Private Sub CommandButton2_Click()

Dim cntrl As Control

For Each cntrl In Me.Controls
If TypeName(cntrl) = "ComboBox" Or TypeName(cntrl) = "TextBox" Then
cntrl.Value = ""
ElseIf TypeName(cntrl) = "CheckBox" Then
cntrl.Value = False
End If
Next cntrl

End Sub

Private Sub CommandButton3_Click()
Unload Me
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub Frame1_Exit(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub Label21_Click()

End Sub

Private Sub Label25_Click()

End Sub

Private Sub Label30_Click()

End Sub

Private Sub Label9_Click()

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub TextBox1_Enter()

End Sub

Private Sub TextBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim inputValue As String
    inputValue = Trim(TextBox1.Value)

    If Len(inputValue) <> 10 And inputValue <> "N/A" And inputValue <> "n/A" And inputValue <> "N/a" And inputValue <> "n/a" Then
        MsgBox "Invalid input. Please enter a valide policy number or 'N/A'.", vbExclamation
       ' TextBox1.SetFocus
        Cancel = True
    End If

End Sub

Private Sub TextBox10_Change()

End Sub

Private Sub TextBox10_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsDate(TextBox10.Value) Then
        MsgBox "Invalid date. Please enter a valid date.", vbExclamation
        TextBox10.Value = Format(CDate(enteredDate), "dd mmmm yyyy")
        TextBox10.SetFocus
        Cancel = True
    End If
End Sub

Private Sub TextBox11_Change()

End Sub

Private Sub TextBox11_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Not IsNumeric(TextBox11.Value) Then
        MsgBox "Invalid currency value. Please enter a numeric value.", vbExclamation
        TextBox11.SetFocus
        Cancel = True
    ' Else
     '   TextBox11.value = Format(TextBox11.value, "R #,##0.00")
    End If

End Sub

Private Sub TextBox12_Change()

End Sub

Private Sub TextBox12_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsDate(TextBox12.Value) Then
        MsgBox "Invalid date. Please enter a valid date.", vbExclamation
        TextBox12.Value = Format(CDate(enteredDate), "dd mmmm yyyy")
        TextBox12.SetFocus
        Cancel = True
    End If
End Sub


Private Sub TextBox14_Change()

End Sub

Private Sub TextBox14_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Trim(TextBox14.Value) = "" Then
        MsgBox "Client's name cannot be left empty. Please enter a value.", vbExclamation
        TextBox14.SetFocus
        Cancel = True
    End If

End Sub

Private Sub TextBox14_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
       KeyAscii = 0
    End If

End Sub

Private Sub TextBox15_Change()

End Sub

Private Sub TextBox15_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsDate(TextBox15.Value) Then
        MsgBox "Invalid date. Please enter a valid date.", vbExclamation
        TextBox15.Value = Format(CDate(enteredDate), "dd mmmm yyyy")
        TextBox15.SetFocus
        Cancel = True
    End If
End Sub

Private Sub TextBox16_Change()

End Sub

Private Sub TextBox16_Exit(ByVal Cancel As MSForms.ReturnBoolean)

   ' If Not IsNumeric(TextBox16.Value) Then
    '    MsgBox "Invalid currency value. Please enter a numeric value.", vbExclamation
     '   TextBox16.SetFocus
      '  Cancel = True
   ' Else
    '    TextBox16.value = Format(TextBox16.value, "R #,##0.00")
    ' End If

End Sub

Private Sub TextBox18_Change()

End Sub

Private Sub TextBox18_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Not IsNumeric(TextBox18.Value) Then
        MsgBox "Invalid currency value. Please enter a numeric value.", vbExclamation
        TextBox18.SetFocus
        Cancel = True
   ' Else
    '    TextBox18.value = Format(TextBox18.value, "R #,##0.00")
    End If

End Sub

Private Sub TextBox19_Change()

End Sub

Private Sub TextBox19_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsDate(TextBox19.Value) Then
        MsgBox "Invalid date. Please enter a valid date.", vbExclamation
        TextBox19.Value = Format(CDate(enteredDate), "dd mmmm yyyy")
        TextBox19.SetFocus
        Cancel = True
    End If
End Sub


Private Sub TextBox20_Change()



End Sub

Private Sub TextBox20_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub TextBox21_Change()

End Sub

Private Sub TextBox21_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsDate(TextBox21.Value) Then
        MsgBox "Invalid date. Please enter a valid date.", vbExclamation
        TextBox21.Value = Format(CDate(enteredDate), "dd mmmm yyyy")
        TextBox21.SetFocus
        Cancel = True
    End If
End Sub


Private Sub TextBox3_Change()

End Sub

Private Sub TextBox3_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsDate(TextBox12.Value) Then
        MsgBox "Invalid date. Please enter a valid date.", vbExclamation
        TextBox3.Value = Format(CDate(enteredDate), "dd mmmm yyyy")
        TextBox3.SetFocus
        Cancel = True
    End If
End Sub



Private Sub TextBox4_Change()

End Sub

Private Sub TextBox4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
       KeyAscii = 0
    End If

End Sub

Private Sub TextBox5_Change()

End Sub

Private Sub TextBox5_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsDate(TextBox5.Value) Then
        MsgBox "Invalid date. Please enter a valid date.", vbExclamation
        TextBox5.Value = Format(CDate(enteredDate), "dd mmmm yyyy")
        TextBox5.SetFocus
        Cancel = True
    End If
End Sub


Private Sub TextBox7_Change()

End Sub

Private Sub TextBox7_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsDate(TextBox7.Value) Then
        MsgBox "Invalid date. Please enter a valid date.", vbExclamation
        TextBox7.Value = Format(CDate(enteredDate), "dd mmmm yyyy")
        TextBox7.SetFocus
        Cancel = True
    End If
End Sub


Private Sub TextBox8_Change()

End Sub

Private Sub TextBox8_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsDate(TextBox8.Value) Then
        MsgBox "Invalid date. Please enter a valid date.", vbExclamation
        TextBox8.Value = Format(CDate(enteredDate), "dd mmmm yyyy")
        TextBox8.SetFocus
        Cancel = True
    End If
End Sub


Private Sub TextBox9_Change()

End Sub

Private Sub TextBox9_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Not IsNumeric(TextBox9.Value) Then
        MsgBox "Invalid currency value. Please enter a numeric value.", vbExclamation
        TextBox9.SetFocus
        Cancel = True
   ' Else
    '    TextBox9.value = Format(TextBox9.value, "R #,##0.00")
    End If

End Sub

Private Sub UserForm_Activate()

ComboBox1.Style = fmStyleDropDownList
ComboBox2.Style = fmStyleDropDownList
ComboBox3.Style = fmStyleDropDownList
ComboBox4.Style = fmStyleDropDownList
ComboBox5.Style = fmStyleDropDownList
ComboBox6.Style = fmStyleDropDownList
ComboBox7.Style = fmStyleDropDownList
ComboBox8.Style = fmStyleDropDownList
ComboBox9.Style = fmStyleDropDownList

Me.ComboBox1.AddItem "Full Life"
Me.ComboBox1.AddItem "ADB"

Me.ComboBox5.AddItem "Full Life"
Me.ComboBox5.AddItem "ADB"

Me.ComboBox4.AddItem "Gauteng"
Me.ComboBox4.AddItem "Mpumalanga"
Me.ComboBox4.AddItem "Limpopo"
Me.ComboBox4.AddItem "Free State"
Me.ComboBox4.AddItem "North West"
Me.ComboBox4.AddItem "Kwa Zulu Natal"
Me.ComboBox4.AddItem "Northern Cape"
Me.ComboBox4.AddItem "Western Cape"
Me.ComboBox4.AddItem "Eastern Cape"

Me.ComboBox3.AddItem "N/A"
Me.ComboBox3.AddItem "Natural"
Me.ComboBox3.AddItem "Suicide"
Me.ComboBox3.AddItem "Terminal Illness Claim"
Me.ComboBox3.AddItem "Unknown"
Me.ComboBox3.AddItem "Unnatural"

Me.ComboBox2.AddItem "Approved"
Me.ComboBox2.AddItem "Cancelled"
Me.ComboBox2.AddItem "Closed"
Me.ComboBox2.AddItem "Pending"
Me.ComboBox2.AddItem "Rejected"

Me.ComboBox6.AddItem "Approved/Pending Payment"
Me.ComboBox6.AddItem "Claim cancelled"
Me.ComboBox6.AddItem "Closed"
Me.ComboBox6.AddItem "MST"
Me.ComboBox6.AddItem "No Cover Rejected"
Me.ComboBox6.AddItem "Paid"
Me.ComboBox6.AddItem "Pending"
Me.ComboBox6.AddItem "Pending-INV"
Me.ComboBox6.AddItem "Policy excl Rejected"

Me.ComboBox7.AddItem "Closed"
Me.ComboBox7.AddItem "Life Assured is still alive according to DHA"
Me.ComboBox7.AddItem "N/A"
Me.ComboBox7.AddItem "OLTI- Case"
Me.ComboBox7.AddItem "Pending"
Me.ComboBox7.AddItem "Pending-INV"
Me.ComboBox7.AddItem "Rejected/OLTI overturn-Approved"

Me.ComboBox8.AddItem "Asiphe"
Me.ComboBox8.AddItem "Bongisipho"
Me.ComboBox8.AddItem "Cristal"
Me.ComboBox8.AddItem "Cyril"
Me.ComboBox8.AddItem "Kaizer"
Me.ComboBox8.AddItem "Kelebogile"
Me.ComboBox8.AddItem "Kreban"
Me.ComboBox8.AddItem "Naledi"
Me.ComboBox8.AddItem "Refiloe"
Me.ComboBox8.AddItem "Sello"
Me.ComboBox8.AddItem "Simone"
Me.ComboBox8.AddItem "Terry"
Me.ComboBox8.AddItem "Tshidi"

Me.ComboBox9.AddItem "Life Assured is still alive according to DHA"
Me.ComboBox9.AddItem "No Cover - ADB Cover Only-Policy sold on ADB"
Me.ComboBox9.AddItem "No Cover - Cover dropped from Full Life to ADB Cover"
Me.ComboBox9.AddItem "No Cover - Death before inception"
Me.ComboBox9.AddItem "No Cover - Fraud / Dishonesty"
Me.ComboBox9.AddItem "No Cover - Not Terminal Illness"
Me.ComboBox9.AddItem "No Cover - Willfully breaking the law"
Me.ComboBox9.AddItem "No Cover - Policy cancelled before date of death"
Me.ComboBox9.AddItem "No Cover - First premium not received for cover to start"
Me.ComboBox9.AddItem "No Cover - Policy lapsed/cancelled before date of death"
Me.ComboBox9.AddItem "Non Compliance with RR"
Me.ComboBox9.AddItem "Non Disclosure - Criminal Activities"
Me.ComboBox9.AddItem "Non Disclosure - Material information"
Me.ComboBox9.AddItem "Non Disclosure - Material Medical Conditions"
Me.ComboBox9.AddItem "Non Disclosure - True, Material and Complete Information"
Me.ComboBox9.AddItem "Not a terminal illness"
Me.ComboBox9.AddItem "Suicide within the 2 year waiting period"
Me.ComboBox9.AddItem "Willfully breaking the law"
Me.ComboBox9.AddItem " "


End Sub

Private Sub UserForm_Click()

End Sub
