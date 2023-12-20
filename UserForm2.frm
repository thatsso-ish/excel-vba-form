VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   13245
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   25275
   OleObjectBlob   =   "UserForm2.frx":0000
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

    ' Assuming "Sheet1" is the name of the worksheet where you want to write the data
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Claims Data")

    ' Find the first empty row in column A
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
    
    
    ' age analysis calculations
    
    Dim yearsDD As Long
    Dim monthsDD As Long
    Dim daysDD As Long

    ' Retrieve values from TextBoxes
    value1 = TextBox1.value
    value2 = TextBox17.value
    value3 = ComboBox1.value
    value4 = ComboBox5.value
    value5 = TextBox13.value
    value6 = TextBox12.value
    value7 = TextBox14.value
    value8 = TextBox2.value
    value9 = ComboBox4.value
    value11 = TextBox5.value
    value13 = ComboBox8.value
    value14 = TextBox4.value
    value15 = TextBox8.value
    value17 = TextBox10.value
    value20 = ComboBox3.value
    value21 = CheckBox1.value
    value22 = ComboBox2.value
    value23 = ComboBox6.value
    value24 = ComboBox7.value
    value25 = TextBox7.value
    value26 = TextBox15.value
    value27 = TextBox3.value
    value33 = ComboBox9.value
    value34 = TextBox11.value
    value38 = TextBox18.value
    value39 = TextBox9.value
    value40 = TextBox16.value
    value41 = TextBox19.value
    value43 = TextBox20.value
    value44 = TextBox21.value
    
    ' Assuming TextBox12 and TextBox10 are the TextBoxes with dates
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
    sheetdate = Sheets("Formula Sheet").Range("$E$26").value
    factor = 1 - Sheets("Formula Sheet").Range("$E$23").value
    

    ' Calculate the date difference
    yearsDD = DateDiff("yyyy", startDate, endDate)
    monthsDD = DateDiff("m", startDate, endDate) Mod 12
    daysDD = DateDiff("d", DateAdd("m", monthsDD, DateAdd("yyyy", yearsDD, startDate)), endDate)

    ' Concatenate the results into one variable
    value10 = yearsDD & " Years " & monthsDD & " Months " & daysDD & " Days"
    value12 = Year(notdate)
    value16 = Application.WorksheetFunction.NetworkDays(ati, datecc)
    value18 = Year(endDate)
    value19 = Application.WorksheetFunction.NetworkDays(endDate, notdate)
    value28 = Application.WorksheetFunction.NetworkDays(dcc, datecc)
    value29 = Year(datecc)
    value39 = value37 / TextBox11 * TextBox18
    value42 = Application.WorksheetFunction.NetworkDays(notdate, datecc)
    
    If CheckBox1.value = True Then
    
    value21 = "Yes"
    
    Else
    
    value21 = "No"
    
    End If
    
    If ComboBox6.value = "Closed" Then
    
        value30 = Application.WorksheetFunction.RoundUp(DateDiff("m", notdate, sheetdate), 0)
        
    End If
    
    If value30 <> "" Then
    
       value31 = Application.VLookup(value30, Sheets("Formula Sheet").Range("$A:$B"), 2, False)
       
    End If
  
    If value31 <> "" Then
        
       value32 = Application.VLookup(value31, Sheets("Formula Sheet").Range("$E$3:$F$19"), 2, False)
       
    End If
  
    If value32 <> "" Then
        
        value35 = value34 * (1 - value32)
       
    End If

    If ComboBox2 = "Pending" Then
    
       value36 = value34 * factor
        
    End If
       
    If value35 = "" And value36 = "" Then
        
        value37 = value34
        
    Else
    
        value37 = value35 + value36
        
    End If
   
        
    ' Write values to the worksheet in the first empty row
    ws.Cells(emptyRow, 1).value = value1
    ws.Cells(emptyRow, 2).value = value2
    ws.Cells(emptyRow, 3).value = value3
    ws.Cells(emptyRow, 4).value = value4
    ws.Cells(emptyRow, 5).value = value5
    ws.Cells(emptyRow, 6).value = value6
    ws.Cells(emptyRow, 7).value = value7
    ws.Cells(emptyRow, 8).value = value8
    ws.Cells(emptyRow, 9).value = value9
    ws.Cells(emptyRow, 10).value = value10
    ws.Cells(emptyRow, 11).value = value11
    ws.Cells(emptyRow, 12).value = value12
    ws.Cells(emptyRow, 13).value = value13
    ws.Cells(emptyRow, 14).value = value14
    ws.Cells(emptyRow, 15).value = value15
    ws.Cells(emptyRow, 16).value = value16
    ws.Cells(emptyRow, 17).value = value17
    ws.Cells(emptyRow, 18).value = value18
    ws.Cells(emptyRow, 19).value = value19
    ws.Cells(emptyRow, 20).value = value20
    ws.Cells(emptyRow, 21).value = value21
    ws.Cells(emptyRow, 22).value = value22
    ws.Cells(emptyRow, 23).value = value23
    ws.Cells(emptyRow, 24).value = value24
    ws.Cells(emptyRow, 25).value = value25
    ws.Cells(emptyRow, 26).value = value26
    ws.Cells(emptyRow, 27).value = value27
    ws.Cells(emptyRow, 28).value = value28
    ws.Cells(emptyRow, 29).value = value29
    ws.Cells(emptyRow, 30).value = value30
    ws.Cells(emptyRow, 31).value = value31
    ws.Cells(emptyRow, 32).value = value32
    ws.Cells(emptyRow, 33).value = value33
    ws.Cells(emptyRow, 34).value = value34
    ws.Cells(emptyRow, 35).value = value35
    ws.Cells(emptyRow, 36).value = value36
    ws.Cells(emptyRow, 37).value = value37
    ws.Cells(emptyRow, 38).value = value38
    ws.Cells(emptyRow, 39).value = value39
    ws.Cells(emptyRow, 40).value = value40
    ws.Cells(emptyRow, 41).value = value41
    ws.Cells(emptyRow, 42).value = value42
    ws.Cells(emptyRow, 43).value = value43
    ws.Cells(emptyRow, 44).value = value44
  
  
    Unload Me
    CommandButton2_Click
    MsgBox ("Claim successfully added")
    
End Sub

Private Sub CommandButton2_Click()

Dim cntrl As Control

For Each cntrl In Me.Controls
If TypeName(cntrl) = "ComboBox" Or TypeName(cntrl) = "TextBox" Then
cntrl.value = ""
ElseIf TypeName(cntrl) = "CheckBox" Then
cntrl.value = False
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
    inputValue = Trim(TextBox1.value)

    If Len(inputValue) <> 10 And inputValue <> "N/A" And inputValue <> "n/A" And inputValue <> "N/a" And inputValue <> "n/a" Then
        MsgBox "Invalid input. Please enter a valide policy number or 'N/A'.", vbExclamation
        TextBox1.SetFocus
        Cancel = True
    End If

End Sub

Private Sub TextBox10_Change()

End Sub

Private Sub TextBox10_Exit(ByVal Cancel As MSForms.ReturnBoolean)

If Not IsDate(TextBox10.value) Then
        TextBox10.value = Format(CDate(enteredDate), "dd mmmm yyyy")
        MsgBox "Invalid date. Please enter a valid date.", vbExclamation
        TextBox10.SetFocus
        Cancel = True
    End If
End Sub

Private Sub TextBox11_Change()

End Sub

Private Sub TextBox11_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Not IsNumeric(TextBox11.value) Then
        MsgBox "Invalid currency value. Please enter a numeric value.", vbExclamation
        TextBox11.SetFocus
        Cancel = True
    Else
        TextBox11.value = Format(TextBox11.value, "R #,##0.00")
    End If

End Sub

Private Sub TextBox12_Change()

End Sub

Private Sub TextBox12_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    
    If Not IsDate(TextBox12.value) Then
        TextBox12.value = Format(CDate(enteredDate), "dd mmmm yyyy")
        MsgBox "Invalid date. Please enter a valid date.", vbExclamation
        TextBox12.SetFocus
        Cancel = True
    End If

End Sub

Private Sub TextBox14_Change()

End Sub

Private Sub TextBox14_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Trim(TextBox14.value) = "" Then
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

Private Sub TextBox15_Exit(ByVal Cancel As MSForms.ReturnBoolean)

If Not IsDate(TextBox15.value) Then
        TextBox15.value = Format(CDate(enteredDate), "dd mmmm yyyy")
        MsgBox "Invalid date. Please enter a valid date.", vbExclamation
        TextBox15.SetFocus
        Cancel = True
    End If
End Sub

Private Sub TextBox16_Change()

End Sub

Private Sub TextBox16_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Not IsNumeric(TextBox16.value) Then
        MsgBox "Invalid currency value. Please enter a numeric value.", vbExclamation
        TextBox16.SetFocus
        Cancel = True
    Else
        TextBox16.value = Format(TextBox16.value, "R #,##0.00")
    End If

End Sub

Private Sub TextBox18_Change()

End Sub

Private Sub TextBox18_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Not IsNumeric(TextBox18.value) Then
        MsgBox "Invalid currency value. Please enter a numeric value.", vbExclamation
        TextBox18.SetFocus
        Cancel = True
    Else
        TextBox18.value = Format(TextBox18.value, "R #,##0.00")
    End If

End Sub

Private Sub TextBox19_Change()

End Sub

Private Sub TextBox19_Exit(ByVal Cancel As MSForms.ReturnBoolean)

If Not IsDate(TextBox19.value) Then
        TextBox19.value = Format(CDate(enteredDate), "dd mmmm yyyy")
        MsgBox "Invalid date. Please enter a valid date.", vbExclamation
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

Private Sub TextBox21_Exit(ByVal Cancel As MSForms.ReturnBoolean)

If Not IsDate(TextBox21.value) Then
        TextBox21.value = Format(CDate(enteredDate), "dd mmmm yyyy")
        MsgBox "Invalid date. Please enter a valid date.", vbExclamation
        TextBox21.SetFocus
        Cancel = True
    End If

End Sub

Private Sub TextBox3_Change()

End Sub

Private Sub TextBox3_Exit(ByVal Cancel As MSForms.ReturnBoolean)

If Not IsDate(TextBox3.value) Then
        TextBox3.value = Format(CDate(enteredDate), "dd mmmm yyyy")
        MsgBox "Invalid date. Please enter a valid date.", vbExclamation
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

Private Sub TextBox5_Exit(ByVal Cancel As MSForms.ReturnBoolean)

If Not IsDate(TextBox5.value) Then
        TextBox5.value = Format(CDate(enteredDate), "dd mmmm yyyy")
        MsgBox "Invalid date. Please enter a valid date.", vbExclamation
        TextBox5.SetFocus
        Cancel = True
    End If
End Sub

Private Sub TextBox7_Change()

End Sub

Private Sub TextBox7_Exit(ByVal Cancel As MSForms.ReturnBoolean)
If Not IsDate(TextBox7.value) Then
        TextBox7.value = Format(CDate(enteredDate), "dd mmmm yyyy")
        MsgBox "Invalid date. Please enter a valid date.", vbExclamation
        TextBox7.SetFocus
        Cancel = True
    End If
End Sub

Private Sub TextBox8_Change()

End Sub

Private Sub TextBox8_Exit(ByVal Cancel As MSForms.ReturnBoolean)

If Not IsDate(TextBox8.value) Then
        TextBox8.value = Format(CDate(enteredDate), "dd mmmm yyyy")
        MsgBox "Invalid date. Please enter a valid date.", vbExclamation
        TextBox8.SetFocus
        Cancel = True
    End If

End Sub

Private Sub TextBox9_Change()

End Sub

Private Sub TextBox9_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Not IsNumeric(TextBox9.value) Then
        MsgBox "Invalid currency value. Please enter a numeric value.", vbExclamation
        TextBox9.SetFocus
        Cancel = True
    Else
        TextBox9.value = Format(TextBox9.value, "R #,##0.00")
    End If

End Sub

Private Sub UserForm_Activate()

    Image1.PictureSizeMode = fmSizeModeStretch
    Image1.Height = 96
    Image1.Width = 828
    Image1.Picture = LoadPicture("C:\Users\ingobeni\Pictures\R.jfif")

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
