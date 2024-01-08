VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   12570
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   25185
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub DeleteData(searchValue As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim foundCell As Range

    Set ws = ThisWorkbook.Sheets("Claims Data")

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Set foundCell = ws.Range("A1:B" & lastRow & ",D1:D" & lastRow).Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole)

    If Not foundCell Is Nothing Then
        foundCell.EntireRow.Delete
        
        TextBox2.Value = ""
        TextBox3.Value = ""
        TextBox4.Value = ""
        TextBox5.Value = ""
        TextBox6.Value = ""
        TextBox7.Value = ""
        TextBox8.Value = ""
        TextBox9.Value = ""
        TextBox10.Value = ""
        Label2.Caption = ""
        TextBox12.Value = ""
        Label3.Caption = ""
        TextBox14.Value = ""
        TextBox15.Value = ""
        TextBox16.Value = ""
        Label4.Caption = ""
        TextBox18.Value = ""
        Label5.Caption = ""
        Label6.Caption = ""
        TextBox21.Value = ""
        TextBox22.Value = ""
        TextBox23.Value = ""
        TextBox24.Value = ""
        TextBox25.Value = ""
        TextBox26.Value = ""
        TextBox27.Value = ""
        TextBox28.Value = ""
        Label7.Caption = ""
        Label8.Caption = ""
        Label9.Caption = ""
        Label10.Caption = ""
        Label11.Caption = ""
        TextBox34.Value = ""
        TextBox35.Value = ""
        Label12.Caption = ""
        Label13.Caption = ""
        Label14.Caption = ""
        TextBox39.Value = ""
        TextBox40.Value = ""
        TextBox41.Value = ""
        TextBox42.Value = ""
        Label15.Caption = ""
        TextBox44.Value = ""
        TextBox45.Value = ""
        
        foundRow = 0
        MsgBox "Data deleted successfully."
    Else
        MsgBox "No data found to delete."
    End If
    
End Sub

Private Sub CommandButton1_Click()
    DeleteData TextBox1.Value
End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub

Private Sub CommandButton3_Click()
UpdateData TextBox1.Value
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label15_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub TextBox2_Change()

End Sub

Private Sub TextBox44_Change()

End Sub

Private Sub UserForm_Click()

End Sub

Sub SearchData(searchValue As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim foundCell As Range

    Set ws = ThisWorkbook.Sheets("Claims Data")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Set foundCell = ws.Range("A1:B" & lastRow & ",E1:E" & lastRow).Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole)

    If Not foundCell Is Nothing Then
      
        TextBox2.Value = ws.Cells(foundCell.Row, 1).Value
        TextBox3.Value = ws.Cells(foundCell.Row, 2).Value
        TextBox4.Value = ws.Cells(foundCell.Row, 3).Value
        TextBox5.Value = ws.Cells(foundCell.Row, 4).Value
        TextBox6.Value = ws.Cells(foundCell.Row, 5).Value
        TextBox7.Value = ws.Cells(foundCell.Row, 6).Value
        TextBox8.Value = ws.Cells(foundCell.Row, 7).Value
        TextBox9.Value = ws.Cells(foundCell.Row, 8).Value
        TextBox10.Value = ws.Cells(foundCell.Row, 9).Value
        Label2.Caption = ws.Cells(foundCell.Row, 10).Value
        TextBox12.Value = ws.Cells(foundCell.Row, 11).Value
        Label3.Caption = ws.Cells(foundCell.Row, 12).Value
        TextBox14.Value = ws.Cells(foundCell.Row, 13).Value
        TextBox15.Value = ws.Cells(foundCell.Row, 14).Value
        TextBox16.Value = ws.Cells(foundCell.Row, 15).Value
        Label4.Caption = ws.Cells(foundCell.Row, 16).Value
        TextBox18.Value = ws.Cells(foundCell.Row, 17).Value
        Label5.Caption = ws.Cells(foundCell.Row, 18).Value
        Label6.Caption = ws.Cells(foundCell.Row, 19).Value
        TextBox21.Value = ws.Cells(foundCell.Row, 20).Value
        TextBox22.Value = ws.Cells(foundCell.Row, 21).Value
        TextBox23.Value = ws.Cells(foundCell.Row, 22).Value
        TextBox24.Value = ws.Cells(foundCell.Row, 23).Value
        TextBox25.Value = ws.Cells(foundCell.Row, 24).Value
        TextBox26.Value = ws.Cells(foundCell.Row, 25).Value
        TextBox27.Value = ws.Cells(foundCell.Row, 26).Value
        TextBox28.Value = ws.Cells(foundCell.Row, 27).Value
        Label7.Caption = ws.Cells(foundCell.Row, 28).Value
        Label8.Caption = ws.Cells(foundCell.Row, 29).Value
        Label9.Caption = ws.Cells(foundCell.Row, 30).Value
        Label10.Caption = ws.Cells(foundCell.Row, 31).Value
        Label11.Caption = ws.Cells(foundCell.Row, 32).Value
        TextBox34.Value = ws.Cells(foundCell.Row, 33).Value
        TextBox35.Value = ws.Cells(foundCell.Row, 34).Value
        Label12.Caption = ws.Cells(foundCell.Row, 35).Value
        Label13.Caption = ws.Cells(foundCell.Row, 36).Value
        Label14.Caption = ws.Cells(foundCell.Row, 37).Value
        TextBox39.Value = ws.Cells(foundCell.Row, 38).Value
        TextBox40.Value = ws.Cells(foundCell.Row, 39).Value
        TextBox41.Value = ws.Cells(foundCell.Row, 40).Value
        TextBox42.Value = ws.Cells(foundCell.Row, 41).Value
        Label15.Caption = ws.Cells(foundCell.Row, 42).Value
        TextBox44.Value = ws.Cells(foundCell.Row, 43).Value
        TextBox45.Value = ws.Cells(foundCell.Row, 44).Value
        
        EnableTextboxes True
        
    Else
        
        TextBox2.Value = ""
        TextBox3.Value = ""
        TextBox4.Value = ""
        TextBox5.Value = ""
        TextBox6.Value = ""
        TextBox7.Value = ""
        TextBox8.Value = ""
        TextBox9.Value = ""
        TextBox10.Value = ""
        Label2.Caption = ""
        TextBox12.Value = ""
        Label3.Caption = ""
        TextBox14.Value = ""
        TextBox15.Value = ""
        TextBox16.Value = ""
        Label4.Caption = ""
        TextBox18.Value = ""
        Label5.Caption = ""
        Label6.Caption = ""
        TextBox21.Value = ""
        TextBox22.Value = ""
        TextBox23.Value = ""
        TextBox24.Value = ""
        TextBox25.Value = ""
        TextBox26.Value = ""
        TextBox27.Value = ""
        TextBox28.Value = ""
        Label7.Caption = ""
        Label8.Caption = ""
        Label9.Caption = ""
        Label10.Caption = ""
        Label11.Caption = ""
        TextBox34.Value = ""
        TextBox35.Value = ""
        Label12.Caption = ""
        Label13.Caption = ""
        Label14.Caption = ""
        TextBox39.Value = ""
        TextBox40.Value = ""
        TextBox41.Value = ""
        TextBox42.Value = ""
        Label15.Caption = ""
        TextBox44.Value = ""
        TextBox45.Value = ""
        
        EnableTextboxes False
        
    End If
End Sub

Sub EnableTextboxes(enable As Boolean)
    TextBox2.Enabled = enable
    TextBox3.Enabled = enable
    TextBox4.Enabled = enable
    TextBox5.Enabled = enable
    TextBox6.Enabled = enable
    TextBox7.Enabled = enable
    TextBox8.Enabled = enable
    TextBox9.Enabled = enable
    TextBox10.Enabled = enable
    TextBox12.Enabled = enable
    TextBox14.Enabled = enable
    TextBox15.Enabled = enable
    TextBox16.Enabled = enable
    TextBox18.Enabled = enable
    TextBox21.Enabled = enable
    TextBox22.Enabled = enable
    TextBox23.Enabled = enable
    TextBox24.Enabled = enable
    TextBox25.Enabled = enable
    TextBox26.Enabled = enable
    TextBox27.Enabled = enable
    TextBox28.Enabled = enable
    TextBox34.Enabled = enable
    TextBox35.Enabled = enable
    TextBox39.Enabled = enable
    TextBox40.Enabled = enable
    TextBox41.Enabled = enable
    TextBox42.Enabled = enable
    TextBox44.Enabled = enable
    TextBox45.Enabled = enable
End Sub
Private Sub TextBox1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    SearchData TextBox1.Value
End Sub

Private Sub UserForm_Initialize()
    DisableAllTextboxes
    TextBox1.Enabled = True
    
    Label2.Caption = ""
    Label3.Caption = ""
    Label4.Caption = ""
    Label5.Caption = ""
    Label6.Caption = ""
    Label7.Caption = ""
    Label8.Caption = ""
    Label9.Caption = ""
    Label10.Caption = ""
    Label11.Caption = ""
    Label12.Caption = ""
    Label13.Caption = ""
    Label14.Caption = ""
    Label15.Caption = ""
End Sub

Sub DisableAllTextboxes()
    TextBox2.Enabled = False
    TextBox3.Enabled = False
    TextBox4.Enabled = False
    TextBox5.Enabled = False
    TextBox6.Enabled = False
    TextBox7.Enabled = False
    TextBox8.Enabled = False
    TextBox9.Enabled = False
    TextBox10.Enabled = False
    TextBox12.Enabled = False
    TextBox14.Enabled = False
    TextBox15.Enabled = False
    TextBox16.Enabled = False
    TextBox18.Enabled = False
    TextBox21.Enabled = False
    TextBox22.Enabled = False
    TextBox23.Enabled = False
    TextBox24.Enabled = False
    TextBox25.Enabled = False
    TextBox26.Enabled = False
    TextBox27.Enabled = False
    TextBox28.Enabled = False
    TextBox34.Enabled = False
    TextBox35.Enabled = False
    TextBox39.Enabled = False
    TextBox40.Enabled = False
    TextBox41.Enabled = False
    TextBox42.Enabled = False
    TextBox44.Enabled = False
    TextBox45.Enabled = False
End Sub

Sub UpdateData(searchValue As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim foundCell As Range

    Set ws = ThisWorkbook.Sheets("Claims Data")

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Set foundCell = ws.Range("A1:B" & lastRow & ",D1:D" & lastRow).Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole)

    If Not foundCell Is Nothing Then
    
    Dim value10 As String
    Dim value12 As Integer
    Dim value16 As Long
    Dim value18 As Integer
    Dim value19 As Long
    Dim value28 As Long
    Dim value29 As Integer
    Dim value30 As Integer
    Dim value31 As String
    Dim value32 As String
    Dim value35 As Currency
    Dim value36 As Currency
    Dim value37 As Currency
    Dim value42 As Long
    
    Dim yearsDD As Long
    Dim monthsDD As Long
    Dim daysDD As Long
    
    Dim startDate As Date
    Dim endDate As Date
    Dim datecc As Date
    Dim dcc As Date
    Dim ati As Date
    Dim notdate As Date
    Dim factor As Double
    
    startDate = DateValue(TextBox7.Value)
    endDate = DateValue(TextBox18.Value)
    notdate = DateValue(TextBox12.Value)
    datecc = DateValue(TextBox27.Value)
    dcc = DateValue(TextBox26.Value)
    ati = DateValue(TextBox16.Value)
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
   ' value39 = value37 / TextBox12.Value * TextBox18
    value42 = Application.WorksheetFunction.NetworkDays(notdate, datecc)
    
    
    If TextBox24.Value = "Closed" Then
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

    If TextBox23.Value = "Pending" Then
       value36 = value34 * factor
    End If
       
    If value35 = 0 And value36 = 0 Then
        value37 = value34
    Else
        value37 = value35 + value36
    End If
    
        ws.Cells(foundCell.Row, 1).Value = TextBox2.Value
        ws.Cells(foundCell.Row, 2).Value = TextBox3.Value
        ws.Cells(foundCell.Row, 3).Value = TextBox4.Value
        ws.Cells(foundCell.Row, 4).Value = TextBox5.Value
        ws.Cells(foundCell.Row, 5).Value = TextBox6.Value
        ws.Cells(foundCell.Row, 6).Value = TextBox7.Value
        ws.Cells(foundCell.Row, 7).Value = TextBox8.Value
        ws.Cells(foundCell.Row, 8).Value = TextBox9.Value
        ws.Cells(foundCell.Row, 9).Value = TextBox10.Value
        ws.Cells(foundCell.Row, 10).Value = value10
        ws.Cells(foundCell.Row, 11).Value = TextBox12.Value
        ws.Cells(foundCell.Row, 12).Value = value12
        ws.Cells(foundCell.Row, 13).Value = TextBox14.Value
        ws.Cells(foundCell.Row, 14).Value = TextBox15.Value
        ws.Cells(foundCell.Row, 15).Value = TextBox16.Value
        ws.Cells(foundCell.Row, 16).Value = value16
        ws.Cells(foundCell.Row, 17).Value = TextBox18.Value
        ws.Cells(foundCell.Row, 18).Value = value18
        ws.Cells(foundCell.Row, 19).Value = value19
        ws.Cells(foundCell.Row, 20).Value = TextBox21.Value
        ws.Cells(foundCell.Row, 21).Value = TextBox22.Value
        ws.Cells(foundCell.Row, 22).Value = TextBox23.Value
        ws.Cells(foundCell.Row, 23).Value = TextBox24.Value
        ws.Cells(foundCell.Row, 24).Value = TextBox25.Value
        ws.Cells(foundCell.Row, 25).Value = TextBox26.Value
        ws.Cells(foundCell.Row, 26).Value = TextBox27.Value
        ws.Cells(foundCell.Row, 27).Value = TextBox28.Value
        ws.Cells(foundCell.Row, 28).Value = value28
        ws.Cells(foundCell.Row, 29).Value = value29
        ws.Cells(foundCell.Row, 30).Value = value30
        ws.Cells(foundCell.Row, 31).Value = value31
        ws.Cells(foundCell.Row, 32).Value = value32
        ws.Cells(foundCell.Row, 33).Value = TextBox34.Value
        ws.Cells(foundCell.Row, 34).Value = TextBox35.Value
        ws.Cells(foundCell.Row, 35).Value = value35
        ws.Cells(foundCell.Row, 36).Value = value36
        ws.Cells(foundCell.Row, 37).Value = value37
        ws.Cells(foundCell.Row, 38).Value = TextBox39.Value
        ws.Cells(foundCell.Row, 39).Value = TextBox40.Value
        ws.Cells(foundCell.Row, 40).Value = TextBox41.Value
        ws.Cells(foundCell.Row, 41).Value = TextBox42.Value
        ws.Cells(foundCell.Row, 42).Value = value42
        ws.Cells(foundCell.Row, 43).Value = TextBox44.Value
        ws.Cells(foundCell.Row, 44).Value = TextBox45.Value

        MsgBox "Data updated successfully."
    Else
        MsgBox "No data found to update."
    End If
End Sub
