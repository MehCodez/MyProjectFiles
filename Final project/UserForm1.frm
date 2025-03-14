VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   5988
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11604
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Unload Me

End Sub

Private Sub CommandButton2_Click()
Dim ws As Worksheet
    Dim lastRow As Long

    ' Set worksheet reference
    Set ws = ThisWorkbook.Sheets("Data") ' Update to your correct sheet name

    ' Find the last used row
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

   ' Transfer data from correct TextBoxes
    ws.Cells(lastRow, 1).Value = Me.TextBox1.Value ' Change to actual TextBox name
    ws.Cells(lastRow, 2).Value = Me.TextBox2.Value ' Change to actual TextBox name
    ws.Cells(lastRow, 3).Value = Me.TextBox3.Value
    ws.Cells(lastRow, 4).Value = Me.TextBox4.Value
    ws.Cells(lastRow, 5).Value = Me.TextBox5.Value
    ws.Cells(lastRow, 6).Value = Me.TextBox6.Value
    ws.Cells(lastRow, 7).Value = Me.TextBox7.Value
    ws.Cells(lastRow, 8).Value = Me.TextBox8.Value
    ws.Cells(lastRow, 9).Value = Me.TextBox9.Value
    ws.Cells(lastRow, 10).Value = Me.TextBox10.Value
    ws.Cells(lastRow, 11).Value = Me.TextBox11.Value
    ws.Cells(lastRow, 12).Value = Me.TextBox12.Value
    ws.Cells(lastRow, 13).Value = Me.TextBox13.Value

    ' Clear the TextBoxes after submission
    Me.TextBox1.Value = ""
    Me.TextBox2.Value = ""
    Me.TextBox3.Value = ""
    Me.TextBox4.Value = ""
    Me.TextBox5.Value = ""
    Me.TextBox6.Value = ""
    Me.TextBox7.Value = ""
    Me.TextBox8.Value = ""
    Me.TextBox9.Value = ""
    Me.TextBox10.Value = ""
    Me.TextBox11.Value = ""
    Me.TextBox12.Value = ""
    Me.TextBox13.Value = ""

    ' Confirmation message
    MsgBox "Data submitted successfully!", vbInformation, "Success"


End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label5_Click()

End Sub

Private Sub TextBox1_Change()

End Sub
