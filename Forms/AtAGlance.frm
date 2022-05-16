
Private Sub txt_startDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next 'Allows code to continue runnning even if an error occurs

Me.txt_startDate = CDate(Me.txt_startDate) 'This converts the value of the entry in the textbox to the "Date" type

'If user types in today, output today's date
If Me.txt_startDate = "today" Then
    Me.txt_startDate = Date
End If


End Sub

Private Sub txt_endDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next 'Allows code to continue runnning even if an error occurs

Me.txt_endDate = CDate(Me.txt_endDate) 'This converts the value of the entry in the textbox to the "Date" type


End Sub

Private Sub cmdSubmit_Click()

'Check for required field is left empty or if inputted data is invalid
    If txt_startDate = "" Then
        MsgBox ("Please enter a start date")
        
    ElseIf txt_endDate = "" Then
        MsgBox ("Please enter an end date")
        
    ElseIf Not IsDate(txt_startDate.Text) Then
        MsgBox "Please enter the start date in the correct format"
        
    ElseIf Not IsDate(txt_startDate.Text) Then
        MsgBox "Please enter the end date in the correct format"
    
    ElseIf txt_startDate.Value > txt_endDate.Value Then
        MsgBox "Please enter an end date that occurs after the start date"
    
    Else
        Application.ScreenUpdating = False
        Call Clear_AtAGlance_Sheet 'Clear any existing values following use
        Call AtAGlance_Output
        Unload Me
        Application.ScreenUpdating = True

    End If

End Sub
