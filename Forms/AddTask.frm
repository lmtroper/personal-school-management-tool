

Private Sub UserForm_Initialize()
    
    'Whenever the UserForm is opened, all data is reset and values are re-added
    Call Reset_AddTaskForm
    
End Sub

Private Sub txtDeadlineTask_Exit(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next 'Allows code to continue runnning even if an error occurs

Me.txtDeadlineTask = CDate(Me.txtDeadlineTask) 'This converts the value of the entry in the textbox to the "Date" type

End Sub

Private Sub cmbReset_Click()
    
    Call Reset_AddTaskForm 'Clear values in UserForm

End Sub

Private Sub cmbSubmit_Click()

Dim tasks As Worksheet
Set tasks = ThisWorkbook.Sheets("Task_List")

'Check for required field is left empty or if inputted data is invalid
Dim msgValue As VbMsgBoxResult
    
        If txtTaskName.Text = "" Then
            msgValue = MsgBox("Please enter a task name")
            
        ElseIf cmbTaskCategory.Value = "" Then
            msgValue = MsgBox("Please select a task category")
            
        ElseIf txtDeadlineTask.Value = "" Or Not IsDate(txtDeadlineTask.Value) Then
            msgValue = MsgBox("Please enter a valid date")
        
        Else
            Call Submit_Task
            Call Reset_AddTaskForm
            Unload Me
            
    End If

End Sub

Private Sub cmbSubmit_Enter()

Dim tasks As Worksheet
Set tasks = ThisWorkbook.Sheets("Task_List")

'Check if a required field is left empty or if inputted data is invalid
Dim msgValue As VbMsgBoxResult
    
        If txtTaskName.Text = "" Then
            msgValue = MsgBox("Please enter a task name")
            
        ElseIf cmbTaskCategory.Value = "" Then
            msgValue = MsgBox("Please select a task category")
            
        ElseIf txtDeadlineTask.Value = "" Or Not IsDate(txtDeadlineTask.Value) Then
            msgValue = MsgBox("Please enter a valid date")
        
        Else
            Call Submit_Task
            Call Reset_AddTaskForm
            Unload Me
            
    End If


End Sub
