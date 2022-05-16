Sub Show_AddTaskUserForm() 'Opens the Add Task User Form when user clicks on the Add Task button

AddTask_UserForm.Show

End Sub

Sub Show_DeleteTaskUserForm() 'Opens the Delete Task User Form when user clicks on the Delete Task button

DeleteTask_UserForm.Show

End Sub

Sub Reset_AddTaskForm() 'Clears the textboxes and resets the combo-box values
   
    With AddTask_UserForm
    
        'Clearing values from combo-box
        .cmbTaskCategory.Clear
        
        'Add default values for combo-box
        .cmbTaskCategory.AddItem "Meeting"
        .cmbTaskCategory.AddItem "Event"
        .cmbTaskCategory.AddItem "Thing to Do"
        
        'Clear textboxes to prevent prior data from re-appearing
        .txtTaskName.Value = ""
        .txtLengthHoursTask.Value = ""
        .txtNoteSection1.Value = ""
        .txtDeadlineTask.Value = ""
    
    End With

End Sub

Sub Submit_Task() 'Adds New Task into Specific Table Depending on Type of Task

'Declare Worksheet to reference
Dim task As Worksheet
Set task = ThisWorkbook.Sheets("Task_List")

'Tables in the Task_List Sheet
Dim tblMeetingRow As Long, tblEventRow As Long, tblThingsToDoRow As Long
tblMeetingRow = task.ListObjects("Table6").ListRows.Count 'Number of rows in Meetings Table
tblEventRow = task.ListObjects("Table4").ListRows.Count 'Number of rows in Events Table
tblThingsToDoRow = task.ListObjects("Table8").ListRows.Count 'Number of rows in Things to Do Table

'If the selected Category = Meeting
If AddTask_UserForm.cmbTaskCategory.Text = "Meeting" Then
        
        'If the table length is equal to 1 and the first entry of the table is blank
        'Then add the UserForm data into the first row
        If tblMeetingRow = 1 And task.ListObjects("Table6").DataBodyRange.Cells(tblMeetingRow, "A").Value = "" Then
            task.ListObjects("Table6").DataBodyRange.Cells(1, "A").Value = AddTask_UserForm.txtTaskName.Text
            task.ListObjects("Table6").DataBodyRange.Cells(1, "B").Value = AddTask_UserForm.txtLengthHoursTask.Text
            task.ListObjects("Table6").DataBodyRange.Cells(1, "C").Value = AddTask_UserForm.txtDeadlineTask.Text
            task.ListObjects("Table6").DataBodyRange.Cells(1, "D").Value = AddTask_UserForm.txtNoteSection1.Text
   
        Else
            task.ListObjects("Table6").ListRows.Add 'Add a new row to prevent re-writing data
            tblMeetingRow = task.ListObjects("Table6").ListRows.Count 'Re-count the rows in the table
            
            'Add the data into the last row of the table (the row that just got added)
            task.ListObjects("Table6").DataBodyRange.Cells(tblMeetingRow, "A").Value = AddTask_UserForm.txtTaskName.Text
            task.ListObjects("Table6").DataBodyRange.Cells(tblMeetingRow, "B").Value = AddTask_UserForm.txtLengthHoursTask.Text
            task.ListObjects("Table6").DataBodyRange.Cells(tblMeetingRow, "C").Value = AddTask_UserForm.txtDeadlineTask.Text
            task.ListObjects("Table6").DataBodyRange.Cells(tblMeetingRow, "D").Value = AddTask_UserForm.txtNoteSection1.Text
   
        End If
    
'If the selected Category = Event
ElseIf AddTask_UserForm.cmbTaskCategory.Text = "Event" Then

        'If the table length is equal to 1 and the first entry of the table is blank
        'Then add the UserForm data into the first row
        If tblEventRow = 1 And task.ListObjects("Table4").DataBodyRange.Cells(tblEventRow, "A").Value = "" Then
            task.ListObjects("Table4").DataBodyRange.Cells(tblEventRow, "A").Value = AddTask_UserForm.txtTaskName.Text
            task.ListObjects("Table4").DataBodyRange.Cells(tblEventRow, "B").Value = AddTask_UserForm.txtLengthHoursTask.Text
            task.ListObjects("Table4").DataBodyRange.Cells(tblEventRow, "C").Value = AddTask_UserForm.txtDeadlineTask.Text
            task.ListObjects("Table4").DataBodyRange.Cells(tblEventRow, "D").Value = AddTask_UserForm.txtNoteSection1.Text
            
        Else
            task.ListObjects("Table4").ListRows.Add 'Add a new row to prevent re-writing data
            tblEventRow = task.ListObjects("Table4").ListRows.Count 'Re-count the rows in the table
            
            'Add the data into the last row of the table (the row that just got added)
            task.ListObjects("Table4").DataBodyRange.Cells(tblEventRow, "A").Value = AddTask_UserForm.txtTaskName.Text
            task.ListObjects("Table4").DataBodyRange.Cells(tblEventRow, "B").Value = AddTask_UserForm.txtLengthHoursTask.Text
            task.ListObjects("Table4").DataBodyRange.Cells(tblEventRow, "C").Value = AddTask_UserForm.txtDeadlineTask.Text
            task.ListObjects("Table4").DataBodyRange.Cells(tblEventRow, "D").Value = AddTask_UserForm.txtNoteSection1.Text
   
        End If

'If the selected Category = Thing to Do
ElseIf AddTask_UserForm.cmbTaskCategory.Text = "Thing to Do" Then

        'If the table length is equal to 1 and the first entry of the table is blank
        'Then add the UserForm data into the first row
        If tblThingsToDoRow = 1 And task.ListObjects("Table8").DataBodyRange.Cells(tblThingsToDoRow, "A").Value = "" Then
            task.ListObjects("Table8").DataBodyRange.Cells(tblThingsToDoRow, "A").Value = AddTask_UserForm.txtTaskName.Text
            task.ListObjects("Table8").DataBodyRange.Cells(tblThingsToDoRow, "B").Value = AddTask_UserForm.txtLengthHoursTask.Text
            task.ListObjects("Table8").DataBodyRange.Cells(tblThingsToDoRow, "C").Value = AddTask_UserForm.txtDeadlineTask.Text
            task.ListObjects("Table8").DataBodyRange.Cells(tblThingsToDoRow, "D").Value = AddTask_UserForm.txtNoteSection1.Text
   
        Else
            task.ListObjects("Table8").ListRows.Add 'Add a new row to prevent re-writing data
            tblThingsToDoRow = task.ListObjects("Table8").ListRows.Count 'Re-count the rows in the table
            
            'Add the data into the last row of the table (the row that just got added)
            task.ListObjects("Table8").DataBodyRange.Cells(tblThingsToDoRow, "A").Value = AddTask_UserForm.txtTaskName.Text
            task.ListObjects("Table8").DataBodyRange.Cells(tblThingsToDoRow, "B").Value = AddTask_UserForm.txtLengthHoursTask.Text
            task.ListObjects("Table8").DataBodyRange.Cells(tblThingsToDoRow, "C").Value = AddTask_UserForm.txtDeadlineTask.Text
            task.ListObjects("Table8").DataBodyRange.Cells(tblThingsToDoRow, "D").Value = AddTask_UserForm.txtNoteSection1.Text
   
        End If

End If

End Sub

Sub DeleteTask() 'Deletes selected task from the specified table

'Delcare worksheet to reference
Dim task As Worksheet
Set task = ThisWorkbook.Sheets("Task_List")

Dim i As Integer 'Used for Looping

'Tables in the Task_List Sheet
Dim tblMeetingRow As Long, tblEventRow As Long, tblThingsToDoRow As Long
tblMeetingRow = task.ListObjects("Table6").ListRows.Count 'Meetings Table
tblEventRow = task.ListObjects("Table4").ListRows.Count 'Events Table
tblThingsToDoRow = task.ListObjects("Table8").ListRows.Count 'Things to Do Table

With DeleteTask_UserForm
    If .cmbMeeting.Value <> "" Then 'Determines if the user selected a task under the Meeting Category
        For i = 1 To tblMeetingRow 'Loop through the rows in the Meetings Table
            If .cmbMeeting.Value = task.ListObjects("Table6").DataBodyRange.Cells(i, "A").Value Then 'Compare selected task name with name in table
                 
                 If tblMeetingRow = 1 Then
                    'If there is only one row, clear the values
                    task.ListObjects("Table6").DataBodyRange.Cells(1, "A").Value = ""
                    task.ListObjects("Table6").DataBodyRange.Cells(1, "B").Value = ""
                    task.ListObjects("Table6").DataBodyRange.Cells(1, "C").Value = ""
                    task.ListObjects("Table6").DataBodyRange.Cells(1, "D").Value = ""
                
                Else
                    task.ListObjects("Table6").ListRows(i).Delete 'Delete the entire row from the table
                
                End If
                
            Exit For 'Exit the for loop after the task is found
            End If
        Next i
        
    ElseIf .cmbEvent.Value <> "" Then 'Determines if the user selected a task under the Event Category
        For i = 1 To tblEventRow 'Loop through the rows in the Events Table
            If .cmbEvent.Value = task.ListObjects("Table4").DataBodyRange.Cells(i, "A").Value Then 'Compare selected task name with name in table
                
                If tblEventRow = 1 Then
                    'If there is only one row, clear the values
                    task.ListObjects("Table4").DataBodyRange.Cells(1, "A").Value = ""
                    task.ListObjects("Table4").DataBodyRange.Cells(1, "B").Value = ""
                    task.ListObjects("Table4").DataBodyRange.Cells(1, "C").Value = ""
                    task.ListObjects("Table4").DataBodyRange.Cells(1, "D").Value = ""
                
                Else
                    task.ListObjects("Table4").ListRows(i).Delete 'Delete the entire row from the table
                
                End If
                
            Exit For 'Exit the for loop after the task is found
            End If
        Next i
        
    ElseIf .cmbThingToDo.Value <> "" Then 'Determines if the user selected a task under the Thing to Do Category
        For i = 1 To tblThingsToDoRow 'Loop through the rows in the Things to Do Table
            If .cmbThingToDo.Value = task.ListObjects("Table8").DataBodyRange.Cells(i, "A").Value Then 'Compare selected task name with name in table
                 
                If tblThingsToDoRow = 1 Then
                    'If there is only one row, clear the values
                    task.ListObjects("Table8").DataBodyRange.Cells(1, "A").Value = ""
                    task.ListObjects("Table8").DataBodyRange.Cells(1, "B").Value = ""
                    task.ListObjects("Table8").DataBodyRange.Cells(1, "C").Value = ""
                    task.ListObjects("Table8").DataBodyRange.Cells(1, "D").Value = ""
                
                Else
                    task.ListObjects("Table8").ListRows(i).Delete 'Delete the entire row from the table
                
                End If
                    
            Exit For 'Exit the for loop after the task is found
            End If
        Next i
        
    End If
End With

End Sub
