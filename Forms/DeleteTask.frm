Private Sub UserForm_Initialize()

'Declare Sheet to reference
Dim task As Worksheet
Set task = ThisWorkbook.Sheets("Task_List")

Dim i As Integer 'Used For Looping

Dim tblMeetingRow As Long, tblEventRow As Long, tblThingsToDoRow As Long 'Tables in Task_List sheet
tblMeetingRow = task.ListObjects("Table6").ListRows.Count 'Meetings Table
tblEventRow = task.ListObjects("Table4").ListRows.Count 'Events Table
tblThingsToDoRow = task.ListObjects("Table8").ListRows.Count 'Things to Do Table

'Clear existing values in combo-boxes
cmbMeeting.Clear
cmbEvent.Clear
cmbThingToDo.Clear

'Loop through each value in Meetings Table
For i = 1 To tblMeetingRow
    cmbMeeting.AddItem task.ListObjects("Table6").DataBodyRange.Cells(i, "A").Value 'Add Meeting Name to Meeting combo-box
Next i

'Loop through each value in Events Table
For i = 1 To tblEventRow
    cmbEvent.AddItem task.ListObjects("Table4").DataBodyRange.Cells(i, "A").Value 'Add Event Name to Events combo-box
Next i

'Loop through each value in Things to Do Table
For i = 1 To tblThingsToDoRow
    cmbThingToDo.AddItem task.ListObjects("Table8").DataBodyRange.Cells(i, "A").Value 'Add Thing to Do Name to Things To Do combo-box
Next i

End Sub

Private Sub cmbEvent_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

cmbEvent.SpecialEffect = fmSpecialEffectEtched
cmbMeeting.SpecialEffect = fmSpecialEffectFlat
cmbThingToDo.SpecialEffect = fmSpecialEffectFlat

End Sub

Private Sub cmbThingToDo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

cmbThingToDo.SpecialEffect = fmSpecialEffectEtched
cmbMeeting.SpecialEffect = fmSpecialEffectFlat
cmbEvent.SpecialEffect = fmSpecialEffectFlat

End Sub

Private Sub cmbMeeting_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

cmbMeeting.SpecialEffect = fmSpecialEffectEtched
cmbEvent.SpecialEffect = fmSpecialEffectFlat
cmbThingToDo.SpecialEffect = fmSpecialEffectFlat

End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

cmbEvent.SpecialEffect = fmSpecialEffectFlat
cmbThingToDo.SpecialEffect = fmSpecialEffectFlat
cmbMeeting.SpecialEffect = fmSpecialEffectFlat

End Sub


Private Sub cmbMeeting_Change()

'If User selects a value in the Meetings combo-box, erase any values in the Events and Things to Do combo-box
'Ensures User only deletes one task

If Me.cmbMeeting.Value <> "" Then
    Me.cmbEvent.Value = ""
    Me.cmbThingToDo.Value = ""

End If

End Sub
Private Sub cmbEvent_Change()

'If User selects a value in the Events combo-box, erase any values in the Meetings and Things to Do combo-box
'Ensures User only deletes one task

If Me.cmbEvent.Value <> "" Then
    Me.cmbMeeting.Value = ""
    Me.cmbThingToDo.Value = ""

End If

End Sub
Private Sub cmbThingToDo_Change()

'If User selects a value in the Things to Do combo-box, erase any values in the Meetings and Events combo-box
'Ensures User only deletes one task

If Me.cmbThingToDo.Value <> "" Then
    Me.cmbMeeting.Value = ""
    Me.cmbEvent.Value = ""
    
End If

End Sub

Private Sub cmdReset_Click() 'Reset Values in the combo-boxes

cmbMeeting.Value = ""
cmbEvent.Value = ""
cmbThingToDo.Value = ""

End Sub
Private Sub cmdDeleteTask_Click()
    
Dim msgValue As VbMsgBoxResult

'Checks if the user selected a value
    If cmbMeeting.Value = "" And cmbEvent.Value = "" And cmbThingToDo.Value = "" Then 'If the user does not select any value
                msgValue = MsgBox("Please select a task to delete")
    Else
        'User Confirmation before deleting the task
        msgValue = MsgBox("Are you sure you want to delete this task?", vbYesNo + vbInformation, "Confirmation")
        
        If msgValue = vbNo Then Exit Sub
    
            Call DeleteTask
            Unload Me

    End If
    
End Sub
