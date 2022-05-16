

Private Sub cmbSubmit_Click()

Dim msgValue As VbMsgBoxResult
    If cmbExistingCourses3.Value = "" Then
        msgValue = MsgBox("Please select a course")
            
    ElseIf txtAssessmentName3.Text = "" Then
        msgValue = MsgBox("Please enter the Deliverable Name")
            
    ElseIf txtDeadline3.Text = "" Then
        msgValue = MsgBox("Please enter a deadline")
        
    ElseIf cmbDeliverableType.Value = "" Then
        msgValue = MsgBox("Please select a deliverable type")
        
    Else
        'User Confirmation before editing the data of a task
        msgValue = MsgBox("Do you want to edit this task?", vbYesNo + vbInformation, "Confirmation")
    
        If msgValue = vbNo Then Exit Sub
        
        Call Submit_DeliverableEditData
        Unload Me
        Unload EditAssessment_UserForm1
        Call SortTablesbyDeadline
    
    End If

End Sub

Private Sub txtDeadline3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next 'Allows code to continue runnning even if an error occurs

Me.txtDeadline3 = CDate(Me.txtDeadline3) 'This converts the value of the entry in the textbox to the "Date" type

End Sub


Private Sub UserForm_Initialize()

Dim deliverables As Worksheet, tests As Worksheet
Set deliverables = ThisWorkbook.Sheets("Deliverables")
Set tests = ThisWorkbook.Sheets("Tests")

Dim iRow As Long, check As Boolean
iRow = [Counta(Deliverables!A:A)]
iRow2 = [Counta(Tests!A:A)]

cmbExistingCourses3.Clear
        
        'Loop through each assessment to add courses to the combo-box
    For i = 2 To iRow
        
        check = True
            
        For k = 0 To cmbExistingCourses3.ListCount - 1 'Loop through the combo-box values
            If deliverables.Cells(i, "B").Value = cmbExistingCourses3.List(k) Then 'If a combo-box value equals a value on the sheet
                check = False
                
            Exit For 'Exit For loop to prevent adding a repeat course name in the combo-box
            End If
        Next k
            
            If check Then
                cmbExistingCourses3.AddItem deliverables.Cells(i, "B").Value 'Add the course to the combo-box
                
            End If
            
        Next i
        
        For i = 2 To iRow2
        check = True
            
            For k = 0 To cmbExistingCourses3.ListCount - 1 'Loop through the combo-box values
                If tests.Cells(i, "B").Value = cmbExistingCourses3.List(k) Then 'If a combo-box value equals a value on the sheet
                check = False
                
                Exit For 'Exit For loop to prevent adding a repeat course name in the combo-box
                End If
            Next k
            
            If check Then
                cmbExistingCourses3.AddItem tests.Cells(i, "B").Value 'Add the course to the combo-box
                
            End If
            
        Next i

cmbDeliverableType.Clear
cmbDeliverableType.AddItem "Assignment"
cmbDeliverableType.AddItem "Essay"
cmbDeliverableType.AddItem "Presentation"
cmbDeliverableType.AddItem "Term Project"
cmbDeliverableType.AddItem "Personal Project"


Dim indexNum As Long
indexNum = EditAssessment_UserForm1.cmbDeliverables.ListIndex

'Retrieves the data
'The row corresponding to the value in the combo-box is (combo-box element value + 2)
cmbDeliverableType.Value = deliverables.Cells(indexNum + 2, 1).Text
cmbExistingCourses3 = deliverables.Cells(indexNum + 2, 2).Text
txtAssessmentName3 = deliverables.Cells(indexNum + 2, 3).Text
txtWeight3 = deliverables.Cells(indexNum + 2, 4).Text
txtAntGrade3 = deliverables.Cells(indexNum + 2, 5).Text
txtAntStudyHours3 = deliverables.Cells(indexNum + 2, 6).Text
txtDeadline3 = deliverables.Cells(indexNum + 2, 7).Text

End Sub
