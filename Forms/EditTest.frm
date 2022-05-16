

Private Sub UserForm_Initialize()

Dim deliverables As Worksheet, tests As Worksheet
Set deliverables = ThisWorkbook.Sheets("Deliverables")
Set tests = ThisWorkbook.Sheets("Tests")

Dim iRow As Long, iRow2 As Long, check As Boolean
iRow = [Counta(Deliverables!A:A)]
iRow2 = [Counta(Tests!A:A)]

cmbExistingCourses2.Clear
        
        'Loop through each assessment to add courses to the combo-box
    For i = 2 To iRow
        
        check = True
            
        For k = 0 To cmbExistingCourses2.ListCount - 1 'Loop through the combo-box values
            If deliverables.Cells(i, "B").Value = cmbExistingCourses2.List(k) Then 'If a combo-box value equals a value on the sheet
                check = False
                
            Exit For 'Exit For loop to prevent adding a repeat course name in the combo-box
            End If
        Next k
            
            If check Then
                cmbExistingCourses2.AddItem deliverables.Cells(i, "B").Value 'Add the course to the combo-box
                
            End If
            
        Next i
        
        For i = 2 To iRow2
        check = True
            
            For k = 0 To cmbExistingCourses2.ListCount - 1 'Loop through the combo-box values
                If tests.Cells(i, "B").Value = cmbExistingCourses2.List(k) Then 'If a combo-box value equals a value on the sheet
                check = False
                
                Exit For 'Exit For loop to prevent adding a repeat course name in the combo-box
                End If
            Next k
            
            If check Then
                cmbExistingCourses2.AddItem tests.Cells(i, "B").Value 'Add the course to the combo-box
                
            End If
            
        Next i


Dim indexNum As Long
indexNum = EditAssessment_UserForm1.cmbTests.ListIndex

'Fills in which value is true
If tests.Cells(indexNum + 2, 1).Text = "Quiz" Then
    optQuiz.Value = True
ElseIf tests.Cells(indexNum + 2, 1).Text = "Midterm" Then
    optMidterm.Value = True
ElseIf tests.Cells(indexNum + 2, 1).Text = "Unit/Term Test" Then
    optUnitTest.Value = True
ElseIf tests.Cells(indexNum + 2, 1).Text = "Final" Then
    optFinal.Value = True
End If

'Retrieve data from test sheet
cmbExistingCourses2 = tests.Cells(indexNum + 2, 2).Text
txtAssessmentName2 = tests.Cells(indexNum + 2, 3).Text
txtLengthHours2 = tests.Cells(indexNum + 2, 4).Text
txtWeight2 = tests.Cells(indexNum + 2, 5).Text
txtAntGrade2 = tests.Cells(indexNum + 2, 6).Text
txtAntStudyHours2 = tests.Cells(indexNum + 2, 7).Text
txtDeadline2 = tests.Cells(indexNum + 2, 8).Text

                 


End Sub

Private Sub cmbSubmit_Click()

'Check for required field is left empty or if inputted data is invalid
Dim msgValue As VbMsgBoxResult
    If cmbExistingCourses2.Value = "" Then
        msgValue = MsgBox("Please select a course or add a new course")
            
    ElseIf txtAssessmentName2.Text = "" Then
        msgValue = MsgBox("Please enter the Assessment Name")
            
    ElseIf txtDeadline2.Text = "" Then
        msgValue = MsgBox("Please enter a deadline")
        
    Else
        'User Confirmation before editing the data of a task
        msgValue = MsgBox("Do you want to edit this task?", vbYesNo + vbInformation, "Confirmation")
    
        If msgValue = vbNo Then Exit Sub
        
        Call Submit_TestEditData
        Unload Me
        Unload EditAssessment_UserForm1
    
    End If
    


End Sub

Private Sub cmbSubmit_Enter()

'Check for required field is left empty or if inputted data is invalid
Dim msgValue As VbMsgBoxResult
    If cmbExistingCourses2.Value = "" Then
        msgValue = MsgBox("Please select a course or add a new course")
            
    ElseIf txtAssessmentName2.Text = "" Then
        msgValue = MsgBox("Please enter the Assessment Name")
            
    ElseIf txtDeadline2.Text = "" Then
        msgValue = MsgBox("Please enter a deadline")
        
    Else
        'User Confirmation before editing the data of a task
        msgValue = MsgBox("Do you want to edit this task?", vbYesNo + vbInformation, "Confirmation")
    
        If msgValue = vbNo Then Exit Sub
        
        Call Submit_TestEditData
        Unload Me
        Unload EditAssessment_UserForm1
    
    End If
    


End Sub
