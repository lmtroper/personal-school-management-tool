
Private Sub cmbSemester_Change()

'Determines the value of the combo-box based on the academic semester (using time-frames)
'If user tries to change the value, the value will stay as the academic semester they are currently in

If Date > 2021 - 1 - 1 And Me.cmbSemester.Value <> "1A" Then
    cmbSemester.Value = "1A"
    
ElseIf Date >= 2021 - 1 - 11 And Date <= 2021 - 8 - 31 And Me.cmbSemester.Value <> "1B" Then
    cmbSemester.Value = "1B"

ElseIf Date >= 2021 - 9 - 1 And Date <= 2021 - 4 - 30 And Me.cmbSemester.Value <> "2A" Then
    cmbSemester.Value = "2A"

ElseIf Date >= 2022 - 5 - 1 And Date <= 2022 - 12 - 31 And Me.cmbSemester.Value <> "2B" Then
    cmbSemester.Value = "2B"

ElseIf Date >= 2023 - 1 - 1 And Date <= 2023 - 8 - 31 And Me.cmbSemester.Value <> "3A" Then
    cmbSemester.Value = "3A"

ElseIf Date >= 2023 - 9 - 1 And Date <= 2024 - 4 - 31 And Me.cmbSemester.Value <> "3B" Then
    cmbSemester.Value = "3B"
    
ElseIf Date >= 2024 - 5 - 1 And Date <= 2024 - 12 - 31 And Me.cmbSemester.Value <> "4A" Then
    cmbSemester.Value = "4A"

ElseIf Date > 2024 - 12 - 31 And Me.cmbSemester.Value <> "4B" Then
    cmbSemester.Value = "4B"

End If


End Sub


Private Sub cmdSemester_Click()

MsgBox "Indicates the academic semester"

End Sub

Private Sub UserForm_Initialize()

Dim deliverables As Worksheet, test As Worksheet
Set deliverables = ThisWorkbook.Sheets("Deliverables")
Set test = ThisWorkbook.Sheets("Tests")

Dim iRow As Long, iRow2 As Long, check As Boolean
iRow = [Counta(Deliverables!A:A)]
iRow2 = [Counta(Tests!A:A)]

cmbADExistingCourses.Clear
        
    For i = 2 To iRow 'Loop through each assessment
        
        check = True
            
        For k = 0 To cmbADExistingCourses.ListCount - 1 'Loop through the combo-box values
            If deliverables.Cells(i, "B").Value = cmbADExistingCourses.List(k) Then 'If a combo-box value equals a value on the sheet
                check = False
                
            Exit For 'Exit For loop to prevent adding a repeat course name in the combo-box
            End If
        Next k
            
            If check Then
                cmbADExistingCourses.AddItem deliverables.Cells(i, "B").Value 'Add the course to the combo-box
                
            End If
            
        Next i
        
        For i = 2 To iRow2
        check = True
            For k = 0 To cmbADExistingCourses.ListCount - 1 'Loop through the combo-box values
                If test.Cells(i, "B").Value = cmbADExistingCourses.List(k) Then 'If a combo-box value equals a value on the sheet
                check = False
                
                Exit For 'Exit For loop to prevent adding a repeat course name in the combo-box
                End If
            Next k
            
            If check Then
                cmbADExistingCourses.AddItem test.Cells(i, "B").Value 'Add the course to the combo-box
                
            End If
            
        Next i
        
 
'Types of Deliverables
cmbDeliverableType.Clear
cmbDeliverableType.AddItem "Assignment"
cmbDeliverableType.AddItem "Essay"
cmbDeliverableType.AddItem "Presentation"
cmbDeliverableType.AddItem "Term Project"
cmbDeliverableType.AddItem "Personal Project"
      
'Clear all text-boxes
txtDeliverableName.Value = ""
txtADAntGrade.Value = ""
txtADAntStudyHours.Value = ""
txtADCourseCode.Value = ""
txtADCourseName.Value = ""
txtADWeight.Value = ""
txtDeadline.Value = ""

'Academic Semesters
cmbSemester.AddItem "1A"
cmbSemester.AddItem "1B"
cmbSemester.AddItem "2A"
cmbSemester.AddItem "2B"
cmbSemester.AddItem "3A"
cmbSemester.AddItem "3B"
cmbSemester.AddItem "4A"
cmbSemester.AddItem "4B"


'Dtermines the value of the combo-box by the academic semester
If Date > 2021 - 1 - 1 Then
    cmbSemester.Value = "1A"
    
ElseIf Date >= 2021 - 1 - 11 And Date <= 2021 - 8 - 31 Then
    cmbSemester.Value = "1B"

ElseIf Date >= 2021 - 9 - 1 And Date <= 2021 - 4 - 30 Then
    cmbSemester.Value = "2A"

ElseIf Date >= 2022 - 5 - 1 And Date <= 2022 - 12 - 31 Then
    cmbSemester.Value = "2B"

ElseIf Date >= 2023 - 1 - 1 And Date <= 2023 - 8 - 31 Then
    cmbSemester.Value = "3A"

ElseIf Date >= 2023 - 9 - 1 And Date <= 2024 - 4 - 31 Then
    cmbSemester.Value = "3B"
    
ElseIf Date >= 2024 - 5 - 1 And Date <= 2024 - 12 - 31 Then
    cmbSemester.Value = "4A"

ElseIf Date > 2024 - 12 - 31 Then
    cmbSemester.Value = "4B"

End If

 
End Sub
Private Sub txtADCourseName_Change()

'Changes the text to all uppercases to maintain consistency between course names
Me.txtADCourseName.Text = UCase(txtADCourseName.Text)

End Sub

Private Sub txtDeadline_Exit(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next 'Allows code to continue runnning even if an error occurs

Me.txtDeadline = CDate(Me.txtDeadline) 'This converts the value of the entry in the textbox to the "Date" type

End Sub

Private Sub cmbADExistingCourses_Change()
'If the user selects an existing course, disable the ability to add a new course
    If Me.cmbADExistingCourses.Value <> "" Then
        Me.txtADCourseCode.Enabled = False
        Me.txtADCourseCode.BackColor = vbGrey
        Me.txtADCourseName.Enabled = False
        Me.txtADCourseName.BackColor = vbGrey
        
        'Clear any values that may have been inputted before selecting a combo-box value
        Me.txtADCourseName.Text = ""
        Me.txtADCourseCode.Text = ""
    
    Else
        Me.txtADCourseCode.Enabled = True
        Me.txtADCourseCode.BackColor = vbWhite
        Me.txtADCourseName.Enabled = True
        Me.txtADCourseName.BackColor = vbWhite
    
    End If

End Sub

Private Sub cmdClear_Click()

Me.cmbADExistingCourses.Value = "" 'Clears the value selected in the combo-box

End Sub
Private Sub cmdClear_Enter()

Me.cmbADExistingCourses.Value = "" 'Clears the value selected in the combo-box

End Sub

Private Sub cmdReset_Click()

Dim deliverables As Worksheet, test As Worksheet
Set deliverables = ThisWorkbook.Sheets("Deliverables")
Set test = ThisWorkbook.Sheets("Tests")

Dim iRow As Long, iRow2 As Long, check As Boolean
iRow = [Counta(Deliverables!A:A)]
iRow2 = [Counta(Tests!A:A)]

cmbADExistingCourses.Clear
        
        'Loop through each assessment to add courses to the combo-box
    For i = 2 To iRow
        
        check = True
            
        For k = 0 To cmbADExistingCourses.ListCount - 1 'Loop through the combo-box values
            If deliverables.Cells(i, "B").Value = cmbADExistingCourses.List(k) Then 'If a combo-box value equals a value on the sheet
                check = False
                
            Exit For 'Exit For loop to prevent adding a repeat course name in the combo-box
            End If
        Next k
            
            If check Then
                cmbADExistingCourses.AddItem deliverables.Cells(i, "B").Value 'Add the course to the combo-box
                
            End If
            
        Next i
        
        For i = 2 To iRow2
        check = True
            
            For k = 0 To cmbADExistingCourses.ListCount - 1 'Loop through the combo-box values
                If test.Cells(i, "B").Value = cmbADExistingCourses.List(k) Then 'If a combo-box value equals a value on the sheet
                check = False
                
                Exit For 'Exit For loop to prevent adding a repeat course name in the combo-box
                End If
            Next k
            
            If check Then
                cmbADExistingCourses.AddItem test.Cells(i, "B").Value 'Add the course to the combo-box
                
            End If
            
        Next i
               
               
'Types of Deliverables
cmbDeliverableType.Clear
cmbDeliverableType.AddItem "Assignment"
cmbDeliverableType.AddItem "Essay"
cmbDeliverableType.AddItem "Presentation"
cmbDeliverableType.AddItem "Term Project"
cmbDeliverableType.AddItem "Personal Project"
    
        
'Clear all text-boxes
txtDeliverableName.Value = ""
txtADAntGrade.Value = ""
txtADAntStudyHours.Value = ""
txtADCourseCode.Value = ""
txtADCourseName.Value = ""
txtADWeight.Value = ""
txtDeadline.Value = ""

End Sub

Private Sub cmdSubmit_Click()
   
'Check for required field is left empty or if inputted data is invalid
   Dim msgValue As VbMsgBoxResult
        If txtADCourseName.Text = "" And cmbADExistingCourses = "" Then
            msgValue = MsgBox("Please select a course or add a new course")
        ElseIf txtADCourseCode.Text = "" And cmbADExistingCourses = "" Then
            msgValue = MsgBox("Please enter a course code")
        ElseIf txtDeliverableName.Text = "" Then
            msgValue = MsgBox("Please enter Deliverable Name")
        ElseIf txtDeadline.Text = "" Then
            msgValue = MsgBox("Please enter a deadline")
        ElseIf Not IsDate(txtDeadline.Text) Then
            msgValue = MsgBox("Please enter a valid date for deadline")
        ElseIf Not IsNumeric(txtADCourseCode.Text) And (cmbADExistingCourses = "" Or cmbExistingCourses = "New Course") Then
            msgValue = MsgBox("Please enter a valid course code")
        ElseIf Not IsNumeric(txtADAntGrade.Text) Then
            msgValue = MsgBox("Please enter a valid expected grade")
        ElseIf Not IsNumeric(txtADWeight.Text) And txtADWeight.Text <> "" Then
            msgValue = MsgBox("Please enter a valid percentage value for weight")
        ElseIf Not IsNumeric(txtADAntStudyHours.Text) Then
            msgValue = MsgBox("Please enter a valid number of hours you plan to study for")

    Else
       
        Call Submit_DeliverableInfo
        Unload Me
                    
    End If
    
End Sub

Private Sub cmdSubmit_Enter()
'Check for required field is left empty or if inputted data is invalid
   Dim msgValue As VbMsgBoxResult
        If txtADCourseName.Text = "" And cmbADExistingCourses = "" Then
            msgValue = MsgBox("Please select a course or add a new course")
        ElseIf txtADCourseCode.Text = "" And cmbADExistingCourses = "" Then
            msgValue = MsgBox("Please enter a course code")
        ElseIf txtDeliverableName.Text = "" Then
            msgValue = MsgBox("Please enter Deliverable Name")
        ElseIf txtDeadline.Text = "" Then
            msgValue = MsgBox("Please enter a deadline")
        ElseIf Not IsDate(txtDeadline.Text) Then
            msgValue = MsgBox("Please enter a valid date for deadline")
        ElseIf Not IsNumeric(txtADCourseCode.Text) And (cmbADExistingCourses = "" Or cmbExistingCourses = "New Course") Then
            msgValue = MsgBox("Please enter a valid course code")
        ElseIf Not IsNumeric(txtADAntGrade.Text) Then
            msgValue = MsgBox("Please enter a valid expected grade")
        ElseIf Not IsNumeric(txtADWeight.Text) And txtADWeight.Text <> "" Then
            msgValue = MsgBox("Please enter a valid percentage value for weight")
        ElseIf Not IsNumeric(txtADAntStudyHours.Text) Then
            msgValue = MsgBox("Please enter a valid number of hours you plan to study for")

    Else
       
        Call Submit_DeliverableInfo
        Unload Me
                    
    End If

End Sub

Private Sub cmdBack_Click()

Application.ScreenUpdating = False
Unload Me
Unload AddAssessment_UserForm
AddAssessment_UserForm.Show

Application.ScreenUpdating = True

End Sub

Private Sub cmdBack_Enter()

Application.ScreenUpdating = False
Unload Me
Unload AddAssessment_UserForm
AddAssessment_UserForm.Show

Application.ScreenUpdating = True

End Sub
