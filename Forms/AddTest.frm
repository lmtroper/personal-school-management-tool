

Private Sub UserForm_Initialize()

Dim deliverables As Worksheet, test As Worksheet
Set deliverables = ThisWorkbook.Sheets("Deliverables")
Set test = ThisWorkbook.Sheets("Tests")

Dim iRow As Long, iRow2 As Long, check As Boolean
iRow = [Counta(Deliverables!A:A)]
iRow2 = [Counta(Tests!A:A)]

cmbATExistingCourses.Clear
        
        'Loop through each assessment to add courses to the combo-box
    For i = 2 To iRow
        
        check = True
            
        For k = 0 To cmbATExistingCourses.ListCount - 1 'Loop through the combo-box values
            If deliverables.Cells(i, "B").Value = cmbATExistingCourses.List(k) Then 'If a combo-box value equals a value on the sheet
                check = False
                
            Exit For 'Exit For loop to prevent adding a repeat course name in the combo-box
            End If
        Next k
            
            If check Then
                cmbATExistingCourses.AddItem deliverables.Cells(i, "B").Value 'Add the course to the combo-box
                
            End If
            
        Next i
        
        For i = 2 To iRow2
        check = True
            
            For k = 0 To cmbATExistingCourses.ListCount - 1 'Loop through the combo-box values
                If test.Cells(i, "B").Value = cmbATExistingCourses.List(k) Then 'If a combo-box value equals a value on the sheet
                check = False
                
                Exit For 'Exit For loop to prevent adding a repeat course name in the combo-box
                End If
            Next k
            
            If check Then
                cmbATExistingCourses.AddItem test.Cells(i, "B").Value 'Add the course to the combo-box
                
            End If
            
        Next i
        
        
'Types of Tests
optQuiz.Value = False
optMidterm.Value = False
optUnitTest.Value = False
optFinal.Value = False

            
'Clear all text-boxes
txtTestName.Value = ""
txtLengthHours.Value = ""
txtATAntGrade.Value = ""
txtATAntStudyHours.Value = ""
txtATCourseCode.Value = ""
txtATCourseName.Value = ""
txtATWeight.Value = ""
txtDate.Value = ""

'Academic Semester
cmbSemester.AddItem "1A"
cmbSemester.AddItem "1B"
cmbSemester.AddItem "2A"
cmbSemester.AddItem "2B"
cmbSemester.AddItem "3A"
cmbSemester.AddItem "3B"
cmbSemester.AddItem "4A"
cmbSemester.AddItem "4B"

'Determines value of combo-box based on academic semester
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


Private Sub cmbSemester_Change()

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

Private Sub optUnitTest_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

'Option button changes shape when you hover over it
optUnitTest.SpecialEffect = fmButtonEffectFlat
optMidterm.SpecialEffect = fmButtonEffectSunken
optFinal.SpecialEffect = fmButtonEffectSunken
optQuiz.SpecialEffect = fmButtonEffectSunken

End Sub

Private Sub optQuiz_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

optUnitTest.SpecialEffect = fmButtonEffectSunken
optMidterm.SpecialEffect = fmButtonEffectSunken
optFinal.SpecialEffect = fmButtonEffectSunken
optQuiz.SpecialEffect = fmButtonEffectFlat

End Sub
Private Sub optMidterm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

optUnitTest.SpecialEffect = fmButtonEffectSunken
optMidterm.SpecialEffect = fmButtonEffectFlat
optFinal.SpecialEffect = fmButtonEffectSunken
optQuiz.SpecialEffect = fmButtonEffectSunken

End Sub
Private Sub optFinal_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

optUnitTest.SpecialEffect = fmButtonEffectSunken
optMidterm.SpecialEffect = fmButtonEffectSunken
optFinal.SpecialEffect = fmButtonEffectFlat
optQuiz.SpecialEffect = fmButtonEffectSunken

End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

optUnitTest.SpecialEffect = fmButtonEffectSunken
optMidterm.SpecialEffect = fmButtonEffectSunken
optFinal.SpecialEffect = fmButtonEffectSunken
optQuiz.SpecialEffect = fmButtonEffectSunken

End Sub
Private Sub txtATCourseName_Change()

'Changes the text to all uppercases to maintain consistency between course names
Me.txtATCourseName.Text = UCase(txtATCourseName.Text)

End Sub

Private Sub txtDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next 'Allows code to continue runnning even if an error occurs

Me.txtDate = CDate(Me.txtDate) 'This converts the value of the entry in the textbox to the "Date" type

End Sub

Private Sub cmbATExistingCourses_Change()

'If the user selects an existing course, disable the ability to add a new course
    If Me.cmbATExistingCourses.Value <> "" Then
        Me.txtATCourseCode.Enabled = False
        Me.txtATCourseCode.BackColor = vbGrey
        Me.txtATCourseName.Enabled = False
        Me.txtATCourseName.BackColor = vbGrey
        
        'Clear any values that may have been inputted before selecting a combo-box value
        Me.txtATCourseName.Text = ""
        Me.txtATCourseCode.Text = ""
    
    Else
        Me.txtATCourseCode.Enabled = True
        Me.txtATCourseCode.BackColor = vbWhite
        Me.txtATCourseName.Enabled = True
        Me.txtATCourseName.BackColor = vbWhite
    
    End If

End Sub

Private Sub cmdSemester_Click()

MsgBox "Indicates the academic semester"

End Sub

Private Sub cmdClear_Click()

Me.cmbATExistingCourses.Value = "" 'Clears the value selected in the combo-box

End Sub

Private Sub cmdClear_Enter()

Me.cmbATExistingCourses.Value = "" 'Clears the value selected in the combo-box

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


Private Sub cmdReset_Click()

Dim deliverables As Worksheet, test As Worksheet
Set deliverables = ThisWorkbook.Sheets("Deliverables")
Set test = ThisWorkbook.Sheets("Tests")

Dim iRow As Long, iRow2 As Long, check As Boolean
iRow = [Counta(Deliverables!A:A)]
iRow2 = [Counta(Tests!A:A)]

cmbATExistingCourses.Clear
        
        'Loop through each assessment to add courses to the combo-box
    For i = 2 To iRow
        
        check = True
            
        For k = 0 To cmbATExistingCourses.ListCount - 1 'Loop through the combo-box values
            If deliverables.Cells(i, "B").Value = cmbATExistingCourses.List(k) Then 'If a combo-box value equals a value on the sheet
                check = False
                
            Exit For 'Exit For loop to prevent adding a repeat course name in the combo-box
            End If
        Next k
            
            If check Then
                cmbATExistingCourses.AddItem deliverables.Cells(i, "B").Value 'Add the course to the combo-box
                
            End If
            
        Next i
        
        For i = 2 To iRow2
        check = True
            
            For k = 0 To cmbATExistingCourses.ListCount - 1 'Loop through the combo-box values
                If test.Cells(i, "B").Value = cmbATExistingCourses.List(k) Then 'If a combo-box value equals a value on the sheet
                check = False
                
                Exit For 'Exit For loop to prevent adding a repeat course name in the combo-box
                End If
            Next k
            
            If check Then
                cmbATExistingCourses.AddItem test.Cells(i, "B").Value 'Add the course to the combo-box
                
            End If
            
        Next i
        
        
  
optQuiz.Value = False
optMidterm.Value = False
optUnitTest.Value = False
optFinal.Value = False

            
'Clear all text-boxes
txtTestName.Value = ""
txtATAntGrade.Value = ""
txtATAntStudyHours.Value = ""
txtATCourseCode.Value = ""
txtATCourseName.Value = ""
txtATWeight.Value = ""
txtDate.Value = ""
txtLengthHours.Value = ""

End Sub

Private Sub cmdReset_Enter()

Dim deliverables As Worksheet, test As Worksheet
Set deliverables = ThisWorkbook.Sheets("Deliverables")
Set test = ThisWorkbook.Sheets("Tests")

Dim iRow As Long, iRow2 As Long, check As Boolean
iRow = [Counta(Deliverables!A:A)]
iRow2 = [Counta(Tests!A:A)]

cmbATExistingCourses.Clear
        
        'Loop through each assessment to add courses to the combo-box
    For i = 2 To iRow
        
        check = True
            
        For k = 0 To cmbATExistingCourses.ListCount - 1 'Loop through the combo-box values
            If deliverables.Cells(i, "B").Value = cmbATExistingCourses.List(k) Then 'If a combo-box value equals a value on the sheet
                check = False
                
            Exit For 'Exit For loop to prevent adding a repeat course name in the combo-box
            End If
        Next k
            
            If check Then
                cmbATExistingCourses.AddItem deliverables.Cells(i, "B").Value 'Add the course to the combo-box
                
            End If
            
        Next i
        
        For i = 2 To iRow2
        check = True
            
            For k = 0 To cmbATExistingCourses.ListCount - 1 'Loop through the combo-box values
                If test.Cells(i, "B").Value = cmbATExistingCourses.List(k) Then 'If a combo-box value equals a value on the sheet
                check = False
                
                Exit For 'Exit For loop to prevent adding a repeat course name in the combo-box
                End If
            Next k
            
            If check Then
                cmbATExistingCourses.AddItem test.Cells(i, "B").Value 'Add the course to the combo-box
                
            End If
            
        Next i
        
        
  
optQuiz.Value = False
optMidterm.Value = False
optUnitTest.Value = False
optFinal.Value = False

            
'Clear all text-boxes
txtTestName.Value = ""
txtATAntGrade.Value = ""
txtATAntStudyHours.Value = ""
txtATCourseCode.Value = ""
txtATCourseName.Value = ""
txtATWeight.Value = ""
txtDate.Value = ""
txtLengthHours.Value = ""

End Sub

Private Sub cmdSubmit_Click()
   
'Check for required field is left empty or if inputted data is invalid
   Dim msgValue As VbMsgBoxResult
        If txtATCourseName.Text = "" And cmbATExistingCourses = "" Then
            msgValue = MsgBox("Please select a course or add a new course")
        ElseIf txtATCourseCode.Text = "" And cmbATExistingCourses = "" Then
            msgValue = MsgBox("Please enter a course code")
        ElseIf txtTestName.Text = "" Then
            msgValue = MsgBox("Please enter the Assessment Name")
        ElseIf txtDate.Text = "" Then
            msgValue = MsgBox("Please enter a deadline")
        ElseIf Not IsDate(txtDate.Text) Then
            msgValue = MsgBox("Please enter a valid date for deadline")
        ElseIf Not IsNumeric(txtATCourseCode.Text) And (cmbATExistingCourses = "" Or cmbATExistingCourses = "New Course") Then
            msgValue = MsgBox("Please enter a valid course code")
        ElseIf Not IsNumeric(txtATAntGrade.Text) Then
            msgValue = MsgBox("Please enter a valid expected grade")
        ElseIf Not IsNumeric(txtATWeight.Text) And txtATWeight.Text <> "" Then
            msgValue = MsgBox("Please enter a valid percentage value for weight")
        ElseIf Not IsNumeric(txtATAntStudyHours.Text) Then
            msgValue = MsgBox("Please enter a valid number of hours you plan to study for")
        ElseIf txtLengthHours.Text <> "" And Not IsNumeric(txtLengthHours) Then
            msgValue = MsgBox("Please enter a valid number of hours for the assesment's length")
        ElseIf optQuiz.Value = False And optMidterm.Value = False And optUnitTest.Value = False And optFinal.Value = False Then
            msgValue = MsgBox("Please select a test category")
        
    Else
       
        Call Submit_TestInfo
        Unload Me
        Call SortTablesbyDeadline
        
    End If
    
End Sub

Private Sub cmdSubmit_Enter()
   
'Check for required field is left empty or if inputted data is invalid
   Dim msgValue As VbMsgBoxResult
        If txtATCourseName.Text = "" And cmbATExistingCourses = "" Then
            msgValue = MsgBox("Please select a course or add a new course")
        ElseIf txtATCourseCode.Text = "" And cmbATExistingCourses = "" Then
            msgValue = MsgBox("Please enter a course code")
        ElseIf txtTestName.Text = "" Then
            msgValue = MsgBox("Please enter the Assessment Name")
        ElseIf txtDate.Text = "" Then
            msgValue = MsgBox("Please enter a deadline")
        ElseIf Not IsDate(txtDate.Text) Then
            msgValue = MsgBox("Please enter a valid date for deadline")
        ElseIf Not IsNumeric(txtATCourseCode.Text) And (cmbATExistingCourses = "" Or cmbATExistingCourses = "New Course") Then
            msgValue = MsgBox("Please enter a valid course code")
        ElseIf Not IsNumeric(txtATAntGrade.Text) Then
            msgValue = MsgBox("Please enter a valid expected grade")
        ElseIf Not IsNumeric(txtATWeight.Text) And txtATWeight.Text <> "" Then
            msgValue = MsgBox("Please enter a valid percentage value for weight")
        ElseIf Not IsNumeric(txtATAntStudyHours.Text) Then
            msgValue = MsgBox("Please enter a valid number of hours you plan to study for")
        ElseIf txtLengthHours.Text <> "" And Not IsNumeric(txtLengthHours) Then
            msgValue = MsgBox("Please enter a valid number of hours for the assesment's length")
        ElseIf optQuiz.Value = False And optMidterm.Value = False And optUnitTest.Value = False And optFinal.Value = False Then
            msgValue = MsgBox("Please select a test category")
        
    Else
       
        Call Submit_TestInfo
        Unload Me
        Call SortTablesbyDeadline
        
    End If
    
End Sub
