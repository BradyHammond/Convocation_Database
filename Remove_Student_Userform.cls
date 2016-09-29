========================= CONVOCATION DATABASE REMOVE STUDENT USERFORM =========================='
'Created By: Brady Hammond
'Created: 05/18/16
'================================================================================================'

'Userform Initialize Subroutine:
Private Sub UserForm_Initialize()

    'Clear all fields
    Student_ID_Input.Value = ""
    Student_Last_Name_Input.Value = ""
    Student_First_Name_Input.Value = ""
    Student_Email_Input.Value = ""
    
End Sub


'Submit Button Click Subroutine:
Private Sub Submit_Button_Click()
    
    'Declare needed variables
    Dim Input_Check
    
    'Check if there is any input
    Input_Check = Check_Input
    If Not Input_Check = "" Then
    
        'If there is some, remove the student
        Call Remove_Student_From_Database(Input_Check)
        
    End If
    
    'Clear the form
    Call UserForm_Initialize

End Sub

'Finish Button Click Subroutine:
Private Sub Finish_Button_Click()

    'Declare needed variables
    Dim Input_Check
    
    'Check if there is any input
    Input_Check = Check_Input
    If Not Input_Check = "" Then
    
        'If there is some, remove the student
        Call Remove_Student_From_Database(Input_Check)
    
        'Close the form
        Unload Remove_Student
        
    End If
    
    
End Sub

'Cancel Button Click Subroutine:
Private Sub Cancel_Button_Click()

    'Close the form
    Unload Remove_Student
    
End Sub

'Check Input Function:
Private Function Check_Input() As String

    'Check if there is any input
    If Remove_Student.Student_ID_Input.Value = "" And Remove_Student.Student_Last_Name_Input.Value = "" And Remove_Student.Student_First_Name_Input.Value = "" And Remove_Student.Student_Email_Input.Value = "" Then
    
        'If not, output an error message and exit the function
        MsgBox "Please give a student to be removed", vbInformation, "No Student Given"
        Check_Input = ""
        Exit Function
        
    'Check if there is a missing first name
    ElseIf Remove_Student.Student_ID_Input.Value = "" And Remove_Student.Student_Last_Name_Input.Value = "" And Remove_Student.Student_Email_Input.Value = "" Then
        
        'If so, output an error message and exit the function
        MsgBox "Please include the student's first and last name", vbInformation, "Invalid Student Name"
        Check_Input = ""
        Exit Function
        
    'Check if there is a missing last name
    ElseIf Remove_Student.Student_ID_Input.Value = "" And Remove_Student.Student_First_Name_Input.Value = "" And Remove_Student.Student_Email_Input.Value = "" Then
    
        'If so, output an error message and exit the function
        MsgBox "Please include the student's first and last name", vbInformation, "Invalid Student Name"
        Check_Input = ""
        Exit Function
        
    Else
    
        'Check if there is input for the student ID
        If Not Remove_Student.Student_ID_Input.Value = "" Then
        
            'If so, return that value and exit the function
            Check_Input = "ID"
            Exit Function
            
        End If
        
        'Check if there is input for the student name
        If Not Remove_Student.Student_Last_Name_Input.Value = "" And Not Remove_Student.Student_First_Name_Input.Value = "" Then
        
            'If so, return that value and exit the function
            Check_Input = "Name"
            Exit Function
            
        End If
        
        'Check if there is input for the student email
        If Not Remove_Student.Student_Email_Input.Value = "" Then
        
            'If so, return that value and exit the function
            Check_Input = "Email"
            Exit Function
            
        End If
        
        'For anything else, return NULL
        Check_Input = ""
    
    End If
        
End Function

'Remove Student Function:
Private Function Remove_Student_From_Database(Input_Check)

    'Declare needed variables
    Dim i As Integer
    Dim Last_Row As Integer
    Dim Response
    Dim First_Name_Found As Boolean
    Dim Last_Name_Found As Boolean
    
    'Find the last row
    Last_Row = Range("A" & Rows.Count).End(xlUp).Row
    
    'Check if there was an ID given
    If Input_Check = "ID" Then
    
        'Search through each row
        For i = 1 To Last_Row
        
            'Check if the student ID number was found
            If Cells(i, 3).Value = Remove_Student.Student_ID_Input.Value Then
            
                'If so, ask if the user really wants to delete the student, and exit the for
                Response = MsgBox("Are you sure you want to remove " & Cells(i, 2).Value & " " & Cells(i, 1).Value & " " & Cells(i, 3).Value, vbInformation + vbYesNo, "Remove Student Confirmation")
                Exit For
                
            End If
            
        Next
        
        'Check if the user wanted to delete the student
        If Response = vbYes Then
        
            'If so, delete the student and delete the blank row
            Range(Cells(i, 1), Cells(i, 8)).Select
            Selection.Delete Shift:=xlUp
            MsgBox "The student has been removed.", vbInformation, "Student Removed"
            
        'Check, if the user didn't want to delete the student
        ElseIf Response = vbNo Then
            
            'If so, exit the function
            Exit Function
        
        Else
        
            'Otherwise, output an error message
            MsgBox "The student you wish to remove was not found.", vbInformation, "Not Found"
            
        End If
        
    'Check if there was a name given
    ElseIf Input_Check = "Name" Then
    
        'Search through each row
        For i = 1 To Last_Row
        
            'Check if the student last name was found
            If UCase(Cells(i, 1).Value) = UCase(Remove_Student.Student_Last_Name_Input.Value) Then
            
                'If so, mark that it was found
                Last_Name_Found = True
                
                'Exit the for loop
                Exit For
                
            End If
            
        Next
        
        'Search through each row
        For i = 1 To Last_Row
        
            'Check if the student first name was found
            If UCase(Cells(i, 2).Value) Like UCase(Remove_Student.Student_First_Name_Input.Value) & "*" Then
            
                'If so, mark that it was found
                First_Name_Found = True
                
                'Exit the for loop
                Exit For
                
            End If
            
        Next
        
        'Check if both names were found
        If Last_Name_Found = True And First_Name_Found = True Then
        
            'If so, check if the user really wants to remove the student
            Response = MsgBox("Are you sure you want to remove " & Cells(i, 2).Value & " " & Cells(i, 1).Value & " " & Cells(i, 3).Value, vbInformation + vbYesNo, "Remove Student Confirmation")
        
            'Check if the user said yes
            If Response = vbYes Then
        
                'If so, delete the student
                Range(Cells(i, 1), Cells(i, 8)).Select
                Selection.Delete Shift:=xlUp
                MsgBox "The student has been removed.", vbInormation, "Student Removed"
            
            'Check if the user said no
            ElseIf Response = vbNo Then
            
                'If so, exit the function
                Exit Function
        
            Else
    
                'Otherwise, output an error message
                MsgBox "The student you wish to remove was not found.", vbInformation, "Not Found"
        
            End If
            
        
        Else
            
            'Otherwise, output an error message
            MsgBox "The student you wish to remove was not found.", vbInformation, "Not Found"
        
        End If
    
    'Check if there was an email given
    ElseIf Input_Check = "Email" Then
    
        'Search through each row
        For i = 1 To Last_Row
        
            'Check if the student email was found
            If UCase(Cells(i, 7).Value) = UCase(Remove_Student.Student_Email_Input.Value) Then
            
                'If so, ask if the user really wants to delete the student, and exit the for
                Response = MsgBox("Are you sure you want to remove " & Cells(i, 2).Value & " " & Cells(i, 1).Value & " " & Cells(i, 3).Value, vbInformation + vbYesNo, "Remove Student Confirmation")
                Exit For
                
            End If
            
        Next
        
        'Check if the user wanted to delete the student
        If Response = vbYes Then
        
            'If so, delete the student and delete the blank row
            Range(Cells(i, 1), Cells(i, 8)).Select
            Selection.Delete Shift:=xlUp
            MsgBox "The student has been removed.", vbInformation, "Student Removed"
        
        'Check, if the user didn't want to delete the student
        ElseIf Response = vbNo Then
            
            'If so, exit the function
            Exit Function
            
        Else
            
            'Otherwise, output an error message
            MsgBox "The student you wish to remove was not found.", vbInformation, "Not Found"
        
        End If
    
    Else
    
        'Otherwise, output an error message
        MsgBox "The student you wish to remove was not found.", vbInformation, "Not Found"
        
    End If
    
End Function

