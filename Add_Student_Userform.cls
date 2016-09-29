'========================= CONVOCATION DATABASE ADD STUDENT USERFORM ========================='
'Created By: Brady Hammond
'Created: 05/18/16
'============================================================================================='

'User Form Initialize Subroutine:
Private Sub UserForm_Initialize()

    'Empty all fields
    With Add_Student
    
        .Student_ID_Input.Value = ""
        .Student_Last_Name_Input.Value = ""
        .Student_First_Name_Input.Value = ""
        .Student_Major_Input.Value = ""
        .Student_Email_Input.Value = ""
        .Summa_Cum_Laude_Check_Box.Value = False
        .Magna_Cum_Laude_Check_Box.Value = False
        .Cum_Laude_Check_Box.Value = False
        .Phi_Kappa_Phi_Check_Box = False
        .None_Check_Box = False
        .Relations_Input = ""
        .Notes_Input.Value = ""
        
    End With
    
    'Fill in combo box options
    Add_Student.Student_Major_Input.List = Worksheets("Master Database").Range("I4:I130").Value
    
End Sub

'Autofill Button Enter Subroutine:
Private Sub Autofill_Button_Enter()

    'Set the focus on the button
    Add_Student.Autofill_Button.SetFocus
    
    'Click the button
    Call Autofill_Button_Click
    
End Sub

'Autofill Button Click Subroutine:
Private Sub Autofill_Button_Click()

    'Declare needed variables
    Dim HTML_Document As HTMLDocument
    Dim HTML_Element As IHTMLElement
    Dim BYU_URL As String
    Dim URL_Check As String
    Dim Extra_Windows_Check_Shell_Window As SHDocVw.ShellWindows
    Dim Extra_Windows_Check_Object As Object
    
    'Create a shell window
    Set Extra_Windows_Check_Shell_Window = New SHDocVw.ShellWindows
    
    'Check if there are open windows
    If Extra_Windows_Check_Shell_Window.Count Then
    
        'Loop through each open window
        For Each Extra_Windows_Check_Object In Extra_Windows_Check_Shell_Window
        
            If InStr(1, Extra_Windows_Check_Object.LocationName, "Brigham Young University Sign-in Service") Then
        
                'Exit the pop up browser
                Extra_Windows_Check_Object.Quit
                        
                'Delete the pop up
                Set Extra_Windows_Check_Object = Nothing
            
            End If
            
        Next Extra_Windows_Check_Object
            
        On Error Resume Next
        Err.Clear
        
    End If
    
    'Define BYU_URL and URL_Check
    BYU_URL = "https://y.byu.edu/ry/ae/prod/person/cgi/personSummary.cgi"
    URL_Check = "https://cas.byu.edu/cas/login?service=https%3A%2F%2Fy.byu.edu%2Fvalidate%3Ftarget%3Dhttps%253A%252F%252Fy.byu.edu%252Fry%252Fae%252Fprod%252Fperson%252Fcgi%252FpersonSummary.cgi"

    On Error GoTo Error_Handling
    
    'Open up internet explorer
    Set IE_Browser = CreateObject("InternetExplorer.Application")
    
    'Disable pop-ups
    IE_Browser.Silent = True
    
    'Navigate to the BYU login page
    IE_Browser.Navigate BYU_URL
    
    'Hide the browser
    IE_Browser.Visible = False
    
    'Wait for the page to load state 4
    Do
    Loop Until IE_Browser.readyState = 4
    
    'Check if the user is already logged into a myBYU account
    If IE_Browser.LocationURL <> BYU_URL Or IE_Browser.LocationURL = URL_Check Then
    
        'If so, exit the browser
        IE_Browser.Quit
        
        'Delete the browser object
        Set IE_Browser = Nothing
        
        'Output that the user is already logged off
        MsgBox "You are not logged into your myBYU account. You must be logged into your myBYU account to use this feature.", vbInformation, "Invalid User"
        
        'Exit the subroutine
        Exit Sub
        
    'Check if there is a student number to look up
    ElseIf Add_Student.Student_ID_Input.Value = "" Then
    
        'If not, output an error message
        MsgBox "There is no student ID to look up. Please enter one and try again.", vbInformation, "No Student ID"
        
        'Exit the subroutine
        Exit Sub
        
    'Check if the student ID is a number
    ElseIf IsNumeric(Add_Student.Student_ID_Input.Value) Then
    
        'If it is, check if the number is the right length
        If Len(Add_Student.Student_ID_Input.Value) <> 9 Then
        
            'If not, output an error message
            MsgBox "The given student number is not the correct length. Please make sure that the student number is 9 numbers long.", vbInformation, "Invalid Student Number"
            
            'Exit the subroutine
            Exit Sub
        
        Else
            
            'Search for the Student Information
            Call Search_ID_Number(Add_Student.Student_ID_Input.Value, IE_Browser)
        
        End If
        
    Else
    
        'Search for the Student Information
        Call Search_ID_Number(Add_Student.Student_ID_Input.Value, IE_Browser)
        
    End If
    
    Err.Clear
    
Error_Handling:

    'Check if there was an error
    If Err.Number <> 0 Then
        
        'If so, output an error message
        MsgBox "The database has lost its connection to the internet. Please wait a moment and try again.", vbCritical, "Failed Internet Connection"
        
    End If
    
End Sub

'Submit Button Click Subroutine:
Private Sub Submit_Button_Click()

    'Upload the data
    If Form_Complete Then
    
        'Clear the form
        Call UserForm_Initialize
        
    End If
    
End Sub

'Finish Button Click Subroutine:
Private Sub Finish_Button_Click()
    
    'Upload the data
    
    If Form_Complete = True Then
    
        'Close the form
        Unload Add_Student
    
    End If
    
End Sub

'Cancel Button Click Subroutine:
Private Sub Cancel_Button_Click()

    'Close the form
    Unload Add_Student
    
End Sub


'Search ID Number Subroutine:
Private Sub Search_ID_Number(Student_ID, IE_Browser)

    
    'Declare needed variables
    Dim Pop_Up_Check_Shell_Window As SHDocVw.ShellWindows
    Dim Pop_Up_Check_Object As Object
    Dim HTML_Document As HTMLDocument
    Dim HTML_Element As IHTMLElement
    Dim BYU_URL As String
    Dim Second_Go
    Dim Pop_Up_Open As Boolean
    Dim Found As Boolean
    Dim Student_Full_Name
    Dim Cell
    Dim Majors_Range As Range
    Dim Major_Found As Boolean
    
    'Define Major_Found
    Major_Found = False
    
    'Define Majors_Range
    Set Majors_Range = Range("I2:I129")
    
    'Define Pop_Up_Open
    Pop_Up_Open = False
    
    'Define BYU URL
    BYU_URL = "https://y.byu.edu/ry/ae/prod/person/cgi/personSummary.cgi"
    
    'Create a shell window
    Set Pop_Up_Check_Shell_Window = New SHDocVw.ShellWindows
    
    'Disable pop-ups
    IE_Browser.Silent = True
    
    'Navigate to the BYU login page
    IE_Browser.Navigate BYU_URL
    
    'Hide the browser
    IE_Browser.Visible = False
    
    'Wait for the page to load state 4
    Do
    Loop Until IE_Browser.readyState = 4
    
    'Check if the browser was able to navigate to the correct page
    If IE_Browser.LocationURL <> BYU_URL Then
        
        'If not, output an error message, quit and destroy the browser, and exit the subroutine
        MsgBox "You are not logged into your myBYU account. Please login and try again.", vbInformation, "Invalid User"
        IE_Browser.Quit
        Set IE_Browser = Nothing
        Exit Sub
        
    End If
    
    'Save the HTML data
    Set HTML_Document = IE_Browser.Document
    
    'Input the given user number into the search box
    HTML_Document.all.NAVpattern.Value = Student_ID
    
    'Search through the HTML for anchor tags
    For Each HTML_Element In HTML_Document.getElementsByTagName("a")
        
        'Skip past the first go button on the page
        If HTML_Element.innerHTML = "GO" And Second_Go = False Then
            
            Second_Go = True
            
        'Click the second go button on the page and exit the for loop
        ElseIf HTML_Element.innerHTML = "GO" And Second_Go = True Then
        
            HTML_Element.Click
            Exit For
            
        End If
        
    Next

    'Wait for the page to load
    Do
    Loop Until IE_Browser.readyState = 3
    
    Do
    Loop Until IE_Browser.readyState = 4
    
    'Wait for any pop ups to show
    Application.Wait (Now + TimeValue("0:00:01"))
    
    'Check if there are open windows
    If Pop_Up_Check_Shell_Window.Count Then
        
        'Loop through each open window
        For Each Pop_Up_Check_Object In Pop_Up_Check_Shell_Window
        
            Do
                'Check if the lookup pop-up is open
                If InStr(1, Pop_Up_Check_Object.LocationName, "Directory lookup/display") Then

                    'If there is a pop up, set pop up open to true
                    Pop_Up_Open = True
                    
                    'Save the pop up browser
                    Set HTML_Document = Pop_Up_Check_Object.Document
                
                    'Loop through the HTML
                    For Each HTML_Element In HTML_Document.getElementsByTagName("var")
                    
                        'Check if the pop up is for an incorrect ID lookup
                        If HTML_Element.innerHTML = "J?n" Then
                        
                            'If it is set found to true
                            Found = True
                            
                        End If
                        
                    Next HTML_Element
                    
                    'There will be an error if a correct student is found from the pop up window. This line circumvents that
                    On Error GoTo Student_Found
                    
                    'Check if the incorrect ID lookup pop up was found
                    If Found = True Then
                    
                        'If so show that the pop up is open
                        Pop_Up_Open = True
                        
                        'Wait for the user to use the pop up
                        Application.Wait (Now + TimeValue("0:00:01"))
                            
                    Else
                        
                        'Show that the pop up is no long open
                        Pop_Up_Open = False
                        
                        'Exit the pop up browser
                        Pop_Up_Check_Object.Quit
                        
                        'Delete the pop up
                        Set Pop_Up_Check_Object = Nothing
                        
                        'Exit the for loop
                        Exit For
                        
                    End If
                        
                End If
            
            Loop Until Pop_Up_Open = False
            
        Next Pop_Up_Check_Object
            
    End If
    
Student_Found:
    
    'Check if there is already a major given
    If Add_Student.Student_Major_Input.Value <> "" Then
        
        'If so, delete it
        Add_Student.Student_Major_Input.Value = ""
    
    End If
    
    'Wait for the page to properly load
    Application.Wait (Now + 5.787e-08)

    'Save the browser
    Set HTML_Document = IE_Browser.Document
    
    'Save the student email, student name, netID, and student number, then exit the for loop
    For Each HTML_Element In HTML_Document.getElementsByTagName("input")
        
        'Check for the restore_name field
        If HTML_Element.Name = "restore_email" Then
        
            'Save that value as student name
            On Error GoTo Error_Handling
            Add_Student.Student_Email_Input.Value = HTML_Element.Value
            
        End If
        
        'Check for the restore_name field
        If HTML_Element.Name = "restore_name" Then
        
            'Save that value as student name
            On Error GoTo Error_Handling
            Student_Full_Name = HTML_Element.Value
            Add_Student.Student_First_Name_Input.Value = Mid(Student_Full_Name, InStr(Student_Full_Name, ",") + 2)
            Add_Student.Student_Last_Name_Input.Value = Left(Student_Full_Name, InStr(Student_Full_Name, ",") - 1)
            
        End If
        
        'Check for the restore_byu_id field
        If HTML_Element.Name = "restore_byu_id" Then
            
            'Save that value as the student ID and exit the for loop
            On Error GoTo Error_Handling
            Student_ID = HTML_Element.Value
            Add_Student.Student_ID_Input.Value = Replace(Student_ID, "-", "")
            Exit For
            
        End If

    Next
    
    'Search for the input tag
    For Each HTML_Element In HTML_Document.getElementsByTagName("input")
    
        'Search for the input box
        If HTML_Element.Name = "pattern" Then
            
            'Input ADV07 in the input box then exit the for loop
            HTML_Element.Value = "ADV07"
            Exit For
            
        End If
    
    Next
    
    'Search for the anchor tag
    For Each HTML_Element In HTML_Document.getElementsByTagName("a")
    
        'Search for the search for quick URL button
        If HTML_Element.Title = "Search for Quick URL" Then
            
            'If found, click it and exit the for loop
            HTML_Element.Click
            Exit For
        
        End If
    
    Next
    
    'Wait for the page to load
    Do
    Loop Until IE_Browser.readyState = 3
    
    Do
    Loop Until IE_Browser.readyState = 4
    
    'Save the browser
    Set HTML_Document = IE_Browser.Document
    
    'Search for the student programs field
    For Each HTML_Element In HTML_Document.getElementsByTagName("b")
     
        For Each Cell In Majors_Range
        
            If Cell.Value = HTML_Element.innerHTML Then
            
                'Save the student major
                On Error GoTo Error_Handling
                Add_Student.Student_Major_Input.Value = HTML_Element.innerHTML
                Major_Found = True
                Exit For
            
            End If
        
        Next
        
        If Major_Found = True Then
            
            Exit For
            
        End If
        
    Next
        
    'Quit and destroy the browser
    IE_Browser.Quit
    Set IE_Browser = Nothing
    
    Err.Clear
    
Error_Handling:

    'Check if there was an error
    If Err.Number <> 0 Then
        
        'If so, output an error message
        MsgBox "The database has lost its connection to the internet. Please wait a moment and try again.", vbCritical, "Failed Internet Connection"
        
    End If
    
End Sub

'Form Complete Function:
Private Function Form_Complete() As Boolean

    'Declare needed variables
    Dim Student_ID
    Dim Student_Last_Name
    Dim Student_First_Name
    Dim Student_Major
    Dim Student_Email
    Dim Honors
    Dim Relations
    Dim Notes
    
    'Define needed variables
    Student_ID = Add_Student.Student_ID_Input.Value
    Student_Last_Name = Add_Student.Student_Last_Name_Input.Value
    Student_First_Name = Add_Student.Student_First_Name_Input.Value
    Student_Major = Add_Student.Student_Major_Input.Value
    Student_Email = Add_Student.Student_Email_Input.Value
    Relations = Add_Student.Relations_Input.Value
    Notes = Add_Student.Notes_Input.Value
    
    'Check for the correct honors
    If Add_Student.Summa_Cum_Laude_Check_Box.Value = True And Add_Student.Phi_Kappa_Phi_Check_Box.Value = False And Add_Student.Magna_Cum_Laude_Check_Box.Value = False And Add_Student.Cum_Laude_Check_Box.Value = False And Add_Student.None_Check_Box.Value = False Then
    
        Honors = "Summa Cum Laude"
        
    ElseIf Add_Student.Magna_Cum_Laude_Check_Box.Value = True And Add_Student.Phi_Kappa_Phi_Check_Box.Value = False And Add_Student.Summa_Cum_Laude_Check_Box.Value = False And Add_Student.Cum_Laude_Check_Box.Value = False And Add_Student.None_Check_Box.Value = False Then
    
        Honors = "Magna Cum Laude"
        
    ElseIf Add_Student.Cum_Laude_Check_Box.Value = True And Add_Student.Phi_Kappa_Phi_Check_Box.Value = False And Add_Student.Summa_Cum_Laude_Check_Box.Value = False And Add_Student.Magna_Cum_Laude_Check_Box.Value = False And Add_Student.None_Check_Box.Value = False Then
    
        Honors = "Cum Laude"
    
    ElseIf Add_Student.Phi_Kappa_Phi_Check_Box.Value = True And Add_Student.Summa_Cum_Laude_Check_Box.Value = False And Add_Student.Magna_Cum_Laude_Check_Box.Value = False And Add_Student.Cum_Laude_Check_Box.Value = False And Add_Student.None_Check_Box.Value = False Then
        
        Honors = "Phi Kappa Phi"
        
    ElseIf Add_Student.None_Check_Box.Value = True And Add_Student.Summa_Cum_Laude_Check_Box.Value = False And Add_Student.Magna_Cum_Laude_Check_Box.Value = False And Add_Student.Cum_Laude_Check_Box.Value = False And Add_Student.Phi_Kappa_Phi_Check_Box.Value = False Then
    
        Honors = ""
        
    ElseIf Add_Student.Summa_Cum_Laude_Check_Box.Value = False And Add_Student.Magna_Cum_Laude_Check_Box.Value = False And Add_Student.Cum_Laude_Check_Box.Value = False And Add_Student.Phi_Kappa_Phi_Check_Box.Value = False And Add_Student.None_Check_Box.Value = False Then
    
        Honors = ""
        
    ElseIf Add_Student.Summa_Cum_Laude_Check_Box.Value = True And Add_Student.Phi_Kappa_Phi_Check_Box.Value = True And Add_Student.Magna_Cum_Laude_Check_Box.Value = False And Add_Student.Cum_Laude_Check_Box.Value = False And Add_Student.None_Check_Box.Value = False Then
    
        Honors = "Summa Cum Laude, Phi Kappa Phi"
        
    ElseIf Add_Student.Magna_Cum_Laude_Check_Box.Value = True And Add_Student.Phi_Kappa_Phi_Check_Box.Value = True And Add_Student.Summa_Cum_Laude_Check_Box.Value = False And Add_Student.Cum_Laude_Check_Box.Value = False And Add_Student.None_Check_Box.Value = False Then
    
        Honors = "Magna Cum Laude, Phi Kappa Phi"
        
    ElseIf Add_Student.Cum_Laude_Check_Box.Value = True And Add_Student.Phi_Kappa_Phi_Check_Box.Value = True And Add_Student.Summa_Cum_Laude_Check_Box.Value = False And Add_Student.Magna_Cum_Laude_Check_Box.Value = False And Add_Student.None_Check_Box.Value = False Then
    
        Honors = "Cum Laude, Phi Kappa Phi"
        
    Else
    
        Honors = ""
        
    End If
        
    
    'Check if there is a value for student ID
    If Student_ID = "" Then
    
        'If there is none, output an error message and exit the function
        MsgBox "Please eneter a student ID for this student.", vbInformation, "Missing Student ID"
        Form_Complete = False
        Exit Function
    
    'Check if there is a value for student last name
    ElseIf Student_Last_Name = "" Then
        
        'If there is none, output an error message and exit the function
        MsgBox "Please eneter a last name for this student.", vbInformation, "Missing Student Name"
        Form_Complete = False
        Exit Function
    
    'Check if there is a value for student first name
    ElseIf Student_First_Name = "" Then
    
        'If there is none, output an error message and exit the function
        MsgBox "Please eneter a first name for this student.", vbInformation, "Missing Student Name"
        Form_Complete = False
        Exit Function
    
    'Check if there is a value for student major
    ElseIf Student_Major = "" Then
    
        'If there is none, output an error message and exit the function
        MsgBox "Please eneter a major for this student.", vbInformation, "Missing Student Major"
        Form_Complete = False
        Exit Function
    
    'Check if there is a value for student email
    ElseIf Student_Email = "" Then
        
        'If there is none, output an error message and exit the function
        MsgBox "Please eneter an email for this student.", vbInformation, "Missing Student Email"
        Form_Complete = False
        Exit Function
    
    Else
    
        'Upload the data
        Form_Complete = True
        Call Upload_Data(Student_ID, Student_Last_Name, Student_First_Name, Student_Major, Student_Email, Honors, Relations, Notes)
        
    End If
    
End Function

'Upload Data Subroutine:
Private Sub Upload_Data(Student_ID, Student_Last_Name, Student_First_Name, Student_Major, Student_Email, Honors, Relations, Notes)

    'Declare needed variables
    Dim Next_Row
    
    'Find the next blank row
    Next_Row = Range("A" & Rows.Count).End(xlUp).Row + 1
    
    'Input the data
    With ActiveWorkbook.ActiveSheet
    
        .Range("A" & Next_Row).Value = Student_Last_Name
        .Range("B" & Next_Row).Value = Student_First_Name
        .Range("C" & Next_Row).Value = Student_ID
        .Range("D" & Next_Row).Value = Student_Major
        .Range("E" & Next_Row).Value = Honors
        .Range("F" & Next_Row).Value = Relations
        .Range("G" & Next_Row).Value = Student_Email
        .Range("H" & Next_Row).Value = Notes
        
    End With
    
End Sub

