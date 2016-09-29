'========================= CONVOCATION DATABASE GET DATA BUTTON ========================='
'Created By: Brady Hammond
'Created: 06/08/16
'========================================================================================'

'Get Data Subroutine:
Sub Get_Data_Button()

    'Declare needed variables
    Dim IE_Browser As Object
    Dim HTML_Document As HTMLDocument
    Dim HTML_Element As IHTMLElement
    Dim BYU_URL_1 As String
    Dim BYU_URL_2 As String
    Dim URL_Check As String
    Dim Extra_Windows_Check_Shell_Window As SHDocVw.ShellWindows
    Dim Extra_Windows_Check_Object As Object
    Dim Drop_Down_Box
    Dim Drop_Down_Option
    Dim Convocation_Date_Part_1
    Dim Convocation_Date_Part_2
    Dim Convocation_Compare_Date
    Dim HTML_Objects
    Dim Row_Data As Object
    Dim Next_Row As Long
    Dim Data_Range As Range
    Dim Cell As Variant
    Dim Second_Go As Boolean
    Dim Current_Cell As String
    Dim Next_Cell As String
    
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
    BYU_URL_1 = "https://y.byu.edu/ry/ae/prod/acad_plan/cgi/graduationStatus.cgi"
    URL_Check = "https://cas.byu.edu/cas/login?service=https%3A%2F%2Fy.byu.edu%2Fvalidate%3Ftarget%3Dhttps%253A%252F%252Fy.byu.edu%252Fry%252Fae%252Fprod%252Facad_plan%252Fcgi%252FgraduationStatus.cgi"
    BYU_URL_2 = "https://y.byu.edu/ry/ae/prod/person/cgi/personSummary.cgi"

    'Open up internet explorer
    Set IE_Browser = CreateObject("InternetExplorer.Application")
    
    'Disable pop-ups
    IE_Browser.Silent = True
    
    'Navigate to the BYU login page
    IE_Browser.Navigate BYU_URL_1
    
    'Hide the browser
    IE_Browser.Visible = False
    
    'Wait for the page to load state 4
    Do
    Loop Until IE_Browser.readyState = 4
    
    'Check if the user is already logged into a myBYU account
    If IE_Browser.LocationURL <> BYU_URL_1 Or IE_Browser.LocationURL = URL_Check Then
    
        'If so, exit the browser
        IE_Browser.Quit
        
        'Delete the browser object
        Set IE_Browser = Nothing
        
        'Output that the user is already logged off
        MsgBox "You are not logged into your myBYU account. Please login and try again.", vbInformation, "Invalid User"
        
        'Exit the subroutine
        Exit Sub
        
    Else
        
        'Show the userform
        Convocation_Select.Show
        
        'Check for the April year 1 convocation
        If Convocation_Select.April_Radio_Button = True And Convocation_Select.Year_1_Radio_Button = True Then
        
            'Save the convocation date
            Convocation_Date_Part_1 = "December " & Convocation_Select.Year_1_Label.Caption - 1
            Convocation_Date_Part_2 = Convocation_Select.April_Label.Caption & " " & Convocation_Select.Year_1_Label.Caption
            Convocation_Compare_Date = "04/30/" & Convocation_Select.Year_1_Label.Caption
            
        'Check for the April year 2 convocation
        ElseIf Convocation_Select.April_Radio_Button = True And Convocation_Select.Year_2_Radio_Button = True Then
            
            'Save the convocation date
            Convocation_Date_Part_1 = "December " & Convocation_Select.Year_1_Label.Caption - 1
            Convocation_Date_Part_2 = Convocation_Select.April_Label.Caption & " " & Convocation_Select.Year_2_Label.Caption
            Convocation_Compare_Date = "04/30/" & Convocation_Select.Year_2_Label.Caption
            
        'Check for the August year 1 convocation
        ElseIf Convocation_Select.August_Radio_Button = True And Convocation_Select.Year_1_Radio_Button = True Then
        
            'Save the convocation date
            Convocation_Date_Part_1 = "June " & Convocation_Select.Year_1_Label.Caption
            Convocation_Date_Part_2 = Convocation_Select.August_Label.Caption & " " & Convocation_Select.Year_1_Label.Caption
            Convocation_Compare_Date = "08/31/" & Convocation_Select.Year_1_Label.Caption
            
        'Check for the August year 2 convocation
        ElseIf Convocation_Select.August_Radio_Button = True And Convocation_Select.Year_2_Radio_Button = True Then
            
            'Save the convocation date
            Convocation_Date_Part_1 = "June " & Convocation_Select.Year_2_Label.Caption
            Convocation_Date_Part_2 = Convocation_Select.August_Label.Caption & " " & Convocation_Select.Year_2_Label.Caption
            Convocation_Compare_Date = "08/31/" & Convocation_Select.Year_2_Label.Caption
            
        Else
            
            'Output an error message, and exit the subroutine
            MsgBox "You did not select a convocation date. The program will now close.", vbInformation, "No Convocation Date"
            Exit Sub
            
        End If
        
        'Unload the userform
        Unload Convocation_Select
        
        'Check if the selected convocation is from a past convocation
        If Date > DateValue(Convocation_Compare_Date) Then
        
            'If so, output an error message and exit the subroutine
            MsgBox "The convocation you selected has already passed. The program is unable to retreive data from past convocations and will now close.", vbInformation, "Passed Convocation"
            Exit Sub
            
        Else
        
            Application.ScreenUpdating = False
            
            'Save the browser
            Set HTML_Document = IE_Browser.Document
            
            'Get the drop down box
            Drop_Down_Box = HTML_Document.getElementsByTagName("select")
        
            'Look through each drop down option
            For Each Drop_Down_Option In Drop_Down_Box.Options
        
                'Check if the option matches the selected convocation
                If Drop_Down_Option.innerHTML = Convocation_Date_Part_1 Then
            
                    'If so, select that option
                    Drop_Down_Option.Selected = True
                
                End If
        
            Next
            
            'Enter in the life science college code
            HTML_Document.forms("schedule_input").elements("acad_unit_id").Value = 1006
            
            'Get each input object
            Set HTML_Objects = HTML_Document.getElementsByTagName("input")
            
            'Loop through each input tage
            For Each HTML_Element In HTML_Objects
                
                'Check if the go button is found
                If (HTML_Element.getAttribute("Name") = "GoButton") Then
                    
                    'If so, click the button and exit the for loop
                    HTML_Element.Click
                    Exit For
                    
                End If
                
            Next HTML_Element
            
            'Wait for the page to load
            Application.Wait (Now + TimeValue("00:00:05"))
            
            Do
            Loop Until IE_Browser.readyState = 4
            
            'Save the browser
            Set HTML_Document = IE_Browser.Document
            
            'Get each data entry in the table
            Set HTML_Objects = HTML_Document.getElementById("stuDegrees").getElementsByTagName("tr")
            
            'Find the next blank row
            Next_Row = Range("A" & Rows.Count).End(xlUp).Row + 1
            
            'Loop through each element in the table
            For Each HTML_Element In HTML_Objects
                
                'Save each row in the table
                Set Row_Data = HTML_Element.getElementsByTagName("td")
                
                'Check if there is an evaluation value  <-- (Doesn't actually work for some reason)
                If Row_Data.Item(4).innerText <> "" Then
                
                    'If so, save the data to the excel sheet
                    ActiveWorkbook.Worksheets("Master Database").Range("A" & Next_Row).Value = Split(Row_Data.Item(2).innerText, ",")(0)
                    ActiveWorkbook.Worksheets("Master Database").Range("B" & Next_Row).Value = Split(Row_Data.Item(2).innerText, ",")(1)
                    ActiveWorkbook.Worksheets("Master Database").Range("C" & Next_Row).Value = Replace(Row_Data.Item(1).innerText, "-", "")
                    
                    If Row_Data.Item(10).innerText = "  " Then
                    
                        ActiveWorkbook.Worksheets("Master Database").Range("D" & Next_Row).Value = Row_Data.Item(9).innerText
                    
                    Else
                    
                        ActiveWorkbook.Worksheets("Master Database").Range("D" & Next_Row).Value = Row_Data.Item(9).innerText & "(" & Row_Data.Item(10).innerText & ") "
                        
                    End If
                    
                    If Row_Data.Item(11).innerText <> "  " And Row_Data.Item(13).innerText <> " " Then
                    
                        ActiveWorkbook.Worksheets("Master Database").Range("E" & Next_Row).Value = "Phi Kappa Phi, " & Row_Data.Item(11).innerText
                        
                    ElseIf Row_Data.Item(11).innerText = "  " And Row_Data.Item(13).innerText <> " " Then
                    
                        ActiveWorkbook.Worksheets("Master Database").Range("E" & Next_Row).Value = "Phi Kappa Phi"
                        
                    ElseIf Row_Data.Item(11).innerText <> "  " And Row_Data.Item(13).innerText = " " Then
                    
                        ActiveWorkbook.Worksheets("Master Database").Range("E" & Next_Row).Value = Row_Data.Item(11).innerText
                    
                    End If
                    
                    'Find the next row
                    Next_Row = Next_Row + 1
                    
                End If
                
            Next HTML_Element
            
            'Get the drop down box
            Drop_Down_Box = HTML_Document.getElementsByTagName("select")
        
            'Look through each drop down option
            For Each Drop_Down_Option In Drop_Down_Box.Options
        
                'Check if the option matches the selected convocation
                If Drop_Down_Option.innerHTML = Convocation_Date_Part_2 Then
            
                    'If so, select that option
                    Drop_Down_Option.Selected = True
                
                End If
        
            Next
            
            'Get each input object
            Set HTML_Objects = HTML_Document.getElementsByTagName("input")
            
            'Loop through each input tage
            For Each HTML_Element In HTML_Objects
                
                'Check if the go button is found
                If (HTML_Element.getAttribute("Name") = "GoButton") Then
                    
                    'If so, click the button and exit the for loop
                    HTML_Element.Click
                    Exit For
                    
                End If
                
            Next HTML_Element
            
            'Wait for the page to load state 4
            Application.Wait (Now + TimeValue("00:00:05"))
            
            Do
            Loop Until IE_Browser.readyState = 4
            
            'Save the browser
            Set HTML_Document = IE_Browser.Document
            
            'Get each data entry in the table
            Set HTML_Objects = HTML_Document.getElementById("stuDegrees").getElementsByTagName("tr")
            
            'Loop through each element in the table
            For Each HTML_Element In HTML_Objects
                
                'Save each row in the table
                Set Row_Data = HTML_Element.getElementsByTagName("td")
                
                'Check if there is an evaluation value
                If Row_Data.Item(4).innerText <> " " Then
                
                    'If so, save the data to the excel sheet
                    ActiveWorkbook.Worksheets("Master Database").Range("A" & Next_Row).Value = Split(Row_Data.Item(2).innerText, ",")(0)
                    ActiveWorkbook.Worksheets("Master Database").Range("B" & Next_Row).Value = Split(Row_Data.Item(2).innerText, ",")(1)
                    ActiveWorkbook.Worksheets("Master Database").Range("C" & Next_Row).Value = Replace(Row_Data.Item(1).innerText, "-", "")
                    
                    'Check if there is an emphasis
                    If Row_Data.Item(10).innerText = "  " Then
                    
                        'If not, just record the major
                        ActiveWorkbook.Worksheets("Master Database").Range("D" & Next_Row).Value = Row_Data.Item(9).innerText
                    
                    Else
                    
                        'If so, record the major and the emphasis
                        ActiveWorkbook.Worksheets("Master Database").Range("D" & Next_Row).Value = Row_Data.Item(9).innerText & "(" & ActiveWorkbook.Worksheets("Raw Data").Range("D" & Next_Row).Value = Row_Data.Item(10).innerText & ") "
                        
                    End If
                    
                    'Check if for honors and phi kappa phi and record the appropriate honors
                    If Row_Data.Item(11).innerText <> "  " And Row_Data.Item(12).innerText <> " " Then
                    
                        ActiveWorkbook.Worksheets("Master Database").Range("F" & Next_Row).Value = "Phi Kappa Phi, " & Row_Data.Item(11)
                        
                    ElseIf Row_Data.Item(11).innerText = "  " And Row_Data.Item(12).innerText <> " " Then
                    
                        ActiveWorkbook.Worksheets("Master Database").Range("F" & Next_Row).Value = "Phi Kappa Phi"
                        
                    ElseIf Row_Data.Item(11).innerText <> "  " And Row_Data.Item(12).innerText = " " Then
                    
                        ActiveWorkbook.Worksheets("Master Database").Range("F" & Next_Row).Value = Row_Data.Item(11).innerText
                    
                    End If

                    'Find the next row
                    Next_Row = Next_Row + 1
                    
                End If
                
            Next HTML_Element
            
            'Go back to the last row with data
            Next_Row = Next_Row - 1
            
            'Get all the cells with data in them
            Set Data_Range = Range("A4:A" & Next_Row)
            
            'Loop through each cell in the data range
            For Each Cell In Data_Range
                
                'Activate the cell
                Cell.Activate
                
                'Check if the cell is empty
                If Cell.Value = "" Then
                
                    'If so, delete that row
                    Range("A" & ActiveCell.Row & ":H" & ActiveCell.Row).Select
                    Selection.Delete Shift:=xlUp
                    
                End If
                    
            Next Cell
            
            'Get the last row with data
            Next_Row = Range("A" & Rows.Count).End(xlUp).Row

            'Get all the cells with data in them
            Set Data_Range = Range("C4:C" & Next_Row)
            
            'Navigate to the persumm page
            IE_Browser.Navigate BYU_URL_2
                
            
            'Wait for the page to load state 4
            Do
            Loop Until IE_Browser.readyState = 4
            
            'Check if the browser was able to navigate to the correct page
            If IE_Browser.LocationURL <> BYU_URL_2 Then
            
                'If not, delete all the data and exit the subroutine
                Set Data_Range = Range("A4:H" & Next_Row)
                Data_Range.Select
                Selection.Delete
                Exit Sub
                
            End If
            
            'Set the second go
            Second_Go = False
                
            Set HTML_Document = IE_Browser.Document
                
            'Input the student number into the search box
            HTML_Document.all.NAVpattern.Value = Cells(4, 3)
                
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
            Loop Until IE_Browser.readyState = 4
                
            Application.Wait (Now + TimeValue("00:00:001"))
                
            'Save the browser
            Set HTML_Document = IE_Browser.Document
    
            'Save the student email
            For Each HTML_Element In HTML_Document.getElementsByTagName("input")
                
                'Check for the student email
                If HTML_Element.Name = "restore_email" Then
                        
                    'If found, save that value in the email column and exit the for loop
                    Cells(Cell.Row, 7).Value = HTML_Element.Value
                    On Error GoTo Error_Handling
                    Exit For
                        
                End If
                
            Next HTML_Element
            
            'Loop through each cell in the data range
            For Each Cell In Data_Range
                
                'Reset the second go
                Second_Go = False
                
                Set HTML_Document = IE_Browser.Document
                
                'Input the student number into the search box
                HTML_Document.all.NAVpattern.Value = Cell.Value
                
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
                Loop Until IE_Browser.readyState = 4
                
                Application.Wait (Now + TimeValue("00:00:001"))
                
                'Save the browser
                Set HTML_Document = IE_Browser.Document
    
                'Save the student email
                For Each HTML_Element In HTML_Document.getElementsByTagName("input")
                
                    'Check for the student email
                    If HTML_Element.Name = "restore_email" Then
                        
                        'If found, save that value in the email column and exit the for loop
                        Cells(Cell.Row, 7).Value = HTML_Element.Value
                        On Error GoTo Error_Handling
                        Exit For
                        
                    End If
                
                Next HTML_Element
                    
            Next Cell
            
            'Get all the cells with data in them
            Set Data_Range = Range("G4:G" & Next_Row)
            
            'Loop through each cell with data
            For Each Cell In Data_Range
                
                'Reset the second go
                Second_Go = False
                
                'Activate the cell
                Cell.Activate
                
                'Save the active cell as the current cell
                Current_Cell = ActiveCell
                
                'Save the next cell
                Next_Cell = ActiveCell.Offset(1, 0)
                
                'Check if the current cell and the next cell are the same
                If Current_Cell = Next_Cell Then
                
                    'If so, get the correct email
                    'Save the browser
                    Set HTML_Document = IE_Browser.Document
                
                    'Input the student number into the search box
                    HTML_Document.all.NAVpattern.Value = Cell.Offset(1, -4).Value
                
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
                    Loop Until IE_Browser.readyState = 4
                
                    Application.Wait (Now + TimeValue("00:00:001"))
                
                    'Save the browser
                    Set HTML_Document = IE_Browser.Document
    
                    'Save the student email
                    For Each HTML_Element In HTML_Document.getElementsByTagName("input")
                
                        'Check for the student email
                        If HTML_Element.Name = "restore_email" Then
                        
                            'If found, save that value in the email column and exit the for loop
                            Cell.Offset(1, 0).Value = HTML_Element.Value
                            On Error GoTo Error_Handling
                            Exit For
                        
                        End If
                
                    Next HTML_Element
                    
                End If
                
            Next Cell
            
            'Get all the cells with data in them
            Set Data_Range = Range("A4:A" & Next_Row)
            
            'Loop through each cell with data
            For Each Cell In Data_Range
            
                'Remove the starting spaces
                Cell.Value = Replace(Cell.Value, "  ", "")
                
            Next Cell
            
            'Get all the cells with data in them
            Set Data_Range = Range("B4:B" & Next_Row)
            
            'Loop through each cell with data
            For Each Cell In Data_Range
                
                'Remove the starting space
                Cell.Value = Right(Cell.Value, Len(Cell.Value) - 1)
                
            Next Cell
            
            'Get all the cells with data in them
            Set Data_Range = Range("D4:D" & Next_Row)
            
            'Loop through each cell with data
            For Each Cell In Data_Range
                
                'Remove the starting space
                Cell.Value = Right(Cell.Value, Len(Cell.Value) - 1)
                
                'Remove the ending space
                Cell.Value = Left(Cell.Value, Len(Cell.Value) - 1)
                
            Next Cell
            
            'Get all the cells with data in them
            Set Data_Range = Range("E4:E" & Next_Row)
            
            'Loop through each cell with data
            For Each Cell In Data_Range
                
                'Remove the starting space
                Cell.Value = Replace(Cell.Value, " C", "Cum Laude")
                Cell.Value = Replace(Cell.Value, " A", "Highest Honors")
                Cell.Value = Replace(Cell.Value, " S", "Summa Cum Laude")
                Cell.Value = Replace(Cell.Value, " M", "Magna Cum Laude")
                
            Next Cell
                    
            
            'Exit and delete the browser
            IE_Browser.Quit
            Set IE_Browser = Nothing
        
            'Clear any errors
            Err.Clear
            
            'Turn on screen updating
            Application.ScreenUpdating = True
        
        End If
        
    End If
    
Error_Handling:

    'Check for errors
    If Err.Number <> 0 Then
        
        'If there are any, output an error message
        MsgBox "The program is having trouble establishing an internet connection. Please wait a moment and try again.", vbCritical, "Failed Internet Connection"
        
    End If
    
End Sub