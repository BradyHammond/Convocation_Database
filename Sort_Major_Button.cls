'========================= CONVOCATION DATABASE SORT MAJORS BUTTON ========================='
'Created By: Brady Hammond
'Created: 06/13/16
'========================================================================================'

'Sort Majors Button Subroutine:
Sub Sort_Majors_Button()

    'Declare needed variables
    Dim Biology_Next_Row As Long
    Dim Exercise_Science_Next_Row As Long
    Dim Health_Science_Next_Row As Long
    Dim MMBIO_Next_Row As Long
    Dim NDFS_Next_Row As Long
    Dim PDBIO_Next_Row As Long
    Dim PWS_Next_Row As Long
    Dim Cell As Variant
    Dim Master_Database_Last_Row As Long
    Dim Data_Range As Range
    Dim Response
    
    'Find the last row with data
    Master_Database_Last_Row = ActiveWorkbook.Worksheets("Master Database").Range("A" & Rows.Count).End(xlUp).Row
    
    'Get the entire range of data
    Set Data_Range = ActiveWorkbook.Worksheets("Master Database").Range("D4:D" & Master_Database_Last_Row)
    
    'Check for the last row on each department page
    Biology_Next_Row = ActiveWorkbook.Worksheets("Biology").Range("E" & Rows.Count).End(xlUp).Row + 1
    Exercise_Science_Next_Row = ActiveWorkbook.Worksheets("Exercise Science").Range("E" & Rows.Count).End(xlUp).Row + 1
    Health_Science_Next_Row = ActiveWorkbook.Worksheets("Health Science").Range("E" & Rows.Count).End(xlUp).Row + 1
    MMBIO_Next_Row = ActiveWorkbook.Worksheets("Micro. & Molecular Biology").Range("E" & Rows.Count).End(xlUp).Row + 1
    NDFS_Next_Row = ActiveWorkbook.Worksheets("NDFS").Range("E" & Rows.Count).End(xlUp).Row + 1
    PDBIO_Next_Row = ActiveWorkbook.Worksheets("Physiology & Dev. Biology").Range("E" & Rows.Count).End(xlUp).Row + 1
    PWS_Next_Row = ActiveWorkbook.Worksheets("Plant & Wildlife Sciences").Range("E" & Rows.Count).End(xlUp).Row + 1
    
    'Check if any of the department pages have data on them already
    If Biology_Next_Row <> 2 Or Exercise_Science_Next_Row <> 2 Or Health_Science_Next_Row <> 2 Or MMBIO_Next_Row <> 2 Or NDFS_Next_Row <> 2 Or PDBIO_Next_Row <> 2 Or PWS_Next_Row <> 2 Then
    
        'If so, ask if the user would like to resort the data
        Response = MsgBox("There is already sorted data present. Would you like to resort the data?", vbYesNo + vbInformation, "Resort Data")
        
        'Check if the user said yes
        If Response = vbYes Then
        
            'Turn off screen updating
            Application.ScreenUpdating = False
        
            'Clear each department page
            ActiveWorkbook.Worksheets("Biology").Range("A2:L" & ActiveWorkbook.Worksheets("Biology").Range("E" & Rows.Count).End(xlUp).Row + 1).Delete
            ActiveWorkbook.Worksheets("Exercise Science").Range("A2:L" & ActiveWorkbook.Worksheets("Exercise Science").Range("E" & Rows.Count).End(xlUp).Row + 1).Delete
            ActiveWorkbook.Worksheets("Health Science").Range("A2:L" & ActiveWorkbook.Worksheets("Health Science").Range("E" & Rows.Count).End(xlUp).Row + 1).Delete
            ActiveWorkbook.Worksheets("Micro. & Molecular Biology").Range("A2:L" & ActiveWorkbook.Worksheets("Micro. & Molecular Biology").Range("E" & Rows.Count).End(xlUp).Row + 1).Delete
            ActiveWorkbook.Worksheets("NDFS").Range("A2:L" & ActiveWorkbook.Worksheets("NDFS").Range("E" & Rows.Count).End(xlUp).Row + 1).Delete
            ActiveWorkbook.Worksheets("Physiology & Dev. Biology").Range("A2:L" & ActiveWorkbook.Worksheets("Physiology & Dev. Biology").Range("E" & Rows.Count).End(xlUp).Row + 1).Delete
            ActiveWorkbook.Worksheets("Plant & Wildlife Sciences").Range("A2:L" & ActiveWorkbook.Worksheets("Plant & Wildlife Sciences").Range("E" & Rows.Count).End(xlUp).Row + 1).Delete
            
            'Loop through each cell in the data range
            For Each Cell In Data_Range
            
                'Find the next Row with data
                Biology_Next_Row = ActiveWorkbook.Worksheets("Biology").Range("E" & Rows.Count).End(xlUp).Row + 1
                Exercise_Science_Next_Row = ActiveWorkbook.Worksheets("Exercise Science").Range("E" & Rows.Count).End(xlUp).Row + 1
                Health_Science_Next_Row = ActiveWorkbook.Worksheets("Health Science").Range("E" & Rows.Count).End(xlUp).Row + 1
                MMBIO_Next_Row = ActiveWorkbook.Worksheets("Micro. & Molecular Biology").Range("E" & Rows.Count).End(xlUp).Row + 1
                NDFS_Next_Row = ActiveWorkbook.Worksheets("NDFS").Range("E" & Rows.Count).End(xlUp).Row + 1
                PDBIO_Next_Row = ActiveWorkbook.Worksheets("Physiology & Dev. Biology").Range("E" & Rows.Count).End(xlUp).Row + 1
                PWS_Next_Row = ActiveWorkbook.Worksheets("Plant & Wildlife Sciences").Range("E" & Rows.Count).End(xlUp).Row + 1
                
                'Check if the student is in the biology department
                If Cell.Value = "Bioinformatics" Or Cell.Value = "Biological Science Education" Or Cell.Value = "Biology" Or Cell.Value = "Conservation Biology" Then
                
                    'If so, paste the student's info on that department's page
                    ActiveWorkbook.Worksheets("Master Database").Range("A" & Cell.Row & ":H" & Cell.Row).Copy
                    ActiveWorkbook.Worksheets("Biology").Cells(Biology_Next_Row, 5).PasteSpecial Paste:=xlPasteValues
                
                'Check if the student is in the exercise science department
                ElseIf Cell.Value = "Athletic Training" Or Cell.Value = "Exercise & Wellness" Or Cell.Value = "Exercise Science" Or Cell.Value = "Physical Education ( Exercise Science )" Then
                    
                    'If so, paste the student's info on that department's page
                    ActiveWorkbook.Worksheets("Master Database").Range("A" & Cell.Row & ":H" & Cell.Row).Copy
                    ActiveWorkbook.Worksheets("Exercise Science").Cells(Exercise_Science_Next_Row, 5).PasteSpecial Paste:=xlPasteValues
                
                'Check if the student is in the health science department
                ElseIf Cell.Value = "Public Health ( Epidemiology )" Or Cell.Value = "School Health Education" Or Cell.Value = "Public Health ( Health Science )" Or Cell.Value = "Public Health ( Health Promotion )" Or Cell.Value = "Public Health ( Environmental/Occupational Health )" Or Cell.Value = "Public Health" Then
                    
                    'If so, paste the student's info on that department's page
                    ActiveWorkbook.Worksheets("Master Database").Range("A" & Cell.Row & ":H" & Cell.Row).Copy
                    ActiveWorkbook.Worksheets("Health Science").Cells(Health_Science_Next_Row, 5).PasteSpecial Paste:=xlPasteValues
                
                'Check if the student is in the MMBIO department
                ElseIf Cell.Value = "Medical Laboratory Science" Or Cell.Value = "Microbiology" Or Cell.Value = "Molecular Biology" Then
                    
                    'If so, paste the student's info on that department's page
                    ActiveWorkbook.Worksheets("Master Database").Range("A" & Cell.Row & ":H" & Cell.Row).Copy
                    ActiveWorkbook.Worksheets("Micro. & Molecular Biology").Cells(MMBIO_Next_Row, 5).PasteSpecial Paste:=xlPasteValues
                
                'Check if the student is in the NDFS department
                ElseIf Cell.Value = "Dietetics" Or Cell.Value = "Food Science" Or Cell.Value = "Nutritional Science" Then
                    
                    'If so, paste the student's info on that department's page
                    ActiveWorkbook.Worksheets("Master Database").Range("A" & Cell.Row & ":H" & Cell.Row).Copy
                    ActiveWorkbook.Worksheets("NDFS").Cells(NDFS_Next_Row, 5).PasteSpecial Paste:=xlPasteValues
                
                'Check if the student is in the PDBIO department
                ElseIf Cell.Value = "Biophysics" Or Cell.Value = "Physiology & Developmental Biology" Then
                
                    'If so, paste the student's info on that department's page
                    ActiveWorkbook.Worksheets("Master Database").Range("A" & Cell.Row & ":H" & Cell.Row).Copy
                    ActiveWorkbook.Worksheets("Physiology & Dev. Biology").Cells(PDBIO_Next_Row, 5).PasteSpecial Paste:=xlPasteValues
                
                'Check if the student is in the PWS department
                ElseIf Cell.Value = "Environmental Science" Or Cell.Value = "Genetics & Biotechnology" Or Cell.Value = "Landscape Management" Or Cell.Value = "Wildlife & Wildlands Conservation" Then
                
                    'If so, paste the student's info on that department's page
                    ActiveWorkbook.Worksheets("Master Database").Range("A" & Cell.Row & ":H" & Cell.Row).Copy
                    ActiveWorkbook.Worksheets("Plant & Wildlife Sciences").Cells(PWS_Next_Row, 5).PasteSpecial Paste:=xlPasteValues
                
                Else
                
                    'Color the cell
                    Cell.Interior.ColorIndex = 28
                
                End If
                
            Next Cell
            
            'Reset the master database as the active page
            ActiveWorkbook.Worksheets("Master Database").Activate
            
            'Turn on screen updating
            Application.ScreenUpdating = True
            
        Else
            
            'Exit the subroutine
            Exit Sub
        
        End If
        
    Else
    
        'Turn off screen updating
        Application.ScreenUpdating = False
    
        'Loop through each cell in the data range
        For Each Cell In Data_Range
        
            'Find the next Row with data
            Biology_Next_Row = ActiveWorkbook.Worksheets("Biology").Range("E" & Rows.Count).End(xlUp).Row + 1
            Exercise_Science_Next_Row = ActiveWorkbook.Worksheets("Exercise Science").Range("E" & Rows.Count).End(xlUp).Row + 1
            Health_Science_Next_Row = ActiveWorkbook.Worksheets("Health Science").Range("E" & Rows.Count).End(xlUp).Row + 1
            MMBIO_Next_Row = ActiveWorkbook.Worksheets("Micro. & Molecular Biology").Range("E" & Rows.Count).End(xlUp).Row + 1
            NDFS_Next_Row = ActiveWorkbook.Worksheets("NDFS").Range("E" & Rows.Count).End(xlUp).Row + 1
            PDBIO_Next_Row = ActiveWorkbook.Worksheets("Physiology & Dev. Biology").Range("E" & Rows.Count).End(xlUp).Row + 1
            PWS_Next_Row = ActiveWorkbook.Worksheets("Plant & Wildlife Sciences").Range("E" & Rows.Count).End(xlUp).Row + 1
            
            'Check if the student is in the biology department
            If Cell.Value = "Bioinformatics" Or Cell.Value = "Biological Science Education" Or Cell.Value = "Biology" Or Cell.Value = "Conservation Biology" Then
                
                'If so, paste the student's info on that department's page
                ActiveWorkbook.Worksheets("Master Database").Range("A" & Cell.Row & ":H" & Cell.Row).Copy
                ActiveWorkbook.Worksheets("Biology").Cells(Biology_Next_Row, 5).PasteSpecial Paste:=xlPasteValues
                
            'Check if the student is in the exercise science department
            ElseIf Cell.Value = "Athletic Training" Or Cell.Value = "Exercise & Wellness" Or Cell.Value = "Exercise Science" Or Cell.Value = "Physical Education ( Exercise Science )" Then
                    
                'If so, paste the student's info on that department's page
                ActiveWorkbook.Worksheets("Master Database").Range("A" & Cell.Row & ":H" & Cell.Row).Copy
                ActiveWorkbook.Worksheets("Exercise Science").Cells(Exercise_Science_Next_Row, 5).PasteSpecial Paste:=xlPasteValues
                
            'Check if the student is in the health science department
            ElseIf Cell.Value = "Public Health ( Epidemiology )" Or Cell.Value = "School Health Education" Or Cell.Value = "Public Health ( Health Science )" Or Cell.Value = "Public Health ( Health Promotion )" Or Cell.Value = "Public Health ( Environmental/Occupational Health )" Or Cell.Value = "Public Health" Then
                    
                'If so, paste the student's info on that department's page
                ActiveWorkbook.Worksheets("Master Database").Range("A" & Cell.Row & ":H" & Cell.Row).Copy
                ActiveWorkbook.Worksheets("Health Science").Cells(Health_Science_Next_Row, 5).PasteSpecial Paste:=xlPasteValues
                
            'Check if the student is in the MMBIO department
            ElseIf Cell.Value = "Medical Laboratory Science" Or Cell.Value = "Microbiology" Or Cell.Value = "Molecular Biology" Then
                    
                'If so, paste the student's info on that department's page
                ActiveWorkbook.Worksheets("Master Database").Range("A" & Cell.Row & ":H" & Cell.Row).Copy
                ActiveWorkbook.Worksheets("Micro. & Molecular Biology").Cells(MMBIO_Next_Row, 5).PasteSpecial Paste:=xlPasteValues
                
            'Check if the student is in the NDFS department
            ElseIf Cell.Value = "Dietetics" Or Cell.Value = "Food Science" Or Cell.Value = "Nutritional Science" Then
                    
                'If so, paste the student's info on that department's page
                ActiveWorkbook.Worksheets("Master Database").Range("A" & Cell.Row & ":H" & Cell.Row).Copy
                ActiveWorkbook.Worksheets("NDFS").Cells(NDFS_Next_Row, 5).PasteSpecial Paste:=xlPasteValues
                
            'Check if the student is in the PDBIO department
            ElseIf Cell.Value = "Biophysics" Or Cell.Value = "Physiology & Developmental Biology" Then
                
                'If so, paste the student's info on that department's page
                ActiveWorkbook.Worksheets("Master Database").Range("A" & Cell.Row & ":H" & Cell.Row).Copy
                ActiveWorkbook.Worksheets("Physiology & Dev. Biology").Cells(PDBIO_Next_Row, 5).PasteSpecial Paste:=xlPasteValues
                
            'Check if the student is in the PWS department
            ElseIf Cell.Value = "Environmental Science" Or Cell.Value = "Genetics & Biotechnology" Or Cell.Value = "Landscape Management" Or Cell.Value = "Wildlife & Wildlands Conservation" Then
                
                'If so, paste the student's info on that department's page
                ActiveWorkbook.Worksheets("Master Database").Range("A" & Cell.Row & ":H" & Cell.Row).Copy
                ActiveWorkbook.Worksheets("Plant & Wildlife Sciences").Cells(PWS_Next_Row, 5).PasteSpecial Paste:=xlPasteValues
            
            Else
            
                'Color the cell
                Cell.Interior.ColorIndex = 28
                
            End If
            
        Next Cell
        
        'Reset the master database as the active page
        ActiveWorkbook.Worksheets("Master Database").Activate
        
        'Turn on screen updating
        Application.ScreenUpdating = True
    
    End If
    
End Sub
