'========================= CONVOCATION DATABASE CONVOCATION SELECT FORM ========================='
'Created By: Brady Hammond
'Created: 05/18/16
'================================================================================================'

Private Sub Convocation_Select_Frame_Click()

End Sub

'Userform Initialize Subroutine:
Private Sub UserForm_Initialize()
    
    'Clear all radio buttons
    Convocation_Select.April_Radio_Button.Value = False
    Convocation_Select.August_Radio_Button.Value = False
    Convocation_Select.Year_1_Radio_Button.Value = False
    Convocation_Select.Year_2_Radio_Button.Value = False
    
    'Input the correct years
    Convocation_Select.Year_1_Label.Caption = Year(Now())
    Convocation_Select.Year_2_Label.Caption = Year(Now()) + 1

    
End Sub

'Submit Button Click Subroutine:
Private Sub Submit_Button_Click()
    
    'Check if the user selected a month
    If Convocation_Select.April_Radio_Button.Value = False And Convocation_Select.August_Radio_Button.Value = False Then
    
        'If not, output an error message and exit the subroutine
        MsgBox "You must select a month before you can proceed.", vbInformation, "No Month"
        Exit Sub
        
    End If
    
    'Check if the use selected a year
    If Convocation_Select.Year_1_Radio_Button.Value = False And Convocation_Select.Year_2_Radio_Button.Value = False Then
        
        'If not, output an error message and exit the subroutine
        MsgBox "You must select a year before you can proceed.", vbInformation, "No Year"
        Exit Sub
        
    End If
    
    'Hide the userform
    Convocation_Select.Hide
    
End Sub

'Cancel Button Click Subroutine:
Private Sub Cancel_Button_Click()

    'Hide the userform
    Convocation_Select.Hide
    
End Sub
