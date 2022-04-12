# SmartView_ReqEssbaseExcelSettings
SmartView功能对接Essbase和Excel的底层语言

Option Explicit



Sub RestoreExcelEnviroment()
'Restore regular excel settings
    With application
        .DisplayFormulaBar = True
        .StatusBar = False
        .Caption = "Microsoft Excel"
        .DisplayStatusBar = True
        .DisplayAlerts = True
        .CellDragAndDrop = True
        .CommandBars("Tools").Controls("Customize...").Enabled = True
        .CommandBars("Toolbar List").Enabled = True
    End With
    
    'hide excel tabs
    ActiveWindow.DisplayWorkbookTabs = True
End Sub

Sub ConfigureExcelEnviroment()
    
    With application
        .WindowState = xlMaximized
        .DisplayFullScreen = False
        .StatusBar = "Please wait while this application open..."
        .Caption = DbName
        .CellDragAndDrop = False
        .CommandBars("Tools").Controls("Customize...").Enabled = False
        .CommandBars("Toolbar List").Enabled = False
    End With
    
    'hide excel tabs
    ActiveWindow.DisplayWorkbookTabs = False
    
End Sub
Sub set_default_user_sheet_options()
Dim X As Integer

On Error Resume Next
     
    'this routine sets default Essbase worksheet options for the sheets
   
    Dim ret As Long
    Dim option_val As Variant
    Dim number_option_val As Integer
    
    'disable both formula preservation and supress options as the starting point for the restoration
    ret = EssVSetSheetOption(Empty, 11, False) 'Disable formula preservation on retrievals
    ret = EssVSetSheetOption(Empty, 6, False)  'Disable suppress #Missing setting
    ret = EssVSetSheetOption(Empty, 7, False)  'Disable suppress zeroes setting
    
    
    For X = 28 To 1 Step -1
        
        'First test to see if the option setting contains a value; if not, skip to the next one
        If ThisWorkbook.Sheets("Control").Range("prior_sheet_options").Cells(X, 2).Value <> "" Then
        
            Select Case X
        
                'Do nothing for these two option settings
                Case 20, 4 '#20 is not used by Hyperion
                    'Do nothing for these
            
                'These are the options that involve a number rather than text
                Case 1, 5 '1 = Drill level; 5 = Indention level
                    number_option_val = ThisWorkbook.Sheets("Control").Range("prior_sheet_options").Cells(X, 2).Value
                    ret = EssVSetSheetOption(Empty, X, number_option_val)

                'these are the textual options that are valid only if Formula Preservation is disabled
                Case 6, 7, 24  '6 = Suppress Missing; 7 = Suppress Zeroes; 24 = Enable both member names and aliases
                    'if Formula Preservation is enabled, don't even try to set these options because it will generate an error
                    If ThisWorkbook.Sheets("Control").Range("prior_sheet_options").Cells(11, 2).Value = False Then
                        option_val = ThisWorkbook.Sheets("Control").Range("prior_sheet_options").Cells(X, 2).Value
                        ret = EssVSetSheetOption(Empty, X, option_val)
                    End If
                    
                'These are the rest of the textual option settings (the validity of each is independent of the settings of other options)
                Case Else
                    option_val = ThisWorkbook.Sheets("Control").Range("prior_sheet_options").Cells(X, 2).Value
                    ret = EssVSetSheetOption(Empty, X, option_val)
        
            End Select
        
        End If
        
    Next X

End Sub


Sub restore_user_sheet_options()
Dim X As Integer
Dim a, b As Boolean

On Error Resume Next
     
    'this routine sets default Essbase worksheet options for report sheets
   
    Dim ret As Long
    Dim option_val As Variant
    Dim number_option_val As Integer
    
    'disable both formula preservation and supress options as the starting point for the restoration
    ret = EssVSetSheetOption(Empty, 11, False) 'Disable formula preservation on retrievals
    
    a = EssVGetSheetOption(Empty, 6)
    b = EssVGetSheetOption(Empty, 7)
    
    If a = True Then
        ret = EssVSetSheetOption(Empty, 6, False)  'Disable suppress #Missing setting
    End If
    If b = True Then
        ret = EssVSetSheetOption(Empty, 7, False)  'Disable suppress zeroes setting
    End If
    
    For X = 28 To 1 Step -1
        
        'First test to see if the option setting contains a value; if not, skip to the next one
        If ThisWorkbook.Sheets("Control").Range("user_sheet_options").Cells(X, 2).Value <> "" Then
        
            Select Case X
        
                'Do nothing for these two option settings
                Case 20, 4 '#20 is not used by Hyperion
                    'Do nothing for these
            
                'These are the options that involve a number rather than text
                Case 1, 5  '1 = Drill level; 5 = Indention level
                    number_option_val = ThisWorkbook.Sheets("Control").Range("user_sheet_options").Cells(X, 2).Value
                    ret = EssVSetSheetOption(Empty, X, number_option_val)

                'these are the textual options that are valid only if Formula Preservation is disabled
                Case 6, 7, 24  '6 = Suppress Missing; 7 = Suppress Zeroes; 24 = Enable both member names and aliases
                    'if Formula Preservation is enabled, don't even try to set these options because it will generate an error
                       'If x = 6 And a = True Then
                        If ThisWorkbook.Sheets("Control").Range("user_sheet_options").Cells(11, 2).Value = False Then
                            option_val = ThisWorkbook.Sheets("Control").Range("user_sheet_options").Cells(X, 2).Value
                            ret = EssVSetSheetOption(Empty, X, option_val)
                        End If
                    
                'These are the rest of the textual option settings (the validity of each is independent of the settings of other options)
                Case Else
                    option_val = ThisWorkbook.Sheets("Control").Range("user_sheet_options").Cells(X, 2).Value
                    ret = EssVSetSheetOption(Empty, X, option_val)
        
            End Select
        
        End If
        
    Next X

End Sub

Sub set_default_user_global_options()
Dim X As Integer
 
On Error Resume Next
     
    'this routine sets default Essbase worksheet options for report sheets
   
    Dim ret As Long
    Dim option_val As Variant
    Dim number_option_val As Integer
    
   
    For X = 11 To 1 Step -1
        
        'First test to see if the option setting contains a value; if not, skip to the next one
        If ThisWorkbook.Sheets("Control").Range("prior_global_options").Cells(X, 2).Value <> "" Then
        
            Select Case X
        
                'Do nothing for this option setting
                Case 4
                    'Do nothing here
            
                'These is the only option that involves a number rather than text
                Case 5
                    number_option_val = ThisWorkbook.Sheets("Control").Range("prior_global_options").Cells(X, 2).Value
                    ret = EssVSetGlobalOption(X, number_option_val)
                
                'These are the textual option settings
                Case Else
                    option_val = ThisWorkbook.Sheets("Control").Range("prior_global_options").Cells(X, 2).Value
                    ret = EssVSetGlobalOption(X, option_val)
        
            End Select
        
        End If
        
    Next X
End Sub

Sub restore_user_global_options()
Dim X As Integer
 
On Error Resume Next
     
    'this routine sets default Essbase worksheet options for report sheets
   
    Dim ret As Long
    Dim option_val As Variant
    Dim number_option_val As Integer
    
   
    For X = 11 To 1 Step -1
        
        'First test to see if the option setting contains a value; if not, skip to the next one
        If ThisWorkbook.Sheets("Control").Range("user_global_options").Cells(X, 2).Value <> "" Then
        
            Select Case X
        
                'Do nothing for this option setting
                Case 4
                    'Do nothing here
            
                'These is the only option that involves a number rather than text
                Case 5
                    number_option_val = ThisWorkbook.Sheets("Control").Range("user_global_options").Cells(X, 2).Value
                    ret = EssVSetGlobalOption(X, number_option_val)
                
                'These are the textual option settings
                Case Else
                    option_val = ThisWorkbook.Sheets("Control").Range("user_global_options").Cells(X, 2).Value
                    ret = EssVSetGlobalOption(X, option_val)
        
            End Select
        
        End If
        
    Next X

End Sub

Sub get_user_sheet_options()
Dim X As Integer

'this routine reads and saves the user's Essbase worksheet options before they are updated by the report workbook
   
On Error Resume Next
   
    Dim ret As Variant
    
    application.ScreenUpdating = False
    
    'add a new worksheet to the workbook -- this sheet's options default to those last set by the user
    'ThisWorkbook.Sheets.Add
    
    'get the Essbase sheet options off of this new worksheet
    For X = 1 To 28
        If X <> 20 Then
            ret = EssVGetSheetOption(Empty, X)
            ThisWorkbook.Sheets("Control").Range("user_sheet_options").Cells(X, 2).Value = ret
        
            'this IF statement is a workaround for an apparent bug related to how the Advanced Interpretation Mode option gets reset
            If X = 16 And ret = False Then
                ThisWorkbook.Sheets("Control").Range("user_sheet_options").Cells(15, 2).Value = "False"
            End If
            
        End If
    Next X

    'NOTE: With the new worksheet still active, Excel will call the get_user_global_options below

End Sub

Sub get_user_global_options()
     
'this routine reads and saves the user's Essbase global options before they are updated by the report workbook
    
On Error Resume Next
    
    Dim ret As Variant
    Dim X As Variant
    
    'get the Essbase global options off of this new worksheet
    For X = 1 To 11
        If X <> 4 Then
            ret = EssVGetGlobalOption(X)
            ThisWorkbook.Sheets("Control").Range("user_global_options").Cells(X, 2).Value = ret
        End If
    Next X

    application.DisplayAlerts = False

End Sub


