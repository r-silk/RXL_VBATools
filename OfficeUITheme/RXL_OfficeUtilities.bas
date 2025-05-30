Attribute VB_Name = "RXL_OfficeUtilities"
'##################################################################################################
' Excel and Office Utility Functions
'  RS 21-Feb-2025, v1.5
'  Copyright (c) 2025, RXL Development. All rights reserved.
'
' Developed by RXL Development, Chelmsford, UK
'   Author: Robert Silk  [ https://excel-bits.net ]
'
' The functions within this module are used to read useful information for the current Windows user
' and Excel instance. This includes determining the Windows UI Theme and Office UI Theme so that
' project styling and UI can be updated to match Office apps and user preferences.
' Additional functions can be used to check the full Excel application version and build number, for
' example where an Excel project requires a minimum version or build in order to provide full
' functionality.
'
'
' Functions:
' ----------
'   ReadOfficeUITheme_Value()
'       Purpose:    returns an integer reflecting the chosen Office UI Theme
'       Parameters: (none)
'   ReadOfficeUITheme_String()
'       Purpose:    returns the string name for the chosen Office UI Theme; if set to
'       Parameters: (none)
'   ReadWindowsUITheme_Value()
'       Purpose:    returns an integer reflecting whether Windows apps should use Light or Dark
'                   theme
'       Parameters: (none)
'   ReadExcelBuildNumber_String()
'       Purpose:    returns a string of the Excel.exe version & build number in the format
'                   [major].[minor].[revision].[Build]
'       Parameters: (none)
'   ReadExcelBuildNumber_Array()
'       Purpose:    returns an array of the full version & and build number string for Excel.exe
'                   converted to Long values: (1)[major];(2)[minor];(3)[revision];(4)[Build]
'       Parameters: (none)
'
'
' Code commenting has been used where possible to aid any reader/user in understanding the below
' code and it's intended function. Whilst care has been taken in the development of these tools
' and the VBA code within this Project, for the specific purpose for which it has been designed,
' each user is advised to carry out their own diligence and checks, and should consider the need
' for independent testing before including it within any Project of their own.
' No responsibility is taken or accepted by RXL Development or the author for the adequacy,
' completeness or accuracy of the output produced by these tools, and all liability is therefore
' expressly excluded. Anyone using the macros contained within this Project, or relying upon the
' outputs produced, does so at their own risk, and no responsibility is accepted for any losses
' which may result from such use, directly or indirectly.
'
'##################################################################################################

Option Explicit

    
Public Function ReadOfficeUITheme_Value() As Integer
    
    '### Function used to read the Office UI Theme which is set for the current user
    '### The setting is common across Office and is set via the Registry
    '### The expected path for current user for Office 365 is found at:
    '###    Path - HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\
    '###    Key  - UI Theme
    '### Known values:  3 = Dark Grey
    '###                4 = Black
    '###                5 = White
    '###                6 = Use system setting (i.e. Windows/OS theme setting)
    '###                7 = Colourful
    '### This function use an On Error wrapper to read the Registry entry therefore an error results
    '### in the function returning a value of 0; if the Registry read value is outside of the
    '### expected values above, the function returns a value of -1
    
    ' Define local variables
    Dim objShell As Object
    Dim intS As Integer, intW As Integer
    
    Set objShell = CreateObject("WScript.Shell")
    
    ' Attempt to read the local user Office UI Theme Regisrty entry
    On Error Resume Next
        intS = objShell.RegRead("HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\UI Theme")
    On Error GoTo 0
    
    ' Clear object as no longer needed
    Set objShell = Nothing
    
    ' Check for a value outside the expected and return result
    If intS > 0 And (intS < 3 Or intS > 7) Then
        ReadOfficeUITheme_Value = -1
    Else
        ReadOfficeUITheme_Value = intS
    End If
    
End Function

Public Function ReadOfficeUITheme_String() As String
    
    '### Function used to read the Office UI Theme which is set for the current user
    '### The setting is common across Office applications and is set via the Registry
    '### The expected path for current user for Office 365 is found at:
    '###    Path - HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\
    '###    Key  - UI Theme
    '### Known values:  3 = Dark Grey
    '###                4 = Black
    '###                5 = White
    '###                6 = Use system setting (i.e. Windows/OS theme setting)
    '###                7 = Colourful
    '### This function use an On Error wrapper to read the registry entry therefore an error results
    '### in a value of 0; If the registry read value is outside of the expected values above, the
    '### function returns an appropriate message.
    
    ' Define local variables
    Dim objShell As Object
    Dim intS As Integer, intW As Integer
    Dim strReturn As String
    
    Set objShell = CreateObject("WScript.Shell")
    
    ' Set an error code as default for registry values to find, so that an error becomes more easily apparent
    intW = -1
    ' Attempt to read the local user Office UI Theme regisrty entry
    On Error Resume Next
        intS = objShell.RegRead("HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\UI Theme")
    On Error GoTo 0
    
    ' Clear object as no longer needed
    Set objShell = Nothing
    
    ' Check for a value outside the expected and return result
    Select Case intS
        Case 0
            strReturn = "ERROR reading Registry"
        Case 3
            strReturn = "Dark Grey"
        Case 4
            strReturn = "Black"
        Case 5
            strReturn = "White"
        Case 6
            strReturn = "System setting"
        Case 7
            strReturn = "Colourful"
        Case Else
            strReturn = "(Not recognised)"
    End Select
    
    If intS = 6 Then
        intW = ReadWindowsUITheme_Value
        If intW < 0 Then
            strReturn = "ERROR reading Windows theme"
        Else
            strReturn = strReturn & " (" & IIf(intW = 1, "Light Theme", "Dark Theme") & ")"
        End If
    End If
    
    ReadOfficeUITheme_String = strReturn
    
End Function

Public Function ReadWindowsUITheme_Value() As Integer
    
    '### Function used to read the Windows UI Theme personalisation which for the current user
    '### The AppsUseLightTheme option is set via the Registry
    '### The expected path for current user for Windows Theme personalisation is found at:
    '###    Path - HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize\
    '###    Key  - AppsUseLightTheme
    '### Known values:  0 = False/No (Dark theme)
    '###                1 = True/Yes (Light theme)
    '### This function use an On Error wrapper to read the Registry entry therefore an error results
    '### in the function defaulting to a value of -1; if the Registry read value is outside of the
    '### expected values above, the function returns a value of 2
    
    ' Define local variables
    Dim objShell As Object
    Dim intW As Integer
    
    Set objShell = CreateObject("WScript.Shell")
    
    ' Set an error code as default for registry values to find, so that an error becomes more easily apparent
    intW = -1
    ' Attempt to read the local user Office UI Theme Regisrty entry
    On Error Resume Next
        intW = objShell.RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme")
    On Error GoTo 0
    
    ' Clear object as no longer needed
    Set objShell = Nothing
    
    ' Check for a value outside the expected and return result
    If Abs(intW) > 1 Then
        ReadWindowsUITheme_Value = 2
    Else
        ReadWindowsUITheme_Value = intW
    End If
    
End Function

Public Function ReadExcelBuildNumber_String() As String
    
    '### Function used to read the full version, build, and revision details for Excel
    '### Uses the FileSystemObject to read the excel executable version number information
    '### This is provided in an expected format of:
    '###    [major].[minor].[revision].[build]
    '###
    '### This should match the build number information which can be found when going to
    '### File >> Account >> About Excel
    '###
    '### Uses an On Error wrapper and will return an error message if an error occurs when trying
    '### to read the build number information, otherwise the build number string is returned in full
    
    ' Define local variables
    Dim objFso As Object
    Dim strVersion As String
    
    ' Create FSO object
    Set objFso = CreateObject("Scripting.FileSystemObject")
    ' Attempt to read file version number information with FSO
    On Error Resume Next
        strVersion = objFso.GetFileVersion(Application.Path & "\EXCEL.EXE")
    On Error GoTo 0
    
    ' Clear object as no longer needed
    Set objFso = Nothing
    
    ' Check if an error occurred (null string)
    If strVersion = vbNullString Then _
        strVersion = "ERROR reading build number"
    ' Return result to calling function/sub-routine
    ReadExcelBuildNumber_String = strVersion
        
End Function

Public Function ReadExcelBuildNumber_Array() As Variant
    
    '### Function used to read the full version, build, and revision details for Excel
    '### Uses the FileSystemObject to read the excel executable version number information
    '### This is provided in an expected format of:
    '###    [major].[minor].[revision].[build]
    '###
    '### This should match the build number information which can be found when going to
    '### File >> Account >> About Excel
    '###
    '### Uses an On Error wrapper and will return a -1 value in the 1st index position (1) of the
    '### return array (and zeros by default in index 2:4); otherwise an array is returned with
    '### values for: index(1) = [major]
    '###             index(2) = [minor]
    '###             index(3) = [revision]
    '###             index(4) = [build]
    
    ' Define local variables
    Dim objFso As Object
    Dim strVersion As String, strArray() As String
    Dim i As Integer, arrReturn(1 To 4) As Long
    
    ' Create FSO object
    Set objFso = CreateObject("Scripting.FileSystemObject")
    ' Attempt to read file version number information with FSO
    On Error Resume Next
        strVersion = objFso.GetFileVersion(Application.Path & "\EXCEL.EXE")
    On Error GoTo 0
    
    ' Clear object as no longer needed
    Set objFso = Nothing
    
    ' Check if an error occurred (null string)
    '  --> if null string return -1 in 1st index position of return array; other indexes will be 0
    If strVersion = vbNullString Then
        arrReturn(1) = -1
    Else
    '  --> if not null, split the string and populate the return array
        strArray = Split(strVersion, ".")
        For i = LBound(strArray) To UBound(strArray)
            arrReturn(i + 1) = CLng(strArray(i))
        Next i
    End If
    ' Return result to calling function/sub-routine
    ReadExcelBuildNumber_Array = arrReturn
    
End Function
