Attribute VB_Name = "RXL_ProgressBar"
'##################################################################################################
' Excel Status Bar Progress Bar
'   RS 21-Feb-2025, v2.0
'   Copyright (c) 2025, RXL Development. All rights reserved.
'
' Developed by RXL Development, Chelmsford, UK
'   Author: Robert Silk  [ https://excel-bits.net ]
'
' This module can be used to display a progress bar in the Excel Status Bar in order to provide
' visual feedback to the user while a macro is running.
' The overhead required for placing the progress bar in the StatusBar as a constructed string is
' less than using a UserForm. In addition, it is easier and quicker to update the string displayed
' in the Status Bar rather than updating and repainting controls within a UserForm.
'
' This module should be included as found below within the VBA Project in which it is intended to
' be used. The Public sub routine declarations allow for the code to be called from other processes
' within the same Project as needed -> e.g. RXL_ProgressBar.InitialiseProgressBar()
' Alternatively, the functions below can be copied into a different module where they are needed
' and called directly. This will require the Global module-level variables and constants to be
' included within the module in question and we would then recommend updating the sub routine
' declarations to Private.
' Where possible any hardcoded values which a user might want to change are included within this
' module as constants and should be changed here if needed although we would advise against it.
'
' Sub-Routines:
' -------------
'   InitialiseProgressBar()
'       Purpose:    sets the required persistent global variables for the progress bar string
'                   creation, and displays the initial progress bar with a 0% value and optional
'                   message string
'       Parameters: pStrMessage [optional] -> string which will be displayed to the right of the
'                                             progress bar
'   UpdateProgressBar()
'       Purpose:    creates the updated progress bar string +/- an optional message, and displays
'                   this within the Excel StatusBar
'       Parameters: pDblProgress           -> the progress value to drive the progress bar;
'                                             expected to be between 0.00 and 1.00,
'                                             i.e. decimalised % (0.10 = 10%)
'                   pStrMessage [optional] -> string which will be displayed to the right of the
'                                             progress bar
'   CloseProgressBar()
'       Purpose:    returns the Status Bar status to the option used before InitialiseProgressBar()
'                   and releases the Status Bar back to Excel
'       Parameters: (none)
'
'
' Code commenting has been used where possible to aid any reader/user in understanding the below
' code and its intended function. Whilst care has been taken in the development of these tools
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


'#############################
'#  Module global variables  #
'#############################
' Constants
Private Const ProgBar_Len As Integer = 20   ' Sets the total size of the progress bar in "blocks"
Private Const ProgBar_Max As Integer = 255  ' The max Excel StatusBar string in characters
Private Const Char_Empty As Long = 9744     ' Unicode character Hex: &H2610  => CLng("&H2610") = 9744
Private Const Char_Full As Long = 9632      ' Unicode character Hex: &H25A0  => CLng("&H25A0") = 9632

' Variables used for the Progress Bar
Dim ProgBar_Empty As String * 1, ProgBar_Full As String * 1
Dim ProgBar_Set As Long
Dim blUserStBar As Boolean



'#############################
'#  Main Functionality       #
'#############################

Public Sub InitialiseProgressBar(Optional ByRef pStrMessage As String = "")
    
    ' Initialise the persistent global variables and set as required
    '  - where the VBA session has initialised or reset the global variables will be empty
    '    and re-defined, the variables will have been re-defined and initialised at 0 (zero)
    '  - by placing the string definitions in the global variables and only setting once at
    '    initialisation reduces the number nof ChrW() function calls
    ProgBar_Empty = ChrW$(Char_Empty)
    ProgBar_Full = ChrW$(Char_Full)
    
    ' Check the length of the max set portion of the progress bar string
    '  - uses the number of blocks to display plus the length of the max size of extra string
    '    which is made up of the progress % in brackets plus space to add the optional message
    ProgBar_Set = ProgBar_Len + Len(" (100%):  ")
    
    ' Capture the currnet user status for displaying the StatusBar
    blUserStBar = Application.DisplayStatusBar
    
    ' Ensure that the StatusBar is being displayed
    If Application.DisplayStatusBar = False Then _
        Application.DisplayStatusBar = True
    
    ' Set the initial output message and empty progress bar
    Call UpdateProgressBar(0, pStrMessage)
    
End Sub

Public Sub UpdateProgressBar(ByVal pDblProgress As Double, _
                             Optional ByRef pStrMessage As String = "")
    
    ' Define local variables
    Dim strOut As String
    Dim lnBlocks As Long, lnPercent As Long
    
    ' Restrict the progress value in case progress value passed is outside of the expected range
    If pDblProgress > 1 Then _
        pDblProgress = 1
    If pDblProgress < 0 Then _
        pDblProgress = 0
    ' Convert the passed progress value (% as Double/decimal) to a whole number
    lnPercent = CLng(pDblProgress * 100)
        
    ' Calculate the number of filled progress bar blocks for the passed progress value (%)
    lnBlocks = CLng(pDblProgress * ProgBar_Len)
    ' Restrict the number of bar blocks in case it is calculated as greater than the set total progress bar length
    If lnBlocks > ProgBar_Len Then _
        lnBlocks = ProgBar_Len
    
    ' Note - max length for a StatusBar message is 255 characters
    '      - the progress bar aspect of the message will be between 28 and 30 characters
    '        [=> 20 boxes + "_(" + 1-3 digits for percentage complete + "%):__"]
    '      - need to check the length of the passed message string [:= pStrMessage] and restrict
    '         -> i.e. If greater than (255 - 30) then restrict to allowable length - 3, and add "..." at
    '                 the end to indicate a truncated message string
    If Len(pStrMessage) > (ProgBar_Max - ProgBar_Set) Then
        pStrMessage = Left$(pStrMessage, ProgBar_Max - ProgBar_Set - 3) & "..."
    End If
    
    ' Construct the output string
    strOut = String(lnBlocks, ProgBar_Full) & _
             String(ProgBar_Len - lnBlocks, ProgBar_Empty) & _
             " (" & lnPercent & "%)"
    If pStrMessage <> vbNullString Then _
        strOut = strOut & ":  " & pStrMessage
        
    ' Display the constructed output string in the Excel StatusBar
    Application.StatusBar = strOut
    
End Sub

Public Sub CloseProgressBar()
    
    ' Reset the StatusBar display to the setting used before initialising and displaying the progress bar
    ' - checks if one of the calculated global variables is 0 (zero) which would indicate that the global
    '   variables have cleared and reset and so the blUserStBar variables is likely a false negative
    ' - where the calculated global variables are reset, default the DisplayStatusBar setting to True
    Application.DisplayStatusBar = IIf(ProgBar_Set = 0, True, blUserStBar)
    
    ' Release the StatusBar back to the application otherwise the final progress bar message will persist
    Application.StatusBar = False

End Sub

