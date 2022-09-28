Attribute VB_Name = "Module6"
Option Explicit
Option Base 1
' Module contains subs/functions for cleaning and formatting.
'
'




' --------------------------------- NEW SUB STARTS
' ----------------
'
Sub clean()
' This macro cleans the workbook ahead of an analysis

'Dims & other initial matters
    Dim sht As Worksheet
    
        On Error GoTo ErrorHandler
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False

'Delete extra worksheets
    For Each sht In Worksheets 'Loops through all worksheets and deletes those not in the original set of sheets
        If sht.Name = "Welcome" Or sht.Name = "Guidance" Or sht.Name = "Dashboard" Then GoTo NextSheet
        sht.Delete
NextSheet:
    Next

' Reset hyperparameters
    'ActiveSheet.Range("C9:C12").ClearContents
    'ActiveSheet.Range("B23:C23").ClearContents

' Final formatting/matters
    Application.ScreenUpdating = True
    Range("A1").Activate
        Application.GoTo ActiveCell, True

' Error handling & gracious exit
Exit Sub
ErrorHandler: MsgBox ("Sorry, something went wrong while cleaning the workbooks. The macro that failed is quite simple. It should have worked just fine. Better contact the administrator.")

End Sub
' --------------------------------- CLEANING SUB ENDS




' --------------------------------- NEW SUB STARTS
' ----------------
'

Sub selectFirstCell()
' This macro selects cell A1 on all sheets

' Dims & other initial matters
    Dim sht As Worksheet
    
        On Error GoTo ErrorHandler
        Application.ScreenUpdating = False

' Loop through sheets and select A1
    For Each sht In Worksheets
        sht.Activate
        Range("A1").Activate
        Application.GoTo ActiveCell, True
    Next

' Error handling & gracious exit
Exit Sub
ErrorHandler: MsgBox ("Sorry, something went wrong with a miscellaneous function. The macro that failed is quite simple. It should have worked just fine. Better contact the administrator.")

End Sub
' --------------------------------- SELECT FIRST CELL SUB ENDS

