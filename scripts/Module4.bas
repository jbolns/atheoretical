Attribute VB_Name = "Module4"
Option Explicit
Option Base 1
' Module contains re-training and final testing functions.
'




' --------------------------------- NEW SUB STARTS
' ----------------
'

Sub testerM5()
' This macro is only for testing

Selection.Offset(1, 2).Interior.Color = 65535

End Sub
' --------------------------------- TESTING SUB ENDS






' --------------------------------- NEW SUB STARTS
' ----------------
'

Sub reTrain()
' This macro takes the chosen model and performs a final test against observations in the test reserve

'Define variables & other initial matters
Dim nvars As Integer, labels() As String, values() As Double, j As Integer, algoType As String
Dim bestModel As Integer, bestAlgo As String, reTrain As Boolean
    
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

'Pick up parameters from dashboard
    Sheets("Dashboard").Select
        If IsEmpty(Range("C23")) Then
            MsgBox ("Sorry, you need to specify the algorithm to use for retraining (cell C23).")
            Exit Sub
        End If
        bestAlgo = Range("C23")

'Run selected algorithm on ReTrain sheet
    Sheets("ReTrain").Select
    Application.Run (bestAlgo)
    
'Move results into RESULTS sheet
    Sheets("ReTrain").Select
        ActiveSheet.UsedRange.Find("Model", LookAt:=xlWhole).Select
            algoType = Range("A2")
            nvars = Range(Selection, Selection.End(xlToRight)).Columns.Count - 1
            ReDim values(nvars + 2) As Double
            ReDim labels(nvars + 2) As String
            For j = 1 To nvars
                labels(j) = ActiveCell.Offset(-1, j)
                values(j) = ActiveCell.Offset(0, j)
            Next j
            labels(nvars + 1) = Range("A7")
            labels(nvars + 2) = Range("A8")
            values(nvars + 1) = Range("B7")
            values(nvars + 2) = Range("B8")
    Sheets("RESULTS").Select
        Range("A" & (ActiveSheet.UsedRange.Rows.Count + 2)).Select
            Selection.Value = "Final Model"
            ActiveCell.Offset(0, 1) = algoType
            For j = nvars + 1 To nvars + 2
                ActiveCell.Offset(1, j - nvars - 1) = labels(j)
                ActiveCell.Offset(2, j - nvars - 1) = values(j)
            Next j
            For j = 1 To nvars
                ActiveCell.Offset(1, j + 2) = labels(j)
                ActiveCell.Offset(2, j + 2) = values(j)
            Next j
        Selection.Font.Size = 16
        Selection.Font.Bold = True
        Range(Selection, ActiveCell.Offset(0, nvars + 2)).Select
        Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Selection.Borders(xlEdgeBottom).Weight = xlMedium
        Range(Selection, ActiveCell.Offset(1, nvars + 2)).Select
        Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Range(Selection, ActiveCell.Offset(2, nvars + 2)).Select
        Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Selection.Borders(xlEdgeBottom).Weight = xlMedium

'Final formatting
    Cells.Select
        Cells.EntireColumn.AutoFit
    Range("A1").Select
    Sheets("Dashboard").Select

'Final message
    Application.ScreenUpdating = True
    MsgBox ("Done! The results of re-training and final testing have been appended to the RESULTS sheet." & vbCrLf & vbCrLf & _
        "Additionally, ReTrain and Test sheets were updated.")
    Application.ScreenUpdating = False

'Error handling & gracious exit
Exit Sub
ErrorHandler: MsgBox ("Sorry, something went wrong. Check your data try again.")

End Sub
' --------------------------------- RETRAIN SUB ENDS



Sub noRetrain()

'Define variables & other initial matters
    Dim i As Integer, nvars As Integer, bestModel As Integer, label As Variant, keepCol As Boolean
                
        On Error GoTo ErrorHandler
        Application.ScreenUpdating = False

'Pick up parameters from dashboard
    Sheets("Dashboard").Select
        If IsEmpty(Range("B23")) Then
            MsgBox ("Sorry, you need to specify the model you want to test (cell B23).")
            Exit Sub
        End If
        bestModel = Range("B23")

'Get the equation for the chosen model
    Sheets("RESULTS").Select
        ActiveSheet.UsedRange.Find("Model " & bestModel, LookAt:=xlWhole).Select
            Selection.Offset(1, 3).Select
                nvars = Range(Selection, Selection.End(xlToRight)).Columns.Count
                ReDim values(nvars) As Double
                ReDim labels(nvars) As String
                For i = 1 To nvars
                    labels(i) = ActiveCell.Offset(0, i - 1)
                    values(i) = ActiveCell.Offset(1, i - 1)
                Next i

'Check if the number of variables in source and destination sheet is the same
    Sheets("Test").Select
        If nvars <> ActiveSheet.UsedRange.Columns.Count - 1 Then
                For i = ActiveSheet.UsedRange.Columns.Count - 1 To 2 Step -1
                    Columns(i).Select
                    For Each label In labels()
                        If Cells(1, i) = label Then keepCol = True
                    Next
                    If keepCol = False Then Columns(i).Delete
                    keepCol = False
                Next i
        End If
        
' Continue in test sheet and calculate predicted outcomes, place next to actual outcomes
    nRows = ActiveSheet.UsedRange.Rows.Count
    nvars = nvars + 1
    Cells(1, nvars + 1) = "Predicted"
    For i = 1 To nRows - 1
        Predicted = values(1) 'Intercept
        For j = 1 To nvars - 2 'Loop through variables, add weight*value to prediction.
            Predicted = Predicted + values(j + 1) * Cells(1 + i, j + 1)
        Next j
        Cells(1 + i, nvars + 1) = Predicted
    Next i

' Continue in test sheet and calculate TSS, place results next to predicted outcomes
    Cells(1, nvars + 2) = "TSSi"
    For i = 1 To nRows - 1
        Cells(1 + i, nvars + 2).Select
        Cells(1 + i, nvars).Select
        Range(Cells(2, nvars), Cells(nRows, nvars)).Select
        Cells(1 + i, nvars + 2) = (Cells(1 + i, nvars) - Application.WorksheetFunction.Average(Range(Cells(2, nvars), Cells(nRows, nvars)))) ^ 2
    Next i
    
' Continue in test sheet and calculate RSS, place results next to TSS
    Cells(1, nvars + 3) = "RSSi"
    For i = 1 To nRows - 1
        Cells(1 + i, nvars + 3) = (Cells(1 + i, nvars) - Cells(1 + i, nvars + 1)) ^ 2
    Next i

' Continue in test sheet and calculate R-squared, placed some columns to the right of data
    Cells(1, nvars + 5) = "R-squared"
    RSS = Application.WorksheetFunction.Sum(Range(Cells(2, nvars + 3), Cells(nRows, nvars + 3)))
    TSS = Application.WorksheetFunction.Sum(Range(Cells(2, nvars + 2), Cells(nRows, nvars + 2)))
    rSquared = 1 - (RSS / TSS)
    Cells(2, nvars + 5) = rSquared
    Range("A1").Select
      
'Back to RESULTS and incorporate final test results
    Sheets("RESULTS").Select
        ActiveSheet.UsedRange.Find("Model " & bestModel, LookAt:=xlWhole).Select
            Selection.Offset(1, 2) = "Final Test R2"
            Selection.Offset(2, 2) = rSquared
            Selection.Offset(1, 2).Interior.Color = 65535
            Selection.Offset(2, 2).Interior.Color = 65535

'Final formatting
    Cells.Select
        Cells.EntireColumn.AutoFit
    Range("A1").Select
    Application.DisplayAlerts = False
    Sheets("ReTrain").Delete
    Application.DisplayAlerts = True
     
'Final message
    Sheets("Dashboard").Select
    Application.ScreenUpdating = True
    MsgBox ("Done! The results of the final test have been added to the RESULTS sheet." & vbCrLf & vbCrLf & _
        "Additionally, Test sheet was updated.")
     
     
'Error handling & gracious exit
Exit Sub
ErrorHandler: MsgBox ("Sorry, something went wrong. Check your data try again.")

End Sub
' --------------------------------- NO RETRAIN MODEL TEST SUB ENDS
