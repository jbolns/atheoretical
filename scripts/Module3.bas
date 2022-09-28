Attribute VB_Name = "Module3"
Option Explicit
Option Base 1
' Module contains algorithms' subs/functions.
'




' --------------------------------- NEW SUB STARTS
' ----------------
'

Sub testerM3()
' This macro is only for testing


End Sub
' --------------------------------- TESTING SUB ENDS





' --------------------------------- NEW SUB STARTS
' ----------------
'

Sub REGR()
' This macro:
    ' (1) runs a simple regression analysis of each IV/DV combination,
    ' (2) chooses IV/DV combo with highest R-squared,
    ' (3) takes the specified model and predicts values in the validation set.
    ' (4) incorporates validation R-squared into summary results.

'Define variables & other initial matters
    Dim nobservations As Integer, nvars As Integer, c As Integer, featureLabels As String, maxR2 As Double, winner As String, validationR2 As Double
    Dim trainSheet As Worksheet, valSheet As Worksheet, valSheetName As String
    
        On Error GoTo ErrorHandler
        Application.ScreenUpdating = False
        Set trainSheet = ActiveSheet

'Get key parameters from dataset
    nvars = trainSheet.UsedRange.Columns.Count 'Number of variables in dataset (incl. features & target, plus index)
    nobservations = trainSheet.UsedRange.Rows.Count 'Number of observations in dataset (plus headers)

'Run multiple regression of each features (and format results).
    For c = 1 To nvars - 2
        featureLabels = Cells(1, nvars - c) 'Picks name of current feature
        Application.Run "ATPVBAEN.XLAM!Regress" _
                , ActiveSheet.Range(Cells(1, nvars), Cells(nobservations, nvars)), ActiveSheet.Range(Cells(1, nvars - c), Cells(nobservations, nvars - c)) _
                , False, True, 95, ActiveSheet.Cells(1, nvars + 1 + (10 * (c - 1))), True, False, False, False, , False
        Application.ScreenUpdating = False
        Cells.Select
        Cells.EntireColumn.AutoFit
        Call writeEquation 'Calls sub to write equation into regression's results
        trainSheet.UsedRange.Find("Coefficients", LookAt:=xlWhole).Select
            Selection.Value = "Coefficients " & featureLabels
        trainSheet.UsedRange.Find("Observation", LookAt:=xlWhole).Select
            Range(Cells(1, nvars), Cells(nobservations, nvars)).Copy
            Selection.Insert Shift:=xlToRight
            Range(Cells(1, nvars - c), Cells(nobservations, nvars - c)).Copy
            Selection.Insert Shift:=xlToRight
            Range(Cells(1, 1), Cells(nobservations, 1)).Copy
            Selection.Insert Shift:=xlToRight
        trainSheet.UsedRange.Find("Observation", LookAt:=xlWhole).Select
            Selection.Resize(nobservations, 1).Select
            Selection.Delete Shift:=xlToLeft
        ActiveCell.Offset(0, -3).Range("A1:C1").Select
            Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
            Selection.Borders(xlEdgeTop).Weight = xlMedium
            Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Selection.Resize(nobservations, 3).Select
            Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
            Selection.Borders(xlEdgeBottom).Weight = xlMedium
        ActiveSheet.UsedRange.Find("RESIDUAL OUTPUT", LookAt:=xlWhole).Select
            Selection.Value = "DATA & OUTPUTS"
        ActiveSheet.UsedRange.Find("SUMMARY OUTPUT", LookAt:=xlWhole).Select
            Selection.Clear
        ActiveCell.Offset(1, 0).Range("A1").Select
            Selection.Value = "Simple regression, " & featureLabels
            Selection.Font.Size = 16
            Selection.Font.Bold = True
            trainSheet.Cells(2, 1).Select
    Next c
    
    For c = 1 To nvars 'Since data is now incorporated in results, delete original columns
        Columns(1).Delete
    Next c

'Choose regression with highest R-squared
    maxR2 = Cells(5, 2) 'Pick first R-squared as first value
    For c = 1 To nvars - 3 'Compare R-squareds for all regressions and choose highest
        If maxR2 < Cells(5, 2 + ((c) * 10)) Then maxR2 = Cells(5, 2 + ((c) * 10))
    Next c

'Format results to show only the winner regression
    Columns(1).Select
        Selection.Insert Shift:=xlLeft
    trainSheet.Rows(5).Find(maxR2, LookAt:=xlWhole).Select
        ActiveCell.Offset(-4, -2).Range("A1").Select
            winner = Split(ActiveCell.Address, "$")(1)
            Columns("A:" & winner).Delete
        nvars = ActiveSheet.UsedRange.Columns.Count + 2
    trainSheet.Cells(1, nvars).Select
        winner = Split(ActiveCell.Address, "$")(1)
        Columns("J:" & winner).Delete
    featureLabels = Cells(18, 1)

'Open space for validation results
    Range(Cells(7, 1), Cells(7, ActiveSheet.UsedRange.Columns.Count)).Select
        Selection.Insert Shift:=xlDown
  
'Transfer results to validation fold & perform validation (using R-squared)
    nvars = 3
    valSheetName = trainSheet.Name
    If trainSheet.Name = "ReTrain" Then
        valSheetName = "Test"
    Else
        valSheetName = Split(valSheetName, "n")(1)
        valSheetName = "Validate" & valSheetName
    End If
    validationR2 = validateFx(nvars, valSheetName)
        
'Transfer validation R-square to summary results
    trainSheet.Select
        If valSheetName = "Test" Then
            Cells(7, 2) = validationR2
            Cells(6, 1) = "ReTraining R2"
            Cells(7, 1) = "Final Test R2"
        Else
            Cells(7, 2) = validationR2
            Cells(6, 1) = "Training R2"
            Cells(7, 1) = "Validation R2"
        End If

'Final formats
    Rows(1).Select
        Selection.Insert Shift:=xlDown
    Cells(1, 1).Select
        Selection.Value = "SUMMARY OUTPUT"
        Selection.Font.Size = 24
        Selection.Font.Bold = True
    Cells(2, 1).Select
        Selection.Value = "REGR"
        Selection.Font.Size = 6
        Selection.Font.Color = RGB(255, 255, 255)
    Range("A1").Activate

'Error handling & gracious exit
Exit Sub
ErrorHandler: MsgBox ("Sorry, something went wrong in the simple regression step. Check your data try again.")

End Sub
' --------------------------------- SIMPLE REGRESSION SUB ENDS




' --------------------------------- NEW SUB STARTS
' ----------------
'

Sub MultiREGR()
' This macro:
    ' (1) runs a multiple regression analysis including all features,
    ' (2) takes the specified model and predicts values in the corresponding validation set,
    ' (3) incorporates validation R-squared into summary results.

'Define variables & other initial matters
    Dim nvars As Integer, nobservations As Integer, c As Integer, valSheetName As String, validationR2 As Double, trainSheet As Worksheet, featureLabels() As String
    
        Application.ScreenUpdating = False
        On Error GoTo ErrorHandler
        
        Set trainSheet = ActiveSheet

'Get key parameters from dataset
    nvars = trainSheet.UsedRange.Columns.Count 'Number of variables in dataset (incl. features & target, plus index)
    nobservations = trainSheet.UsedRange.Rows.Count 'Number of observations in dataset (plus headers)

'Run multiple regression including all features.
    Application.Run "ATPVBAEN.XLAM!Regress" _
            , ActiveSheet.Range(Cells(1, nvars), Cells(nobservations, nvars)), ActiveSheet.Range(Cells(1, 2), Cells(nobservations, nvars - 1)) _
            , False, True, 95, ActiveSheet.Cells(1, nvars + 1), True, False, False, False, , False
    Application.ScreenUpdating = False
    Cells.Select
        Cells.EntireColumn.AutoFit
    Call writeEquation

'Merge original data with regression's results
    Range(Cells(1, 1), Cells(nobservations, nvars)).Select
        Selection.Copy
    trainSheet.UsedRange.Find("Observation", LookAt:=xlWhole).Select
        Selection.Insert Shift:=xlToRight
    trainSheet.UsedRange.Find("Observation", LookAt:=xlWhole).Select
        Selection.Resize(nobservations, 1).Select
        Selection.Delete Shift:=xlToLeft

'Formatting
    trainSheet.UsedRange.Find("RESIDUAL OUTPUT", LookAt:=xlWhole).Select
        Selection.Value = "DATA & OUTPUTS"
    ActiveCell.Offset(1, 0).Range(Cells(1, 1), Cells(1, nvars)).Select
        Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
        Selection.Borders(xlEdgeTop).Weight = xlMedium
        Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Resize(nobservations, nvars).Select
        Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Selection.Borders(xlEdgeBottom).Weight = xlMedium
        For c = 1 To nvars 'Since data is now incorporated in results, delete original columns
            Columns(1).Delete
        Next c

'Open space for results from validation
    Range(Cells(7, 1), Cells(7, trainSheet.UsedRange.Columns.Count)).Select
    Selection.Insert Shift:=xlDown

'Transfer results to validation fold & perform validation (using R-squared)
    valSheetName = trainSheet.Name
    If trainSheet.Name = "ReTrain" Then
        valSheetName = "Test"
    Else
        valSheetName = Split(valSheetName, "n")(1)
        valSheetName = "Validate" & valSheetName
    End If
    validationR2 = validateFx(nvars, valSheetName)

'Transfer validation R-square to summary results
    trainSheet.Select
        If valSheetName = "Test" Then
            Cells(7, 2) = validationR2
            Cells(6, 1) = "ReTraining R2"
            Cells(7, 1) = "Final Test R2"
        Else
            Cells(7, 2) = validationR2
            Cells(6, 1) = "Training R2"
            Cells(7, 1) = "Validation R2"
        End If

'Final formatting/matters
    trainSheet.Select
    Cells(1, 1).Select
        Selection.Font.Size = 24
        Selection.Font.Bold = True
    Rows(3).Select
        Selection.Insert Shift:=xlDown
    Cells(2, 1).Select
        Selection.Value = "MultiREGR"
        Selection.Font.Size = 6
        Selection.Font.Color = RGB(255, 255, 255)
    Cells(3, 1).Select
        Selection.Value = "Multiple regression"
        Selection.Font.Size = 16
        Selection.Font.Bold = True
    Range("A1").Activate

'Error handling & gracious exit
Exit Sub
ErrorHandler: MsgBox ("Something went wrong in the multiple regression step. Most likely, data is not appropriately formatted." & vbCrLf & vbCrLf & "Please check and re-import if needed.")

End Sub
' --------------------------------- MULTIPLE REGRESSION SUB ENDS





' --------------------------------- NEW SUB STARTS
' ----------------
'

Sub adjMultiREGR()
' This macro:
    ' (1) performs a correlation analysis including all IVs in an active training sheet (the macro ignores indexes in column A),
    ' (2) eliminates correlations between IVs (if two IVs correlate, the one with highest correlation to outcome is chosen),
    ' (2) and runs a multiple regression on surviving IVs.

'Define variables & other initial matters
    Dim nvars As Integer, nobservations As Integer
    Dim c As Range, c1 As Integer, c2 As Integer
    Dim trainSheet As Worksheet, valSheet As Worksheet, valSheetName As String, validationR2 As Double
        
        Application.ScreenUpdating = False
        On Error GoTo ErrorHandler
        
' Set corresponding training/validation sheets
    Set trainSheet = ActiveSheet
    
    valSheetName = trainSheet.Name
    If trainSheet.Name = "ReTrain" Then
        valSheetName = "Test"
    Else
        valSheetName = Split(valSheetName, "n")(1)
        valSheetName = "Validate" & valSheetName
    End If
    Sheets(valSheetName).Select
    Set valSheet = ActiveSheet
    
    trainSheet.Select

'Pick up parameters
    nvars = trainSheet.UsedRange.Columns.Count 'Number of variables in dataset (incl. features & target, plus index)
    nobservations = trainSheet.UsedRange.Rows.Count 'Number of observations in dataset (plus headers)
    
'Run a correlation analysis on all features originally in dataset.
    Application.Run "ATPVBAEN.XLAM!Mcorrel", ActiveSheet.Range(Cells(1, 2), Cells(nobservations, nvars)) _
            , ActiveSheet.Cells(1, nvars + 1), "C", True
    Application.ScreenUpdating = False

'Bring results to top of data.
    Range(Cells(1, 1), Cells(nvars, nvars)).Select
        Selection.Insert Shift:=xlDown
        Selection.Delete Shift:=xlToLeft

'Get rid of correlated IVs
    Range(Cells(2, 2), Cells(nvars - 1, nvars - 1)).Select
    For Each c In Selection.Cells
        If c.Value = 1 Then c.ClearContents
    Next
    For Each c In Selection.Cells
        If c.Value > 0.8 Then
            c1 = c.Column
            c2 = c.Row
            If Abs(Cells(nvars, c1).Value) > Abs(Cells(nvars, c2).Value) Then
                trainSheet.Columns(c2).Delete
                valSheet.Columns(c2).Delete
            Else
                trainSheet.Columns(c1).Delete
                valSheet.Columns(c2).Delete
            End If
        End If
    Next

'Get rid of correlation analysis
    Range(Cells(1, 1), Cells(nvars, trainSheet.UsedRange.Columns.Count)).Select
        Selection.Delete Shift:=xlUp
    Range("A1").Select
    Application.ScreenUpdating = True

'Perform multiple regression on remaining variables.
    Call MultiREGR

'Relabel summary output
    trainSheet.Cells(2, 1).Select
        Selection.Value = "adjMultiREGR"
        Selection.Font.Size = 6
        Selection.Font.Color = RGB(255, 255, 255)
    trainSheet.Cells(3, 1).Select
        Selection.Value = "Correlation-adjusted multiple regression"
    Application.ScreenUpdating = False

'Error handling & gracious exit
Exit Sub
ErrorHandler: MsgBox ("Sorry, something went wrong during the collinearity adjusted regression step. Check your data try again.")

End Sub
' --------------------------------- CORRELATION-ADJUSTED MULTIPLE REGRESSION SUB ENDS
