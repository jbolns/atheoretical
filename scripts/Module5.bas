Attribute VB_Name = "Module5"
Option Explicit
Option Base 1
' Module contains support subs/functions (incl. cleaning and formatting).
'
'





' --------------------------------- NEW SUB STARTS
' ----------------
'

Sub writeEquation()
' This macro writes down the equation into the summary of results from Excel's Data Analysis Toolpack

'Define variables & other initial matters
    Dim i As Integer, nFeatures As Integer, coefficients() As String, equation() As Double, mSpec() As String, spec As Variant, fullThing As String
        On Error GoTo ErrorHandler

' Write down equation into a single row
ActiveSheet.UsedRange.Find("Coefficients", LookAt:=xlWhole).Select
    Range(Selection, Selection.End(xlDown)).Select
        nFeatures = Selection.Rows.Count - 1 'Get number of coefficients in regression results (including intercept)
        ReDim equation(nFeatures) As Double
        ReDim coefficients(nFeatures) As String
        ReDim mSpec(nFeatures, 4) As String
    
    For i = 1 To nFeatures
        equation(i) = ActiveCell.Offset(i, 0)
    Next i
    
    For i = 1 To nFeatures
        coefficients(i) = ActiveCell.Offset(i, -1)
    Next i
    
    For i = 1 To nFeatures
            mSpec(i, 1) = Round(equation(i), 2)
            mSpec(i, 2) = "*"
            mSpec(i, 3) = coefficients(i)
            mSpec(i, 4) = "+"
    Next i
    mSpec(1, 2) = ""
    mSpec(1, 3) = ""
   
' Format a bunch of small things so the result looks nice
    ActiveCell.Offset(nFeatures + 4, -1).Select
        Selection.Resize(2, 3).Select
        Selection.Insert Shift:=xlDown
        ActiveCell.FormulaR1C1 = "Model"
    ActiveCell.Offset(-1, 0).Select
        ActiveCell.Offset(0, 1) = coefficients(1)
        ActiveCell.Offset(1, 1) = equation(1)
        For i = 1 To nFeatures - 1
            ActiveCell.Offset(0, i + 1) = coefficients(i + 1)
            ActiveCell.Offset(1, i + 1) = equation(i + 1)
        Next i
        ActiveCell.Offset(-1, 0) = "EQUATION"
    Range(Selection, ActiveCell.Offset(0, nFeatures)).Select
        Selection.HorizontalAlignment = xlCenter
        Selection.Font.Italic = True
    Range(Selection, ActiveCell.Offset(1, 0)).Select
        Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
        Selection.Borders(xlEdgeTop).Weight = xlMedium
        Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Selection.Borders(xlEdgeBottom).Weight = xlMedium
        Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        
' Write down model specification to a single cells (will be needed to show model in dashboard)
    ActiveCell.Offset(0, nFeatures + 2).Select
        Selection.Value = "Specification"
        Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
        Selection.Borders(xlEdgeTop).Weight = xlMedium
        Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Selection.Borders(xlEdgeBottom).Weight = xlThin
    ActiveCell.Offset(1, 0).Select
        Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Selection.Borders(xlEdgeBottom).Weight = xlMedium
        Selection.ShrinkToFit = True
        For i = 1 To nFeatures
            fullThing = fullThing & " " & mSpec(i, 1) & mSpec(i, 2) & mSpec(i, 3) & " " & " " & mSpec(i, 4) & " "
        Next i
        fullThing = Replace(fullThing, "   ", " ")
        fullThing = Replace(fullThing, "  ", " ")
        fullThing = Left(fullThing, Len(fullThing) - 2)
        fullThing = "Y = " & Trim(fullThing)
        Selection = fullThing
   
' More formats
ActiveSheet.UsedRange.Find("Observation", LookAt:=xlWhole).Select
    ActiveCell.Offset(-1, 0).Range("A1:C1").Select
      Selection.Delete Shift:=xlUp
    
'Error handling & gracious exit
Exit Sub
ErrorHandler: MsgBox ("Sorry, something went wrong. Check your data try again.")

End Sub
' --------------------------------- WRITE EQUATION SUB ENDS





' --------------------------------- NEW FUNCTION STARTS
' ----------------
'

Function validateFx(nvars As Integer, valSheetName As String)
' This function validates a trained model against a validation (or test) set and returns the validation's R-squared

'Define variables & other initial matters
Dim nRows As Integer, i As Integer, j As Integer, keepCol As Boolean
Dim equation() As Double, featureLabels() As String, label As Variant
Dim Predicted As Double, RSS As Double, TSS As Double, rSquared As Double
Dim sourceSheet As Worksheet, destinationSheet As Worksheet
ReDim equation(nvars - 1) As Double
ReDim featureLabels(nvars - 1) As String
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    
    Set sourceSheet = ActiveSheet
    
    Sheets(valSheetName).Select
    Set destinationSheet = ActiveSheet

' Get equation and featurelabels from source sheet
sourceSheet.Select
    sourceSheet.UsedRange.Find("Model", LookAt:=xlWhole).Select
        For i = 1 To nvars - 1
            equation(i) = ActiveCell.Offset(0, i)
        Next i
    sourceSheet.UsedRange.Find("Coefficients", LookAt:=xlPart).Select
        For i = 1 To nvars - 1
            featureLabels(i) = ActiveCell.Offset(i, -1)
        Next i

' Check if the number of variables in source and destination sheet is the same
destinationSheet.Select
    If nvars <> destinationSheet.UsedRange.Columns.Count Then
            For i = destinationSheet.UsedRange.Columns.Count - 1 To 2 Step -1
                Columns(i).Select
                For Each label In featureLabels()
                    If Cells(1, i) = label Then keepCol = True
                Next
                If keepCol = False Then Columns(i).Delete
                keepCol = False
            Next i
    End If

' Continue in validate sheet and calculate predicted outcomes, place next to actual outcomes
    nRows = destinationSheet.UsedRange.Rows.Count
    Cells(1, nvars + 1) = "Predicted"
    For i = 1 To nRows - 1
        Predicted = equation(1) 'Intercept
        For j = 1 To nvars - 2 'Loop through variables, add weight*value to prediction.
            Predicted = Predicted + equation(j + 1) * Cells(1 + i, j + 1)
        Next j
        Cells(1 + i, nvars + 1) = Predicted
    Next i

' Continue in validation sheet and calculate TSS, place results next to predicted outcomes
    Cells(1, nvars + 2) = "TSSi"
    For i = 1 To nRows - 1
        Cells(1 + i, nvars + 2) = (Cells(1 + i, nvars) - Application.WorksheetFunction.Average(Range(Cells(2, nvars), Cells(nRows, nvars)))) ^ 2
    Next i
    
' Continue in validation sheet and calculate RSS, place results next to TSS
    Cells(1, nvars + 3) = "RSSi"
    For i = 1 To nRows - 1
        Cells(1 + i, nvars + 3) = (Cells(1 + i, nvars) - Cells(1 + i, nvars + 1)) ^ 2
    Next i

' Continue in validation sheet and calculate R-squared, placed some columns to the right of data
    Cells(1, nvars + 5) = "R-squared"
    RSS = Application.WorksheetFunction.Sum(Range(Cells(2, nvars + 3), Cells(nRows, nvars + 3)))
    TSS = Application.WorksheetFunction.Sum(Range(Cells(2, nvars + 2), Cells(nRows, nvars + 2)))
    rSquared = 1 - (RSS / TSS)
    Cells(2, nvars + 5) = rSquared

' Send R-square as main function result
validateFx = rSquared

'Error handling & gracious exit
Exit Function
ErrorHandler: MsgBox ("Sorry, something went wrong. Check your data try again.")
End Function
' --------------------------------- VALIDATE FX FUNCTION ENDS






' --------------------------------- NEW FUNCTION STARTS
' ----------------
'

Function headersFx(nvars As Integer)
' This function reads headers on any data sheet with headers in row(1) and writes them out to an array

'Define variables & other initial matters
    Dim i As Integer, header() As String
        On Error GoTo ErrorHandler
        ReDim header(nvars) As String

' Get headers from current sheet
    For i = 1 To nvars
        header(i) = ActiveSheet.Cells(1, i)
    Next i

'Send header array as main function results
    headersFx = header()

'Error handling & gracious exit
Exit Function
ErrorHandler: MsgBox ("Sorry, something went wrong. Check your data try again.")

End Function
' --------------------------------- HEADERS FX FUNCTION ENDS





' --------------------------------- NEW SUB STARTS
' ----------------
'

Sub summaryResults()
' This macro brings the results of all training/validation folds into a single summary results sheet

'Define variables & other initial matters
    Dim nModels As Integer, nvars As Integer, i As Integer, j As Integer, algoType As String, values() As Double, labels() As String
        
        Application.ScreenUpdating = False
        On Error GoTo ErrorHandler

' Get number of available models from dashboard
    Sheets("Dashboard").Select
        nModels = Range("C11") 'The number of available models is the same as the number of desired folds

' Change to a new results sheet and write equations down
    Sheets.Add After:=Sheets("Dashboard") 'Create new results sheet
        ActiveSheet.Name = "RESULTS"
    For i = 1 To nModels
        Sheets("Train" & i).Select
        algoType = Range("A2")
        ActiveSheet.UsedRange.Find("Model", LookAt:=xlWhole).Select
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
        If i = 1 Then Range("A1").Select
        If i <> 1 Then Range("A" & (ActiveSheet.UsedRange.Rows.Count + 2)).Select
            Selection.Value = "Model " & i
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
    Next i

'Final formatting
    Cells.Select
        Cells.EntireColumn.AutoFit
    Range("A1").Select
    Sheets("Dashboard").Select

'Error handling & gracious exit
Exit Sub
ErrorHandler: MsgBox ("Sorry, something went wrong while importing results from across folds into a single summary sheet. Check your data try again.")

End Sub
' --------------------------------- SUMMARY RESULTS SUB ENDS




' --------------------------------- NEW SUB STARTS
' ----------------
'

Sub checkWinner()
' This function automatically selects the best performing model/algorithm type

'Define variables & other initial matters
    Dim nModels As Integer, i As Integer, bestModel As Integer, bestAlgo As String, c1 As Double, c2 As Double, reTrain As Boolean
        
        Application.ScreenUpdating = False
        On Error GoTo ErrorHandler
        
'Get retraining strategy from dashboard
    Sheets("Dashboard").Select
        If Range("C17") = "No ReTrain (Best Model)" Then
            reTrain = False
        ElseIf Range("C17") = "ReTrain (Best Algorithm)" Then
            reTrain = True
        Else
            MsgBox ("Sorry, we cannot perform the last step in the process because you must select a model selection strategy. Check hyper-parameters in the dashboard and try again")
            Exit Sub
        End If
    

'Get number of available models from dashboard
    nModels = Range("C15") 'The number of available models is the same as the number of desired folds

'Compare validation R2 across models
    Sheets("RESULTS").Select
        For i = 1 To nModels - 1
            c1 = ActiveSheet.UsedRange.Find("Model " & i, LookAt:=xlWhole).Offset(2, 1).Value
            c2 = ActiveSheet.UsedRange.Find("Model " & i + 1, LookAt:=xlWhole).Offset(2, 1).Value
            If c1 > c2 Then
                bestModel = i
            Else
                bestModel = i + 1
            End If
        Next i
        bestAlgo = ActiveSheet.UsedRange.Find("Model " & bestModel, LookAt:=xlWhole).Offset(0, 1)
       

'Error handling & gracious exit
Exit Sub
ErrorHandler: MsgBox ("Sorry, something went wrong while comparing the performance of the different algorithms/models. Check your data try again.")

End Sub
' --------------------------------- CHECK WINNER SUB ENDS

