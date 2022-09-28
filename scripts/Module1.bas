Attribute VB_Name = "Module1"
Option Explicit
' Module contains subs/functions to manage the general process flow.
'
'




' --------------------------------- NEW SUB STARTS
' ----------------
'

Sub testerM1()
' This macro is solely for testing.

    
End Sub
' --------------------------------- TESTING SUB ENDS






' --------------------------------- NEW SUB STARTS
' ----------------
'

Sub mainFlow()
' This macro starts the whole process and administers the general workflow.

'Dims & other initial matters
    Dim i As Integer
    Dim dash As Worksheet
    
        On Error GoTo ErrorHandler
        Set dash = ActiveSheet

'Exit sub if there are missing hyper-parameters
    For i = 1 To 4
        If IsEmpty(dash.Range("C" & 8 + i).Value) Then
            MsgBox ("Sorry, you need to specify all hyper-parameters. See guidance worksheet for an overview of what each hyper-parameter does.")
            Exit Sub
        End If
    Next i
    dash.Range("A1").Select

'Import dataset
    Call importCSV
    dash.Select
    
    Application.ScreenUpdating = True
    MsgBox ("Data has been imported. Click OK to continue.")
    Application.ScreenUpdating = False

'Cross-validation
    If dash.Range("C10").Value = "Segmentation" Then
        Call foldsBySegmentation
    ElseIf dash.Range("C10").Value = "Randomisation" Then
        Call foldsByRandomisation
    Else: MsgBox ("Please check that the segmentation strategy hyper-parameter is set correctly and try again.")
    End If
    
    Application.ScreenUpdating = True
    MsgBox ("Training/validation folds were created and will be stored in corresponding sheets. Click OK to continue.")
    Application.ScreenUpdating = False

'Identify algorithm strategy and perform regression for each fold
    Call algoFlow
    
'Summarise results across folds into a single worksheet
    Call summaryResults

'Final message
    Application.ScreenUpdating = True
    MsgBox ("Training is done! Careful, the computer might be sweating!" & vbCrLf & vbCrLf & _
        "A summary of model specifications is available in the RESULTS sheet.")
    
'Error handling & gracious exit
Exit Sub
ErrorHandler: MsgBox ("Something went wrong and we don't actually know where. Sorry.")

End Sub
' --------------------------------- MAIN FLOW SUB ENDS





' --------------------------------- NEW SUB STARTS
' ----------------
'

Sub algoFlow()
' This macro IDs and calls algorithms for each fold

'Define variables & other initial matters
    Dim i As Integer, k As Integer, s As Integer, availableAlgos() As String, kFolds As Integer, algoStrategy As String, algoCall As String
    
        On Error GoTo ErrorHandler

'Pick up parameters from Dashboard
    Sheets("Dashboard").Select
    kFolds = Range("C11")
    algoStrategy = Range("C12")

'Pick up list of available algorithms
    k = Sheets("Guidance").Range("rng_availableAlgos").Count() - 1
    ReDim availableAlgos(k) As String
    For i = 1 To k
        availableAlgos(i) = Sheets("Guidance").Range("rng_availableAlgos").Cells(i, 1)
    Next i

'Manage algorithm calls
    For s = 1 To kFolds
        If algoStrategy = "REGR" Then
            Sheets("Train" & s).Select
            Call REGR
        ElseIf algoStrategy = "MultiREGR" Then
            Sheets("Train" & s).Select
            Call MultiREGR
        ElseIf algoStrategy = "adjMultiREGR" Then
            Sheets("Train" & s).Select
            Call adjMultiREGR
        ElseIf algoStrategy = "MIX" And s <= 3 Then
            algoCall = availableAlgos(s)
            Sheets("Train" & s).Select
            Application.Run (algoCall)
        ElseIf algoStrategy = "MIX" And s > 3 Then
            Sheets("Train" & s).Select
            Randomize
            algoCall = availableAlgos(Int((k * Rnd) + 1))
            Application.Run (algoCall)
        End If
    Next s

'Final formatting/matters
    Call selectFirstCell
    Application.ScreenUpdating = True
    Sheets("Dashboard").Select

'Error handling & gracious exit
Exit Sub
ErrorHandler: MsgBox ("Sorry, something went wrong while selecting algorithms and managing the flow of training. Check your data try again.")

End Sub
' --------------------------------- ALGORITHM FLOW SUB ENDS
