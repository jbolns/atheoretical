Attribute VB_Name = "Module2"
Option Explicit
Option Base 1
' Module contains subs/functions to import data and create cross-validation folds.
'
'
'





' --------------------------------- NEW SUB STARTS
' ----------------
'

Sub testerM2()

Sheets.Add Before:=Sheets("ReTrain")

End Sub
' --------------------------------- TESTING SUB ENDS





' --------------------------------- NEW SUB STARTS
' ----------------
'

Sub importCSV()
' This macro imports data from a .csv file into "DATA" worksheet.

'Dims & other initial matters
    Dim i As Integer, nRows As Integer, nCols, sourceFileName As String
    Dim thisW As Workbook, sourceW As Workbook, dataSht As Worksheet
        
        On Error GoTo ErrorHandler
        Application.ScreenUpdating = False
        Set thisW = ActiveWorkbook

'Prepare destination sheet
    Sheets.Add After:=Sheets(Sheets.Count) 'Create new sheet to import dataset
    ActiveSheet.Name = "DATA"
    Set dataSht = ActiveSheet

'Open import file and copy stuff
    sourceFileName = thisW.Path & "\data.csv"
    Workbooks.OpenText fileName:=sourceFileName, DataType:=xlDelimited, Comma:=True
    Set sourceW = ActiveWorkbook
    Cells.Select
    Selection.Copy

'Import data
    thisW.Activate
    dataSht.Select
    Cells.Select
    ActiveSheet.Paste
    Cells.EntireColumn.AutoFit

'Add index column
    Columns("A:A").Select
    nRows = ActiveSheet.UsedRange.Rows.Count
    Selection.Insert Shift:=xlToRight
    Range("A1").FormulaR1C1 = "Index"
    For i = 1 To nRows - 1 'Adds observation/index number to each entry
        Range("A" & i + 1).FormulaR1C1 = i
    Next i

'Create table from data
    Range("A1").Select
    nCols = ActiveSheet.UsedRange.Columns.Count
    ActiveSheet.ListObjects.Add(xlSrcRange, Range(Selection, Cells(nRows, nCols)), , xlYes).Name = "Data"
    ActiveSheet.ListObjects("Data").TableStyle = "TableStyleMedium1"
    Range("A1").Select

'Final formatting/matters
    sourceW.Close
    Application.ScreenUpdating = True

'Error handling & gracious exit
Exit Sub
ErrorHandler: MsgBox ("Something went wrong. Likely possibilities are:" & vbCrLf & vbCrLf & _
    "- Workbook was not initialised correctly. Go to the dashboard and start with Step 1. " & vbCrLf & _
    "- Source data file is incorrectly named. Rename as 'data.csv'." & vbCrLf & _
    "- Source data file is not in the correct location. Move to same folder as this workbook.") & vbCrLf & _
    "- Source data is not appropriately formatted. Visit the 'guidance' sheet for guidelines." & vbCrLf & vbCrLf & _
    "If none of these solutions is applicable, do contact the administrator."

End Sub
' --------------------------------- DATA IMPORT SUB ENDS




' --------------------------------- NEW SUB STARTS
' ----------------
'

Sub foldsBySegmentation()
' This macro creates folds using a segmentation approach.

'Dims & other initial matters
    Dim kFolds As Integer, testReserve As Double, nobservations As Integer, nvars As Integer, testObservations As Integer, trainObservations As Integer, nValidation As Integer
    Dim i As Integer, j As Integer, k As Integer, lowIndex As Integer, upIndex As Integer
    Dim sourceSht As Worksheet, header() As String
    
        On Error GoTo ErrorHandler
        Application.ScreenUpdating = False
        Sheets("Data").Select
        Set sourceSht = ActiveSheet

'Pick up hyper-parameters from "Dashboard" and "Data" sheet
    Sheets("Dashboard").Select
    testReserve = Range("C9") 'Size of test reserve (%)
    kFolds = Range("C11") 'Number of desired folds

'Calculate additional parameters from "Data" sheet
    sourceSht.Select
        sourceSht.ListObjects("Data").DataBodyRange.Select
            nvars = Selection.Columns.Count 'Number of variables in dataset (incl. features & target, plus index)
            nobservations = Selection.Rows.Count 'Number of observations in dataset
        
            testObservations = WorksheetFunction.RoundDown(nobservations * testReserve, 0) 'Number of observations to reserve for testing
            nValidation = WorksheetFunction.RoundDown((nobservations - testObservations) / kFolds, 0) 'Number of observations per validation subset
            trainObservations = nValidation * kFolds 'Number of observations per training subset
            lowIndex = trainObservations + 1 'Sets the cut point between data for training/validation and test reserve.

'Copy header into an array of its own (for future usage)
    ReDim header(nvars) As String
    header() = headersFx(nvars)

'Create ReTraining Set.
    sourceSht.Select 'Return to main dataset and copy test observations
    sourceSht.ListObjects("Data").DataBodyRange.Rows("1:" & trainObservations).Select
        Selection.Copy
        Sheets.Add After:=Sheets(Sheets.Count) 'Create new sheet for test observations
        ActiveSheet.Name = "ReTrain"
        Range("A1").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Rows(1).Select 'Add header
        Selection.Insert Shift:=xlDown
        For i = 1 To nvars
            Cells(1, i) = header(i)
        Next i

'Create Test Reserve.
    sourceSht.Select 'Return to main dataset and copy test observations
    sourceSht.ListObjects("Data").DataBodyRange.Rows(lowIndex & ":" & nobservations).Select
    Selection.Copy
    Sheets.Add After:=Sheets(Sheets.Count) 'Create new sheet for test observations
    ActiveSheet.Name = "Test"
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Rows(1).Select 'Add header
    Selection.Insert Shift:=xlDown
    For i = 1 To nvars
        Cells(1, i) = header(i)
    Next i

'Create k training/validation folds.
    lowIndex = 1 'Re-initialise lower index (first observation of 1st fold)
    upIndex = nValidation 'Reinitialise upper index (last observation of 1st fold)
    For k = 1 To kFolds 'Main iteration: copy all train/validate entries to train sheet, validate entries to validate sheet, delete validate entries from train sheet, repeat.
        Sheets.Add Before:=Sheets("ReTrain")
        ActiveSheet.Name = "Train" & k
        sourceSht.Select
        sourceSht.ListObjects("Data").DataBodyRange.Rows("1:" & trainObservations).Select
        Selection.Copy
        Sheets("Train" & k).Select
        Range("A1").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Sheets.Add Before:=Sheets("ReTrain")
        ActiveSheet.Name = "Validate" & k
        Sheets("Train" & k).Select
        Rows(lowIndex & ":" & upIndex).Select
        Selection.Copy
        Sheets("Validate" & k).Select
        Range("A1").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Rows(1).Select
        Selection.Insert Shift:=xlDown
        For j = 1 To nvars
            Cells(1, j) = header(j)
        Next j
        Sheets("Train" & k).Select
        Rows(lowIndex & ":" & upIndex).Select
        Selection.Delete Shift:=xlUp
        Rows(1).Select
        Selection.Insert Shift:=xlDown
        For j = 1 To nvars
            Cells(1, j) = header(j)
        Next j
        lowIndex = lowIndex + nValidation
        upIndex = nValidation * (k + 1)
    Next k ' End of main iteration

'Final formatting/matters
    Call selectFirstCell 'Calls a formatting function that selects A1 cell in all sheets
    Sheets("Dashboard").Select
    
'Error handling & gracious exit
Exit Sub
ErrorHandler: MsgBox ("Something went wrong. Most likely, data is not appropriately formatted." & vbCrLf & vbCrLf & "Please check and re-import if needed.")

End Sub
' --------------------------------- SEGMENTATION SUB ENDS




' --------------------------------- NEW SUB STARTS
' ----------------
'

Sub foldsByRandomisation()
' This macro creates folds using a randomisation approach.

'Dims & other initial matters
    Dim kFolds As Integer, testReserve As Double, nobservations As Integer, nvars As Integer, testObservations As Integer, nValidation As Integer, trainObservations
    Dim i As Integer, j As Integer, k As Integer, lowIndex As Integer, upIndex As Integer, randomVal As Integer
    Dim sourceSht As Worksheet, trainSht As Worksheet, valSht As Worksheet
    Dim header() As String, data() As Double
    
        On Error GoTo ErrorHandler
        Application.ScreenUpdating = False
        Sheets("Data").Select
        Set sourceSht = ActiveSheet

'Pick up hyper-parameters from "Dashboard" and "Data" sheet
    Sheets("Dashboard").Select
    testReserve = Range("C9") 'Size of test reserve (%)
    kFolds = Range("C11") 'Number of desired folds

'Calculate additional parameters from "Data" sheet
    sourceSht.Select
    sourceSht.ListObjects("Data").DataBodyRange.Select
    nvars = Selection.Columns.Count 'Number of variables in dataset (incl. features & target, plus index)
    nobservations = Selection.Rows.Count 'Number of observations in dataset
    testObservations = WorksheetFunction.RoundDown(nobservations * testReserve, 0) 'Number of observations to reserve for testing
    nValidation = WorksheetFunction.RoundDown((nobservations - testObservations) / kFolds, 0) 'Number of observations per validation subset
    trainObservations = nValidation * kFolds 'Number of observations per training subset

'Copy header into an array of its own (for future usage)
    ReDim header(nvars) As String
    header() = headersFx(nvars)

'Create ReTrain sheet and set as new source
    Sheets.Add Before:=Sheets("ReTrain")
    ActiveSheet.Name = "ReTrain"
    sourceSht.ListObjects("Data").DataBodyRange.Copy 'Transfer data before changing sourceSht
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    Set sourceSht = ActiveSheet 'Set ReTrain as new sourceSht

'Create testing subset by taking observations away from ReTrain sheet.
    ReDim data(testObservations, nvars) As Double 'Array to handle randomised selection
    lowIndex = 1 'Set lower bound for random selection
    upIndex = nobservations 'Set upper bound for random selection'
    sourceSht.Select 'Select observations randomly and copy each into array
    For i = 1 To testObservations
        Randomize
        randomVal = Int((upIndex - lowIndex + 1) * Rnd + lowIndex)
        For j = 1 To nvars
            data(i, j) = sourceSht.Cells(randomVal, j)
        Next j
        sourceSht.Rows(randomVal).Delete
        upIndex = upIndex - 1
    Next i
    
    Sheets.Add After:=Sheets(Sheets.Count) 'Transfer array to test sheet
    ActiveSheet.Name = "Test"
    For i = 1 To testObservations
        For j = 1 To nvars
            Cells(i, j) = data(i, j)
        Next j
    Next i
    
    Rows(1).Insert Shift:=xlDown 'Insert header
    For j = 1 To nvars
       Cells(1, j) = header(j)
    Next j

'Create k training/validation folds.
    ReDim data(nValidation, nvars) As Double 'Array to handle randomised selection
    For k = 1 To kFolds 'Main iteration: copy all train/validate entries to train sheet, randomly select validate entries and transfer to validate sheet, delete selected validate entries from train sheet, repeat.
        lowIndex = 1 'Set lower bound for randomisation
        upIndex = nobservations - testObservations 'Set upper bound for randomisation
        Sheets.Add After:=Sheets(Sheets.Count) 'Add train sheet
        ActiveSheet.Name = "Train" & k
        Set trainSht = ActiveSheet
        sourceSht.Select 'Copy train/validate observations
        Range("A1").CurrentRegion.Select
        Selection.Copy
        trainSht.Select 'Import observations into train sheet
        Cells(1, 1).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
        Sheets.Add Before:=Sheets("ReTrain") 'Add validate sheet
        ActiveSheet.Name = "Validate" & k
        Set valSht = ActiveSheet
        trainSht.Select 'Copy random observations in train sheet, delete said observation from train sheet, update randomisation bound, repeat.
        For i = 1 To nValidation
            Randomize
            randomVal = Int((upIndex - lowIndex + 1) * Rnd + lowIndex)
            For j = 1 To nvars
                data(i, j) = trainSht.Cells(randomVal, j)
            Next j
            trainSht.Rows(randomVal).Delete
            upIndex = upIndex - 1
        Next i
        valSht.Select 'Dump array of randomly selected observations into validate sheet
        For i = 1 To nValidation
            For j = 1 To nvars
                Cells(i, j) = data(i, j)
            Next j
        Next i
        Rows(1).Insert Shift:=xlDown 'Insert header into validate sheet
        For j = 1 To nvars
           Cells(1, j) = header(j)
        Next j
        trainSht.Select
        Rows(1).Insert Shift:=xlDown ' Insert header into train sheet
        For j = 1 To nvars
            Cells(1, j) = header(j)
        Next j
    Next k

' Final formatting/matters
Call selectFirstCell 'Calls a formatting function that selects A1 cell in all sheets
Sheets("Dashboard").Select

' Error handling & gracious exit
Exit Sub
ErrorHandler: MsgBox ("Something went wrong. Most likely, data is not appropriately formatted." & vbCrLf & vbCrLf & "Please check and re-import if needed.")
End Sub
' --------------------------------- RANDOMISATION SUB ENDS
