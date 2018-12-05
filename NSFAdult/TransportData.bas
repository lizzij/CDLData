Attribute VB_Name = "Module9"
' transport participant number, condition and ratio to overall worksheet
'    NSF Exp 1 Adult FaceAOI. and
'    NSF Exp 1 Adult EyesMouthAOI
' and format the data into
'    one row per participant
'    one column per conditio

' column 1 of Data contains participant number
Const participantNumCol = 1
' column 11 of Data contains ratio
Const ratioCol = 11
' column 13 of Data contains condition
Const conditionCol = 13
' last row containing data
Const inputLastRow = 3073
' match participant and condition, print ratio

Sub TransportData()

    ' Configure workbook
    Dim NSFExp1 As Workbook

    ' Configure input worksheet
    Dim Data As Worksheet

    ' Configure output worksheets
    Dim FaceEyesAOI As Worksheet
    Dim EyesMouthAOI As Worksheet

    ' Specify name of the workbook and worksheet
    Set NSFExp1 = Workbooks("NSF Exp 1 Adult Random 1 Trial Summary (AOI).xlsm")
    Set Data = NSFExp1.Worksheets("NSF Exp 1 Adult Random 1 Trial ")
    Set FaceAOI = NSFExp1.Worksheets("NSF Exp 1 Adult FaceAOI")
    Set EyesMouthAOI = NSFExp1.Worksheets("NSF Exp 1 Adult EyesMouthAOI")

    ' Set input row from 2
    Dim inputRow As Integer
    inputRow = 2

    While inputRow <= 3073

        ' Get condition and ratio from inputRow
        Dim condition As String
        condition = Data.Cells(inputRow, conditionCol).Value
        Dim ratio As Double
        ratio = Data.Cells(inputRow, ratioCol).Value

        ' check for corresponding participant and condition in aoi forms
        Dim targetRow As Integer
        Dim row As Integer
        row = inputRow - 2
        targetRow = Application.Floor(row / 384, 1) + 2

        Dim targetFaceCol As Integer
        Dim targetFaceCondition As String

        Dim targetEyesMouthCol As Integer
        Dim targetEyesMouthCondition As String

        Dim aoi As String
        aoi = Left(condition, 1)

        If aoi = "f" Then
            For targetFaceCol = 2 To 129
                targetFaceCondition = FaceAOI.Cells(1, targetFaceCol).Value
                If condition = targetFaceCondition Then
                    FaceAOI.Cells(targetRow, targetFaceCol).Value = ratio
                End If
            Next targetFaceCol
        ElseIf aoi = "e" Then
            For targetEyesMouthCol = 2 To 129
                targetEyesMouthCondition = EyesMouthAOI.Cells(1, targetEyesMouthCol).Value
                If condition = targetEyesMouthCondition Then
                    EyesMouthAOI.Cells(targetRow, targetEyesMouthCol).Value = ratio
                End If
            Next targetEyesMouthCol
        ElseIf aoi = "m" Then
            For targetEyesMouthCol = 130 To 257
                targetEyesMouthCondition = EyesMouthAOI.Cells(1, targetEyesMouthCol).Value
                If condition = targetEyesMouthCondition Then
                    EyesMouthAOI.Cells(targetRow, targetEyesMouthCol).Value = ratio
                End If
            Next targetEyesMouthCol
        End If
        
    ' print the participant number in the first column
    Dim participantNum As Integer
    participantNum = Application.Floor((inputLastRow - 1) / 384, 1)
    Dim participant As Integer
    participant = 1
    
    While participant <= participantNum
        FaceAOI.Cells((participant + 1), 1).Value = Data.Cells((participant * 384 + 1), 1).Value
        EyesMouthAOI.Cells((participant + 1), 1).Value = Data.Cells((participant * 384 + 1), 1).Value
        
        ' move onto the next participant
        participant = participant + 1
    Wend
    
        ' move onto the next input row
        inputRow = inputRow + 1
    Wend
    
End Sub
