Attribute VB_Name = "Module5"
' Transport distractor-target ratio

' column 1 of Data contains participant number
Const participantNumCol = 1

Sub TransportFaceDistractorTarget()

    ' Configure workbook
    Dim NSFExp1 As Workbook

    ' Configure input worksheet
    Dim FaceSorted As Worksheet
    Dim EyesMouthSorted As Worksheet

    ' Configure output worksheets
    Dim FaceDistractorTarget As Worksheet
    Dim EyesMouthDistractorTarget As Worksheet

    ' Specify name of the workbook and worksheet
    Set NSFExp1 = Workbooks("NSF Exp 1 Adult Random 1 Trial Summary (AOI).xlsm")
    Set FaceSorted = NSFExp1.Worksheets("NSF Exp 1 Adult FaceAOI")
    Set FaceDistractorTarget = NSFExp1.Worksheets("NSF Exp 1 Adult FaceAOI dt")
    Set EyesMouthSorted = NSFExp1.Worksheets("NSF Exp 1 Adult EyesMouthAOI")
    Set EyesMouthDistractorTarget = NSFExp1.Worksheets("NSF Exp 1 Adult EyesMouthAOI dt")
    
    ' Set input row from 2
    Dim inputRow As Integer
    inputRow = 2
    ' Set output row from 2
    Dim outputRow As Integer
    outputRow = 2
    
    While Not IsEmpty(FaceSorted.Cells(inputRow, participantNumCol))
    
        ' transport the participant name to the participant column
        FaceDistractorTarget.Cells(outputRow, participantNumCol) = FaceSorted.Cells(inputRow, participantNumCol).Value
        EyesMouthDistractorTarget.Cells(outputRow, participantNumCol) = EyesMouthSorted.Cells(inputRow, participantNumCol).Value

        ' For rest of the rows, compile d1, d2, d3 into d, and transport target
        Dim totalFaceRatio As Double
        totalFaceRatio = 0
        Dim averageFaceRatio As Double
        
        Dim inputCol1 As Integer
        inputCol1 = 2
        Dim outputCol As Integer
        outputCol = 2
        For inputCol1 = 2 To 129
            If inputCol1 Mod 4 = 2 Then
                totalFaceRatio = FaceSorted.Cells(inputRow, inputCol1).Value _
                + FaceSorted.Cells(inputRow, inputCol1 + 1).Value _
                + totalFaceRatio + FaceSorted.Cells(inputRow, inputCol1 + 2).Value
                averageFaceRatio = totalFaceRatio / 3
                outputCol = Application.Floor((inputCol1 - 1) / 2, 1) + 2
            ElseIf inputCol1 Mod 4 = 1 Then
                averageFaceRatio = FaceSorted.Cells(inputRow, inputCol1).Value
                totalFaceRatio = 0
                outputCol = outputCol + 1
            End If
            
            FaceDistractorTarget.Cells(outputRow, outputCol).Value = averageFaceRatio
            ' FaceDistractorTarget.Cells(10, outputCol).Value = inputCol1 ' for debugging
        Next inputCol1
        
        Dim totalEyesMouthRatio As Double
        totalEyesMouthRatio = 0
        Dim averageEyesMouthRatio As Double

        ' move onto the next input row
        inputRow = inputRow + 1
        outputRow = inputRow
    Wend
    
End Sub
