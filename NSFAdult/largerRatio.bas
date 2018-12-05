Attribute VB_Name = "Module3"
' input ouput column can be configured for future usage
Const outputColumn = 11               ' output ratio on column 11
Const inputAOIColumn = 9              ' input AOI fixation time from column 9
Const inputFaceColumn = 10           ' input face fixation time from column 10

' get the larger of the input AOI and face fixation for a given row
Public Function largerRatio(inputRow As Integer) As Double

    Dim AOIRatio
    Dim faceRatio
    
    AOIRatio = Cells(inputRow, inputAOIColumn).Value
    faceRatio = Cells(inputRow, inputFaceColumn).Value
    
    If AOIRatio > faceRatio Then
        largerRatio = AOIRatio
    Else
        largerRatio = faceRatio
    End If
    
End Function

' loops and output ratio to output column
Public Sub getLargerRatio()

    Dim inputRow As Integer
    inputRow = 2   ' input data starts from row 2
    
    ' output ratio until the last data in input column
    While Not IsEmpty(Cells(inputRow, inputAOIColumn))
        Cells(inputRow, outputColumn).Value = largerRatio(inputRow)
    
        ' move onto the next input row
        inputRow = inputRow + 1
    Wend
    
End Sub
