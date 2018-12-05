Attribute VB_Name = "Module1"
' input ouput column can be configured for future usage
Const outputColumn = 9  ' output ratio on column 9
Const inputColumn = 6    ' input fixation time from column 6

' calculate ratio of AOI/face for a given fixation time in input row
Public Function ratio(inputRow As Integer) As Double

    Dim faceFixTime
    Dim inputFixTime
    
    ' AOI is eyes
    If inputRow Mod 3 = 2 Then
        ' next row contains face fixation time
        faceFixTime = Cells(inputRow + 1, inputColumn).Value
        If faceFixTime = 0 Then
            ratio = 0
        Else
            inputFixTime = Cells(inputRow, inputColumn).Value
            ratio = inputFixTime / faceFixTime
        End If
    
    ' AOI is mouth
    ElseIf inputRow Mod 3 = 1 Then
        ' previous row contains face fixation time
        faceFixTime = Cells(inputRow - 1, inputColumn).Value
        If faceFixTime = 0 Then
            ratio = 0
        Else
            inputFixTime = Cells(inputRow, inputColumn).Value
            ratio = inputFixTime / faceFixTime
        End If
    
    ' AOI is face itself
    Else
        ratio = 0

End If

End Function

' loops and output ratio to output column
Public Sub getAOIRatio()

    Dim inputRow As Integer
    inputRow = 2   ' input data starts from row 2
    
    ' output ratio until the last data in input column
    While Not IsEmpty(Cells(inputRow, inputColumn))
        Cells(inputRow, outputColumn).Value = ratio(inputRow)
    
        ' move onto the next input row
        inputRow = inputRow + 1
    Wend
    
End Sub
