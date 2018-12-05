Attribute VB_Name = "Module2"
' input ouput column can be configured for future usage
Const outputColumn = 10 ' output ratio on column 10
Const inputColumn = 6    ' input fixation time from column 6

' output specific face/all 4 faces ratio for a given fixation time in input row
Sub getFaceRatio()
    
    Dim inputRow As Integer ' input row number
    inputRow = 2                  ' data starts from row 2
    
    Dim trialNum As Integer  ' represents trial number minus 2 (counting trial from 0)
    trialNum = Int(inputRow / 12)
    
    Dim totalFixTime             ' total fixation time of all 4 faces in a trial
    Dim inputFixTime            ' fixation time of one face in a trial
    Dim faceRatio As Double  ' specific face/all 4 faces ratio
    
    ' output face ratio until the last data in input column
    While Not IsEmpty(Cells(inputRow, inputColumn))
        
        ' AOI is face for every 3nd, 6th, 9th, and 12th row
        If inputRow Mod 3 = 0 Then
            totalFixTime = Cells(12 * trialNum + 3, inputColumn).Value _
                                + Cells(12 * trialNum + 6, inputColumn).Value _
                                + Cells(12 * trialNum + 9, inputColumn).Value _
                                + Cells(12 * trialNum + 12, inputColumn).Value
            inputFixTime = Cells(inputRow, inputColumn).Value
            
            If totalFixTime = 0 Then
                faceRatio = 0
            Else
                faceRatio = inputFixTime / totalFixTime
            End If
            
            'output face ratio into the output column
            ' Cells(inputRow, outputColumn).Value = faceRatio
        
        ' AOI is not face (eyes or mouth)
        Else
            faceRatio = 0
            ' Cells(inputRow, outputColumn).Value = faceRatio
            
       End If
    
    'output face ratio into the output column
    Cells(inputRow, outputColumn).Value = faceRatio
    
    ' move onto the next input row
    inputRow = inputRow + 1
    
    Wend
            
End Sub
