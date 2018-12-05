Attribute VB_Name = "Module10"
' to input condition row titile as compiled distractor vs target
' from f01-d f01-t ... f32-d f32-t
Const rowNum = 1

Dim conditionTitle As String

Dim trialNum As Integer
Dim colNum As Integer
Public Sub getFaceDistractorTargetConditionTitle()

    colNum = 2 ' start column number from 2
    While colNum < 66

        Dim col As Integer
        col = colNum - 2 ' start col from 0, easier to compute
        
        ' set AOI to face
        conditionTitle = "f"

        ' determine random trial number
        trialNum = Application.Floor((col Mod 128) / 2, 1) + 1
        ' add a "0" at the front if trial is a single digit
        If trialNum <= 9 Then
            conditionTitle = conditionTitle & "0" & CStr(trialNum)
        Else
            conditionTitle = conditionTitle & CStr(trialNum)
        End If

        ' determine d/t
        If col Mod 2 = 0 Then
            conditionTitle = conditionTitle & "-d"
        ElseIf col Mod 2 = 1 Then
            conditionTitle = conditionTitle & "-t"
        End If

        ' output condition title to the corresponding column
        Cells(rowNum, colNum).Value = conditionTitle

        ' move onto the next column
        colNum = colNum + 1
    Wend

End Sub


