Attribute VB_Name = "Module4"
' to input condition row titile as compiled distractor vs target
' from e01-d e01-t ... e32-d e32-t
' then m01-d m01-t ... m32-d m32-t
Const rowNum = 1

Dim conditionTitle As String

Dim trialNum As Integer
Dim colNum As Integer
Public Sub getEyesMouthDistractorTargetConditionTitle()

    colNum = 2 ' start column number from 2
    While colNum <= 127

        Dim col As Integer
        col = colNum - 2 ' start col from 0, easier to compute

        ' determine aoi
        If col < 64 Then
            conditionTitle = "e"
        ElseIf col <= 128 Then
            conditionTitle = "m"
        End If

        ' determine random trial number
        trialNum = Application.Floor((col Mod 128) / 4, 1) + 1
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

