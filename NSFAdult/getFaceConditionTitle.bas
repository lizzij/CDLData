Attribute VB_Name = "Module12"
' to input face aoi condition row title
' from f01-d1   f01-d2  f01-d3  f01-t to f32-d1 f32-d2  f32-d3  f32-t
Const rowNum = 1

Dim conditionTitle As String

Dim trialNum As Integer
Dim colNum As Integer

Public Sub getFaceConditionTitle()

    colNum = 2 ' start column number from 2
    While colNum <= 129

        Dim col As Integer
        col = colNum - 2 ' start col from 0, easier to compute

        ' aoi is face
        conditionTitle = "f"

        ' determine random trial number
        trialNum = Application.Floor(col / 4, 1) + 1
        ' add a "0" at the front if trial is a single digit
        If trialNum <= 9 Then
            conditionTitle = conditionTitle & "0" & CStr(trialNum)
        Else
            conditionTitle = conditionTitle & CStr(trialNum)
        End If

        ' determine d1/d2/d3/t
        If col Mod 4 = 0 Then
            conditionTitle = conditionTitle & "-d1"
        ElseIf col Mod 4 = 1 Then
            conditionTitle = conditionTitle & "-d2"
        ElseIf col Mod 4 = 2 Then
            conditionTitle = conditionTitle & "-d3"
        ElseIf col Mod 4 = 3 Then
            conditionTitle = conditionTitle & "-t"
        End If

        ' output condition title to the corresponding column
        Cells(rowNum, colNum).Value = conditionTitle

        ' move onto the next column
        colNum = colNum + 1
    Wend

End Sub

