Attribute VB_Name = "Module11"
' to input condition row title
' from e01-d1   e01-d2  e01-d3  e01-t to e32-d1 e32-d2  e32-d3  e32-t
' then followed by m01-d1 ... m32-t
Const rowNum = 1

Dim conditionTitle As String

Dim trialNum As Integer
Dim colNum As Integer

Public Sub getConditionTitle()

    colNum = 2 ' start column number from 2
    While colNum <= 257

        Dim col As Integer
        col = colNum - 2 ' start col from 0, easier to compute

        ' determine aoi
        If col < 128 Then
            conditionTitle = "e"
        ElseIf col <= 255 Then
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
