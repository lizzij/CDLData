Attribute VB_Name = "Module1"
' to input condition row title
' from b1-e 1a-e 2a-e ... 10a-e 11a-e 13a-e b2-e (no 12-a)
' then b1-m 1a-m 2a-m ... 10a-m 11a-m 13a-m b2-m (no 12-a)
Const rowNum = 1

Dim conditionTitle As String

Dim trialNum As Integer
Dim colNum As Integer

Public Sub getConditionTitle()

    colNum = 2 ' start column number from 2
    While colNum <= 29

        Dim col As Integer
        col = colNum - 2 ' start col from 0, easier to compute

        ' condition number
        If col = 0 Or col = 14 Then
            conditionTitle = "b1"
        ElseIf col <= 11 Then
            conditionTitle = col & "a"
        ElseIf col >= 15 And col <= 25 Then
            conditionTitle = (col - 14) & "a"
        ElseIf col = 12 Or col = 26 Then
            conditionTitle = "13a"
        ElseIf col = 13 Or col = 27 Then
            conditionTitle = "b2"
        End If
        
        ' determine AOI -e, or -m
        If col <= 13 Then
            conditionTitle = conditionTitle & "-e"
        Else
            conditionTitle = conditionTitle & "-m"
        End If

        ' output condition title to the corresponding column
        Cells(rowNum, colNum).Value = conditionTitle

        ' move onto the next column
        colNum = colNum + 1
    Wend

End Sub

