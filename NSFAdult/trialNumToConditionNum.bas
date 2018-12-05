Attribute VB_Name = "Module8"
' From Trial# output Condition# according to experiment design Random # 1

Const conditionColumn = 14  ' output condition #  on column 14
Const trialColumn = 2  ' input trial # from column 2

Public Sub trialNumToConditionNum()
    
    Dim trialRow As Integer
    Dim trial As String
    Dim condition As String

    trialRow = 2  'trial data starts from row 2

    ' ouput condition until the last data in the trialRow column
    While Not IsEmpty(Cells(trialRow, trialColumn))

        trial = Cells(trialRow, trialColumn).Text

        If InStr(1, trial, "002") Then
            condition = "19"
        ElseIf InStr(1, trial, "003") Then
            condition = "16"
        ElseIf InStr(1, trial, "004") Then
            condition = "20"
        ElseIf InStr(1, trial, "005") Then
            condition = "29"
        ElseIf InStr(1, trial, "006") Then
            condition = "21"
        ElseIf InStr(1, trial, "007") Then
            condition = "8"
        ElseIf InStr(1, trial, "008") Then
            condition = "28"
        ElseIf InStr(1, trial, "009") Then
            condition = "2"
        ElseIf InStr(1, trial, "010") Then
            condition = "7"
        ElseIf InStr(1, trial, "011") Then
            condition = "3"
        ElseIf InStr(1, trial, "012") Then
            condition = "15"
        ElseIf InStr(1, trial, "013") Then
            condition = "30"
        ElseIf InStr(1, trial, "014") Then
            condition = "1"
        ElseIf InStr(1, trial, "015") Then
            condition = "9"
        ElseIf InStr(1, trial, "016") Then
            condition = "10"
        ElseIf InStr(1, trial, "017") Then
            condition = "26"
        ElseIf InStr(1, trial, "018") Then
            condition = "17"
        ElseIf InStr(1, trial, "019") Then
            condition = "13"
        ElseIf InStr(1, trial, "020") Then
            condition = "12"
        ElseIf InStr(1, trial, "021") Then
            condition = "27"
        ElseIf InStr(1, trial, "022") Then
            condition = "32"
        ElseIf InStr(1, trial, "023") Then
            condition = "23"
        ElseIf InStr(1, trial, "024") Then
            condition = "4"
        ElseIf InStr(1, trial, "025") Then
            condition = "18"
        ElseIf InStr(1, trial, "026") Then
            condition = "22"
        ElseIf InStr(1, trial, "027") Then
            condition = "5"
        ElseIf InStr(1, trial, "028") Then
            condition = "25"
        ElseIf InStr(1, trial, "029") Then
            condition = "24"
        ElseIf InStr(1, trial, "030") Then
            condition = "6"
        ElseIf InStr(1, trial, "031") Then
            condition = "11"
        ElseIf InStr(1, trial, "032") Then
            condition = "31"
        ElseIf InStr(1, trial, "033") Then
            condition = "14"
        End If

        'output new stimulus name
        Cells(trialRow, conditionColumn).Value = condition
        
        ' move onto the next input row
        trialRow = trialRow + 1
    Wend
    
End Sub
