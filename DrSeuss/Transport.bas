Attribute VB_Name = "Module2"
' Transport participant number and ratio to DrSeuss Export worksheet
' and format them into one row per participant, one col per condition

' column 1 of input sheets contains participant number
Const participantNumCol = 1
' column 14 of input sheets contains rato
Const ratioCol = 14
' column 15 of input sheets contains condition
Const conditionCol = 15
' match participant and condition, print ratio

' Configure workbook
Dim DrSeuss As Workbook

' Configure input worksheets
Dim Data As Worksheet
Dim Condition1a As Worksheet
Dim Condition2a As Worksheet
Dim Condition3a As Worksheet
Dim Condition4a As Worksheet

' Configure output worksheets
Dim DrSeussExport As Worksheet
Dim inputRow As Integer
Dim lastRow As Integer
    

' Transport data from one input worksheet
Public Function Transport(Data As Worksheet) As Worksheet

    ' Find the first empty row
    inputRow = 2
    
    ' Specify name of the workbook and worksheet
    Set DrSeuss = Workbooks("DrSeuss_all conditions.xlsm")
    Set Condition1a = DrSeuss.Worksheets("Condition 1a")
    Set Condition2a = DrSeuss.Worksheets("Condition 2a")
    Set Condition3a = DrSeuss.Worksheets("Condition 3a")
    Set Condition4a = DrSeuss.Worksheets("Condition 4a")
    Set DrSeussExport = DrSeuss.Worksheets("DrSeuss Export")
    
    While Not IsEmpty(Data.Cells(inputRow, ratioCol))
        
        ' Get condition and ratio from inputRow
        Dim condition As String
        condition = Data.Cells(inputRow, conditionCol).Value
        Dim ratio As Double
        ratio = Data.Cells(inputRow, ratioCol).Value
        
        ' Check for corresponding participant and condition in export worksheet
        Dim targetRow As Integer
        Dim row As Integer
        row = inputRow - 2
        targetRow = Application.Floor(row / 42, 1) + 2
        
        Dim targetCol As Integer
        Dim targetCondition As String
        
        For targetCol = 2 To 29
            targetCondition = DrSeussExport.Cells(1, targetCol).Value
            If condition = targetCondition Then
                DrSeussExport.Cells(targetRow, targetCol).Value = ratio
            End If
        Next targetCol
        
        ' move onto the next input row
        inputRow = inputRow + 1
        lastRow = inputRow
        
     Wend
     
     ' print participant number in the first column
     Dim participantNum As Integer
     participantNum = Application.Floor((lastRow - 1) / 42, 1)
     Dim participant As Integer
     participant = 1
     While participant <= participantNum
        DrSeussExport.Cells((participant + 1), 1).Value = Data.Cells((participant * 42 + 1), 1).Value
        ' move onto the next participant
        participant = participant + 1
     Wend
     
    Set Transport = DrSeussExport
    
End Function
Sub transport1aData()
    
    Set DrSeussExport = Transport(Condition1a)
    
End Sub
    
