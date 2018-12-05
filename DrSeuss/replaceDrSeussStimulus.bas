Attribute VB_Name = "Module5"
' DrSeuss replace stimulus and AOI column name
' e.g.,
' Baseline_1  Mouth
' b1              -1
' 6a_Match    Eyes
' 6a -1          -2

' - Baseline => b

' - Match => 1 (not used, commented out)
' - NoMatch => 0  (not used, commented out)

' - FaceAOI => f
' - MouthAOI => m
' - EyesAOI => e

' input ouput column can be configured for future usage
Const outputColumn = 15  ' output simplified condition on column 15
Const stimulusColumn = 3  ' input stimulus name from column 3
Const aoiColumn = 5         ' input AOI name from column 3

Public Sub replaceDrSeussStimulus()
    
    Dim inputRow As Integer
    Dim stimulus As String
    Dim aoi As String
    Dim newStimulusName As String
    Dim WrdArray() As String

    inputRow = 2   ' input data starts from row 2
    
    ' output ratio until the last data in input column
    While Not IsEmpty(Cells(inputRow, stimulusColumn))
    
        stimulus = Cells(inputRow, stimulusColumn).Text ' get stimulus name
        aoi = Cells(inputRow, aoiColumn).Text                 ' get AOI name
        newStimulusName = ""
       
        ' replace "Baseline" with "0", followed by "-[number])
        If InStr(1, stimulus, "Baseline") Then
            newStimulusName = "b" & Right(stimulus, 1)
        Else
            WrdArray() = Split(stimulus, "_")
            ' replace "NoMatch" with "0" (not used, commented out)
            If InStr(1, WrdArray(1), "NoMatch") Then
                newStimulusName = WrdArray(0) ' & "-" & "0"
            ' replace "Match" with "1" (not used, commented out)
            Else
                newStimulusName = WrdArray(0) ' & "-" & "1"
            End If
        End If
        
        ' replace AOI
        ' Face - 0
        ' Mouth - 1
        ' Eyes - 2
        If InStr(1, aoi, "Face") Then
            newStimulusName = newStimulusName & "-f"
        ElseIf InStr(1, aoi, "Mouth") Then
            newStimulusName = newStimulusName & "-m"
        Else
            newStimulusName = newStimulusName & "-e"
        End If
        
        'output new stimulus name
        Cells(inputRow, outputColumn).Value = newStimulusName
        
        ' move onto the next input row
        inputRow = inputRow + 1
    Wend
    
End Sub
