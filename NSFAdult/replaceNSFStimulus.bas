Attribute VB_Name = "Module6"
' replace NSF stimulus name with simplified code
' e.g.
' Alyssa_Alltalk_4diff VisAsync LeftBottom -We can..- 19 s1.avi   LeftTopEyes
' A                           -v             -3              -1                              -2        -2

' - Alyssa_Alltalk_4diff => A
' - Megan_Alltalk_4diff => M

' - VisAsync => a
' - RightTop => 1
' - LeftTop => 2
' - LeftBottom => 3
' - RightBottom => 4

' - -We can..- 19 s1.avi => 1
' - -Good morning.- 19 s.avi => 2
' - --They like to ice-.- 19 s.avi => 3
' - -But..- 19s.avi => 4

' - Face => 0
' - Mouth => 1
' - Eyes => 2

' - Target => *

' input ouput column can be configured for future usage
Const outputColumn = 12  ' output simplified condition on column 12
Const stimulusColumn = 3  ' input stimulus name from column 3
Const aoiColumn = 4         ' input AOI name from column 4

Public Sub replaceNSFStimulus()
    
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
        WrdArray() = Split(stimulus, "_")
       
        ' - Alyssa_Alltalk_4diff => A
        ' - Megan_Alltalk_4diff => M
        If InStr(1, WrdArray(0), "Alyssa") Then
            newStimulusName = "A"
        Else
            newStimulusName = "M"
        End If
        
        ' - VisAsync => v
        ' - AudSync => a
        If InStr(1, WrdArray(2), "VisAsync") Then
            newStimulusName = newStimulusName & "-v"
        ElseIf InStr(1, WrdArray(2), "ViisAsync") Then
            newStimulusName = newStimulusName & "-v"
       ElseIf InStr(1, WrdArray(2), "VisAync") Then
            newStimulusName = newStimulusName & "-v"
        Else
            newStimulusName = newStimulusName & "-a"
        End If
        
        ' - RightTop => 1
        ' - LeftTop => 2
        ' - LeftBottom => 3
        ' - RightBottom => 4
        If InStr(1, WrdArray(2), "RightTop") Then
            newStimulusName = newStimulusName & "-1"
        ElseIf InStr(1, WrdArray(2), "LeftTop") Then
            newStimulusName = newStimulusName & "-2"
        ElseIf InStr(1, WrdArray(2), "LeftBottom") Then
            newStimulusName = newStimulusName & "-3"
        ElseIf InStr(1, WrdArray(2), "LefttBottom") Then
            newStimulusName = newStimulusName & "-3"
        Else
            newStimulusName = newStimulusName & "-4"
        End If
        
        ' - -We can..- 19 s1.avi => 1
        ' - -Good morning.- 19 s.avi => 2
        ' - --They like to ice-.- 19 s.avi => 3
        ' - -But..- 19s.avi => 4
        If InStr(1, WrdArray(2), "We can") Then
            newStimulusName = newStimulusName & "-1"
        ElseIf InStr(1, WrdArray(2), "Good morning") Then
            newStimulusName = newStimulusName & "-2"
        ElseIf InStr(1, WrdArray(2), "They like to ice") Then
            newStimulusName = newStimulusName & "-3"
        Else
            newStimulusName = newStimulusName & "-4"
        End If
        
        ' replace AOI
        ' - RightTop => 1
        ' - LeftTop => 2
        ' - LeftBottom => 3
        ' - RightBottom => 4
        If InStr(1, aoi, "RightTop") Then
            newStimulusName = newStimulusName & "-1"
        ElseIf InStr(1, aoi, "LeftTop") Then
            newStimulusName = newStimulusName & "-2"
        ElseIf InStr(1, aoi, "LeftBottom") Then
            newStimulusName = newStimulusName & "-3"
        Else
            newStimulusName = newStimulusName & "-4"
        End If
        
        ' - Face => 0
        ' - Mouth => 1
        ' - Eyes => 2
        If InStr(1, aoi, "Face") Then
            newStimulusName = newStimulusName & "-0"
        ElseIf InStr(1, aoi, "Mouth") Then
            newStimulusName = newStimulusName & "-1"
        Else
            newStimulusName = newStimulusName & "-2"
        End If
        
        ' - Target => *
         If InStr(1, aoi, "Target") Then
            newStimulusName = newStimulusName & "*"
        End If
        
        'output new stimulus name
        Cells(inputRow, outputColumn).Value = newStimulusName
        
        ' move onto the next input row
        inputRow = inputRow + 1
    Wend
    
End Sub

