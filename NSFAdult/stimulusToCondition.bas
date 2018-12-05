Attribute VB_Name = "Module7"
' Transform the stimulus condition into the format:
' eye/mouth/faceAOI Condition-distractor/distractor.
'
'  - AOI (e for eyes, m for mouth, f for face);
'  - design trial number (1 to 32, different for random order # 1/2/3/4);
'  - distractor or target (d1, d2, d3, t - where d1/2/3 order does not matter).
'
' i.e. f01-d1 = face AOI, design trial number 1, distractor1;
'      e10-d3 = eyes AOI, design trial number 10, distractor3;
'      m31-t = mouth AOI, design trial number 31, target.
'
' To transform a stimulus representation to a condition representation, like:
'   A-v-3-1-2-0    =>    f19-d
'   M-a-1-4-1-2*   =>    e08-t
'
' First,
' Last number in stimulus (the 11th character, like 0 in M-a-1-3-4-0) represents AOI
' - Face => 0
' - Mouth => 1
' - Eyes => 2
'
' Second,
' First 7 characters correspond to a trial number in experiment design
' According to the Experiment 1 Design
' Trial # | Stimulus
' 01      | M-v-4-4
' 02      | M-a-4-4
' 03      | M-v-3-4
' 04      | M-a-3-4
' 05      | M-v-2-4
' 06      | M-a-2-4
' 07      | M-v-1-4
' 08      | M-a-1-4
' 09      | M-v-4-3
' 10      | M-a-4-3
' 11      | M-v-3-3
' 12      | M-a-3-3
' 13      | M-v-2-3
' 14      | M-a-2-3
' 15      | M-v-1-3
' 16      | M-a-1-3
' 17      | A-v-4-1
' 18      | A-a-4-1
' 19      | A-v-3-1
' 20      | A-a-3-1
' 21      | A-v-2-1
' 22      | A-a-2-1
' 23      | A-v-1-1
' 24      | A-a-1-1
' 25      | A-v-4-2
' 26      | A-a-4-2
' 27      | A-v-3-2
' 28      | A-a-3-2
' 29      | A-v-2-2
' 30      | A-a-2-2
' 31      | A-v-1-2
' 32      | A-a-1-2
' Where
' - Alyssa_Alltalk_4diff => A
' - Megan_Alltalk_4diff => M
' and
' - VisAsync => v
' - AudSync => a
' and
' - RightTop => 1
' - LeftTop => 2
' - LeftBottom => 3
' - RightBottom => 4
' and
' - -We can..- 19 s1.avi => 1
' - -Good morning.- 19 s.avi => 2
' - --They like to ice-.- 19 s.avi => 3
' - -But..- 19s.avi => 4
'
' Lastly,
' If contains * like M-a-1-3-1-2* then target, else distractor
' - Target => *

' input ouput column can be configured for future usage
Const conditionColumn = 13  ' output simplified condition on column 13
Const stimulusColumn = 12  ' input stimulus name from column 12

Public Sub stimulusToCondition()
    
    Dim stimulusRow As Integer
    Dim stimulus As String
    Dim condition As String

    stimulusRow = 2 ' input stimulus data from row 2

    ' output condition name until the last stimulus in stimulus column
    While Not IsEmpty(Cells(stimulusRow, stimulusColumn))

            stimulus = Cells(stimulusRow, stimulusColumn).Text
            trial = Left(stimulus, 7) ' First 7 characters correspond to a trial number
            aoi = Mid(stimulus, 11, 1) 'Last number (the 11th character) represents AOI

            ' from AOI number representation to letter representation
            If aoi = "0" Then
                condition = "f"
            ElseIf aoi = "1" Then
                condition = "m"
            ElseIf aoi = "2" Then
                condition = "e"
            Else
                condition = "ERROR: wrong aoi"
            End If

            ' from random trial number to standard exp design condition number
            If trial = "M-v-4-4" Then
                condition = condition & "01"
            ElseIf trial = "M-a-4-4" Then
                condition = condition & "02"
            ElseIf trial = "M-v-3-4" Then
                condition = condition & "03"
            ElseIf trial = "M-a-3-4" Then
                condition = condition & "04"
            ElseIf trial = "M-v-2-4" Then
                condition = condition & "05"
            ElseIf trial = "M-a-2-4" Then
                condition = condition & "06"
            ElseIf trial = "M-v-1-4" Then
                condition = condition & "07"
            ElseIf trial = "M-a-1-4" Then
                condition = condition & "08"
            ElseIf trial = "M-v-4-3" Then
                condition = condition & "09"
            ElseIf trial = "M-a-4-3" Then
                condition = condition & "10"
            ElseIf trial = "M-v-3-3" Then
                condition = condition & "11"
            ElseIf trial = "M-a-3-3" Then
                condition = condition & "12"
            ElseIf trial = "M-v-2-3" Then
                condition = condition & "13"
            ElseIf trial = "M-a-2-3" Then
                condition = condition & "14"
            ElseIf trial = "M-v-1-3" Then
                condition = condition & "15"
            ElseIf trial = "M-a-1-3" Then
                condition = condition & "16"
            ElseIf trial = "A-v-4-1" Then
                condition = condition & "17"
            ElseIf trial = "A-a-4-1" Then
                condition = condition & "18"
            ElseIf trial = "A-v-3-1" Then
                condition = condition & "19"
            ElseIf trial = "A-a-3-1" Then
                condition = condition & "20"
            ElseIf trial = "A-v-2-1" Then
                condition = condition & "21"
            ElseIf trial = "A-a-2-1" Then
                condition = condition & "22"
            ElseIf trial = "A-v-1-1" Then
                condition = condition & "23"
            ElseIf trial = "A-a-1-1" Then
                condition = condition & "24"
            ElseIf trial = "A-v-4-2" Then
                condition = condition & "25"
            ElseIf trial = "A-a-4-2" Then
                condition = condition & "26"
            ElseIf trial = "A-v-3-2" Then
                condition = condition & "27"
            ElseIf trial = "A-a-3-2" Then
                condition = condition & "28"
            ElseIf trial = "A-v-2-2" Then
                condition = condition & "29"
            ElseIf trial = "A-a-2-2" Then
                condition = condition & "30"
            ElseIf trial = "A-v-1-2" Then
                condition = condition & "31"
            ElseIf trial = "A-a-1-2" Then
                condition = condition & "32"
            Else
                condition = "ERROR: wrong condition"
            End If

            ' from * to target (t), and non-star to distractor (d)
            If InStr(1, stimulus, "*") Then
                condition = condition & "-t"
            Else
                condition = condition & "-d"
            End If
            
            ' add 1/2/3 for distractor
            Dim row As Integer
            row = stimulusRow - 2 ' start row with 0 for easier computation
            If row Mod 12 = 0 Then
                condition = condition & "1"
            ElseIf row Mod 12 = 1 Then
                condition = condition & "1"
            ElseIf row Mod 12 = 2 Then
                condition = condition & "1"
            ElseIf row Mod 12 = 3 Then
                condition = condition & "2"
            ElseIf row Mod 12 = 4 Then
                condition = condition & "2"
            ElseIf row Mod 12 = 5 Then
                condition = condition & "2"
            ElseIf row Mod 12 = 6 Then
                condition = condition & "3"
            ElseIf row Mod 12 = 7 Then
                condition = condition & "3"
            ElseIf row Mod 12 = 8 Then
                condition = condition & "3"
            End If
            
        'output condition name
        Cells(stimulusRow, conditionColumn).Value = condition
        
        ' move onto the next input row
        stimulusRow = stimulusRow + 1
    Wend

End Sub
