VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPhotoPeriod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'This sub works as a table that loads sunset/sunraise data
'3 options are available:
'        1) A fixed PM and AM value with fixed twilights can be used for all days of the year.
'        2) A value can be calculated internally using the AstroMech module for each day.
'        3) Values for all 365 days can be imported from an external table.  Only the table format from the Naval Observatory is accepted.
'

'Movement Definitions from Beaudreau (2009?)
'1) Dawn: from nautical dawn to sunrise
'2) Day AM: from sunrise to the midpoint between sunrise and sunset
'3) Day PM: from the midpoint between sunrise and sunset to sunset
'4) Dusk, from sunset to nautical dusk
'5) Night: from nautical dusk to nautical dawn.


Const FIXED_TWILIGHT = 60 '1 hour fixed twilight

Dim PhaseTime(TotalDielPhases, 365) As Integer

Dim DayOfYear As Long
Dim UseAstroMech As Boolean
Public Function FindDielPhase(D As Long, T As Integer) As Long
'returns the diel phase of the detection time and day
'
'use enumerated type DielPhase to interpret output

Dim Phase As Long
Dim Found As Boolean
Dim DayIndex As Long
 
'get day of year
DayIndex = ConvertMJD(D)

'first check if it has been initiated for this day
If PhaseTime(DielPhase.Sunset, DayIndex) = -1 Then
    UpdateEntry DayIndex
End If

'compare time to phase times
Do
    If T >= PhaseTime(Phase, DayIndex) And T < PhaseTime(Phase + 1, DayIndex) Then
        'time is within this interval!
        Found = True
    End If
    
    'inc phase
    Phase = Phase + 1
    
Loop Until Phase >= TotalDielPhases Or Found


'if not found within intervals check boundary intervals (intervals located at both extremes)
If Not Found Then
    Phase = 0
Else
    'correct phase number to reflect inc before loop exit
    Phase = Phase + 1
End If

'rts
FindDielPhase = Phase


End Function
Private Sub UpdateEntry(D As Long)
'Updates entry by requesting the information for sunrise/sunset or twilight from the astromech object
'or uses the AM/PM and TWILIGHT variables to set a fix value
'
Dim DayIndex As Long

DayIndex = D

If UseAstroMech Then
    'set day
    AstroMech.SetDay = DayIndex
    
    
    'Update using astromech object
    'request sunrise & sunset
    PhaseTime(DielPhase.Sunrise, DayIndex) = AstroMech.EventTime(AstralFunction.Sunrise)
    PhaseTime(DielPhase.Sunset, DayIndex) = AstroMech.EventTime(AstralFunction.Sunset)
    
    'get midpoint
    PhaseTime(DielPhase.Midpoint, DayIndex) = (PhaseTime(DielPhase.Sunset, DayIndex) + PhaseTime(DielPhase.Sunrise, DayIndex)) / 2
    
    'get twilights
    'use nautical
    PhaseTime(DielPhase.Dawn, DayIndex) = AstroMech.EventTime(AstralFunction.Begin_Nautical_Twilight)
    PhaseTime(DielPhase.Dusk, DayIndex) = AstroMech.EventTime(AstralFunction.End_Nautical_Twilight)
Else
    'use fix values
    PhaseTime(DielPhase.Sunrise, DayIndex) = AM_TH
    PhaseTime(DielPhase.Sunset, DayIndex) = PM_TH
    
    'get midpoint
    PhaseTime(DielPhase.Midpoint, DayIndex) = (PhaseTime(DielPhase.Sunset, DayIndex) - PhaseTime(DielPhase.Sunrise, DayIndex)) / 2
    
    'get twilights
    PhaseTime(DielPhase.Dawn, DayIndex) = AM_TH - FIXED_TWILIGHT
    PhaseTime(DielPhase.Dusk, DayIndex) = PM_TH + FIXED_TWILIGHT
End If

End Sub

Public Function StampIsInPhase(Phase As Long) As Boolean
Dim StartPhase As Long
Dim EndPhase As Long
Dim Criteria_H As Long
Dim Criteria_L As Long
Static CH As Long
Static CL As Long
Static D As Long
Static S As Long

'Phase:
'1) Night: from nautical dusk to nautical dawn
'2) Dawn: from nautical dawn to sunrise
'3) Day AM: from sunrise to the midpoint between sunrise and sunset
'4) Day PM: from the midpoint between sunrise and sunset to sunset
'5) Dusk, from sunset to nautical dusk

'get time interval


StartPhase = Choose(Phase, DielPhase.Dusk, DielPhase.Dawn, DielPhase.Sunrise, DielPhase.Midpoint, DielPhase.Sunset)
EndPhase = Choose(Phase, DielPhase.Dawn, DielPhase.Sunrise, DielPhase.Midpoint, DielPhase.Sunset, DielPhase.Dusk)

If D = Stamp.Date And S = Stamp.Site Then
    Criteria_H = CH
    Criteria_L = CL
Else

    'Fire up astromech calc and get values
    AstroMech.Latitude = Receiver.LA(Stamp.Site)
    AstroMech.Longitude = Receiver.LO(Stamp.Site)
    Criteria_L = ReturnPhaseTime(StartPhase, Stamp.Date)
    Criteria_H = ReturnPhaseTime(EndPhase, Stamp.Date)
    CL = Criteria_L
    CH = Criteria_H
End If

'compare (night is special case)
StampIsInPhase = False
If Criteria_L > Criteria_H Then
    If Stamp.Time <= Criteria_H Or Stamp.Time >= Criteria_L Then
        StampIsInPhase = True
    End If
Else
    If Stamp.Time >= Criteria_L And Stamp.Time <= Criteria_H Then
        StampIsInPhase = True
    End If
End If
 
End Function
Public Property Get ReturnPhaseTime(p As Long, D As Long) As Long
Dim DayIndex As Long

DayIndex = ConvertMJD(D)

'first check if it has been initiated for this day
If PhaseTime(DielPhase.Sunset, DayIndex) = -1 Then
    UpdateEntry DayIndex
End If

'return phase time in memory
If p <= TotalDielPhases Then ReturnPhaseTime = PhaseTime(p, DayIndex)

End Property
Private Function ConvertMJD(D As Long) As Long
Dim NumberOfLeapYearDays As Long
Dim N As Long
Dim Y As Long

Const AVG_DAYS_YR = 365.25

'correct for epoch
N = N + EPOCH
'get year by first calculating leap year days
NumberOfLeapYearDays = Fix(N / (400# * AVG_DAYS_YR)) - Fix(N / (100# * AVG_DAYS_YR)) + Fix(N / (4# * AVG_DAYS_YR))
'get residual days
N = N - NumberOfLeapYearDays

'year
Y = Fix(N / 365#)
'residual is number of days in year
ConvertMJD = N - (Y * 365)

End Function
Public Sub ReadFile(FileName As String)
'Reads File with information to table
'
'File is from http://aa.usno.navy.mil/data/docs/RS_OneYear.php
' SPACE separated file
'Format:
'  [for Day 1 Jan] SR SS, Day 1 Feb SR SS, ...
' Day 2 Jan SR SS, Day 2 Feb SR SS, ...
'TWO SPACES BETWEEN DAYS FROM DIFFERENT MONTHS
'ONE SPACE BETWEEN TIMES
'

Dim FileNumber As Long
Dim T As Long
Dim D As Long
Dim S As String
Dim SpreadSheet(31, 31) As String
Dim c() As String
Dim TimeString() As String
Dim i As Long
Dim Max As Long
Dim column As Long
Dim row As Long
Dim MaxC As Long
FileNumber = FreeFile

Open FileName For Input As #FileNumber

Line Input #FileNumber, S
MaxC = 0

'Read file
Do
    c = Split(S, Space(2))
    If UBound(c) > MaxC Then
        MaxC = UBound(c)
    End If
    
    For i = 0 To UBound(c)
        SpreadSheet(Max, i) = c(i)
    Next i
    Line Input #FileNumber, S
    Max = Max + 1
Loop Until EOF(FileNumber)

Max = i - 1
'file no longer needed
Close #FileNumber

D = 0
'Read each row, column by column
For column = 1 To MaxC
    For row = 0 To Max
        If Val(SpreadSheet(row, column)) > 0 Then
            TimeString = Split(SpreadSheet(row, column), Space(1))
            PhaseTime(DielPhase.Sunrise, D) = ConvertTime(FormatTime(TimeString(0)))
            PhaseTime(DielPhase.Sunset, D) = ConvertTime(FormatTime(TimeString(1)))
            D = D + 1
        End If
    Next row
Next column

End Sub
Private Function FormatTime(T As String)
Dim L As Long
Dim H As String

Const Separator = ":"

L = Len(Trim(T))

'valid?
If L <= 2 Then Exit Function

H = Left(T, L - 2)
'insert :
FormatTime = H & Separator & Right(T, 2)
End Function
Private Sub Class_Initialize()
Dim D As Long

'set all values to default
For D = 0 To 365
    PhaseTime(DielPhase.Sunset, D) = -1
Next D

'default to use of the astromech module
UseAstroMech = True
End Sub
