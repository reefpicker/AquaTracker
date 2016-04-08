VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Calculator Class
'Calculates all data needed to be exported or displayed for the Tracking and Export functions
'Copyright (c) 2009 - 2015 by Jose J. Reyes-Tomassini and BE Team Manchester Lab


Dim CodedTrackString(MAX_TRACK_STRING) As Integer

Dim AM_PM_Threshold As Long

Dim Distance As Double
Dim MaxStay_ID As Long
Dim MaxStay_Time As Long

Dim Speed As Single
Dim TTL_Distance As Double
Dim Time_Traveled As Long
Dim Days_Traveled As Long
Dim Cumulative_Speed As Double

Dim Number_Of_Detections As Long
Dim Receivers_Active As Long

Dim LastSite As Long
Dim LastSite_Long As Double
Dim LastSite_Lat As Double
Dim ActiveSite As Long
Dim ActiveSite_Lat As Double
Dim ActiveSite_Long As Double

Dim DayZero As Long
Dim LastSite_Day As Long
Dim ActiveSite_Day As Long
Dim ActiveSite_Time As Long
Dim LastSite_Time As Long

Dim Max_Lat As Double
Dim Max_Long As Double
Dim Min_Lat As Double
Dim Min_Long As Double

Dim Stay As Boolean
Dim StayTime As Long
Dim StayStation As Integer
Dim MaxStayTime As Long
Dim ReportedStay As Integer
Dim ReportedStayTime As Long

Dim NumberOfTimesTagVisited(MAX_TAGS) As Long

Dim PSI_La As Double
Dim PSI_Lo As Double


Dim Visited(MAX_RECEIVERS) As Boolean
Dim Times(MAX_RECEIVERS) As Integer
Dim Ticks(MAX_RECEIVERS) As Integer
Dim Transitions_Total As Long
Dim Transitions_New As Long
Dim TTL_AM_Travel As Long
Dim TTL_PM_Travel As Long
Dim Move_Status As Long
Dim NumberOfMoves As Long


Dim MoveByType(25, 25) As Long
Dim StayByType(25) As Integer
Dim StayTimeThreshold As Long
Dim StayPassedThreshold As Boolean
Dim PrevStaytime As Long
Dim ResidentBeforeDeparture As Boolean
Dim ResidentAfterArrival As Boolean
Dim ValidMove As Boolean

Dim StayIndex As Long
Dim ReceiverStays(MAX_DETECTIONS_PER_FISH) As Long
Private Type TrackParameters
    LA As Single
    LO As Single
    Linearity As Single
    Meandering As Single
    Distance As Single
End Type

Dim NumberOfTracks As Long
Dim Diel As Boolean
Dim Tide As Boolean
Dim DistanceBetweenReceivers(MAX_RECEIVERS, MAX_RECEIVERS) As Double
Public Property Let OnlyCountMovesIfResidentBeforeDeparture(a As Boolean)
ResidentBeforeDeparture = a
End Property
Public Property Let OnlyCountMovesIfResidentAfterArrival(a As Boolean)
ResidentAfterArrival = a
End Property
Public Property Get OnlyCountMovesIfResidentBeforeDeparture() As Boolean
OnlyCountMovesIfResidentBeforeDeparture = ResidentBeforeDeparture
End Property
Public Property Get OnlyCountMovesIfResidentAfterArrival() As Boolean
OnlyCountMovesIfResidentAfterArrival = ResidentAfterArrival
End Property

Public Property Get AnalyzeDielMoves() As Boolean
AnalyzeDielMoves = Diel
End Property
Public Property Let AnalyzeDielMoves(a As Boolean)
Dim i As Long
Dim ii As Long
ReDim K(4) As Variant
ReDim NN(4) As Variant

If a And Not Diel Then
    'key and names for phases
    K = Array("N", "DW", "AM", "PM", "DK")
    NN = Array("Night", "Dawn", "AM", "PM", "Dusk")
    
    'Add fish numbers
    FactorStay.Assign_Field("fid") = "Fish ID"
    FactorMove.Assign_Field("fid") = "Fish ID"
    'store them
    For i = 0 To 4
        'stays
        FactorStay.Assign_Field(K(i) & "") = NN(i) & ""
        For ii = 0 To 4
            'moves
            FactorMove.Assign_Field(K(i) & "->" & K(ii)) = NN(i) & "->" & NN(ii)
        Next ii
    Next i
End If

'turn analysis on!
Diel = a

End Property

Public Property Let AnalyzeTideMoves(a As Boolean)
Tide = a
End Property
Public Sub Reset()
'reset ranges and averages
With Min
    .LA = 9999
    .Linearity = 9999
    .LO = 9999
    .Meandering = 9999
    .Distance = 999999999
End With

With Max
    .LA = 0
    .Linearity = 0
    .LO = 0
    .Meandering = 0
    .Distance = 0
End With

With Average_For_All_Tracks
    .LA = 0
    .Linearity = 0
    .LO = 0
    .Meandering = 0
    .Distance = 0
End With

NumberOfTracks = 0

End Sub
Public Sub Clear()
Dim R As Integer
Dim i As Long
Dim ii As Long

'clear movetypes and staytypes
For i = 0 To 25
    For ii = 0 To 25
        MoveByType(i, ii) = 0
    Next ii
    StayByType(i) = 0
Next i

'clear stay related vars
StayPassedThreshold = False
PrevStaytime = 0
Stay = False
StayTime = 0
StayStation = 0
MaxStayTime = 0
StayIndex = 0

'clear accumulators for receiver "tags"
For i = 0 To MAX_TAGS
    NumberOfTimesTagVisited(i) = 0
Next i

'clear flags
For R = 0 To MAX_RECEIVERS
    Visited(R) = False
Next R

'clear all
PSI_La = 0
PSI_Lo = 0

Receivers_Active = 0
TTL_Distance = 0
DayZero = 0
Cumulative_Speed = 0
Time_Traveled = 0
Number_Of_Detections = 0
ActiveSite = 0
ActiveSite_Long = 0
ActiveSite_Lat = 0
LastSite = 0
LastSite_Long = 0
LastSite_Lat = 0

ActiveSite_Day = 0
LastSite_Day = 0

Max_Lat = 0
Max_Long = 0
Min_Lat = 9999
Min_Long = 9999


Transitions_Total = 0
Transitions_New = 0
TTL_AM_Travel = 0
TTL_PM_Travel = 0
NumberOfMoves = 0
ValidMove = False

End Sub
Public Sub AnalyzeOverlaps(Threshold As Long, NumberOfFish As Long, ListOfOverlaps As ListBox, ListOfPairs As ListBox, ListOfFish As ListBox)
'Entry: Threshold in minutes/km; numberoffish; recepient list of overlap cases; recepients with parsed detector pairs
'

'Algorithm:
'Look for entries with more than 2 detections
'Among these, look for delta ctime <= TH
'mark as a case and send to listbox
'
'Pairs are the receiver pairs.
'for each pair, there is a list of codetected fish

'TTL Number of possible pairs = Number of Receivers ^ 2

Const MAX_PAIRS = 2000


Dim Pair_1(MAX_PAIRS) As Integer
Dim Pair_2(MAX_PAIRS) As Integer
Dim Count(MAX_PAIRS) As Long
Dim Fishes(MAX_PAIRS) As String
Dim Speed As Single
Dim Detections As Long
Dim Fish As Integer
Dim DetectionStamps As Long
Dim DateBefore As Single
Dim DateAfter As Single
Dim TimeBefore As Integer
Dim TimeAfter As Integer
Dim SiteBefore As Integer
Dim SiteAfter As Integer
Dim CTimeAfter As Long
Dim CTimeBefore As Long

Dim NumberOfPairs As Long
Dim Index As Long
Dim Found As Boolean
Dim s As String
Dim response As Variant


For Fish = 0 To NumberOfFish - 1
    Detections = FishDatabase.NumberOfStamps(Fish)
    If Detections > 1 Then
        FishTable.ReadStamp Fish, 0
        If Stamp.Valid Then
            DateAfter = Stamp.Date
            TimeAfter = Stamp.Time
            SiteAfter = Stamp.Site
            CTimeAfter = Stamp.CTime
            
            For DetectionStamps = 1 To Detections - 1
                DateBefore = DateAfter
                TimeBefore = TimeAfter
                CTimeBefore = CTimeAfter
                SiteBefore = SiteAfter
                
                FishTable.ReadStamp Fish, DetectionStamps
                DateAfter = Stamp.Date
                TimeAfter = Stamp.Time
                SiteAfter = Stamp.Site
                CTimeAfter = Stamp.CTime
                If SiteBefore <> SiteAfter And CTimeAfter - CTimeBefore <= Threshold Then
                    Index = 0
                    Found = False
                    Do
                        If (Pair_1(Index) = SiteBefore And Pair_2(Index) = SiteAfter) Or (Pair_1(Index) = SiteAfter And Pair_2(Index) = SiteBefore) Then
                            Count(Index) = Count(Index) + 1
                            If InStr(1, Fishes(Index), FishDatabase.Code(Stamp.Fish)) < 1 Then
                                Fishes(Index) = Fishes(Index) & ", " & FishDatabase.Code(Stamp.Fish)
                            End If
                            Found = True
                        End If
                        Index = Index + 1
                    Loop While (Index < NumberOfPairs) And Not Found
                    If Not Found Then
                        Pair_1(NumberOfPairs) = SiteBefore
                        Pair_2(NumberOfPairs) = SiteAfter
                        Count(NumberOfPairs) = 1
                        Fishes(Index) = FishDatabase.Code(Stamp.Fish)
                        NumberOfPairs = NumberOfPairs + 1
                        If NumberOfPairs >= MAX_PAIRS Then
                            'throw exception
                            response = MsgBox("The number of overlaps is too high! Analysis aborted premarturely.  Consult with manual for ways to workaround this issue.", vbOKOnly, "Warning: Program limit exceeded")
                            Exit For
                        End If
                    End If
                End If
            Next DetectionStamps
        End If
    End If
Next Fish

For Index = 0 To NumberOfPairs
    If Pair_1(Index) > 0 And Pair_2(Index) > 0 Then
        s = Receiver.ID(Pair_1(Index)) & " overlaps with " & Receiver.ID(Pair_2(Index)) & "(" & Str$(Count(Index)) & " times)"
        ListOfOverlaps.AddItem s
        ListOfPairs.AddItem Str$(Pair_1(Index)) & "-" & Str$(Pair_2(Index))
        ListOfFish.AddItem Fishes(Index)
    End If
Next Index
End Sub

Public Property Let Site(i As Long)
Dim R As Integer
Dim Found As Boolean
Dim L As Long
Dim TagID As Long


If ActiveSite <> 0 Then
    'Active site is now last site
    LastSite = ActiveSite
    LastSite_Long = ActiveSite_Long
    LastSite_Lat = ActiveSite_Lat
End If

'This is the new site
ActiveSite = i

'redirect if group
If Receiver.GroupNumber(ActiveSite) Then
    ActiveSite = Receiver.GroupReceiverNumber(Receiver.GroupNumber(ActiveSite))
End If

ActiveSite_Long = Receiver.LO(CByte(ActiveSite))
ActiveSite_Lat = Receiver.LA(CByte(ActiveSite))




'Calculate bounding box as it expands
'max andmin
If ActiveSite_Long > Max_Long Then
    Max_Long = ActiveSite_Long
End If

If ActiveSite_Long < Min_Long Then
    Min_Long = ActiveSite_Long
End If

If ActiveSite_Lat > Max_Lat Then
    Max_Lat = ActiveSite_Lat
End If

If ActiveSite_Lat < Min_Lat Then
    Min_Lat = ActiveSite_Lat
End If

'Calculate Distance
Distance = Calculate_Distance_Between_Sites(LastSite, ActiveSite)
TTL_Distance = TTL_Distance + Distance

'Find if this receiver has been visited
'or flag as visited!
If Not Visited(ActiveSite) Then
    Receivers_Active = Receivers_Active + 1
    Transitions_New = Transitions_New + 1
    Visited(ActiveSite) = True
Else
    Visited(ActiveSite) = True
End If

'check if at same station
Stay = False
If ActiveSite = LastSite Then
    Stay = True
Else
    'count transitioning
    Transitions_Total = Transitions_Total + 1
    StayIndex = StayIndex + 1
    'save info
    If StayTime >= StayTimeThreshold Then SaveStayInformation
    Number_Of_Detections = Number_Of_Detections + 1
    PSI_La = PSI_La + ActiveSite_Lat
    PSI_Lo = PSI_Lo + ActiveSite_Long
End If


If StayPassedThreshold And Not Stay Then
    PrevStaytime = StayTime
    StayPassedThreshold = False
End If

'erase counter if previous was a stay and this one is a move
If Not Stay And StayTime > 0 Then StayTime = 0


'check and count tags
TagID = Receiver.Tag(ActiveSite)
NumberOfTimesTagVisited(TagID) = NumberOfTimesTagVisited(TagID) + 1

End Property
Public Property Get PercentVisitsToTag(TagID As Long) As Single
Dim i As Long
Dim ttl As Long

'count all visits to tags
For i = 0 To Receiver.NumberOfTags
    ttl = ttl + NumberOfTimesTagVisited(i)
Next i

'no /0 allowed!
If ttl > 0 Then
    PercentVisitsToTag = CSng(NumberOfTimesTagVisited(TagID) / ttl)
End If

End Property
Public Property Get Percent_Receivers_Active() As Single
Dim ttl As Long
Dim result As Single

ttl = Receiver.TotalReceivers

If ttl > 0 Then
    result = Receivers_Active / ttl
End If
Percent_Receivers_Active = result

End Property
Private Sub SaveStayInformation()
'comp to max stay time
If StayTime > MaxStayTime And StayTime > StayTimeThreshold Then
    ReportedStayTime = StayTime
    ReportedStay = LastSite
    MaxStayTime = StayTime
End If
End Sub
Public Property Get Total_Displacement() As Double
Total_Displacement = TTL_Distance

End Property
Public Property Get Displacement_Rate() As Double
'same as speed
'm/h
If Time_Traveled > 0 Then
    Displacement_Rate = (TTL_Distance * 1000) / (Time_Traveled / 60)
End If

End Property
Public Property Get Distance_From_Last_Site() As Double
Distance_From_Last_Site = Distance
End Property
Public Property Get Instant_Speed() As Single
Instant_Speed = Speed

End Property
Public Property Let Day(d As Long)
'store first day
If LastSite_Day = 0 And ActiveSite_Day = 0 Then
    DayZero = d
End If

'save last
LastSite_Day = ActiveSite_Day

'load new
'day # for this site
ActiveSite_Day = d

'calculate
Days_Traveled = DayZero - d

End Property
Public Property Get Meandering_Index()

If Transitions_Total > 0 Then Meandering_Index = Transitions_New / Transitions_Total Else Meandering_Index = 0

End Property
Public Property Get Total_Time_Traveled() As Single
'time travel in fractional days
Total_Time_Traveled = Time_Traveled / 1440
End Property
Public Property Let Time(t As Long)
LastSite_Time = ActiveSite_Time
ActiveSite_Time = t
End Property
Public Property Get Total_Days_Traveled() As Long
Total_Days_Traveled = Days_Traveled
End Property
Public Sub Calculate()
Dim dummy
dummy = Time_Elapsed_Between_Detections

'get the cumulative values for averages and the min/max for the ranges
'

End Sub
Public Property Get PercentTravelByCategory(C1 As Long, C2 As Long) As Single
'Returns Percent travel time during phase/category transition C1->C2
'percent is from cumulative travel time

If NumberOfMoves > 0 Then
    PercentTravelByCategory = (MoveByType(C1, C2) / NumberOfMoves) * 100
    
End If

End Property
Public Property Get PercentResidenceByCategory(C1 As Long) As Single
'Returns Percent Residence time during phase/category C1
'percent is from cumulative travel time

If StayTime > 0 Then
    PercentResidenceByCategory = (StayByType(C1) / StayTime) * 100
End If

End Property

Public Function Time_Elapsed_Between_Detections() As Long

Dim TE As Long
Dim Days_in_Minutes As Long
Dim AM As Long
Dim PM As Long

'do not perform calculations if on first site!
If LastSite_Day = 0 Then Exit Function

'calculate
Days_in_Minutes = (ActiveSite_Day - LastSite_Day) * 1440

TE = Days_in_Minutes + Abs(ActiveSite_Time - LastSite_Time)

If TE > 0 Then
    Speed = Distance / TE
Else
    Speed = 0
End If

'stay?
If Stay Then
    CategorizeStay TE
Else
    CategorizeMove
End If

'accumulate time between detections as time traveled
Time_Traveled = Time_Traveled + TE
Time_Elapsed_Between_Detections = TE

End Function
Private Sub CategorizeStay(TE As Long)
'establishes when the stay occurred
'any abiotic factor in analysis is considered
'
Dim Category As Long
Dim response As Variant

Static LastFishToRaiseError As Long

Const RETURN_RECEIVER_NAME_ONLY = True
On Error GoTo RaiseError

'accumulate staytime
StayTime = StayTime + TE

If StayTime >= StayTimeThreshold And Not StayPassedThreshold Then
    StayPassedThreshold = True
    ReceiverStays(StayIndex) = StayTime
    'Diel/Photoperiod
    If Diel Then
        Category = PhotoPeriodCalculator.FindDielPhase(ActiveSite_Day, CInt(ActiveSite_Time))
    End If
    
    'tides
    If Tide Then
    End If
    
     'now count
    StayByType(Category) = StayByType(Category) + 1
    
    'count if prev. move
    If ResidentAfterArrival And StayIndex > 0 Then
        If ResidentBeforeDeparture Then
            If ReceiverStays(StayIndex) > StayTimeThreshold Then
            End If
        End If
    End If
    
    
Else
    'set to 0 if not above TE
    If StayIndex > MAX_DETECTIONS_PER_FISH Then GoTo RaiseError
    ReceiverStays(StayIndex) = 0
End If

Exit Sub

RaiseError:

If CURRENT_FISH = LastFishToRaiseError Then Exit Sub
response = MsgBox("Warning: Fish detections too many.  Can't calculate stay statistics for track " & CURRENT_FISH & frmFloater.cmbFishCode.List(CInt(CURRENT_FISH)), vbOKOnly, "Warning")
LastFishToRaiseError = CURRENT_FISH

End Sub

Private Sub CategorizeMove()
'establishes when the move occurred
'any abiotic factor in analysis is considered
'
Dim ToCategory As Long
Dim FromCategory As Long

'total number of moves
NumberOfMoves = NumberOfMoves + 1

'Diel/Photoperiod
If Diel Then
    ToCategory = PhotoPeriodCalculator.FindDielPhase(ActiveSite_Day, CInt(ActiveSite_Time))
    FromCategory = PhotoPeriodCalculator.FindDielPhase(LastSite_Day, CInt(LastSite_Time))
End If

'now count
'according to move definitions

If ResidentBeforeDeparture And StayIndex > 0 Then
    If ReceiverStays(StayIndex - 1) > StayTimeThreshold And Not ResidentAfterArrival Then
        MoveByType(FromCategory, ToCategory) = MoveByType(FromCategory, ToCategory) + 1
    End If
Else

'Else
 '   If ResidentBeforeDeparture And ResidentAfterArrival Then
    
End If

End Sub
Public Property Get Range_Bounding_Box() As Double
Range_Bounding_Box = GetRange

End Property
Private Function GetRange() As Double
If (Max_Lat > 0 Or Min_Lat > 0) And (Max_Long > 0 Or Min_Long > 0) Then
    GetRange = Calculate_Distance(Max_Long, Min_Long, Max_Lat, Min_Lat)
End If

End Function
Public Property Get Linearity() As Double
Dim range As Double
Dim t As Double
range = GetRange

If range > 0 Then
    t = TTL_Distance / range
End If

If t <> 0 Then Linearity = 1 / t Else Linearity = 0

End Property
Public Property Get Residence_Time_Days() As Long
Dim t As Long
End Property
Public Property Get Longest_Stay_Duration() As Long
'duration in hrs
Longest_Stay_Duration = ReportedStayTime / 60
End Property
Public Property Get Longest_Stay_Station() As String
Longest_Stay_Station = Receiver.ID(ReportedStay)
End Property
Public Property Get Path_Similarity_Index_Lo(Optional Switch As Long = 0) As Double
Dim result As Double

If Number_Of_Detections > 0 Then
    result = PSI_Lo / Number_Of_Detections
End If

'get x/y coord instead of la/lo
If Switch Then
    result = (Origin_Long - result) / Scale_X
End If

Path_Similarity_Index_Lo = result
End Property
Public Property Get Path_Similarity_Index_La(Optional Switch As Long = 0) As Double
Dim result As Double

If Number_Of_Detections > 0 Then
    result = PSI_La / Number_Of_Detections
End If

'get x/y coord instead of la/lo
If Switch Then
    result = (Origin_Lat - result) / Scale_Y
End If

Path_Similarity_Index_La = result
End Property
Private Function Calculate_Distance_Between_Sites(Site_A As Long, Site_B As Long) As Double
'Calculates distance between two sites
'
Dim Site_A_Long As Double
Dim Site_B_Long As Double
Dim Site_A_Lat As Double
Dim Site_B_Lat As Double
Dim Distance As Double

Const LAND_FOUND = True

'first check if this is a stay (distance will then = 0)
'also if this is the first move
If Site_A = Site_B Or Site_A = 0 Or Site_B = 0 Then Exit Function

'any land between sites?
If Land_Avoidance = AvoidanceMode.NotActive Or (Receiver.LandLocked(Site_A) Or Receiver.LandLocked(Site_B)) Then
    'get lats and longs for each site
    Site_A_Long = Receiver.LO(CByte(Site_A))
    Site_A_Lat = Receiver.LA(CByte(Site_A))
    Site_B_Long = Receiver.LO(CByte(Site_B))
    Site_B_Lat = Receiver.LA(CByte(Site_B))
    Distance = Calculate_Distance(Site_A_Long, Site_B_Long, Site_A_Lat, Site_B_Lat)
Else
    If Receiver.NavigateBetweenSites(Site_B, Site_A) = LAND_FOUND Then
        Distance = Receiver.RouteDistance(Site_B, Site_A)
    Else
        'get lats and longs for each site
        Site_A_Long = Receiver.LO(CByte(Site_A))
        Site_A_Lat = Receiver.LA(CByte(Site_A))
        Site_B_Long = Receiver.LO(CByte(Site_B))
        Site_B_Lat = Receiver.LA(CByte(Site_B))
        Distance = Calculate_Distance(Site_A_Long, Site_B_Long, Site_A_Lat, Site_B_Lat)
    End If
End If

DistanceBetweenReceivers(Site_A, Site_B) = Distance

Calculate_Distance_Between_Sites = Distance

End Function

Public Function Calculate_DistanceXY(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long) As Double
Dim LA As Double
Dim LO As Double
Dim La2 As Double
Dim Lo2 As Double
Dim Distance As Double

LA = Origin_Lat - (Y1 * Scale_Y)
LO = Origin_Long - (X1 * Scale_X)
La2 = Origin_Lat - (Y2 * Scale_Y)
Lo2 = Origin_Long - (X2 * Scale_X)

Distance = Calculate_Distance(LO, Lo2, LA, La2)

Calculate_DistanceXY = Distance

End Function
Public Function Calculate_Distance(Longitude As Double, longitude2 As Double, Latitude As Double, latitude2 As Double) As Double
Dim X As Double
Dim d As Double
Dim INTER As Double
Dim R As Double

On Error GoTo ExitNow

Const N = 57.2958

d = 0
R = 6378.7 'kilometers

INTER = Sin(Latitude / N) * Sin(latitude2 / N)

X = INTER + (Cos(Latitude / N) * Cos(latitude2 / N) * Cos(longitude2 / N - Longitude / N))

If X ^ 2 = 1 Then
    d = 0
Else
    d = R * Atn(Sqr(1 - X ^ 2) / X)
End If


ExitNow:
Calculate_Distance = d
End Function
Public Property Get ReceiverInTrackString(Position As Integer) As Integer
If Position <= MAX_TRACK_STRING Then ReceiverInTrackString = CodedTrackString(Position)
End Property
Public Function CreateVerboseTrackString(FishNumber As Long) As String
Dim Lag_0 As String
Dim Lag_1 As String
Dim Lag_2 As String
Dim R0 As Integer
Dim R1 As Integer
Dim R2 As Integer
Dim R(MAX_TRACK_STRING) As Integer
Dim L(MAX_TRACK_STRING) As Integer
Dim m(MAX_TRACK_STRING) As Integer
Dim N As Long
Dim i As Long
Dim ii As Long
Dim s As Long
Dim p As Long
Dim TrackString As String
Dim Count As Long
Dim Bidirectional As Boolean
Dim Prev As String
Dim RPrev As Integer
Dim FAILED As Boolean
Dim LastStamp As Long

On Error GoTo ExitWithError

'validate
If FishNumber < 0 Then Exit Function

'Get FishNumber
FishDatabase.Fish = FishNumber
LastStamp = FishDatabase.NumberOfStamps(CInt(FishNumber)) - 1
'prefetch first VALID stamp
Do
    FishTable.ReadStamp FishNumber, s
    s = s + 1
Loop While (Not Stamp.Valid) And s <= LastStamp

Lag_0 = Receiver.ID(Stamp.Site)
TrackString = Lag_0
R0 = Stamp.Site
L(0) = Len(Lag_0)
R(0) = R0
N = 0

'Read stamps
For i = s To LastStamp
    FishTable.ReadStamp FishNumber, i
    'truncate if too long
    If Len(TrackString) > MAX_TRACK_STRING Then
        TrackString = TrackString & "::Truncated (too long)::"
        Exit For
    End If
    
    If Stamp.Valid Then
        Prev = Lag_0
        RPrev = R0
        R0 = Stamp.Site
        Lag_0 = Receiver.ID(R0)
        'moved?
        If Lag_0 <> Prev Then
            'store move
            Lag_2 = Lag_1
            Lag_1 = Prev
            R2 = R1
            R1 = RPrev
            'did it just came from there, and its going back?
            If Lag_0 = Lag_2 And Lag_2 <> "" Then
                If Bidirectional Then
                    Count = Count + 1
                Else
                    TrackString = ChangeToBidirectional(Lag_0, Lag_1, TrackString)
                    m(N - 1) = -2
                    Bidirectional = True
                    Count = Count + 1
                End If
            Else
                If Bidirectional Then
                    If (Count Mod 2) = 0 Then
                        TrackString = TrackString & " --> " & Lag_0
                        N = N + 1
                        R(N) = R0
                        L(N) = Len(Lag_0)
                        m(N - 1) = -1
                    Else
                        TrackString = TrackString & " --> " & Lag_1 & " --> " & Lag_0
                        N = N + 1
                        R(N) = R1
                        L(N) = Len(Lag_1)
                        m(N - 1) = -1
                        N = N + 1
                        R(N) = R0
                        L(N) = Len(Lag_0)
                        m(N - 1) = -1
                    End If
                    Bidirectional = False
                    Count = 0
                Else
                    TrackString = TrackString & " --> " & Lag_0
                    N = N + 1
                    R(N) = R0
                    L(N) = Len(Lag_0)
                    m(N - 1) = -1
                End If
            End If
        End If
    End If
Next i

'code it!
'first receiver
For i = 1 To L(0)
    CodedTrackString(p) = R(0)
    p = p + 1
Next i
''rest of receivers
For i = 1 To N
    For ii = 1 To 5
        CodedTrackString(p) = m(i - 1)
        p = p + 1
    Next ii
    For ii = 1 To L(i)
        CodedTrackString(p) = R(i)
        If p + 5 > MAX_TRACK_STRING Then
            FAILED = True
            Exit For
        Else
            p = p + 1
        End If
    Next ii
    If FAILED Then Exit For
Next i

If FAILED Then GoTo ExitWithError
CreateVerboseTrackString = TrackString
Exit Function

ExitWithError:
TrackString = TrackString & "Truncated (too long)"
CreateVerboseTrackString = TrackString
End Function
Private Function ChangeToBidirectional(L0 As String, L1 As String, s As String) As String
Dim L As Long
Dim p As Long
Dim ll As Long
Dim result As String

If s = "" Or L0 = "" Then Exit Function

'get length of track string
L = Len(s)
'get length of receiver id
ll = Len(L1)

'change to arrow from - symbol
result = s
Mid$(result, L - ll - 3) = "<"

ChangeToBidirectional = result
End Function
Friend Sub Average()
'calculate
'AverageForAllTracks

If NumberOfTracks = 0 Then Exit Sub

With Average_For_All_Tracks
    .LA = .LA / NumberOfTracks
    .LO = .LO / NumberOfTracks
    .Linearity = .Linearity / NumberOfTracks
    .Meandering = .Meandering / NumberOfTracks
    .Distance = .Distance / NumberOfTracks
End With

'display track averages on combo box
With frmFloater
    .txtLinearity = Format(Average_For_All_Tracks.Linearity, "##.###")
    .txtAvgLocation = Format(Average_For_All_Tracks.LA, "##.##") & "," & Format(Average_For_All_Tracks.LO, "##.##")
    .txtTTLDistance = Format(Average_For_All_Tracks.Distance, "##.###")
    .txtMeandering = Format(Average_For_All_Tracks.Meandering, "##.###")
End With

End Sub
Public Sub ComputeSummary()

With Average_For_All_Tracks
    .LA = Path_Similarity_Index_La + .LA
    .LO = Path_Similarity_Index_Lo + .LO
    .Linearity = Linearity + .Linearity
    .Meandering = Meandering_Index + .Meandering
    .Distance = TTL_Distance + .Distance
End With

If Linearity > Max.Linearity Then Max.Linearity = Linearity
If Linearity < Min.Linearity Then Min.Linearity = Linearity
If Meandering_Index > Max.Meandering Then Max.Meandering = Meandering_Index
If Meandering_Index < Min.Meandering Then Min.Meandering = Meandering_Index
If TTL_Distance > Max.Distance Then Max.Distance = TTL_Distance
If TTL_Distance < Min.Distance Then Min.Distance = TTL_Distance
If Max_Lat > Max.LA Then Max.LA = Max_Lat
If Min_Lat < Min.LA Then Min.LA = Min_Lat
If Max_Long > Max.LO Then Max.LO = Max_Long
If Min_Long < Min.LO Then Min.LO = Min_Long


'update ticker
NumberOfTracks = NumberOfTracks + 1

End Sub
Public Sub ComputeResidence(TimeAtReceiver() As Long, ByVal FishNumber As Integer)

'Returns time at each receiver

Dim TotalResidence As Long
Dim GN As Integer
Dim i As Long
Dim s As Integer
Dim GroupReceiverNumber As Integer
Dim TimeDetected(MAX_RECEIVERS) As Long
Dim DeltaTime As Long
Dim MovedAway As Boolean
Dim ps As Integer
'Const ResidenceThreshold = 60


s = -1
'visited receivers
For i = 0 To FishDatabase.NumberOfStamps(FishNumber) - 1
    FishTable.ReadStamp FishNumber, i
    If Stamp.Valid Then
        s = Stamp.Site
        GN = Receiver.GroupNumber(s)
        'assign to fake receiver
        If GN <> 0 Then s = Receiver.GroupReceiverNumber(GN)
        'check if moved
        If ps = s Then MovedAway = False Else MovedAway = True
        
        If TimeDetected(s) <> 0 Then
            DeltaTime = Stamp.CTime - TimeDetected(s)
            If DeltaTime <= ResidenceThreshold And Not MovedAway Then TimeAtReceiver(s) = TimeAtReceiver(s) + DeltaTime
        End If
        TimeDetected(s) = Stamp.CTime
        ps = s
    End If
Next i
End Sub
Private Function GetTimeBetweenPings(ByVal R As Integer, ByVal FishNumber As Integer, TimeBetweenPings() As Long) As Long
Dim G As Integer
Dim TotalPairs As Long

If Receiver.GroupNumber(R) Then
    TotalPairs = GetTimeBetweenPings_Group(R, FishNumber, TimeBetweenPings())
Else
    TotalPairs = GetTimeBetweenPings_Receiver(R, FishNumber, TimeBetweenPings())
End If

GetTimeBetweenPings = TotalPairs

End Function
Private Function GetTimeBetweenPings_Receiver(ByVal R As Integer, ByVal FishNumber As Integer, TimeBetweenPings() As Long) As Long
Dim i As Long
Dim Count As Long
Dim FirstTime As Long
Dim SecondTime As Long
Dim TotalTime As Long
Dim TotalPings As Long
Dim Done As Boolean

'get total pings
TotalPings = Receiver.Detection_TTL(R)

'start at stamp 1
i = 0
'preconds
FirstTime = -1
SecondTime = -1
Do
    'preconds
    Done = False
    'get OVERLAPPING pairs
    Do
        Do
            i = i + 1
            Receiver.ReadStamp R, i
        Loop Until Stamp.Valid Or i = TotalPings
        'right fish?
        If Stamp.Fish = FishNumber Then
            'get first or second of the pair
            If FirstTime = -1 Then
                FirstTime = Stamp.CTime
            Else
                SecondTime = Stamp.CTime
                Done = True
            End If
        End If
    Loop Until Done Or i = TotalPings
    'store pairs. if lone ping, then no comp.
    If Done Then
        TimeBetweenPings(Count) = SecondTime - FirstTime + 1
        Count = Count + 1
        FirstTime = SecondTime
        SecondTime = -1
    End If
Loop Until i = TotalPings

GetTimeBetweenPings_Receiver = Count

End Function
Private Function GetTimeBetweenPings_Group(R As Integer, FishNumber As Integer, TimeBetweenPings() As Long) As Long
Dim i As Long
Dim Count As Long
Dim FirstTime As Long
Dim SecondTime As Long
Dim TotalTime As Long
Dim TotalPings As Long
Dim TotalFishPings As Long
Dim Done As Boolean
Dim SizeOfTable(1) As Long

'get total pings
TotalPings = Receiver.Detection_TTL(R)



'"GroupTable" is a special table used to temporarily sort and read stamps from groups
'By passing totalpings as argument and using only one entry (0), it allows me to write the table
'however, we wont be passing the entire stamps to the table, just those belonging to the fish
'this was a quick solution that allowed me to use the database I have already built without modification
'but this is a memory drain.  For smaller systems, the better solution is to create a copy of clsTable, change the name to
'clsMiniTable, and change the MAX_ENTRY for the class itself to be a smaller number.... Just sayin...

'Partition group table
SizeOfTable(0) = TotalPings

GroupTable.PartitionTable SizeOfTable()

'read fish stamps into it
For i = 0 To TotalPings - 1
    Receiver.ReadStamp R, i
    If Stamp.Fish = FishNumber And Stamp.Valid Then
        GroupTable.WriteStamp 0, TotalFishPings
        TotalFishPings = TotalFishPings + 1
    End If
Next i

'if 0 or only 1 stamp, no residence in this receiver!
If TotalFishPings <= 1 Then Exit Function

SizeOfTable(0) = TotalFishPings

GroupTable.PartitionTable SizeOfTable()
'sort! (Will sort when last stamp is written!)
GroupTable.WriteStamp 0, TotalFishPings

'start at stamp 1
i = 1
'preconds
FirstTime = -1
SecondTime = -1
Do
    'preconds
    Done = False
    'get OVERLAPPING pairs
    Do
        GroupTable.ReadStamp 0, i
        'right fish?
        'get first or second of the pair
        If Stamp.Fish = FishNumber Then
            'get first or second of the pair
            If FirstTime = -1 Then
                FirstTime = Stamp.CTime
            Else
                SecondTime = Stamp.CTime
                Done = True
            End If
        End If
        i = i + 1
    Loop Until Done Or i > TotalFishPings
    'store pairs. if lone ping, then no comp.
    If Done Then
        TimeBetweenPings(Count) = SecondTime - FirstTime + 1
        Count = Count + 1
        FirstTime = SecondTime
        SecondTime = -1
    End If
Loop Until i > TotalFishPings

GetTimeBetweenPings_Group = Count

End Function
