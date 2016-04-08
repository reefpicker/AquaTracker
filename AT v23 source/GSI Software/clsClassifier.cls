VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClassifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Classifier class
'for AquaTracker
'
'v0.1 by Jose J. Reyes
'
'
'This class will accept any parameter and it will evaluate it and throw it into bins, then count the items in each bin
'it is a simple classifier but it allows the user to define the classification scheme
'
'There are 3 important STATES: STAY, MOVE, and TRANSIENT (neither STAY nor MOVE)
'Moves are incoming or outgoing
'Transient is neither STAY nor MOVE

'To classify the state of the fish, the class accepts a fish number and a STAMP index, it then applies the rules for STATE classification
'and looks ahead or before the index to perform the classification.

'Receiver table can be setup to use either spatial or spatiotemporal information with two arrays:
'Receiver_Data(R): R=receiver number
'Receiver_TimeData(R,T): R=Receiver number, T=1 hour intervals

Const BINS_PER_DAY = 23

Enum State
    Stay = 1
    Move = 2
    Transient = 0
End Enum

Enum SourceArray
    ReceiverData
    ReceiverTimeData
End Enum

Enum Trend
    Up
    High
    Down
    Low
    NoChange
End Enum

Dim Fish_State(MAX_FISH) As Long
Dim SortingBins(10) As BinType
Dim TotalBins As Long
Dim StateDefs(3) As StateDefinition
Dim DATASOURCE As Long
'database of values
Dim Receiver_Data(MAX_RECEIVERS) As Single
Dim Receiver_TimeData(MAX_RECEIVERS, BINS_PER_DAY) As Single
Dim TrendCount(4) As Long
Public Sub LoadReceiverData(FilePath As String)
DATASOURCE = SourceArray.ReceiverData
End Sub
Public Sub LoadReceiverTimeData(FilePath As String)
DATASOURCE = SourceArray.ReceiverTimeData
End Sub
Public Sub AddBin(BinName As String, RuleHigh As Single, RuleLow As Single)
With SortingBins(TotalBins)
    .Name = BinName
    .High = RuleHigh
    .Low = RuleLow
End With
TotalBins = TotalBins + 1
End Sub
Public Sub ClearAll()
Dim i As Long

'clear bins
For i = 0 To TotalBins
    With SortingBins(i)
        .Name = ""
        .High = 0
        .Low = 0
        .Count = 0
    End With
Next i
            
'clear trend bins
For i = 0 To 4
    TrendCount(i) = 0
Next i

TotalBins = 0

End Sub
Public Property Let StateRule_Stay_Time(T As Long)
StateDefs(State.Stay).During = T
End Property
Public Property Let StateRule_Move_BeforeTime(T As Long)
StateDefs(State.Move).Before = T
End Property
Public Property Let StateRule_Move_AfterTime(T As Long)
StateDefs(State.Move).After = T
End Property
Public Function Classify_by_Trend(FishNumber As Long, StampNumber As Long, Optional State As Long = -1) As String
'Classify fish move or stay or transient move based on location information
'
Dim ReceiverNumber As Integer
Dim Value_0 As Single 'before
Dim Value_1 As Single 'actual or center value (value of fish at point in time on request)
Dim Value_2 As Single 'after
Dim BinNumber As Long
Dim TrendingString As String
Dim TimeOfDay As Long

'validate
If TotalBins = 0 Then Exit Function

'get location
FishTable.ReadStamp FishNumber, StampNumber
ReceiverNumber = Stamp.Site

'get location information
'depending on source of data
If DATASOURCE = SourceArray.ReceiverData Then
    Value_1 = Receiver_Data(ReceiverNumber)
    'get before
    If StampNumber > 0 Then
        FishTable.ReadStamp FishNumber, StampNumber - 1
        Value_0 = Receiver_Data(Stamp.Site)
    End If
    
    'get after
    If StampNumber > FishDatabase.NumberOfStamps Then
        FishTable.ReadStamp FishNumber, StampNumber + 1
        Value_2 = Stamp.Site
    End If
Else
    'get stamped time of day
    TimeOfDay = CLng(Stamp.Time / 60)
    Value_1 = Receiver_TimeData(ReceiverNumber, TimeOfDay)
    
    'value before
    If TimeOfDay > 0 Then
        Value_0 = Receiver_TimeData(ReceiverNumber, TimeOfDay - 1)
    Else
        Value_0 = Receiver_TimeData(ReceiverNumber, 23)
    End If
    
    'value after
    If TimeOfDay < 23 Then
        Value_2 = Receiver_TimeData(ReceiverNumber, TimeOfDay + 1)
    Else
        Value_2 = Receiver_TimeData(ReceiverNumber, 0)
    End If
End If

'is the state of the fish important?
If State <> -1 Then
    'get state
    If Fish_State(FishNumber) <> State Then
        Classify_by_Trend = ""
        Exit Function
    End If
End If

'trending up
If Value_0 > Value_1 And Value_1 > Value_2 Then
    TrendingString = "UP"
    TrendCount(Trend.Up) = TrendCount(Trend.Up) + 1
End If

If Value_0 < Value_1 And Value_1 < Value_2 Then
    TrendingString = "DOWN"
    TrendCount(Trend.Down) = TrendCount(Trend.Down) + 1
End If

If Value_0 < Value_1 And Value_1 > Value_2 Then
    TrendingString = "HIGH"
    TrendCount(Trend.High) = TrendCount(Trend.High) + 1
End If

If Value_0 > Value_1 And Value_1 < Value_2 Then
    TrendingString = "LOW"
    TrendCount(Trend.Low) = TrendCount(Trend.Low) + 1
End If

If Value_0 = Value_1 Or Value_1 = Value_2 Then
    If Value_0 > Value_1 Then
        TrendingString = "HIGH"
        TrendCount(Trend.High) = TrendCount(Trend.High) + 1
    End If
    
    If Value_1 > Value_2 Then
        TrendingString = "LOW"
        TrendCount(Trend.Low) = TrendCount(Trend.Low) + 1
    End If
End If

If Value_0 = Value_1 And Value_1 = Value + 2 Then
    TrendingString = "NOCHANGE"
    TrendCount(Trend.NoChange) = TrendCount(Trend.NoChange) + 1
End If


Classify_by_Trend = TrendingString

End Function
Public Function Classify(FishNumber As Long, StampNumber As Long, Optional State As Long = -1) As String
'Classify fish move or stay or transient move based on location information
'
Dim ReceiverNumber As Integer
Dim Value As Single
Dim BinNumber As Long
Dim ClassificationString As String
Dim TimeOfDay As Long

'validate
If TotalBins = 0 Then Exit Function

'get location
FishTable.ReadStamp FishNumber, StampNumber
ReceiverNumber = Stamp.Site

'get location information
'depending on source of data
If DATASOURCE = SourceArray.ReceiverData Then
    Value = Receiver_Data(ReceiverNumber)
Else
    'get stamped time of day
    TimeOfDay = CLng(Stamp.Time / 60)
    Value = Receiver_TimeData(ReceiverNumber, TimeOfDay)
End If

'is the state of the fish important?
If State <> -1 Then
    'get state
    If Fish_State(FishNumber) <> State Then
        Classify = ""
        Exit Function
    End If
End If

'get bin number and count
BinNumber = 0
ClassificationString = ""
Do
    With SortingBins(BinNumber)
        If .Low >= Value And .High < Value Then
            ClassificationString = .Name
            .Count = .Count + 1
            Exit Do
        End If
    End With
    TotalBins = TotalBins + 1
Loop Until BinNumber >= TotalBins

'if a bin is not found, default is to throw it on last bin defined
If BinNumber >= TotalBins And ClassificationString = "" Then
    ClassificationString = SortingBins(TotalBins - 1).Name
End If

Classify = ClassificationString
End Function

