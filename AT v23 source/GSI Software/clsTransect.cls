VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTransect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Const DETECTIONS_PER_TRANSECT = 300

Private Type TransectStamp
    ReceiverAssigned As Long
    FishID As Long
    Direction As Single
    CTime As Long
End Type

Private Type TransectRouteType
    Name As String
    Waypoints(100) As ExternalWayPointStamp
    TotalWayPoints As Long
    CTime As Long
    TotalDetections As Long
    Detections(DETECTIONS_PER_TRANSECT) As TransectStamp
End Type
Dim TransectNameTable(MAX_TRANSECTS) As String
Dim TransectRoute(MAX_TRANSECTS) As TransectRouteType
Public Property Get Name(T As Long) As String
Name = TransectRoute(T).Name
End Property
Public Property Get NumberOfTransectsLoaded() As Long
Dim TransectNumber As Long
Dim LastOne As Boolean

TransectNumber = -1

Do
    TransectNumber = TransectNumber + 1
    If TransectRoute(TransectNumber).TotalDetections = 0 Then
        LastOne = True
    End If
Loop Until LastOne Or (TransectNumber > MAX_TRANSECTS)

NumberOfTransectsLoaded = TransectNumber - 1
End Property
Public Sub AddStamp(Code As Long, Direction As Single, TransectNumber As Long, ReceiverNumber As Long, CTimeString As String)
'adds a detection (as a receiver) to the table
Dim TotalDetections As Long
Dim T() As String

'convert time in stamp to Ctime, put back in stamp
T = Split(CTimeString)

With TransectRoute(TransectNumber)
    TotalDetections = .TotalDetections
    With .Detections(TotalDetections)
        .CTime = ConsolidateTime(DayNumber(T(0)), ConvertTime(T(1)))
        .ReceiverAssigned = ReceiverNumber
        .Direction = Direction
        .FishID = Code
    End With
    .TotalDetections = .TotalDetections + 1
End With
End Sub
Public Property Get TotalDetections(Name As String) As Long
Dim TransectNumber As Long
TransectNumber = AssignTo(Name)
TotalDetections = TransectRoute(TransectNumber).TotalDetections
End Property
Public Function AssignTo(Name As String) As Long
'Creates a table entry then returns number
Dim Success As Boolean
Dim TransectNumber As Long
Dim TableEntry As String
Success = False
TransectNumber = -1
Do
    TransectNumber = TransectNumber + 1
    TableEntry = UCase(TransectNameTable(TransectNumber))
    If TableEntry = UCase(Name) Then
        Success = True
    End If
Loop While (Not Success) And (TransectNumber < MAX_TRANSECTS) And TableEntry <> ""

If Success Then
    AssignTo = TransectNumber
Else
    If TransectNumber <= MAX_TRANSECTS Then
        TransectNameTable(TransectNumber) = Name
        TransectRoute(TransectNumber).Name = Name
        AssignTo = TransectNumber
    End If
End If

End Function

Public Sub AddWayPoint()
Dim WayPoint_count As Long
Dim TransectNumber As Long
'Since waypoints are not as numerous, this program uses insertion sorting
'We calculate the CTime and then based on that, we compare to the last point inserted in table, then push one before or after up or down....
Dim T() As String
Dim Response As Variant

On Error GoTo ExitWithErrors

'get number by attempting to create,
'if already exist, create func should return number
TransectNumber = AssignTo(WayPointStamp.Name)

'make sure that dates are in order, otherwise, do an insert sort
'
'convert time in stamp to Ctime, put back in stamp
T = Split(WayPointStamp.DayTimeString)
WayPointStamp.CTime = ConsolidateTime(DayNumber(T(0)), ConvertTime(T(1)))

'expected order?
With TransectRoute(TransectNumber)
    WayPoint_count = .TotalWayPoints
    'first waypoint or not?
    If WayPoint_count = 0 Then
        InsertAtBottom TransectNumber
        .TotalWayPoints = WayPoint_count + 1
    Else
        If .Waypoints(WayPoint_count - 1).CTime >= WayPointStamp.CTime Then
            InsertAtBottom TransectNumber
            .TotalWayPoints = WayPoint_count + 1
        Else
            'insert-sort
            InsertSort TransectNumber
        End If
    End If
    'add point to ttl
    
End With

Exit Sub

ExitWithErrors:
Response = MsgBox("Error: Wrong transect stamp format", vbOKOnly, "Import File Error")

End Sub
Private Sub InsertSort(TransectNumber As Long)
'Performs a sorted insert
'uses the classic method of first starting comparison at the middle, then moves up or down depending on comparison results
'the inserted data shifts down the rest of the data
'data moves are costly, so this is not very efficient but the stamp data is expected to be mostly in order
'

Dim Count As Long
Dim Middle As Long
Dim Direction As Long
Dim Highest As Long
Dim InsertPosition As Long
Dim CTime As Long
Dim Found As Boolean
Dim First As Long
Dim Goal As Long

'get last item in table
Highest = TransectRoute(TransectNumber).TotalWayPoints

'first is 0
First = 0

'get time to compare
CTime = WayPointStamp.CTime

'search until item(position)<item to insert, then insert item at position+1

'this version of the algorithm goes one by one:
Direction = -1
Goal = First
InsertPosition = Highest
Do
    InsertPosition = InsertPosition + Direction
    If TransectRoute(TransectNumber).Waypoints(InsertPosition).CTime <= CTime Then
        Found = True
    End If
Loop Until Found Or InsertPosition = Goal

'insert on insertposition+1, so push down items from insertposition+1 to end
'start on the bottom
For Count = Highest To InsertPosition + 1 Step -1
    With TransectRoute(TransectNumber)
        .Waypoints(Count + 1) = .Waypoints(Count)
    End With
Next Count

'insert
'uses external waypoint stamp to transfer data to this object
With TransectRoute(TransectNumber).Waypoints(InsertPosition + 1)
    .Lat = WayPointStamp.Lat
    .Lon = WayPointStamp.Lon
    .CTime = WayPointStamp.CTime
    'get x and y from lat and long
    If .Lat <> 0 Then .x = Int((Origin_Long - Abs(.Lon)) / Scale_X)
    If .Lon <> 0 Then .y = Int((Origin_Lat - Abs(.Lat)) / Scale_Y)
End With
End Sub
Private Sub InsertAtBottom(TransectNumber As Long)
'inserts stamp into bottom of table
'--regular insertion--
Dim Response As Variant

On Error GoTo ExitWithErrors

'uses external waypoint stamp to transfer data to this object
With TransectRoute(TransectNumber)
    With .Waypoints(TransectRoute(TransectNumber).TotalWayPoints)
        .Lat = WayPointStamp.Lat
        .Lon = WayPointStamp.Lon
        .CTime = WayPointStamp.CTime
        'get x and y from lat and long
        If .Lat <> 0 Then .x = Int((Origin_Long - Abs(.Lon)) / Scale_X)
        If .Lon <> 0 Then .y = Int((Origin_Lat - Abs(.Lat)) / Scale_Y)
    End With
End With

Exit Sub
ExitWithErrors:
Response = MsgBox("Error: Wrong transect stamp format", vbOKOnly, "Import File Error")


End Sub
Public Sub Draw(TransectNumber As Long, Drawingboard As PictureBox)
'draws route
Const WAYPOINT_MARKER = vbRed
Const ROUTE_LINE = vbBlack

Dim WayPoint_count As Long
Dim P_X As Long
Dim P_Y As Long

For WayPoint_count = 0 To TransectRoute(TransectNumber).TotalWayPoints
    With TransectRoute(TransectNumber).Waypoints(WayPoint_count)
        Drawingboard.Circle (.x, .y), 3, WAYPOINT_MARKER
        If P_X <> 0 And P_Y <> 0 Then
            Drawingboard.Line (.x, .y)-(P_X, P_Y), ROUTE_LINE
        End If
        P_X = .x
        P_Y = .y
    End With
Next WayPoint_count
End Sub

