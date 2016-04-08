VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReceiver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Navigator constants
Const MAX_ROUTES = 2550
Const MAX_SEGMENTS = 100
Const ROUTE_LEN = 4000
Const RECEIVER_RADIUS = 2
Private Type RouteSegments
    NumberOfSegments As Integer
    RouteNumber(MAX_SEGMENTS) As Integer
    StepForward As Boolean
    Distance As Single
End Type

Private Type NP
    X As Integer
    Y As Integer
    Count As Integer
    direction(8) As Integer
End Type

Private Type Route
    Index As Integer
    X(ROUTE_LEN) As Integer
    Y(ROUTE_LEN) As Integer
End Type

Private Type School
    Count As Long
    Time As String
    Day As String
    GroupOfReceivers As Boolean
    ID As Long
    Fishes(MAX_FISH) As Long
End Type

Private Type typeZone
    TagNumber As Long
    UpperX As Integer
    UpperY As Integer
    LowerX As Integer
    LowerY As Integer
    UpperLat As Single
    UpperLong As Single
    LowerLat As Single
    LowerLong As Single
End Type

'Navigational Waypoints/Groups
Private Type NavigationalGroup
    ID As String
    Description As String
    Receiver(MAX_RECEIVERS_PERGROUP) As Long
    ReceiverIndex As Long
    Centroid_X As Integer
    Centroid_Y As Integer
    Centroid_Lo As Single
    Centroid_La As Single
    TTL_X As Long
    TTL_Y As Long
    TTL_LO As Single
    TTL_LA As Single
    TICK As Long
    Detections As Long
    FakeReceiver As Integer
End Type

Private Type Recursive_Node
    Segment_X1 As Long
    Segment_Y1 As Long
    Segment_X2 As Long
    Segment_Y2 As Long
    Verified As Boolean
End Type

Dim IgnorePreviousTrackRouteDefinition As Boolean


Dim K_X(8) As Long
Dim K_Y(8) As Long
Dim UserDefinedTrack(MAX_RECEIVERS) As Long
Dim Index_to_UserDefinedTrack As Long

Dim TTL_Size(MAX_GROUPS) As Long

Dim Fish_Code(MAX_FISH) As String
Dim FishGroup(MAX_RECEIVERS, MAX_SUBGROUPS) As School
Dim FishNumber_GroupAssigment(MAX_FISH, MAX_RECEIVERS) As Long

Dim LastNode As Long

Dim CurrentStation As Long


Dim NewRoute(MAX_RECEIVERS, MAX_RECEIVERS) As RouteSegments
Dim RouteMap(MAX_ROUTES) As Route
Dim RouteMap_Index As Long

Dim Nav_Group(MAX_GROUPS) As NavigationalGroup
Dim Nav_Group_Index As Long

Dim Dark_Time As Long
'Nodes and pixel contour variables... for land-avoidance
Dim Node(MAX_NODES) As Recursive_Node
Dim Node_TTL As Long
Dim ReX As Integer
Dim ReY As Integer
Dim PixelMap_X(CONTOUR_SIZE) As Long
Dim PixelMap_Y(CONTOUR_SIZE) As Long
Dim Receiver_Tag(MAX_TAGS) As String

'Zone definition
Dim NumberOfAssignedTags As Long
Dim Receiver_Zones(MAX_ZONES) As typeZone

'Upper and Lower bounds of Receiver Coverage
Dim UPPER_La As Single
Dim LOWER_La As Single
Dim UPPER_Lo As Single
Dim LOWER_Lo As Single
Dim TotalNumberOfReceivers As Long
Dim DepthFirstSearchNode(200) As NP
Dim IndexedArray(MAX_FISH, MAX_RECEIVERS) As String
Dim IndexCodes(MAX_FISH, MAX_RECEIVERS) As Long
Dim OldValue(MAX_X, MAX_Y) As Long
Public Property Let IgnorePreviousTrackRoute(v As Boolean)
IgnorePreviousTrackRouteDefinition = v
End Property
Public Property Get Detection_TTL(R As Integer, Optional OVERRIDE As Boolean = False) As Long
Dim t As Boolean

'total detections
'if grouped then return detections from group
Dim GN As Long
Dim Total As Long


GN = Receiver_Table.GroupNumber(R)

If GN <> 0 Then
    Total = Nav_Group(GN).Detections
Else
    Total = Receiver_Table.Detection_TTL(R)
End If

'set value for invisible receivers
If OVERRIDE Then
    Total = Receiver_Table.Detection_TTL_UnGrouped(R)
End If
Detection_TTL = Total
End Property
Public Property Get Color(R As Integer) As Long
Color = Receiver_Table.Color(R)
End Property
Public Property Let Color(R As Integer, c As Long)
Dim GN As Long
Dim iR As Integer

'chk if it belongs to a group
GN = Receiver_Table.GroupNumber(R)
If GN > 0 Then
    'if it belogns to a group, assign to all rcvrs in group
    With Nav_Group(GN)
        Receiver_Table.Color(.FakeReceiver) = c
        For iR = 0 To .ReceiverIndex - 1
            Receiver_Table.Color(CInt(.Receiver(iR))) = c
        Next iR
    End With
Else
    'assign just to single receiver
    Receiver_Table.Color(R) = c
End If
End Property

Public Property Get CoverageArea_Upper_La() As Single
CoverageArea_Upper_La = UPPER_La
End Property
Public Property Get CoverageArea_Upper_Lo() As Single
CoverageArea_Upper_Lo = UPPER_Lo
End Property
Public Sub CalculateCoverageArea()
Dim R As Integer
Dim LO As Single
Dim LA As Single

'init variables that need to have initial values that are non-zero
LOWER_La = 100000
LOWER_Lo = 100000
UPPER_La = 0
UPPER_Lo = 0

For R = 1 To TTL_Receivers
    LO = Receiver_Table.LO(R)
    LA = Receiver_Table.LA(R)
    'Get Boundaries for Receiver Coverage
    If LA > UPPER_La Then UPPER_La = LA
    If LA < LOWER_La Then LOWER_La = LA
    If LO > UPPER_Lo Then UPPER_Lo = LO
    If LO < LOWER_Lo Then LOWER_Lo = LO
Next R

End Sub

Public Property Get CoverageArea_Lower_La() As Single
CoverageArea_Lower_La = LOWER_La
End Property
Public Property Get CoverageArea_Lower_Lo() As Single
CoverageArea_Lower_Lo = LOWER_Lo
End Property
Public Property Get ZoneLatRange(ByVal TagID As Long) As String
'Get zone lat. range in the form (XX1-XX2)
Dim L1 As Single
Dim L2 As Single
Dim Concatenated As String

With Receiver_Zones(TagID)
    L1 = .LowerLat
    L2 = .UpperLat
End With

If L1 > L2 Then
    Concatenated = Str$(L1) & "-" & Str$(L2)
Else
    Concatenated = Str$(L2) & "-" & Str$(L1)
End If

ZoneLatRange = Concatenated

End Property
Public Property Get ZoneLongRange(ByVal TagID As Long) As String
'Get zone long. range in the form (XX1-XX2)
Dim L1 As Single
Dim L2 As Single
Dim Concatenated As String

With Receiver_Zones(TagID)
    L1 = .LowerLong
    L2 = .UpperLong
End With

If L1 > L2 Then
    Concatenated = Str$(L1) & "-" & Str$(L2)
Else
    Concatenated = Str$(L2) & "-" & Str$(L1)
End If

ZoneLongRange = Concatenated
End Property
Public Property Get GroupNumber(ByVal R As Integer) As Integer
GroupNumber = Receiver_Table.GroupNumber(R)
End Property
Public Property Get GroupReceiverNumber(G As Integer) As Integer
If G <> 0 Then
    GroupReceiverNumber = Nav_Group(G).FakeReceiver
End If

End Property
Public Property Get TagIsGeographic(ByVal TagID As Long) As Boolean
'tests if tag belongs to a geo zone or not
Dim ReturnValue As Boolean

If Receiver_Zones(TagID).TagNumber = TagID And TagID > 0 Then
    ReturnValue = True
Else
    ReturnValue = False
End If

TagIsGeographic = ReturnValue

End Property
Public Property Get Notes(ByVal R As Integer) As String
Dim NewNote As String
Dim PreviousNote As String
Dim NL As String
Dim GN As Long
Dim ReturnString As String
Dim iR As Long

NL = Chr$(13) & Chr$(10)
If R > 0 And R <= MAX_RECEIVERS Then
    GN = Receiver_Table.GroupNumber(R)
    'treat as a group
    If GN > 0 Then
        Do While Nav_Group(GN).Receiver(iR) <> 0 And iR < MAX_RECEIVERS_PERGROUP
            NewNote = Receiver_Table.Notes(CInt(Nav_Group(GN).Receiver(iR)))
            If PreviousNote <> NewNote Then
                ReturnString = ReturnString & NL & NewNote
            End If
            iR = iR + 1
        Loop
    Else
        ReturnString = Receiver_Table.Notes(R)
    End If
End If

Notes = ReturnString

End Property
Public Property Let Notes(ByVal R As Integer, note As String)
If R > 0 And R <= MAX_RECEIVERS Then
    Receiver_Table.Notes(R) = note
End If

End Property

Public Sub DefineZone(ByVal TagNumber As Long, ByVal TopX As Long, ByVal TopY As Long, ByVal BottomX As Long, ByVal BottomY As Long)
'defines zones and mark zone in geo map
Dim X As Long
Dim Y As Long
Dim R As Integer

'convert here
TopX = DeScaleThisX(TopX)
TopY = DeScaleThisY(TopY)
BottomX = DeScaleThisX(BottomX)
BottomY = DeScaleThisY(BottomY)

With Receiver_Zones(TagNumber)
'Screen coordinates
    .LowerX = BottomX
    .UpperX = TopX
    .LowerY = BottomY
    .UpperY = TopY
'Long/Lats
    'Lo = Origin_Long - (X1 * Scale_X)
    .LowerLat = Origin_Lat - (BottomY * Scale_Y)
    .LowerLong = Origin_Long - (BottomX * Scale_X)
    .UpperLat = Origin_Lat - (TopY * Scale_Y)
    .UpperLong = Origin_Long - (TopX * Scale_X)
'for future use
    .TagNumber = TagNumber
End With

'place a flag to mark coordinates within bounding box region as belonging to geographic area
For X = TopX To BottomX Step Sgn(BottomX - TopX)
    For Y = TopY To BottomY Step Sgn(BottomY - TopY)
        GeographicZones(X, Y) = TagNumber
        'mark all receivers within this boundary as belonging to this geo zone
        R = MapImage(X, Y)
        If R > 1 Then
            R = R - 1
            Receiver_Table.Tag(R) = TagNumber
        End If
    Next Y
Next X

End Sub
Public Property Get TranslateReceiverTag(ByVal TagID As Long) As String
If TagID > 0 And TagID <= NumberOfAssignedTags Then
    TranslateReceiverTag = Receiver_Tag(TagID)
End If
End Property
Public Property Get Tag(ByVal R As Integer) As Long
'get tag id
Tag = Receiver_Table.Tag(R)
End Property
Public Property Let AssignTag(ByVal R As Integer, ByVal TagID As Long)
'adds tag number to receiver
Receiver_Table.Tag(R) = TagID
End Property
Public Property Let NewReceiverTag(Tag As String)
'adds tags
NumberOfAssignedTags = NumberOfAssignedTags + 1
'add it there
Receiver_Tag(NumberOfAssignedTags) = Tag
'add column to file
DeviceBuffer.Assign_Field("TAG" & Str$(NumberOfAssignedTags)) = Tag

End Property
Public Property Get NumberOfTags() As Long
NumberOfTags = NumberOfAssignedTags
End Property

Public Sub ShowTagList(List As ComboBox)
'adds names of tags used into list
Dim i As Long

For i = 1 To NumberOfAssignedTags
    List.AddItem Receiver_Tag(i)
Next i

End Sub
Public Sub Show(ByVal R As Integer, Pic As PictureBox, Optional ScaleX As Single = 1, Optional ScaleY As Single = 1)
Dim X As Single
Dim Y As Single
Dim Color As Long

'special case: draw on a different scale
'if scale is <>1
If ScaleX = 1 And ScaleY = 1 Then
    X = Receiver.X(R)
    Y = Receiver.Y(R)
Else
    X = DeScaleThisX(Receiver.X(R)) * ScaleX
    Y = DeScaleThisY(Receiver.Y(R)) * ScaleY
End If
Color = Not (Receiver_Table.Color(R)) + 50
Pic.Circle (X, Y), 1, Color
Pic.Circle (X, Y), 2, Color
Pic.Circle (X, Y), 3, Not Color
Pic.Circle (X, Y), 4, Not Color
End Sub
Public Property Let Set_Dark_Time(temp As Long)
Dark_Time = temp
End Property
Public Sub AssignRoute(Site_To, Site_From, Size)
'Assigns route number to latest route in pixelmap
'also assigns reverse route
'uses PIXEL MAP

Dim i As Long
Dim temp_array(MAX_RECEIVERS) As Long
Dim Pixel As Long
Dim Total As Long
Dim TTL_KM As Double
Dim Segment As Integer
Dim LastSegment As Integer
Dim ID As Integer
Dim X1 As Long
Dim X2 As Long
Dim Y1 As Long
Dim Y2 As Long

'find route number to assign
ID = RouteMap_Index
RouteMap_Index = RouteMap_Index + 1
'copy route
RouteMap(ID).Index = Size
i = 0

For Pixel = 0 To Size 'Step SegmentLength
    'get previous
    X1 = X2: Y1 = Y2
    'store in route map
    X2 = PixelMap_X(Pixel): Y2 = PixelMap_Y(Pixel)
    RouteMap(ID).X(Pixel) = X2: RouteMap(ID).Y(Pixel) = Y2
    'distance between points
    If Pixel > 1 Then
        TTL_KM = TTL_KM + DistancePixels(X1, X2, Y1, Y2)
    End If
Next Pixel

'assign forward route first
'find number of segments to assign to this specific route
If Land_Avoidance = AvoidanceMode.RandomWalk Then
 Segment = 0
End If

'assign to forward route
With NewRoute(Site_To, Site_From)
    .RouteNumber(Segment) = ID
    .StepForward = True
End With
'add the last segment if shore line based
NewRoute(Site_To, Site_From).NumberOfSegments = Segment + 1

LastSegment = Segment

'now do the reverse route
For Segment = 0 To LastSegment
    temp_array(Segment) = NewRoute(Site_From, Site_To).RouteNumber(Segment)
Next Segment

'copy and insert, pushing down older elements in table
NewRoute(Site_From, Site_To).RouteNumber(0) = ID

For Segment = 0 To LastSegment
    NewRoute(Site_From, Site_To).RouteNumber(Segment + 1) = temp_array(Segment)
Next Segment

'add reverse route!
With NewRoute(Site_From, Site_To)
    .NumberOfSegments = LastSegment + 1
    .StepForward = False
End With

'store distance for segment and total with other segments
'forward route

NewRoute(Site_To, Site_From).Distance = NewRoute(Site_To, Site_From).Distance + TTL_KM

'store distance for segment and total with other segments
'reverse route
NewRoute(Site_From, Site_To).Distance = NewRoute(Site_To, Site_From).Distance

End Sub
Public Sub Route_Pixels(ByVal R As Integer, ByVal RR As Long, ByVal number As Long, ByRef PX() As Long, ByRef PY() As Long)
Dim i As Long
Dim stepping As Long
Dim RouteID As Long
Dim pixels As Long
Dim StartPixel As Long
Dim EndPixel As Long


'First validate input
If R <= MAX_RECEIVERS And RR <= MAX_RECEIVERS Then
        'get route map id from table
        RouteID = NewRoute(R, RR).RouteNumber(number)
        If RouteID > 0 Then
            With RouteMap(RouteID)
                'is this the forward or reverse route???
                If NewRoute(R, RR).StepForward Then
                    StartPixel = 0
                    EndPixel = .Index
                    stepping = 1
                Else
                    StartPixel = .Index
                    EndPixel = 0
                    stepping = -1
                End If
                'transfer to array
                For pixels = StartPixel To EndPixel Step stepping
                    i = i + 1
                    PX(i) = .X(pixels)
                    PY(i) = .Y(pixels)
                Next pixels
            End With
        End If
End If
End Sub
Public Property Get LandLocked(ByVal ReceiverNumber As Integer) As Boolean
Dim ReturnValue As Boolean
Dim Code As Integer
'returns if receiver is landlocked....

'if code=1 then its not landlocked
'if code=2 then its landlocked
'if code=0 then its not initiated yet

Code = Receiver_Table.LandLockedCode(ReceiverNumber)

Select Case Code
    Case 1
        ReturnValue = False
    Case 2
        ReturnValue = True
    Case 0
        ReturnValue = ImageProcessingEngine.LandLockedReceiver(ReceiverNumber)
        Receiver_Table.LandLockedCode(ReceiverNumber) = CInt(IIf(ReturnValue, 2, 1))
End Select

LandLocked = ReturnValue

End Property
Public Property Get Y(ByVal ReceiverNumber As Integer) As Long
Dim temp As Long
Dim GroupNumber As Long

GroupNumber = Receiver_Table.GroupNumber(ReceiverNumber)
If GroupNumber > 0 Then
    temp = Nav_Group(GroupNumber).Centroid_Y
Else
    temp = Receiver_Table.Y(ReceiverNumber)
End If

'modify for scale
temp = (temp - ZoomRegion.OriginY) * ZoomRegion.ScaleY

'pass calculated longitude
Y = temp

End Property
Public Property Let Y(ByVal ReceiverNumber As Integer, YValue As Long)
Dim temp As Long
Dim GroupNumber As Long

'always modify for scale
YValue = (YValue / ZoomRegion.ScaleY) + ZoomRegion.OriginY

GroupNumber = Receiver_Table.GroupNumber(ReceiverNumber)
If GroupNumber > 0 Then
     Nav_Group(GroupNumber).Centroid_Y = YValue
Else
    Receiver_Table.Y(ReceiverNumber) = YValue
End If

Receiver_Table.LandLockedCode(ReceiverNumber) = 0

End Property
Public Property Get X(ByVal ReceiverNumber As Integer) As Long
Dim temp As Long
Dim GroupNumber As Integer
GroupNumber = Receiver_Table.GroupNumber(ReceiverNumber)
If GroupNumber > 0 Then
    temp = Nav_Group(GroupNumber).Centroid_X
Else
    temp = Receiver_Table.X(ReceiverNumber)
End If

'modify for scale
temp = (temp - ZoomRegion.OriginX) * ZoomRegion.ScaleX
'pass calculated longitude
X = temp
End Property
Public Property Let X(ByVal ReceiverNumber As Integer, XValue As Long)
Dim temp As Long
Dim GroupNumber As Integer
GroupNumber = Receiver_Table.GroupNumber(ReceiverNumber)

'always modify for scale
XValue = (XValue / ZoomRegion.ScaleX) + ZoomRegion.OriginX

If GroupNumber > 0 Then
     Nav_Group(GroupNumber).Centroid_X = XValue
Else
     Receiver_Table.X(ReceiverNumber) = XValue
End If

Receiver_Table.LandLockedCode(ReceiverNumber) = 0

End Property
Public Property Get LO(ReceiverNumber As Integer) As Single
Dim Longitude As Single

If Receiver_Table.GroupNumber(ReceiverNumber) > 0 Then
    Longitude = Nav_Group(Receiver_Table.GroupNumber(ReceiverNumber)).Centroid_Lo
Else
    Longitude = Receiver_Table.LO(ReceiverNumber)
End If

'pass calculated longitude
LO = Longitude

End Property
Public Property Get LA(ReceiverNumber As Integer) As Single
Dim Latitude As Single
Dim GroupNumber As Long


GroupNumber = Receiver_Table.GroupNumber(ReceiverNumber)
If Receiver_Table.GroupNumber(ReceiverNumber) > 0 Then
    Latitude = Nav_Group(GroupNumber).Centroid_La
Else
    Latitude = Receiver_Table.LA(ReceiverNumber)
End If

'pass calculated latitude
LA = Latitude
End Property
Public Function CreateNewReceiverGroup(GroupName As String, Optional GN As Long = -1) As Long
'This property adds a group to the Nav Group list
'Handles duplicates names by adding the word NEW to the group
'
Dim i As Long

'make sure its all upper case
GroupName = UCase(GroupName)

If GN = -1 Then
    'make sure group name does not currently exists
    'if it does, change it to NEW something
    For i = 1 To Nav_Group_Index
        If Nav_Group(i).ID = GroupName Then
            GroupName = "New " & GroupName
        End If
    Next i
    
    'add new name to list
    'group_name_index is a base 1
    'update upper bound of array
    Nav_Group_Index = Nav_Group_Index + 1
    GN = Nav_Group_Index
    'or put group name under a specific #
    'this is used when loading AQN files
Else
    'the index should be set to the highest navgroup #
    If Nav_Group_Index < GN Then
        Nav_Group_Index = GN
    End If
End If

'group name assign
Nav_Group(GN).ID = GroupName
'return number used
CreateNewReceiverGroup = GN

End Function
Public Function GetGroupNames(ByRef ListOfGroups() As String) As Long
Dim i As Long
For i = 1 To Nav_Group_Index
    ListOfGroups(i) = Nav_Group(i).ID
Next i
GetGroupNames = Nav_Group_Index
End Function
Public Function GetReceiversInGroup(GN As Long, ByRef ListOfMembers() As Integer) As Long
Dim i As Long

With Nav_Group(GN)
    Do Until i >= .ReceiverIndex Or i = MAX_RECEIVERS_PERGROUP
        ListOfMembers(i) = .Receiver(i)
        i = i + 1
    Loop
End With
'mark end
GetReceiversInGroup = i
End Function
Public Property Let AddReceiverToGroup(ByVal GroupNumber As Long, Optional FR As Integer = -1, Receiver_Number As Integer)
'adds receiver to receiver list of group groupname
Dim i As Long
Dim N As Long
Dim FakeReceiverNumber As Integer
Dim dummy As Long
Dim Found As Boolean

'validate: DO NOT ADD IF ALREADY PART OF A GROUP
If Receiver.GroupNumber(Receiver_Number) <> 0 Then Exit Property

'SPECIAL CASE: If a receiver is being loaded from file, then wait until its # comes up to add it...
If FR = Receiver_Number Then
    With Nav_Group(GroupNumber)
        FakeReceiverNumber = Receiver.Store(.Centroid_Lo, .Centroid_La, Nav_Group(GroupNumber).ID, -1, FR)
        Receiver_Table.GroupNumber(FakeReceiverNumber) = GroupNumber
        .FakeReceiver = FakeReceiverNumber
        Receiver.Color(Receiver_Number) = Receiver.Color(FakeReceiverNumber)
        Receiver.MakeVisible FakeReceiverNumber
    End With
Else
    'Receiver list ends when array value is empty, thus there is no index
    Do
        N = Nav_Group(GroupNumber).Receiver(i)
        'if already in list, abort!
        
        If N = Receiver_Number Then
            Exit Property
        End If
        
        If N = 0 Then
            Found = True
        Else
            i = i + 1
        End If
    Loop While Not Found And i <= MAX_RECEIVERS
    
    'add to end of list
    With Nav_Group(GroupNumber)
        .Receiver(i) = Receiver_Number
        .ReceiverIndex = .ReceiverIndex + 1
        'get totals to find new center
        .TTL_LA = .TTL_LA + Receiver_Table.LA(Receiver_Number)
        .TTL_LO = .TTL_LO + Receiver_Table.LO(Receiver_Number)
        .TICK = .TICK + 1
        'divide
        If .TTL_LA > 0 Then .Centroid_La = .TTL_LA / .TICK
        If .TTL_LO > 0 Then .Centroid_Lo = .TTL_LO / .TICK
        .Centroid_X = LoToX(.Centroid_Lo)
        .Centroid_Y = LaToY(.Centroid_La)
        
        .Detections = .Detections + Receiver.Detection_TTL(Receiver_Number, True)
        
        'Deploy this as a new receiver,
        'Note: This receiver should not link to any other data, only the other receivers in the group.
        If FR = -1 Then
            If .ReceiverIndex = 1 Then
                FakeReceiverNumber = Receiver.Store(.Centroid_Lo, .Centroid_La, Nav_Group(GroupNumber).ID, -1)
                Receiver_Table.Color(FakeReceiverNumber) = Receiver.Color(Receiver_Number)
                Receiver_Table.GroupNumber(FakeReceiverNumber) = GroupNumber
                TotalNumberOfReceivers = TotalNumberOfReceivers + 1
                .FakeReceiver = FakeReceiverNumber
                Receiver.Color(Receiver_Number) = Receiver.Color(FakeReceiverNumber)
                Receiver.MakeVisible FakeReceiverNumber
            Else
                Receiver_Table.LA(.FakeReceiver) = .Centroid_La
                Receiver_Table.LO(.FakeReceiver) = .Centroid_Lo
                Receiver_Table.X(.FakeReceiver) = .Centroid_X
                Receiver_Table.Y(.FakeReceiver) = .Centroid_Y
                Receiver.MoveReceiver_InMemory .FakeReceiver, CSng(ScaleThisX(.Centroid_X)), CSng(ScaleThisY(.Centroid_Y))
            End If
            
        End If
    End With
    
    'vanish receiver from screen and data analysis
    Receiver.MakeInvisible Receiver_Number

End If

'if all is good, then add number to receiver itself
Receiver_Table.GroupNumber(Receiver_Number) = GroupNumber

End Property
Public Sub DeleteReceiver(GroupNumber As Long, R As Integer)
'Deletes receiver from list of group groupname

'find receiver in group first
'if not found, exit, nothing you can do

Dim Found As Boolean
Dim i As Long

With Nav_Group(GroupNumber)
    Do
        If .Receiver(i) = R Then
            Found = True
        End If
        i = i + 1
    Loop Until i > .ReceiverIndex Or i >= MAX_RECEIVERS_PERGROUP Or Found

    'not found? get out of here!
    If Found Then
        'delete receiver and recalculate
        'compact list
        Do Until i > .ReceiverIndex Or i >= MAX_RECEIVERS_PERGROUP
            .Receiver(i - 1) = .Receiver(i)
            i = i + 1
        Loop
        .ReceiverIndex = .ReceiverIndex + 1
        'find new center
        .TTL_X = .TTL_X - Receiver_Table.X(R)
        .TTL_Y = .TTL_Y - Receiver_Table.Y(R)
        .TTL_LA = .TTL_LA - Receiver_Table.LA(R)
        .TTL_LO = .TTL_LO - Receiver_Table.LO(R)
        .TICK = .TICK - 1
        'divide
        If .TTL_X > 0 Then .Centroid_X = .TTL_X / .TICK
        If .TTL_Y > 0 Then .Centroid_Y = .TTL_Y / .TICK
        If .TTL_LA > 0 Then .Centroid_La = .TTL_LA / .TICK
        If .TTL_LO > 0 Then .Centroid_Lo = .TTL_LO / .TICK
        .Detections = .Detections - Receiver_Table.Detection_TTL(R)
        Receiver_Table.LA(.FakeReceiver) = .Centroid_La
        Receiver_Table.LO(.FakeReceiver) = .Centroid_Lo
        Receiver_Table.X(.FakeReceiver) = .Centroid_X
        Receiver_Table.Y(.FakeReceiver) = .Centroid_Y
    End If
End With
End Sub
Public Sub DeleteGroup(GroupNumber As Long)
'deletes group from name list and unlinks it from receiver
Dim i As Long
Dim R As Integer

'unlink
Do While Nav_Group(GroupNumber).Receiver(i) > 0
    R = Nav_Group(GroupNumber).Receiver(i)
    Receiver_Table.GroupNumber(R) = 0
    Receiver.MakeVisible R
    Nav_Group(GroupNumber).Receiver(i) = 0
    i = i + 1
Loop

'delete
Nav_Group(GroupNumber).ID = ""

'delete fake receiver
R = Nav_Group(GroupNumber).FakeReceiver
Receiver.X(R) = 0
Receiver.Y(R) = 0
Receiver_Table.Detection_TTL(R) = 0
Receiver.MakeInvisible R


End Sub

Public Property Get ReceiverNumber(ID As String) As Integer
Dim N As Long
Dim Found As Boolean

'find receiver number in table
Do
    N = N + 1
    If N > MAX_RECEIVERS Then Exit Function 'critical error
    If Receiver_Table.ID(N) = ID Then Found = True
Loop While Found = False And ID <> ""

If Found Then
    ReceiverNumber = N
End If

End Property


Public Property Let CurrentStation_Number(N As Integer)
CurrentStation = N

End Property
Public Property Get CurrentStation_Number() As Integer
CurrentStation_Number = CurrentStation

End Property
Public Property Get Detection_Total(ByVal R As Integer) As Long
Detection_Total = Receiver_Table.Detection_TTL(R)
End Property
Public Property Get Pings(R As Integer) As Long
Dim p As Long
Dim detection As Long
Dim Totalstamps As Long

'counts valid pings

If R > MAX_RECEIVERS Or Receiver_Table.Invisible(R) Then
    p = 0
Else
    Totalstamps = Receiver.Detection_TTL(R)
    If Receiver_Table.Invisible(R) Then Totalstamps = 0
    'count valid stamps
    For detection = 1 To Totalstamps
        ReadStamp R, detection
        If Stamp.Valid And FishDatabase.IsVisible(Stamp.Fish) Then p = p + 1
    Next detection
End If

'return # of valid pings
Pings = p

End Property
Public Property Let Pings(R As Integer, p As Long)
Receiver_Table.Pings(R) = p
End Property
Public Sub PingUp(R As Integer)
Receiver_Table.Pings(R) = Receiver_Table.Pings(R) + 1
End Sub
Public Property Get ID(i As Integer, Optional RETURN_RECEIVER_NAME_ONLY As Boolean = False) As String
Dim ReceiverName As String
Dim GroupNumber As Long
GroupNumber = Receiver_Table.GroupNumber(i)
If GroupNumber > 0 And Not RETURN_RECEIVER_NAME_ONLY Then
    ReceiverName = Nav_Group(GroupNumber).ID
Else
    ReceiverName = Receiver_Table.ID(i)
End If

ID = ReceiverName

End Property
Public Property Get Date_Active(R As Integer, i As Long) As Integer
Dim ReceiverNumber As Long
Dim GN As Integer

ReceiverNumber = R
ReadStamp ReceiverNumber, 1
Date_Active = Stamp.Date
End Property
Public Sub Reset()
Dim i As Long
Dim ii As Long

'hashtable
For i = 0 To MAX_FISH
    For ii = 0 To MAX_RECEIVERS
        IndexedArray(i, ii) = ""
        IndexCodes(i, ii) = 0
    Next ii
Next i

'routes
For i = 0 To MAX_RECEIVERS
    For ii = 0 To MAX_RECEIVERS
        With NewRoute(i, ii)
            .Distance = 0
            .NumberOfSegments = 0
        End With
    Next ii
Next i

'groups
For i = 0 To MAX_GROUPS
    With Nav_Group(i)
        .FakeReceiver = 0
        .Centroid_La = 0
        .Centroid_Lo = 0
        .Centroid_X = 0
        .Centroid_Y = 0
        .Detections = 0
        .TTL_LA = 0
        .TTL_LO = 0
        .TICK = 0
        .TTL_X = 0
        .TTL_Y = 0
        .ID = ""
        .ReceiverIndex = 0
    End With
Next i
    
NumberOfAssignedTags = 0
TotalNumberOfReceivers = 0
Nav_Group_Index = 0

End Sub
Public Sub Clear()
Receiver_Table.ClearPingCount
End Sub
Public Sub Restore()
'Restores data to include all receivers
'(makes all more simple to me but makes a bit of a sloppy program maybe)
Receiver_Table.RestoreTable
End Sub
Public Sub Backup()
Receiver_Table.BackupTable
End Sub
Public Property Get TotalReceivers() As Long
TotalReceivers = TTL_Receivers

End Property
Private Function TTL_Receivers() As Long
Dim ReceiverNumber As Long
Dim ID As String

'run once
If TotalNumberOfReceivers = 0 Then
    'Count receivers
    Do
        ReceiverNumber = ReceiverNumber + 1
        If ReceiverNumber > MAX_RECEIVERS Then Exit Function 'critical error
        ID = Receiver_Table.ID(ReceiverNumber)
    Loop While ID <> ""
    TotalNumberOfReceivers = ReceiverNumber - 1
End If

'return to caller
TTL_Receivers = TotalNumberOfReceivers

End Function
Public Sub ShowSingletons(Threshold As Long, List As ListBox, ReceiverList() As Integer, EntryList() As Long, FishList() As Integer)
'Detects singleton detections, defined by a single detection in one receiver from one fish.
'In some instances, these are assumed to be random, and may be deleted to purge the database of unwanted false detections
'
Dim R As Integer
Dim s As Long
Dim LastDetection(MAX_FISH) As Long
Dim StampNumber(MAX_FISH) As Long
Dim Singleton(MAX_FISH) As Boolean
Dim FishNumber As Long
Dim EntryIndex As Long

'make sure all receivers are visible
'(checks entire array)
Receiver.MakeVisible

For R = 1 To TTL_Receivers
    If Receiver.Detection_TTL(R) = 1 Then
        'singleton by definition!
        'show on list
        Receiver.ReadStamp R, 0
        If Stamp.Valid Then
            List.AddItem Str$(FishDatabase.Code(Stamp.Fish)) & " at " & Receiver_Table.ID(R) & " on " & Convert_DayNumber(Stamp.Date)
            FishList(EntryIndex) = Stamp.Fish
            ReceiverList(EntryIndex) = R
            EntryList(EntryIndex) = 1
            EntryIndex = EntryIndex + 1
        End If
    Else
        'scan all entries in receiver
        'but don't do groups until the end
        If Receiver.GroupNumber(R) = 0 Or (Receiver.GroupNumber(R) <> 0 And Receiver.GroupReceiverNumber(R) = R) Then
            For s = 0 To Receiver.Detection_TTL(R) - 1
                Receiver.ReadStamp R, s
                If Stamp.Valid Then
                    'see if fish already in table
                    FishNumber = Stamp.Fish
                    If (Stamp.CTime - LastDetection(FishNumber)) >= Threshold Then
                        'valid?
                        If Singleton(FishNumber) Then
                            'mark
                            List.AddItem Str$(FishDatabase.Code(FishNumber)) & " at " & Receiver_Table.ID(R) & " on " & Convert_DayNumber(Fix(LastDetection(FishNumber) / 1440))
                            FishList(EntryIndex) = FishNumber
                            ReceiverList(EntryIndex) = R
                            EntryList(EntryIndex) = StampNumber(FishNumber)
                            EntryIndex = EntryIndex + 1
                        End If
                        
                        'wait for next round to validate
                        Singleton(FishNumber) = True
                    Else
                        'if within Threshold then discard
                        'no validation needed until next time
                        Singleton(FishNumber) = False
                    End If
                    
                    LastDetection(FishNumber) = Stamp.CTime
                    StampNumber(FishNumber) = s
                End If
            Next s
        End If
        'at end discard all entries that still need validation
        'AND clear table
        '(Loop with two functions? Yes, but its just b/c of speed concerns)
        For FishNumber = 0 To MAX_FISH
            If Singleton(FishNumber) Then
                'mark
                List.AddItem Str$(FishDatabase.Code(FishNumber)) & " at " & Receiver_Table.ID(R) & " on " & Convert_DayNumber(Fix(LastDetection(FishNumber) / 1440))
                FishList(EntryIndex) = FishNumber
                ReceiverList(EntryIndex) = R
                EntryList(EntryIndex) = StampNumber(FishNumber)
                EntryIndex = EntryIndex + 1
            End If
            
            'clear entry
            Singleton(FishNumber) = False
            LastDetection(FishNumber) = 0
        Next FishNumber
    End If
Next R
            
End Sub
Public Function Store(LO As Single, LA As Single, Site As String, FishCode As Long, Optional ReceiverNumber As Integer = 0) As Integer
Dim X As Long
Dim Y As Long
Dim L As Long
Dim i As Long

Dim Found As Boolean
Dim ReceiverTimeBinFound As Boolean

Dim Receiver_ID As String
Dim response As Variant
Dim R As Long
Static AssignedNumber As Long
On Error GoTo RaiseError

'Stores receiver's info in a table
'

'Scan table.

If ReceiverNumber <> 0 Then
    Found = True
End If

If Not Found Then
    If FishCode = -1 Then
        AssignedNumber = Receiver.TotalReceivers
    Else
        'super simple hashing
        L = (Len(Site) - 1)
        L = L + Asc(Right(Site, 1)) + Asc(Left(Site, 1))
        If L > MAX_FISH Then L = MAX_FISH
        Found = False
        'find valid receiver number or used forced number
        Do
            R = R + 1
            Receiver_ID = IndexedArray(L, R)
            If Receiver_ID = Site Then Found = True
        Loop While (Not Found) And Receiver_ID <> ""
        ReceiverNumber = IndexCodes(L, R)
        'if receiver not found store information in table and mark map with code for receiver
    End If
End If

'if found, update ping info
If Found Then
    Receiver_Table.Detection_TTL(ReceiverNumber) = Receiver_Table.Detection_TTL(ReceiverNumber) + 1
Else
    'If not found and there is an empty place in the table,
    'fill entry
    AssignedNumber = AssignedNumber + 1
    ReceiverNumber = AssignedNumber
    IndexCodes(L, R) = AssignedNumber
    IndexedArray(L, R) = Site
    Receiver_Table.ID(ReceiverNumber) = Site
    
    Receiver_Table.Pings(ReceiverNumber) = 1
    
    'store long and lat for distance calculations
    Receiver_Table.LO(ReceiverNumber) = LO
    Receiver_Table.LA(ReceiverNumber) = LA
        
    'use coordinates to draw receiver in in-memory map and to store that info here
    Receiver.X(ReceiverNumber) = LoToX(LO)
    Receiver.Y(ReceiverNumber) = LaToY(LA)
    'give it a marker
    Receiver_Table.AssignDefaultMarker CLng(ReceiverNumber)
    Receiver.MakeVisible ReceiverNumber
    Receiver_Table.Detection_TTL(ReceiverNumber) = 1
    
    'Receiver code in map is +1 from code in table
    DrawReceiver_InMemory ReceiverNumber

End If

'return receiver number
Store = ReceiverNumber
Exit Function

RaiseError:
response = MsgBox("Row #" & Str$(row) & ": An internal error ocurred while attempting to process receiver information.", vbOKOnly, "Error")
End Function
Private Function LoToX(L As Single) As Long
LoToX = Int((Origin_Long - L) / Scale_X)
End Function
Private Function LaToY(L As Single) As Long
LaToY = Int((Origin_Lat - L) / Scale_Y)
End Function
Public Sub ReDraw()
Dim ReceiverNumber As Integer
Dim X As Long
Dim Y As Long
Dim G As Long
Dim t As Long
Dim response As Variant

On Error GoTo ExitWithError

'Redraws all receivers using a new scale

'clear receivers in map but preserve land water
For X = 0 To MAX_X
    For Y = 0 To MAX_Y
        If MapImage(X, Y) > 1 Then MapImage(X, Y) = 0
    Next Y
Next X

'Receivers
For ReceiverNumber = 1 To TTL_Receivers
    'use coordinates to draw receiver in in-memory map and to store that info here
    Receiver_Table.X(ReceiverNumber) = Int((Origin_Long - Receiver_Table.LO(ReceiverNumber)) / Scale_X)
    Receiver_Table.Y(ReceiverNumber) = Int((Origin_Lat - Receiver_Table.LA(ReceiverNumber)) / Scale_Y)
    'Receiver code in map is +1 from code in table
    DrawReceiver_InMemory ReceiverNumber
Next ReceiverNumber

'Groups
If Nav_Group_Index > 0 Then
    For G = 1 To Nav_Group_Index
        With Nav_Group(G)
            .Centroid_X = Int((Origin_Long - .Centroid_Lo) / Scale_X)
            .Centroid_Y = Int((Origin_Lat - .Centroid_La) / Scale_Y)
        End With
    Next G
End If

'Zones
'any tags assigned?
If NumberOfAssignedTags > 0 Then
    For t = 1 To NumberOfAssignedTags
        With Receiver_Zones(t)
            .LowerX = Int((Origin_Long - .LowerLong) / Scale_X)
            .LowerY = Int((Origin_Lat - .LowerLat) / Scale_Y)
            .UpperX = Int((Origin_Long - .UpperLong) / Scale_X)
            .UpperY = Int((Origin_Lat - .UpperLat) / Scale_Y)
        End With
    Next t
End If


Exit Sub

ExitWithError:
response = MsgBox("Error redrawing receivers! Invalid control points!", vbOKOnly)

End Sub
Private Function Format_To_Output(s As String, Optional FL As Long = 0) As String
Dim L As Long
Dim Tabulation As String
Static LastValue As Long

'retrieve last used if set to 0
If FL = 0 Then FL = LastValue

L = Len(s)
If L > FL Then
    Tabulation = ""
Else
    Tabulation = Space(FL - L)
End If

'save value for next one
LastValue = FL
Format_To_Output = s & Tabulation

End Function
Public Function ShowAllPings(ReceiverNumber As Integer, picPings As PictureBox, Optional FishNumber As Long = -1, Optional Start As Long = 1) As Long

Const FIELD_LEN = 20

Dim Receiver_Name As String
Dim p As Long
Dim GroupName As String
Dim Date_Stamp As String
Dim Time_Stamp As String

Dim i As Long

Dim Concatenated As String
Dim Header(4) As String
Dim Header_X(4) As Long

Dim CurrentFish As String
Dim R As Long
Dim LastDetection As Long



Dim PageEnds As Long
Static c As Single

R = CLng(ReceiverNumber)
LastDetection = Receiver.Detection_Total(ReceiverNumber)


'heading
Header(0) = "Fish"
Header(1) = "Ping"
Header(2) = "Receiver"
Header(3) = "Date"
Header(4) = "Time"

With picPings
    .Cls
    .CurrentX = 0
    .CurrentY = 0
    .ForeColor = vbBlack
    .BackColor = vbWhite
    .FontBold = True
End With

For i = 0 To UBound(Header)
    Header_X(i) = picPings.CurrentX
    picPings.Print Format_To_Output(Header(i), FIELD_LEN);
Next i
picPings.Print ""
picPings.FontBold = False

'print receiver information

If Receiver.GroupNumber(ReceiverNumber) > 0 Then
    GroupName = Receiver.ID(ReceiverNumber)
    If Len(GroupName) > 7 Then GroupName = Left(GroupName, 5) & "..."
    GroupName = "   (" & GroupName & ")"
End If

If FishNumber = -1 Then
    p = Start - 1
Else
    p = 0
    i = 0
    Do
        Receiver.ReadStamp R, p
        If Stamp.Fish = FishNumber Then i = i + 1
        p = p + 1
    Loop Until i = Start Or p > LastDetection
End If
        
Do

        Receiver.ReadStamp R, p
        Receiver_Name = Receiver.ID(Stamp.Site, True) & GroupName
        
        If (FishNumber = -1 Or FishNumber = Stamp.Fish) And Stamp.Valid Then
        
            If CurrentFish <> "" And CurrentFish <> FishDatabase.Code(Stamp.Fish) Then
                c = c + 1
                If c > 2 Then c = 1
                picPings.ForeColor = Choose(c, vbBlue, vbBlack)
            End If
                
            picPings.CurrentX = Header_X(0)
            picPings.Print FishDatabase.Code(Stamp.Fish);
            picPings.CurrentX = Header_X(1)
            picPings.Print Trim(Str$(p));
            picPings.CurrentX = Header_X(2)
            picPings.Print Receiver_Name;
            picPings.CurrentX = Header_X(3)
            picPings.Print Convert_DayNumber(Stamp.Date);
            picPings.CurrentX = Header_X(4)
            picPings.Print Convert_ToStandardTime(CSng(Stamp.Time))
                    
            CurrentFish = FishDatabase.Code(Stamp.Fish)
        End If
    If picPings.CurrentY > picPings.ScaleHeight Then
        PageEnds = p - 1
    End If
    p = p + 1
Loop Until PageEnds = p Or p > LastDetection


ShowAllPings = PageEnds

End Function
Public Function ReadFishStampOnReceiver(R As Integer, F As Integer, d As Long) As Long
'Reads a stamp from receiver, serially
'Filter by fishnumber
'if receiver is part of group, reads all stamps in group
'
'Besides returning its stamp
'also returns a number that is the last stamp read
'
'R is the receiver, F is the fishnumber to filter it by, and D is the start stamp

'NOTE: IF F=-1, it will read all detections, equivalent to the much simpler ReadStamp
'Returns -1 when last stamp is reached

Dim ReturnValue As Long
Dim Totalstamps As Long
Dim Found As Boolean

Totalstamps = Receiver.Detection_TTL(R)

If d < Totalstamps And d >= 0 Then
    Do
        d = d + 1
        ReadStamp R, d
        If Stamp.Fish = F Or F = -1 Then
            Found = True
        End If
    Loop Until Found Or d >= Totalstamps
End If

If Not Found Then d = -1

ReadFishStampOnReceiver = d
End Function
Public Sub ReadStamp(ByVal R As Long, ByVal e As Long)
'this version of readstamp can handle groupped receivers behind the black box
'
Dim GN As Long
Dim ReceiversInGroup(MAX_RECEIVERS_PERGROUP) As Long
Dim i As Long
Dim Bound As Long

GN = CLng(Receiver_Table.GroupNumber(R))

If GN <> 0 Then
    With Nav_Group(GN)
        Bound = .ReceiverIndex - 1
        For i = 0 To Bound
            ReceiversInGroup(i) = Nav_Group(GN).Receiver(i)
        Next i
    End With
        
    'read
    Receiver_Table.ReadGroupStamp ReceiversInGroup, Bound, e
Else
    ReceiverTable.ReadStamp R, e
End If

End Sub
Public Sub Draw(MapPicture As PictureBox, R As Integer, c As Long)
'draws receiver in picture
MapPicture.Circle (Receiver.X(R), Receiver.Y(R)), 3, c
End Sub
Public Sub DrawReceiver(MapPicture As PictureBox, R As Integer, Optional s As Long = 1, Optional ColorFill As Long = -1, Optional X As Long = -1, Optional Y As Long = -1)
Const DefaultSize = 2

Dim MarkerSize As Long
Dim i As Long
Dim LineY As Long
Dim Color As Long
Dim RN As Integer

'override receiver group number
If Receiver.GroupNumber(R) <> 0 Then RN = Receiver.GroupReceiverNumber(Receiver.GroupNumber(R)) Else RN = R


MarkerSize = CLng(DefaultSize * s)
If X = -1 Or Y = -1 Then
    X = Receiver.X(RN)
    Y = Receiver.Y(RN)
End If

'Draw receiver in picbox, option scale to S size
With Receiver_Table
    'exit if not visible
    If .Invisible(RN) Then Exit Sub
    
    'get color. use default fill if not given
    Color = .Color(RN)
    If ColorFill = -1 Then ColorFill = Color
        
    Select Case .Marker(RN)
        Case MarkerType.Circular
            'fill
            For i = DefaultSize To MarkerSize
                MapPicture.Circle (X, Y), i, ColorFill
            Next i
            
            'fill origin
            MapPicture.Line (X - 1, Y - 1)-(X + 1, Y + 1), ColorFill, BF
            
            'marker
            MapPicture.Circle (X, Y), MarkerSize + 1, Color
            
        Case MarkerType.Rectangle
            'filled box
            MapPicture.Line (X - MarkerSize, Y - MarkerSize)-(X + MarkerSize, Y + MarkerSize), ColorFill, BF
            MarkerSize = MarkerSize + 1
            MapPicture.Line (X - MarkerSize, Y - MarkerSize)-(X + MarkerSize, Y + MarkerSize), Color, B
        Case MarkerType.Triangle
            'outside
            MarkerSize = MarkerSize + 3
            LineY = Y + CLng(MarkerSize / 2) - 1
            For i = 0 To MarkerSize
                MapPicture.Line (X - MarkerSize + i, LineY - i)-(X + MarkerSize - i, LineY - i), Color
            Next i
            'fill
            MarkerSize = MarkerSize - 1
            For i = 1 To MarkerSize
                MapPicture.Line (X - MarkerSize + i, LineY - i)-(X + MarkerSize - i, LineY - i), ColorFill
            Next i
        Case MarkerType.InvertedTriangle
            'outside
            MarkerSize = MarkerSize + 3
            LineY = Y - CLng(MarkerSize / 2) - 1
            For i = MarkerSize To 0 Step -1
                MapPicture.Line (X - MarkerSize + i, LineY + i)-(X + MarkerSize - i, LineY + i), Color
            Next i
            'fill
            MarkerSize = MarkerSize - 1
            For i = MarkerSize To 1 Step -1
                MapPicture.Line (X - MarkerSize + i, LineY + i)-(X + MarkerSize - i, LineY + i), ColorFill
            Next i
         Case MarkerType.Diamond
            'outside
            MarkerSize = MarkerSize + 3
            LineY = Y - CLng(MarkerSize) - 1
            For i = MarkerSize / 2 To 0 Step -1
                MapPicture.Line (X - MarkerSize + i, LineY - i)-(X + MarkerSize - i, LineY - i), Color
            Next i
            LineY = Y - CLng(MarkerSize) - 1
            For i = 0 To MarkerSize / 2
                MapPicture.Line (X - MarkerSize + i, LineY + i)-(X + MarkerSize - i, LineY + i), Color
            Next i
            MarkerSize = MarkerSize - 1
            'fill
            LineY = Y - CLng(MarkerSize) - 1
            For i = MarkerSize / 2 To 0 Step -1
                MapPicture.Line (X - MarkerSize + i, LineY - i)-(X + MarkerSize - i, LineY - i), ColorFill
            Next i
            LineY = Y - CLng(MarkerSize) - 1
            For i = 0 To MarkerSize / 2
                MapPicture.Line (X - MarkerSize + i, LineY + i)-(X + MarkerSize - i, LineY + i), ColorFill
            Next i
    End Select
End With
   
End Sub
Public Sub MoveReceiver_InMemory(ByVal R As Integer, X As Single, Y As Single)
'delete receiver option of drawreceiver
'leaves water behind
DrawReceiver_InMemory R, 1

'move it to new coords
Receiver.X(R) = X
Receiver.Y(R) = Y
'Draw on new coords
DrawReceiver_InMemory R
End Sub
Private Sub DrawReceiver_InMemory(ByVal R As Integer, Optional Code As Long = -1)
Dim X As Long
Dim Y As Long
Dim Xr As Long
Dim Yr As Long


Xr = DeScaleThisX(Receiver.X(R))
Yr = DeScaleThisY(Receiver.Y(R))


'Receiver code in map is +1 from code in table
'if code <>-1, deletes receiver
If Xr >= 0 And Xr <= MAX_X And Yr >= 0 And Yr <= MAX_Y Then
    If Code = -1 Then
        Code = R + 1
        'OldValue(Xr, Yr) = 1
    Else
        Code = 1
    End If
End If

'no bound check!
For X = Xr - RECEIVER_RADIUS To Xr + RECEIVER_RADIUS
    For Y = Yr - RECEIVER_RADIUS To Yr + RECEIVER_RADIUS
        If (X >= 0 And X <= MAX_X) And (Y >= 0 And Y <= MAX_Y) Then MapImage(X, Y) = Code
    Next Y
Next X

End Sub
Public Sub LoadReceiver(ReceiverNumber As Integer, p As Long, s As String, LO As Single, LA As Single, X As Long, Y As Long)
'visible
Receiver.MakeVisible ReceiverNumber
'pings
Receiver_Table.Pings(ReceiverNumber) = p
Receiver_Table.Detection_TTL(ReceiverNumber) = p
'name
Receiver_Table.ID(ReceiverNumber) = s

'coords
'store long and lat for distance calculations
Receiver_Table.LO(ReceiverNumber) = LO
Receiver_Table.LA(ReceiverNumber) = LA
    
'use x/y unless no valid x/y
If X = -1 Or Y = -1 Then
    'valid scale?
    With ProgramSettings
        If .Scale_X <> 0 And .Scale_Y <> 0 Then   'use coordinates to draw receiver in in-memory map and to store that info here
            Receiver.X(ReceiverNumber) = Int((.Origin_Long - LO) / .Scale_X)
            Receiver.Y(ReceiverNumber) = Int((.Origin_Lat - LA) / .Scale_Y)
        End If
    End With
Else
    Receiver.X(ReceiverNumber) = X
    Receiver.Y(ReceiverNumber) = Y
End If

'Receiver code in map is +1 from code in table
DrawReceiver_InMemory ReceiverNumber
End Sub
Private Function DistancePixels(X1 As Long, X2 As Long, Y1 As Long, Y2 As Long) As Double
Dim d As Double
Dim LA As Single
Dim LO As Single
Dim La2 As Single
Dim Lo2 As Single



LA = Origin_Lat - (Y2 * Scale_Y)
LO = Origin_Long - (X2 * Scale_X)
La2 = Origin_Lat - (Y1 * Scale_Y)
Lo2 = Origin_Long - (X1 * Scale_X)
d = Distance(LO, Lo2, LA, La2)
DistancePixels = d
End Function
Private Function Distance(Longitude As Single, longitude2 As Single, Latitude As Single, latitude2 As Single) As Double
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
    Distance = d

End Function

Public Sub OutputFishGroups(ListOfGroups As ListBox, ListOfFishInGroup As ListBox, ListGroupsPerse As ListBox, ListFormattedForCSV As ListBox)
Dim i As Long
Dim Name As String
Dim IDNumber As Long
Dim Concatenated As String
Dim F As Long
Dim Subgroup As Long
Dim Counter As Long
Dim CSV As String

'outputs to list boxes in window
'CSV format: FG, Receiver(s), Number of fish, Detection time (last), Detection date (last), Fish List
For i = 0 To (MAX_GROUPS)
    For Subgroup = 0 To MAX_SUBGROUPS
        With FishGroup(i, Subgroup)
            If .Count > 0 Then
                Counter = Counter + 1
                IDNumber = .ID
                'get name of station OR group
                If .GroupOfReceivers Then
                    Name = "FG#" & Str$(Counter) & ": " & Str$(.Count) & " fishes at " & Nav_Group(IDNumber).ID & Space(10) & "(" & .Day & Space(2) & "at " & .Time & Space(2) & ")"
                    CSV = Str$(Counter) & "," & Nav_Group(IDNumber).ID & "," & Str$(.Count) & "," & .Time & "," & .Day
                Else
                    Name = "FG#" & Str$(Counter) & ": " & Str$(.Count) & " fishes at " & Receiver_Table.ID(IDNumber) & Space(10) & "(" & .Day & Space(2) & "at " & .Time & Space(2) & ")"
                    CSV = Str$(Counter) & "," & Receiver_Table.ID(IDNumber) & "," & Str$(.Count) & "," & .Time & "," & .Day
                End If
                
                'show
                ListOfGroups.AddItem Name
                
                'get names of fish in that group
                Concatenated = "Time " & .Time & "-->"
                For F = 0 To .Count - 1
                    Concatenated = Concatenated & FishDatabase.Code(CInt(.Fishes(F))) & ",  "
                    CSV = CSV & "," & FishDatabase.Code(CInt(.Fishes(F)))
                Next F
                
                
                'show this names
                ListOfFishInGroup.AddItem Concatenated
                ListFormattedForCSV.AddItem CSV
            End If
        End With
    Next Subgroup
Next i


End Sub
Public Sub FindFishGroups(Threshold As Long)
'Finds and defines fish groups
Dim GroupName As Long
Dim R As Integer
Dim Group As Long
Dim G As Long
Dim FishGroupNumber As Long
Dim Subgroup As Long
Dim dummy As Long
Dim response As Variant
Dim sg As Long
Dim MaxSG As Long

'Create table
dummy = CreateGroupTable
On Error GoTo ExitWithError

'fish arriving together at group
For G = 1 To MAX_GROUPS
    If Nav_Group(G).ID <> "" Then
        'set group
        Group = Group + 1
        For Subgroup = 0 To MAX_SUBGROUPS
            With FishGroup(Group, Subgroup)
               .Count = 0
               .GroupOfReceivers = True
               .ID = Group
            End With
            'count fish in group
        Next Subgroup
        FishGroupNumber = FishGroupNumber + 1
        sg = CountFish(Threshold, GroupTable, Group, TTL_Size(Group), FishGroupNumber)
        If sg > MaxSG Then MaxSG = sg
    End If
Next G

'and fish arriving together at receivers
For R = 1 To TTL_Receivers
    If Receiver_Table.GroupNumber(R) = 0 Then
        FishGroupNumber = FishGroupNumber + 1
        For Subgroup = 0 To MAX_SUBGROUPS
            With FishGroup(FishGroupNumber, Subgroup)
                .Count = 0
                .GroupOfReceivers = False
                .ID = R
            End With
        Next Subgroup
        sg = CountFish(Threshold, ReceiverTable, CLng(R), Receiver.Detection_TTL(CLng(R)), FishGroupNumber)
        If sg > MaxSG Then MaxSG = sg
    End If
Next R

If MaxSG >= MAX_SUBGROUPS Then GoTo ExitWithError

Exit Sub


ExitWithError:

response = MsgBox("Error: You have exceeded the number of fish groups! Change the grouping criteria or split your file into two or more files!", vbExclamation, "Warning")

End Sub
Public Function CreateGroupTable(Optional GroupNumber As Long = -1) As Long
'Creates a Table with Receiver Data Organized by GROUPS
'table is GroupTable
Dim R As Integer
Dim s As Long
Dim Group As Long
Dim G As Long
Dim Index As Long
Dim FirstGroup As Long
Dim LastGroup As Long

'You can do one group only
If GroupNumber = -1 Then
    FirstGroup = 1
    LastGroup = MAX_GROUPS
Else
    FirstGroup = GroupNumber
    LastGroup = GroupNumber
End If

'Start with groups
'Place group in a separate table (as if all belonged to the same receiver, then sort)
For G = FirstGroup To LastGroup
    If Nav_Group(G).ID <> "" Then
        TTL_Size(Group) = Nav_Group(G).Detections
        Group = Group + 1
    End If
Next G

'partition table
GroupTable.PartitionTable TTL_Size()
Group = 0

'copy stamps (they are sorted once all are copied!
For G = FirstGroup To LastGroup
    If Nav_Group(G).ID <> "" Then
        Index = 0
        R = Nav_Group(G).Receiver(0)
        For s = 0 To TTL_Size(Group) - 1
            'read from receiver table
            ReadStamp R, s
            'copy to group table so its sorted
            GroupTable.WriteStamp Group, s
        Next s
        Group = Group + 1
    End If
Next G

'return number of groups in table
CreateGroupTable = Group

End Function
Private Function CountFish(Threshold As Long, Table As clsTable, Entry As Long, Size As Long, StoreGroup As Long) As Long
'returns subgroup count
Dim PrevFishNumber As Long
Dim FishNumber As Long
Dim StampNumber As Long
Dim TimePrevious As Long
Dim TimeNow As Long
Dim TimeDifference As Long
Dim GroupName As Long
Dim SubGroupIsEmpty As Boolean
Dim Subgroup As Long


On Error GoTo ExitWithError

For StampNumber = 1 To Size
    TimePrevious = TimeNow
    PrevFishNumber = FishNumber
    Table.ReadStamp Entry, StampNumber
    If Stamp.Date > 0 And Stamp.Valid Then
        FishNumber = Stamp.Fish
        TimeNow = Stamp.CTime
        
        'if at first fishnumber, make sure you compare it with itself...!
        If PrevFishNumber = 0 Then
            PrevFishNumber = FishNumber
            TimePrevious = TimeNow
        End If
        
        'Test if detection of different fish are within a threshold
        If FishNumber <> PrevFishNumber Then
            TimeDifference = TimeNow - TimePrevious
            If TimeDifference <= Threshold Then
                'Store in a table
                'keep count of number of times found
                InsertFish FishNumber, PrevFishNumber, StoreGroup, Subgroup, Stamp.Time, Stamp.Date
                SubGroupIsEmpty = False
            Else
            'create a new subgroup if not time-contigous with last one
                If Not SubGroupIsEmpty Then
                    Subgroup = Subgroup + 1
                    If Subgroup > MAX_SUBGROUPS Then
                        GoTo ExitWithError
                    End If
                    SubGroupIsEmpty = True
                End If
            End If
        End If
    End If
Next StampNumber

ExitWithError:
'nop

'error or no error, return sg
CountFish = Subgroup

End Function
Private Sub InsertFish(FishNumber_1 As Long, FishNumber_2 As Long, StoreGroup As Long, Subgroup As Long, TimeStamp As Integer, DayStamp As Long)
'Insert fish in table if fish is not in table already!
Dim i As Long
Dim Counts As Long
Dim Fish_1_AlreadyExists As Boolean
Dim Fish_2_AlreadyExists As Boolean

On Error GoTo ExitWithError

If StoreGroup > MAX_RECEIVERS Or Subgroup > MAX_SUBGROUPS Then
    GoTo ExitWithError
End If


With FishGroup(StoreGroup, Subgroup)
'check if it exist
    For i = 0 To .Count - 1
        If .Fishes(i) = FishNumber_1 Then
            Fish_1_AlreadyExists = True
        End If
        
        If .Fishes(i) = FishNumber_2 Then
            Fish_2_AlreadyExists = True
        End If
        
    Next i
    
    'no?
    'insert
    If Fish_1_AlreadyExists And Fish_2_AlreadyExists Then
        Exit Sub
    Else
        .Time = Convert_ToStandardTime(TimeStamp)
        .Day = Convert_DayNumber(DayStamp)
    End If
    If Not Fish_1_AlreadyExists Then
        .Fishes(.Count) = FishNumber_1
        .Count = .Count + 1
    End If
    
    If Not Fish_2_AlreadyExists Then
        .Fishes(.Count) = FishNumber_2
        .Count = .Count + 1
    End If

End With

ExitWithError:
Exit Sub

End Sub
Public Property Get RouteDistance(Site_To As Long, Site_From As Long) As Double
'get distance from route map
RouteDistance = NewRoute(Site_To, Site_From).Distance

End Property
Public Function NavigateBetweenCorridorPoints(ByVal Site_To As Long, ByVal Site_From As Long) As Boolean
Dim X1 As Long
Dim X2 As Long
Dim Y1 As Long
Dim Y2 As Long
Dim ST As Integer
Dim SF As Integer
Dim RouteFound As Boolean
Dim LandFound As Boolean
Dim ForwardDirection As Boolean
Dim p As Long
Dim ReturnValue As Boolean
Dim ID As Long
Dim RouteExists As Boolean

Dim N As Long
Dim i As Long

'check validity
If Site_To = Site_From Then
    Exit Function
End If

'here its different
ST = CInt(Site_To)
SF = CInt(Site_From)


'check if land present
X2 = DeScaleThisX(Receiver.X(ST))
Y2 = DeScaleThisY(Receiver.Y(ST))
X1 = DeScaleThisX(Receiver.X(SF))
Y1 = DeScaleThisY(Receiver.Y(SF))
With ImageProcessingEngine
    If .LandBetweenPositions(X1, Y1, X2, Y2) Then
        'check if already known results
        LandFound = True
        If NewRoute(Site_To, Site_From).NumberOfSegments = 0 Then
            RouteFound = FindFishCorridorForRoute(Site_To, Site_From)
            ForwardDirection = True
            If Not RouteFound Then
                RouteFound = FindFishCorridorForRoute(Site_From, Site_To)
                ForwardDirection = False
            End If
        Else
            RouteExists = True
        End If
    End If
End With

'draw it if found, or if not found or no-land simple draw instead
If RouteFound Then
    
    If Not ForwardDirection Then
        Site_To = SF
        Site_From = ST
    End If
    
    'set route
    With NewRoute(Site_To, Site_From)
        .RouteNumber(0) = RouteMap_Index
        ID = RouteMap_Index
    End With
    
    RouteMap_Index = RouteMap_Index + 1
    With RouteMap(ID)
        NewRoute(Site_To, Site_From).NumberOfSegments = 1
        'next segment goes along the corridor
        p = 1
        .X(0) = DeScaleThisX(Receiver.X(SF))
        .Y(0) = DeScaleThisY(Receiver.Y(SF))
        For N = 0 To LastNode
            .X(p) = ScaleThisX(DepthFirstSearchNode(N).X): .Y(p) = ScaleThisY(DepthFirstSearchNode(N).Y)
            NewRoute(Site_To, Site_From).Distance = NewRoute(Site_To, Site_From).Distance + TrackCalculator.Calculate_DistanceXY(CLng(.X(p - 1)), CLng(.Y(p - 1)), CLng(.X(p)), CLng(.Y(p)))
            p = p + 1
        Next N
        NewRoute(Site_To, Site_From).Distance = NewRoute(Site_To, Site_From).Distance + TrackCalculator.Calculate_DistanceXY(CLng(.X(p - 1)), CLng(.Y(p - 1)), X2, Y2)
        .Index = p
        .X(p) = DeScaleThisX(Receiver.X(ST))
        .Y(p) = DeScaleThisY(Receiver.Y(ST))
    End With
    'ensure these values are stored for reverse route too
    NewRoute(Site_From, Site_To).NumberOfSegments = NewRoute(Site_To, Site_From).NumberOfSegments
    NewRoute(Site_From, Site_To).Distance = NewRoute(Site_To, Site_From).Distance
    If ForwardDirection Then
        NewRoute(Site_To, Site_From).StepForward = True
        NewRoute(Site_From, Site_To).StepForward = False
    Else
        NewRoute(Site_To, Site_From).StepForward = False
        NewRoute(Site_From, Site_To).StepForward = True
    End If
End If

'return value
ReturnValue = False
If RouteFound Then ReturnValue = True
If LandFound And RouteExists Then ReturnValue = True

NavigateBetweenCorridorPoints = ReturnValue


End Function
Public Function NavigateBetweenSites(ByVal Site_To As Long, ByVal Site_From As Long) As Boolean
'See notes on NavigateBetweenReceivers
'
Dim ReturnValue As Boolean

Select Case Land_Avoidance
    Case AvoidanceMode.NotActive
        ReturnValue = False
    Case AvoidanceMode.FishCorridor
        ReturnValue = NavigateBetweenCorridorPoints(Site_To, Site_From)
    Case AvoidanceMode.RandomWalk
        ReturnValue = NavigateBetweenReceivers(Site_To, Site_From)
    Case AvoidanceMode.Shoreline
        ReturnValue = NavigateBetweenReceivers(Site_To, Site_From)
End Select



NavigateBetweenSites = ReturnValue


End Function

Public Function NavigateBetweenReceivers(ByVal Site_To As Long, ByVal Site_From As Long) As Boolean

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'THIS CODE NEEDS TO BE CLEANED AND REFACTORED
'its filthy and unsanitary ;)
'
'I need to do a SELECT CASE type tree and have
'the different kinds of land avoidance on it
'so that this routine accomplishes the selection, and sends the choice
'to 4 different subs (one for shoreline, one for corridors, one for random walks, one for none)
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

'returns TRUE if land is found, false if no land found
Dim result As Boolean
Dim X1 As Long
Dim X2 As Long
Dim Y1 As Long
Dim Y2 As Long
Dim ToX As Long
Dim ToY As Long
Dim FromX As Long
Dim FromY As Long

Dim Found As Boolean
Dim N As Long
Dim w As Long
Dim Size As Long
Dim PrevX As Long
Dim PrevY As Long

Dim LA As Single
Dim LO As Single
Dim La2 As Single
Dim Lo2 As Single

'check if already known results
If NewRoute(Site_To, Site_From).Distance > 0 And Not IgnorePreviousTrackRouteDefinition Then
    result = True
Else
    'show hour glass
    Form1.MousePointer = vbHourglass

    NewRoute(Site_To, Site_From).Distance = 0
    NewRoute(Site_From, Site_To).Distance = 0
    
    'define to and from
    ToX = DeScaleThisX(Receiver.X(CInt(Site_To)))
    ToY = DeScaleThisY(Receiver.Y(CInt(Site_To)))
    FromX = DeScaleThisX(Receiver.X(CInt(Site_From)))
    FromY = DeScaleThisY(Receiver.Y(CInt(Site_From)))
    
    'Clear nodes
    Node_TTL = 0
    
    For N = 0 To MAX_NODES
        Node(N).Segment_X2 = 0
        Node(N).Segment_Y2 = 0
        Node(N).Verified = False
    Next N
    

    'if land found, hatch new line nodes
    Follow_Line ToX, ToY, FromX, FromY
    
    If Node_TTL > 0 Then
        'land found!
        result = True
        
        'do a random walk or follow contour?
        If Land_Avoidance = AvoidanceMode.RandomWalk Then
            Size = RandomWalk(Site_To, Site_From)
            If Size > 0 Then
                AssignRoute Site_To, Site_From, Size
            Else
                'no valid path found
                result = False
            End If
        Else
           'start at node 1
           N = 0
           'first segment starts with receiver
           PrevX = FromX
           PrevY = FromY
           
           Do
               N = N + 1
           
               'get land entry points
               With Node(N)
                   X1 = .Segment_X1
                   X2 = .Segment_X2
                   Y1 = .Segment_Y1
                   Y2 = .Segment_Y2
               End With
               
               'Notify user of attempt
               'Make sure user knows where we are at
               'Form1.StatusBar.Panels(1).Text = Receiver.ID(CInt(Site_From)) & "->" & Receiver.ID(CInt(Site_To))

                'use image analyzer to look at map and follow map contour from land entry to land exit
                Size = ImageProcessingEngine.Redirect(X2, Y2, X1, Y1, PixelMap_X(), PixelMap_Y(), Site_From)
               If Size > 0 Then
                    AssignRoute Site_To, Site_From, Size
                End If
               'store the end points of segment to connect to next segment
               PrevX = X2
               PrevY = Y2
           Loop Until N >= Node_TTL Or Not ImageProcessingEngine.LandBetweenPositions(X2, Y2, Receiver.X(Site_From), Receiver.Y(Site_From))
            
            'Now display track # to show user its ok
            'Form1.StatusBar.Panels(1).Text = "Track " & FishDatabase.Code(CURRENT_FISH)
            
            'add final segment
            If Size > 0 Then
                'forward
                NewRoute(Site_To, Site_From).Distance = NewRoute(Site_To, Site_From).Distance + DistancePixels(ToX, PrevX, ToY, PrevY)
                'reverse
                NewRoute(Site_From, Site_To).Distance = NewRoute(Site_To, Site_From).Distance
            Else
                'no valid path found
                result = False
            End If
        End If
    End If
End If

'show hour glass
Form1.MousePointer = vbArrow

NavigateBetweenReceivers = result

End Function
Public Sub DrawRoute(ByVal Site_To As Long, ByVal Site_From As Long, Box As PictureBox, ColorC As Long)
Dim G As Integer
Dim ReceiversLockedInLand As Boolean

'same?
If Site_To = Site_From Then
    Exit Sub
End If

'part of group
G = Receiver.GroupNumber(Site_To)
If G <> 0 Then Site_To = Receiver.GroupReceiverNumber(G)
G = Receiver.GroupNumber(Site_From)
If G <> 0 Then Site_From = Receiver.GroupReceiverNumber(G)

'Draws route between receivers and navigates to avoid land

'BOTH Receivers have to be "connected" to the water
'so verify first!
ReceiversLockedInLand = Receiver.LandLocked(Site_To) Or Receiver.LandLocked(Site_From)

If ReceiversLockedInLand Then
    JustDrawSimpleTrack Site_To, Site_From, Box, ColorC
Else
    'check if need to reroute
    
    Select Case Land_Avoidance
        Case AvoidanceMode.Shoreline
            DrawRouteBySegments Site_To, Site_From, Box, ColorC
        Case AvoidanceMode.FishCorridor
            DrawRouteBySegments Site_To, Site_From, Box, ColorC
        Case AvoidanceMode.RandomWalk
            DrawRouteBySegments Site_To, Site_From, Box, ColorC
        Case Else
            JustDrawSimpleTrack Site_To, Site_From, Box, ColorC
    End Select
End If

End Sub
Private Function ScaleThisY(ByVal c As Long) As Long
With ZoomRegion
    c = (c - .OriginY) * .ScaleY
End With

ScaleThisY = c

End Function
Private Function ScaleThisX(ByVal c As Long) As Long
With ZoomRegion
    c = (c - .OriginX) * .ScaleX
End With

ScaleThisX = c

End Function
Private Function DeScaleThisX(ByVal c As Long) As Long
'if x is from a zoom area it converts it to regular x on mapimage
With ZoomRegion
    c = (c / .ScaleX) + .OriginX
End With

DeScaleThisX = c

End Function
Private Function DeScaleThisY(ByVal c As Long) As Long
'if x is from a zoom area it converts it to regular x on mapimage
With ZoomRegion
    c = (c / .ScaleY) + .OriginY
End With

DeScaleThisY = c

End Function
Private Sub DrawRouteFishCorridor(ByVal Site_To As Long, ByVal Site_From As Long, Box As PictureBox, ColorC As Long)
Dim X1 As Long
Dim X2 As Long
Dim Y1 As Long
Dim Y2 As Long
Dim ST As Integer
Dim SF As Integer
Dim Found As Boolean

Dim N As Long
Dim i As Long
'check validity
If Site_To = Site_From Then
    Exit Sub
End If

'here its different
ST = CInt(Site_To)
SF = CInt(Site_From)


'check if land present
X2 = DeScaleThisX(Receiver.X(ST))
Y2 = DeScaleThisY(Receiver.Y(ST))
X1 = DeScaleThisX(Receiver.X(SF))
Y1 = DeScaleThisY(Receiver.Y(SF))

With ImageProcessingEngine
    If .LandBetweenPositions(X1, Y1, X2, Y2) Then
        'check if already known results
        If NewRoute(Site_To, Site_From).NumberOfSegments = 0 Then
            Found = FindFishCorridorForRoute(Site_To, Site_From)
            If Not Found Then
                Found = FindFishCorridorForRoute(Site_From, Site_To)
            End If
        Else
            Found = True
        End If
    Else
        Found = False
    End If
End With
'draw it if found, or if not found or no-land simple draw instead
If Found Then
    '0 it first
    'arcaic but just avoiding one more if...then
    With NewRoute(Site_To, Site_From)
        .Distance = 0
        .NumberOfSegments = 0
    End With
    
    'Draw by segments
    'First segment is from receiver to fish corridor
    With ImageProcessingEngine
        X1 = ScaleThisX(.CorridorPoint_X(SF))
        Y1 = ScaleThisX(.CorridorPoint_Y(SF))
    End With
    Box.Line (Receiver.X(SF), Receiver.Y(SF))-(X1, Y1), vbBlue
    
    'add distance
    NewRoute(Site_To, Site_From).Distance = NewRoute(Site_To, Site_From).Distance + TrackCalculator.Calculate_DistanceXY(Receiver.X(SF), Receiver.Y(SF), X1, Y1)
    NewRoute(Site_To, Site_From).NumberOfSegments = NewRoute(Site_To, Site_From).NumberOfSegments + 1
    'next segment goes along the corridor
    For N = 0 To LastNode
        With DepthFirstSearchNode(N)
            X2 = ScaleThisX(.X): Y2 = ScaleThisY(.Y)
        End With
        Box.Line (X1, Y1)-(X2, Y2), ColorC
        NewRoute(Site_To, Site_From).Distance = NewRoute(Site_To, Site_From).Distance + TrackCalculator.Calculate_DistanceXY(X1, Y1, X2, Y2)
        NewRoute(Site_To, Site_From).NumberOfSegments = NewRoute(Site_To, Site_From).NumberOfSegments + 1
        X1 = X2: Y1 = Y2
    Next N
    

    With ImageProcessingEngine
        'then to receiver
        .DrawLineWithArrow Box, X1, Receiver.X(ST), Y1, Receiver.Y(ST), ColorC
        NewRoute(Site_To, Site_From).Distance = NewRoute(Site_To, Site_From).Distance + TrackCalculator.Calculate_DistanceXY(X1, Y1, Receiver.X(ST), Receiver.Y(ST))
        NewRoute(Site_To, Site_From).NumberOfSegments = NewRoute(Site_To, Site_From).NumberOfSegments + 1
    End With
Else
    'normal draw
    JustDrawSimpleTrack Site_To, Site_From, Box, ColorC
End If

'ensure these values are stored
NewRoute(Site_From, Site_To).NumberOfSegments = NewRoute(Site_To, Site_From).NumberOfSegments
NewRoute(Site_From, Site_To).Distance = NewRoute(Site_To, Site_From).Distance
End Sub
Private Function FindFishCorridorForRoute(ByVal Site_To As Long, ByVal Site_From As Long) As Boolean
Dim Found As Boolean
Dim X1 As Long
Dim X2 As Long
Dim Y1 As Long
Dim Y2 As Long
Dim X As Long
Dim Y As Long
Dim XC As Long
Dim YC As Long
Dim direction As Long
Dim N As Long
Dim i As Long
Dim BranchPointFound As Boolean
Dim CorridorEnds As Boolean
Dim Successful As Boolean
Dim VisitedMap(MAX_X, MAX_Y) As Long



'clear stuff
If LastNode = 0 Then LastNode = MAX_NODES
'validate
If Not MapWasScanned Then Exit Function

'Make sure user knows where we are at
'Form1.StatusBar.Panels(1).Text = Receiver.ID(CInt(Site_From)) & "->" & Receiver.ID(CInt(Site_To))


'Prepare nodes
For i = 0 To LastNode
    With DepthFirstSearchNode(i)
        .Count = 0
        .direction(0) = -1
    End With
Next i

LastNode = 0


'
'Get both receiver-to and receiver-from, connecting points to the corridor
'both receivers need to be connected for the algorithm to be successfull
'otherwise return false and do a normal draw

X1 = ImageProcessingEngine.CorridorPoint_X(CInt(Site_From))
Y1 = ImageProcessingEngine.CorridorPoint_Y(CInt(Site_From))
X2 = ImageProcessingEngine.CorridorPoint_X(CInt(Site_To))
Y2 = ImageProcessingEngine.CorridorPoint_Y(CInt(Site_To))

If X1 = -1 Or Y1 = -1 Or X2 = -1 Or Y2 = -1 Then
    Found = False
Else
    X2 = DeScaleThisX(Receiver.X(CInt(Site_To)))
    Y2 = DeScaleThisY(Receiver.Y(CInt(Site_To)))
    'start seartch at x1,y1
    X = X1
    Y = Y1
    Do
        Do
            direction = 0
            BranchPointFound = False
            CorridorEnds = True
            
            'mark as visited!
            VisitedMap(X, Y) = 1
            Do
                XC = X + K_X(direction): YC = Y + K_Y(direction)
                
                If FishCorridorMap(XC, YC) = 1 And VisitedMap(XC, YC) = 0 Then
                'found a line-of-sight?
                Found = Not (ImageProcessingEngine.LandBetweenPositions(XC, YC, X2, Y2))
                    'if limit is passed, abort
                    If LastNode > 200 Then
                        Found = False
                        Exit Do
                    End If
                    
                    'store direction
                    With DepthFirstSearchNode(LastNode)
                        .X = X
                        .Y = Y
                        .direction(.Count) = direction
                        If .Count = 1 Then BranchPointFound = True
                        .Count = .Count + 1
                        CorridorEnds = False
                    End With
                End If
                direction = direction + 1
            Loop Until direction > 8 Or Found
        
            'if more to explore then explore at last known direction that works
            If Not CorridorEnds And Not Found Then
                With DepthFirstSearchNode(LastNode)
                    X = X + K_X(.direction(0))
                    Y = Y + K_Y(.direction(0))
                    .direction(0) = -1
                    'is this the only direction on this node?
                    If .Count = 1 Then .Count = 0
                End With
            End If
            'if branch point advance to next node
            If BranchPointFound Then
                LastNode = LastNode + 1
            End If
        
        Loop Until Found Or CorridorEnds

        If Not Found Then
            Successful = False
            'get direction from last node
            N = LastNode - 1
            Do
                With DepthFirstSearchNode(N)
                    i = 0
                    Do
                        'get a direction
                        If .direction(i) <> -1 Then
                            Successful = True
                            X = .X + K_X(.direction(i))
                            Y = .Y + K_Y(.direction(i))
                            direction = .direction(i)
                            .direction(i) = -1
                        End If
                        i = i + 1
                    Loop Until i >= .Count Or Successful
                End With
                
                If Not Successful Then
                    'kill node
                    DepthFirstSearchNode(N).Count = 0
                    LastNode = LastNode - 1
                    N = N - 1
                End If
            
            Loop Until Successful Or N < 0
            
        End If
        
    Loop Until Found Or Not Successful

End If

'Now display track # to show user its ok
'Form1.StatusBar.Panels(1).Text = "Track " & FishDatabase.Code(CURRENT_FISH)

FindFishCorridorForRoute = Found
        
End Function
Private Sub JustDrawSimpleTrack(ByVal Site_To As Long, ByVal Site_From As Long, Box As PictureBox, ColorC As Long)
Dim X1 As Long
Dim X2 As Long
Dim Y1 As Long
Dim Y2 As Long
'Normal Draw
X1 = Receiver.X(Site_To)
Y1 = Receiver.Y(Site_To)
X2 = Receiver.X(Site_From)
Y2 = Receiver.Y(Site_From)

ImageProcessingEngine.DrawLineWithArrow Box, X2, X1, Y2, Y1, ColorC
End Sub
Private Function RandomWalk(ByVal Site_To As Integer, ByVal Site_From As Integer) As Integer
'performs a couple of random walks and chooses the shortest

Dim Success As Boolean
Dim i As Long
Dim RouteLen As Long
Dim Max As Long
Dim Attempt As Long
Dim SelectedRoute As Long

Const MaxAttempts = 100


Dim SuccessfulRouteX(2500, MaxAttempts) As Long
Dim SuccessfulRouteY(2500, MaxAttempts) As Long

'set value of max base
Max = 3000
SelectedRoute = -1

Do
    RouteLen = ImplementRandomWalk(Site_To, Site_From, PixelMap_X(), PixelMap_Y())
    If RouteLen = -1 Then
        Success = False
    Else
        Success = True
        'store
        For i = 0 To RouteLen
            SuccessfulRouteX(i, Attempt) = PixelMap_X(i)
            SuccessfulRouteY(i, Attempt) = PixelMap_Y(i)
        Next i
        If RouteLen < Max Then
            SelectedRoute = Attempt
            Max = RouteLen
        End If
    End If
    Attempt = Attempt + 1
    'Release to the OS between tries
    DoEvents
Loop Until Attempt > MaxAttempts

'was a route selected?
If SelectedRoute = -1 Then Success = False Else Success = True

'if so, store as a new route
If Success Then
    For i = 0 To Max
        PixelMap_X(i) = SuccessfulRouteX(i, SelectedRoute)
        PixelMap_Y(i) = SuccessfulRouteY(i, SelectedRoute)
    Next i
    Max = Max + 1
    PixelMap_X(Max) = DeScaleThisX(Receiver.X(Site_To))
    PixelMap_Y(Max) = DeScaleThisY(Receiver.Y(Site_To))
Else
    Max = 0
End If
        
RandomWalk = Max

End Function
Private Function ImplementRandomWalk(ByVal Site_To As Integer, ByVal Site_From As Integer, ByRef RouteX() As Long, ByRef RouteY() As Long) As Long
Dim X1 As Long
Dim X2 As Long
Dim Y1 As Long
Dim Y2 As Long
Dim Tries As Long
Dim Success As Boolean
Dim Failure As Boolean
Dim direction As Currency
Dim StepsOverLand As Long
Dim RX As Long
Dim RY As Long
Dim i As Long

Dim TotalSteps As Long
Dim TurnAround As Boolean
Dim Persistance As Single
Dim MAX_STEPS As Long

'constants
Dim StepSize As Long
'show hour glass
Form1.MousePointer = vbHourglass
'Make sure user knows where we are at
'Form1.StatusBar.Panels(1).Text = Receiver.ID(CInt(Site_From)) & "->" & Receiver.ID(CInt(Site_To))

StepSize = Nav_Segment
If StepSize = 0 Then StepSize = 1
MAX_STEPS = SEARCHRADIUS / StepSize

Const MaxStepOverLand = 4
Const TotalTries = 25

RND_MT.Randomize Timer
'destination
RX = DeScaleThisX(Receiver.X(Site_To))
RY = DeScaleThisY(Receiver.Y(Site_To))
'inform user
Randomize Timer
Do Until Tries > TotalTries Or Success
    'check if land present
    X1 = DeScaleThisX(Receiver.X(Site_From))
    Y1 = DeScaleThisY(Receiver.Y(Site_From))
    TotalSteps = 0
    RouteX(TotalSteps) = X1
    RouteY(TotalSteps) = Y1
    Failure = False
    Tries = Tries + 1
    
    Do Until Success Or Failure
        'define x2,y2
        
        'if hit land, turnaround
        'otherwise, randomize direction
        If TurnAround Then
            direction = direction + 1
            If direction > 8 Then direction = 1
            TurnAround = False
        Else
            If Rnd(Timer) > Persistance_TH Then direction = RND_MT.Generate_Number(8) + 1
        End If
        
        X2 = X1 + StepSize * K_X(direction)
        Y2 = Y1 + StepSize * K_Y(direction)
        
        'land?
        With ImageProcessingEngine
            If .LandBetweenPositions(X1, Y1, X2, Y2) Then
                StepsOverLand = StepsOverLand + 1
                If StepsOverLand > MaxStepOverLand Then Failure = True
                TurnAround = True
            Else
                'if no land
                'test direct route to receiver now
                If Not .LandBetweenPositions(X2, Y2, RX, RY) Then
                    Success = True
                End If
                TotalSteps = TotalSteps + 1
                If TotalSteps > MAX_STEPS Then Failure = True: Success = False
                RouteX(TotalSteps) = X2
                RouteY(TotalSteps) = Y2
                X1 = X2
                Y1 = Y2
                StepsOverLand = 0
            End If
        End With
    Loop
Loop

If Not Success Then TotalSteps = -1

'show arrow
Form1.MousePointer = vbArrow
'Now display track # to show user its ok
'Form1.StatusBar.Panels(1).Text = "Track " & FishDatabase.Code(CURRENT_FISH)

ImplementRandomWalk = TotalSteps
End Function
Private Sub DrawRouteBySegments(ByVal Site_To As Long, ByVal Site_From As Long, Box As PictureBox, ColorC As Long)
Dim X1 As Long
Dim X2 As Long
Dim Y1 As Long
Dim Y2 As Long
Dim i As Long
Dim Pixel As Long
Dim RouteID As Long
Dim t As Long
Dim LAND_ENCOUNTERD As Boolean

'land?
LAND_ENCOUNTERD = NavigateBetweenSites(Site_To, Site_From)

'uses shoreline land avoidance
'to draw lines connecting receivers
If LAND_ENCOUNTERD And NewRoute(Site_To, Site_From).NumberOfSegments > 0 Then
    'set for rnd walk or shoreline
    If Land_Avoidance = AvoidanceMode.RandomWalk Or Land_Avoidance = AvoidanceMode.FishCorridor Then
        DrawSingleSegment Site_To, Site_From, Box, ColorC
    Else
        DrawMultipleSegments Site_To, Site_From, Box, ColorC
    End If
Else
    'normal draw
    X1 = Receiver.X(Site_To)
    Y1 = Receiver.Y(Site_To)
    X2 = Receiver.X(Site_From)
    Y2 = Receiver.Y(Site_From)
    ImageProcessingEngine.DrawLineWithArrow Box, X2, X1, Y2, Y1, ColorC
End If
End Sub
Private Sub DrawSingleSegment(ByVal Site_To As Long, ByVal Site_From As Long, Box As PictureBox, ColorC As Long)
Dim X1 As Long
Dim X2 As Long
Dim Y1 As Long
Dim Y2 As Long
Dim i As Long
Dim StartPixel As Long
Dim EndPixel As Long
Dim direction As Long
Dim FirstPixel As Boolean
Dim RouteID As Long
Dim Pixel As Long
With NewRoute(Site_To, Site_From)
    For i = 0 To .NumberOfSegments - 1
        RouteID = .RouteNumber(i)
        'draw direction
        If .StepForward Then
            StartPixel = 0
            EndPixel = RouteMap(RouteID).Index - 1
            direction = 1
        Else
            StartPixel = RouteMap(RouteID).Index
            EndPixel = 1
            direction = -1
        End If
        'draw segments from receiver
        With RouteMap(RouteID)
            FirstPixel = True
            For Pixel = StartPixel To EndPixel Step direction
                X2 = ScaleThisX(CLng(.X(Pixel))): Y2 = ScaleThisY(CLng(.Y(Pixel)))
                If Not FirstPixel Then Box.Line (X1, Y1)-(X2, Y2), ColorC
                X1 = X2: Y1 = Y2: FirstPixel = False
            Next Pixel
            'last segment has arrow
            X2 = ScaleThisX(.X(EndPixel + direction))
            Y2 = ScaleThisY(.Y(EndPixel + direction))
            ImageProcessingEngine.DrawLineWithArrow Box, X1, X2, Y1, Y2, ColorC
        End With
    Next i
End With
End Sub
Private Sub DrawMultipleSegments(ByVal Site_To As Long, ByVal Site_From As Long, Box As PictureBox, ColorC As Long)
Dim X1 As Long
Dim X2 As Long
Dim Y1 As Long
Dim Y2 As Long
Dim i As Long
Dim Pixel As Long
Dim SegmentLength As Long
Dim RouteID As Integer

'length of segments
SegmentLength = Nav_Segment

'first segment goes from receiver to land entry
X1 = Receiver.X(Site_From)
Y1 = Receiver.Y(Site_From)
'reroute!
With NewRoute(Site_To, Site_From)
    For i = 0 To .NumberOfSegments - 1
        RouteID = .RouteNumber(i)
        With RouteMap(RouteID)
            'connecting segment goes from last segment to start of next
            If NewRoute(Site_To, Site_From).StepForward Then
                X2 = ScaleThisX(.X(.Index - 1))
                Y2 = ScaleThisY(.Y(.Index - 1))
            Else
                X2 = ScaleThisX(.X(0))
                Y2 = ScaleThisY(.Y(0))
            End If
            
            Box.Line (X1, Y1)-(X2, Y2), ColorC
            
            X2 = ScaleThisX(.X(0))
            Y2 = ScaleThisY(.Y(0))
        
            For Pixel = 1 To .Index - 1 Step SegmentLength
                'each line starts where next one ended
                X1 = X2
                Y1 = Y2
                'and ends on next
                X2 = ScaleThisX(.X(Pixel))
                Y2 = ScaleThisY(.Y(Pixel))
               Box.Line (X1, Y1)-(X2, Y2), ColorC
            Next Pixel
            
            Box.Line (ScaleThisX(.X(.Index - 1)), ScaleThisY(.Y(.Index - 1)))-(X2, Y2), ColorC
            
            'start next segment where this one ended
            X1 = X2
            Y1 = Y2
        End With
    Next i
    'last segment goes to receiver
    X2 = Receiver.X(Site_To)
    Y2 = Receiver.Y(Site_To)
    With RouteMap(RouteID)
        If NewRoute(Site_To, Site_From).StepForward Then
            ImageProcessingEngine.DrawLineWithArrow Box, ScaleThisX(.X(0)), X2, ScaleThisY(.Y(0)), Y2, ColorC
        Else
            ImageProcessingEngine.DrawLineWithArrow Box, ScaleThisX(.X(.Index - 1)), X2, ScaleThisY(.Y(.Index - 1)), Y2, ColorC
        End If
    End With
End With
End Sub
Private Function Node_Is_Valid(N As Long) As Boolean
'pythagorean distance has to be larger than a threshold value
Dim hyp As Double
Dim a As Single
Dim B As Single
Dim Validity As Boolean


With Node(N)
    a = (.Segment_X1 - .Segment_X2) ^ 2
    B = (.Segment_Y1 - .Segment_Y2) ^ 2
    hyp = Sqr(a + B)

    If hyp >= SEGMENT_SIZE_THRESHOLD And .Segment_X2 > 0 And .Segment_Y2 > 0 Then
        Validity = True
    Else
        Validity = False
    End If
End With

Node_Is_Valid = Validity

    
End Function
Private Sub Follow_Line(ByVal New_x As Single, ByVal New_y As Single, ByVal Old_X As Single, ByVal Old_Y As Single)
'Alerts if line is going over LAND
Dim m As Single
Dim X As Long
Dim Y As Long

Dim xi As Long
Dim yi As Long

Dim increment_x As Single
Dim increment_y As Single


Dim Sign_X As Long
Dim Sign_Y As Long

Dim Navigation As Boolean
Dim i As Long
Dim Past_Results As Boolean
Dim p As Long


Const OVER_WATER = False
Const OVER_LAND = True

'check if move occured; if not, no line exists
If New_x = Old_X And New_y = Old_Y Then
    Exit Sub
End If


'all begin in same point
X = Old_X
Y = Old_Y

'loop direction
Sign_X = Sgn(New_x - Old_X)
Sign_Y = Sgn(New_y - Old_Y)




If New_x = Old_X Then
    increment_x = 0
    increment_y = Sgn(New_y - Old_Y)
Else
    m = (New_y - Old_Y) / (New_x - Old_X)
    Select Case Abs(m)
        Case 0
            increment_x = Sgn(New_x - Old_X)
            increment_y = 0
        Case 1
            increment_x = Sgn(New_x - Old_X)
            increment_y = Sgn(New_y - Old_Y)
        Case Is > 1
            increment_y = Sgn(New_y - Old_Y)
            X = New_x
        Case Is < 1 And (m <> 0)
            increment_x = Sign_X
            Y = New_y
    End Select
End If
    
Do
    Do
        Select Case Abs(m)
            Case Is > 1
                increment_x = -1 * (New_y - Y) / m
                X = New_x
            Case Is < 1
                increment_y = -1 * (New_x - X) * m
                Y = New_y
        End Select
        X = CLng(X + increment_x)
        Y = CLng(Y + increment_y)
          
    'check position to see if over land or water
    
    p = MapImage(X, Y)
    Past_Results = Navigation
    If p = 0 Then
        Navigation = OVER_LAND
    Else
        Navigation = OVER_WATER
    End If
      
    'compare to last result.  any change is an exit or entry point to land
    If Navigation = OVER_LAND And Past_Results = OVER_WATER Then
        Node_TTL = Node_TTL + 1
        With Node(Node_TTL)
            .Segment_X1 = X
            .Segment_Y1 = Y
            .Verified = True
        End With
    End If
    If Navigation = OVER_WATER And Past_Results = OVER_LAND Then
        With Node(Node_TTL)
            .Segment_X2 = X
            .Segment_Y2 = Y
        End With
        'validate node
        'if not "valid" (i.e. > TH) then delete and ignore entry and exit points
        If Not Node_Is_Valid(Node_TTL) Then Node_TTL = Node_TTL - 1
        
    End If
    
    Loop Until Sign_X * X >= Sign_X * New_x
Loop Until Sign_Y * Y >= Sign_Y * New_y


End Sub
Public Sub Load()
'Loads receiver table from loaded table file in receiverentry class
Dim ReceiverNumber As Integer
Dim LO As Single
Dim LA As Single

'Stores receiver's info in a table
For ReceiverNumber = 1 To Receiver_Table.TotalReceiversLoaded
    LO = Receiver_Table.LO(ReceiverNumber)
    LA = Receiver_Table.LA(ReceiverNumber)
    
    'Receiver code in map is +1 from code in table
    DrawReceiver_InMemory ReceiverNumber
    
    'Get Boundaries for Receiver Coverage
    If LA > UPPER_La Then UPPER_La = LA
    If LA < LOWER_La Then LOWER_La = LA
    If LO > UPPER_Lo Then UPPER_Lo = LO
    If LO < LOWER_Lo Then LOWER_Lo = LO
    'ttl
    SizeReceiverData(ReceiverNumber) = Receiver_Table.Pings(ReceiverNumber)
Next ReceiverNumber

End Sub
Private Sub ReadReceiverTable(s As String)
Dim Fields() As String
Dim N As Long
Dim Pos As Long
Dim R As Integer

'first find if N has been defined
'for now do nothing with it.
Pos = InStr(1, s, "N=")
If Pos > 0 Then Exit Sub

'break into fields
Fields = Split(s, ",")

'store in table
R = Store(CSng(Fields(3)), CSng(Fields(2)), Fields(1), "", CInt(Fields(0)))

'load other variables into it
Receiver_Table.Notes(R) = Fields(7)
Receiver_Table.DeployDate(R) = Fields(8)
Receiver_Table.RetrieveDate(R) = Fields(9)
Receiver_Table.Model(R) = Fields(10)


End Sub
Public Sub Save(FileName As String)
'Saves Receiver Table to file path FileName
'
'(Extension .RT)
Dim FileNumber As Long
Dim R As Integer
Dim NumberOfReceivers As Long
Dim Concatenated As String
Dim G As Long
Dim t As Long


Const COMMA = ","

'Get number of receivers
NumberOfReceivers = TTL_Receivers

'Choose any file number that is free
FileNumber = FreeFile

'Open for output (append)
Open FileName For Append As #FileNumber

'Write control block
Print #FileNumber, "[/CONTROL POINTS/]"
Print #FileNumber, "[/Scale X=" & Str$(Scale_X) & "/]"
Print #FileNumber, "[/Scale Y=" & Str$(Scale_Y) & "/]"
Print #FileNumber, "[/Origin Lat=" & Str$(Origin_Lat) & "/]"
Print #FileNumber, "[/Origin Long=" & Str$(Origin_Long) & "/]"
Print #FileNumber, "[/RECEIVER TABLE/]"
Print #FileNumber, "[/N=" & Str$(NumberOfReceivers) & "/]"

'Write a header
Print #FileNumber, "&&RECEIVER NUMBER, RECEIVER NAME, LAT, LONG, X, Y, GROUP NUMBER, NOTES, DEPLOY DATE, RETRIEVE DATE, MODEL, TAG, ZONE"

'Write Receiver information
For R = 1 To NumberOfReceivers
    Concatenated = Str$(R)
    Concatenated = Concatenated & COMMA & Receiver_Table.ID(R) & COMMA & Str$(Receiver_Table.LA(R)) & COMMA & Str$(Receiver_Table.LO(R)) & COMMA & Str$(Receiver_Table.X(R)) & COMMA & Str$(Receiver_Table.Y(R))
    Concatenated = Concatenated & COMMA & Str$(Receiver_Table.GroupNumber(R)) & COMMA & Receiver_Table.Notes(R) & COMMA & Receiver_Table.DeployDate(R) & COMMA & Receiver_Table.RetrieveDate(R) & COMMA & Receiver_Table.Model(R) & COMMA & Str$(Receiver_Table.Tag(R)) & COMMA & Str$(Receiver_Table.Zone(R))
    Print #FileNumber, Concatenated
Next R

'Groups
'any active navigational groups?
If Nav_Group_Index > 0 Then
    'show in another table
    Print #FileNumber, "[/GROUPS/]"
    Print #FileNumber, "[/N=" & Str$(Nav_Group_Index) & "/]"
    Print #FileNumber, "&&GROUP NUMBER, GROUP NAME, RECEIVERS IN GROUP"
    For G = 1 To Nav_Group_Index
        With Nav_Group(G)
            Concatenated = Str$(G) & COMMA & .ID
            For R = 0 To .ReceiverIndex - 1
                Concatenated = Concatenated & COMMA & Str$(.Receiver(R))
            Next R
            'output
            Print #FileNumber, Concatenated
        End With
    Next G
End If

'Geo/Tags
'any tags assigned?
If NumberOfAssignedTags > 0 Then
    Print #FileNumber, "[/TAGS/]"
    Print #FileNumber, "&&TAG NUMBER, UX, UY, LX, LY, ULa, ULo, LLa, LLo"
    For t = 1 To NumberOfAssignedTags
        With Receiver_Zones(t)
            Concatenated = Str$(.TagNumber)
            Concatenated = Concatenated & COMMA & Str$(.UpperX) & COMMA & Str$(.UpperY) & COMMA & Str$(.LowerX) & COMMA & Str$(.LowerY) & COMMA & Str$(.UpperLat) & COMMA & Str$(.UpperLong) & COMMA & Str$(.LowerLat) & COMMA & Str$(.LowerLong)
            Print #FileNumber, Concatenated
        End With
    Next t
End If

Print #FileNumber, "[/End Section/]"

Close #FileNumber

End Sub
Public Function CountUniqueEntriesInReceiver_ByDielPhase(ReceiverNumber As Integer, Phase As Long)
Dim ItemToAdd As Integer
Dim FishIsInList(MAX_FISH) As Boolean
Dim Count As Long
Dim Found As Boolean
Dim s As Long

'read stamps
For s = 1 To Receiver.Detection_TTL(ReceiverNumber)
    Receiver.ReadStamp ReceiverNumber, s
    If PhotoPeriodCalculator.StampIsInPhase(Phase) And Stamp.Valid Then
        If (Not FishIsInList(Stamp.Fish)) And FishDatabase.IsVisible(Stamp.Fish) Then
            FishIsInList(Stamp.Fish) = True
            Count = Count + 1
        End If
    End If
 Next s
 
'return number of fish in receiver that are within time slot
CountUniqueEntriesInReceiver_ByDielPhase = Count
    
End Function
Public Sub AddReceiverToUserDefinedTrack(R As Long)

UserDefinedTrack(Index_to_UserDefinedTrack) = R
Index_to_UserDefinedTrack = Index_to_UserDefinedTrack + 1

End Sub
Public Sub UserDefinedTrack_Reset()
Index_to_UserDefinedTrack = 0
End Sub
Public Property Get UserDefinedTrack_Receiver(i As Long) As Long
Dim R As Long


If i < Index_to_UserDefinedTrack Then
    R = UserDefinedTrack(i)
Else
    R = -1 'no more receivers
End If

UserDefinedTrack_Receiver = R
End Property
Public Property Get UserDefinedTrack_NumberOfReceiversInTrack() As Long
'Returns the index to user defined track
UserDefinedTrack_NumberOfReceiversInTrack = Index_to_UserDefinedTrack
End Property
Public Sub DrawUserDefinedTrack(Pic As PictureBox)
Dim R As Integer
Dim SiteFrom As Long
Dim SiteTo As Long
If Index_to_UserDefinedTrack > 1 Then
    Pic.Cls
    'draws user defined track
    For R = 1 To Index_to_UserDefinedTrack - 1
        DrawRoute UserDefinedTrack(R), UserDefinedTrack(R - 1), Pic, vbRed
    Next R
End If
End Sub
Public Sub DrawZone(ByVal TagID As Long, Box As PictureBox, Source As PictureBox)
'draw zone in a window
Dim LocalScaleX As Single
Dim LocalScaleY As Single
Dim BoxWidth As Long
Dim BoxHeight As Long

'get scales
BoxWidth = Box.ScaleWidth
BoxHeight = Box.ScaleHeight

LocalScaleX = BoxWidth / Source.ScaleWidth
LocalScaleY = BoxHeight / Source.ScaleHeight

'draw
With Receiver_Zones(TagID)
    Box.Line (.LowerX * LocalScaleX, .LowerY * LocalScaleY)-(.UpperX * LocalScaleX, .UpperY * LocalScaleY), vbBlue, B
End With

End Sub
Public Sub MakeInvisible(Optional R As Integer = -1)
If R = -1 Then Receiver_Table.SetAllReceiversToINVISIBLE Else Receiver_Table.Invisible(R) = True
End Sub
Public Sub MakeVisible(Optional R As Integer = -1)
If R = -1 Then Receiver_Table.SetAllReceiversToVISIBLE Else Receiver_Table.Invisible(R) = False
End Sub
Private Sub Class_Initialize()
'screen scale
With ZoomRegion
    .ScaleX = 1: .ScaleY = 1: .OriginX = 0: .OriginY = 0
End With

'define cells around central pixel
K_X(0) = 0: K_X(1) = 1: K_X(2) = 1: K_X(3) = 0: K_X(4) = -1: K_X(5) = -1: K_X(6) = -1: K_X(7) = 0: K_X(8) = 1
K_Y(0) = 0: K_Y(1) = 0: K_Y(2) = -1: K_Y(3) = -1: K_Y(4) = -1: K_Y(5) = 0: K_Y(6) = 1: K_Y(7) = 1: K_Y(8) = 1
End Sub
Public Function CountUniqueEntries(EntryNumber As Integer, QueryField As Long, Optional range As Long = 0) As Long
'
'
'Depending on the queryfield, the search engine is optimized to search faster...
'Yay for Polymorphism and total encapsulation! ;)
Dim Counts As Long

Select Case QueryField
    Case FieldNames.Fish
        Counts = CountFishInReceiver(EntryNumber)
    Case Else
        Counts = CountGeneric(EntryNumber, QueryField, range)
End Select

'return value
CountUniqueEntries = Counts
End Function
Private Function CountFishInReceiver(R As Integer) As Long
'Search the receiver field (or any other stamp-formatted table entry) for # of fish
Dim FishIsInTable(MAX_FISH) As Boolean
Dim range As Long
Dim i As Long
Dim Count As Long

For i = 1 To Receiver.Detection_TTL(R)
    Receiver.ReadStamp R, i
    
    If (Not FishIsInTable(Stamp.Fish)) And Stamp.Valid And FishDatabase.IsVisible(Stamp.Fish) Then
        FishIsInTable(Stamp.Fish) = True
        Count = Count + 1
    End If
Next i

'return value
CountFishInReceiver = Count

End Function

Private Function CountGeneric(ByVal EntryNumber As Long, QueryField As Long, range As Long) As Long
'count unique entries in a table range
'adapted to receiver only
Dim i As Long
Dim ItemToAdd As String
Dim TempList(MAX_ENTRIES) As String
Dim LastItem As Long
Dim Item As Long
Dim Found As Boolean

For i = 1 To range
    Receiver.ReadStamp EntryNumber, i
    If Stamp.Valid Then
        ItemToAdd = Choose(QueryField, FishDatabase.Code(Stamp.Fish), Receiver.ID(Stamp.Site), Convert_DayNumber(Stamp.Date), Convert_ToStandardTime(Stamp.Time), Convert_DayNumber(Stamp.Date) & "  " & Convert_ToStandardTime(Stamp.Time), Str$(Stamp.Fish) & " on " & Convert_DayNumber(Stamp.Date) & " at " & Convert_ToStandardTime(Stamp.Time))
            
        'search item in list
        Found = False
        Item = 0
        Do
            If ItemToAdd = TempList(Item) Then
                'if already in list, do not add
                Found = True
            End If
            Item = Item + 1
        Loop While (Not Found) And Item <= LastItem
        'else add
        If Not Found Then
            TempList(LastItem) = ItemToAdd
            LastItem = LastItem + 1
        End If
    End If
Next i

CountGeneric = LastItem
End Function
Public Sub QueryEntry_To_GetField(ByVal R As Integer, Query As String, ByVal QueryField As Long, ByVal GetField As Long, ByRef List() As String)
'Queries Field QueryField on EntryNumber thru Range for Field GetField
'
Dim range As Long
Dim i As Long
Dim response As String
Dim ItemToAdd As String
Dim LastItem As Long
Dim Item As Long
Dim Found As Boolean

Const RETURN_RECEIVER_NAME_ONLY = True

range = Receiver.Detection_TTL(R)
For i = 1 To range
    Receiver.ReadStamp R, i
    If Stamp.Valid Then
        response = Choose(QueryField, FishDatabase.Code(Stamp.Fish), Receiver.ID(Stamp.Site, RETURN_RECEIVER_NAME_ONLY), Convert_DayNumber(Stamp.Date), Convert_ToStandardTime(Stamp.Time), Convert_DayNumber(Stamp.Date) & "  " & Convert_ToStandardTime(Stamp.Time), Str$(Stamp.Fish) & " on " & Convert_DayNumber(Stamp.Date) & " at " & Convert_ToStandardTime(Stamp.Time), Str$(Stamp.Date))
        If Query = response Then
            ItemToAdd = Choose(GetField, FishDatabase.Code(Stamp.Fish), Receiver.ID(Stamp.Site, RETURN_RECEIVER_NAME_ONLY), Convert_DayNumber(Stamp.Date), Convert_ToStandardTime(Stamp.Time), Convert_DayNumber(Stamp.Date) & "  " & Convert_ToStandardTime(Stamp.Time), Str$(Stamp.Fish) & " on " & Convert_DayNumber(Stamp.Date) & " at " & Convert_ToStandardTime(Stamp.Time), Str$(Stamp.Date))
        End If
        'search item in list
        Found = False
        Item = 0
        Do
            'if already in list or item deleted then do not add
            If ItemToAdd = List(Item) Or ItemToAdd = "" Then
                Found = True
            End If
            Item = Item + 1
        Loop While (Not Found) And Item <= LastItem
        'else add
        If Not Found Then
            List(LastItem) = ItemToAdd
            LastItem = LastItem + 1
        End If
    End If
Next i

End Sub
Public Sub TransferUniqueFishEntriesToList(ByVal R As Integer, FishIsInTable() As Boolean)
'Search the receiver field for # of fish
'transfers to a list/updates list
'
'List is a T/F of if fish with code INDEX is present or not....

Dim FishNumber As Integer
Dim range As Long
Dim i As Long
Dim GN As Long


range = Receiver.Detection_TTL(R)

For i = 0 To range - 1
    Receiver.ReadStamp R, i
    FishNumber = Stamp.Fish
    If (Not FishIsInTable(FishNumber)) And Stamp.Valid And FishDatabase.IsVisible(FishNumber) Then
        FishIsInTable(FishNumber) = True
    End If
Next i

End Sub
Public Sub TransferDatesToList(R As Integer, List As ListBox)

'Transfer contents of date field to a list
'
Dim i As Long
Dim range As Long
Dim ItemToAdd As String
Dim TempList(MAX_ENTRIES) As String
Dim LastItem As Long
Dim Item As Long
Dim Found As Boolean
Dim LastItemAdded As String

'if list contains something, transfer contents
If List.ListCount >= 1 Then
    For i = 0 To List.ListCount - 1
        TempList(i) = List.List(i)
    Next i
End If

'last item in list
LastItem = i

'get range
range = Receiver.Detection_TTL(R)


'step thru list
For i = 0 To range - 1
    Receiver.ReadStamp R, i
    If Stamp.Valid Then
        ItemToAdd = Convert_DayNumber(Stamp.Date)
        
        'search item in list
        Found = False
        Item = 0
        If ItemToAdd = LastItemAdded Then
            Found = True
        Else
            Do
                If ItemToAdd = TempList(Item) Then
                    'if already in list, do not add
                    Found = True
                End If
                Item = Item + 1
            Loop While (Not Found) And Item <= LastItem
        End If
        
        'else add
        If Not Found Then
            TempList(LastItem) = ItemToAdd
            LastItem = LastItem + 1
            'also add to list
            List.AddItem ItemToAdd
            LastItemAdded = ItemToAdd
        End If
    End If
Next i


End Sub
Public Function LookUpMapImage(ByVal X As Single, ByVal Y As Single) As Long
'gets map image value
'makes sure to scale it up/down to actual x,y

If ZoomRegion.Zoomed Then
    X = DeScaleThisX(CLng(X))
    Y = DeScaleThisY(CLng(Y))
End If

LookUpMapImage = MapImage(X, Y)

End Function
Public Sub SaveRoute(R As Integer, RR As Integer, FN As Long)
'saves route R,RR to path FN
'Each route is labeled as R--->RR
'followed by each segment
'Each pixel is written as X/Y and each segment is a line
'Buffer number for file is FN
Dim i As Long
Dim d As Single
Dim s As String
Dim p As Long
Dim RouteID As Integer

d = NewRoute(R, RR).Distance

'only output something to buffer is there is a route to output, otherwise NOP
If d > 0 Then
    'only do forward route
    'reverse route will be done automatically when loaded
    With NewRoute(R, RR)
        If .StepForward Then
            For i = 0 To .NumberOfSegments - 1
                'label each route
                Print #FN, Str$(R) & ">" & Str$(RR)
                Print #FN, Str$(d)
                RouteID = .RouteNumber(i)
                s = s & "[" & Str$(RouteID) & ":"
                'concatenate string with route map
                For p = 0 To RouteMap(RouteID).Index
                    s = s & Str$(RouteMap(RouteID).X(p)) & "/" & Str$(RouteMap(RouteID).Y(p))
                    'if last one, don't add ,
                    If p <> RouteMap(RouteID).Index Then s = s & ","
                Next p
                s = s & "]"
            Next i
            Print #FN, s
        End If
    End With
End If
End Sub
Public Property Let LoadRoute(R As Integer, RR As Integer, d As Single, Buffer As String)
Dim i As Long
Dim s As String
Dim p As Long
Dim RouteID As Integer
Dim L As Long
Dim Segment As Long
Dim first_char As Long
Dim last_char As Long
ReDim SplitString(ROUTE_LEN) As String
ReDim coords(2) As String

'forward route
NewRoute(R, RR).Distance = d
'backroute
NewRoute(RR, R).Distance = d

p = 1
L = Len(Buffer)

'First segment is 0
Segment = 0
'get position of next segment block
p = InStr(p, Buffer, "[")
Do While p > 0
    first_char = p + 1
    p = InStr(p, Buffer, "]")
    last_char = p - 1
    'segment begins with route id
    p = InStr(first_char, Buffer, ":")
    'route id
    RouteID = CInt(Mid(Buffer, first_char, p - first_char))
    'forward
    NewRoute(R, RR).RouteNumber(Segment) = RouteID
    NewRoute(R, RR).StepForward = True
    'reverse
    NewRoute(RR, R).RouteNumber(Segment) = RouteID
    NewRoute(RR, R).StepForward = False
    'get segment out of buffer string
    s = Mid(Buffer, p + 1, last_char - p)
    SplitString = Split(s, ",")
    'store coordinates in route map of routeid
    With RouteMap(RouteID)
        .Index = UBound(SplitString)
        For i = 0 To UBound(SplitString)
            coords = Split(SplitString(i), "/")
            .X(i) = CInt(coords(0))
            .Y(i) = CInt(coords(1))
        Next i
    End With
    'advance to next segment
    Segment = Segment + 1
    'advance carret
    p = InStr(last_char, Buffer, "[")
Loop
'forward route
NewRoute(R, RR).NumberOfSegments = Segment
'backroute
NewRoute(RR, R).NumberOfSegments = Segment
End Property
