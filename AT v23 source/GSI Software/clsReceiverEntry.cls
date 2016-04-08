VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReceiverEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Receiver table manager
'gets info from a direct file that contains the receiver information
'
Const TableFileName = "AT_ReceiverTable.RET"
Const BackupTableFileName = "Backup.RET"
Const HeaderSize = 3 'Header is 3 lines
Dim TableOpen As Boolean
Const FileNumberForAnimation = 100
Dim CurrentSourceDataFile As String
Dim TotalEntries As Long

Const NumberofMarkerTypes = 4

Const Len_eID = 40
Const Len_eNotes = 25
Const Len_eDeployDate = 10
Const Len_eRetrieveDate = 10
Const Len_eModel = 10

Private Type XY_Entry
    eID As Long
    ePings As Long
    eDetection_TTL As Long
    eX As Integer
    eY As Integer
    eLo As Single
    eLa As Single
    eMarker As Byte
    eColor As Long
    eZone As Integer
    eTag As Integer
    eInvisible As Boolean
End Type
    
Private Type Entry
    eID As String * Len_eID
    eGroupNumber As Integer
    ePings As Long
    eDetection_TTL As Long
    eX As Integer
    eY As Integer
    eLo As Single
    eLa As Single
    eNotes As String * Len_eNotes
    eDeployDate As String * Len_eDeployDate
    eRetrieveDate As String * Len_eRetrieveDate
    eModel As String * Len_eModel
    eTag As Integer
    eZone As Integer
    eColor As Long
    eMarker As Byte
    eInvisible As Boolean
End Type

Dim DoNotOverwrite As Boolean
Dim BlankReceiverData As Entry
Dim NewReceiverData As Entry
Dim OldReceiverData As Entry
Dim RecordSize As Long
Dim Header_Line As Entry
Dim ReceiverDataInMemory(MAX_RECEIVERS) As Entry
Dim LandLockedConditionCode(MAX_RECEIVERS) As Integer

Public Property Get TotalReceiversLoaded() As Long
TotalReceiversLoaded = TotalEntries
End Property
Public Property Let ID(ByVal R As Integer, s As String)
Dim dummy As String
dummy = Space(Len_eID)
LSet dummy = s
ReceiverDataInMemory(R).eID = dummy
End Property
Public Property Get ID(ByVal R As Integer) As String
Dim temp As String
If R < 0 Then Exit Property
temp = Trim$(ReceiverDataInMemory(R).eID)
If temp = String(Len_eID, Chr$(0)) Then temp = ""
ID = temp
End Property
Public Property Let GroupNumber(ByVal R As Integer, i As Integer)
If R <> -1 Then ReceiverDataInMemory(R).eGroupNumber = i
End Property
Public Property Get GroupNumber(ByVal R As Integer) As Integer
If R <> -1 Then GroupNumber = ReceiverDataInMemory(R).eGroupNumber
End Property
Public Property Get LandLockedCode(ByVal R As Integer) As Integer
LandLockedCode(R) = LandLockedConditionCode(R)
End Property
Public Property Let LandLockedCode(ByVal R As Integer, c As Integer)
 LandLockedConditionCode(R) = c
End Property
Public Property Let Pings(R As Integer, L As Long)
ReceiverDataInMemory(R).ePings = L
End Property
Public Property Get Pings(R As Integer) As Long
Dim Ping As Long
Dim TotalPings As Long
Dim PingCount As Long
'count only valid pings
TotalPings = ReceiverDataInMemory(R).ePings
For Ping = 0 To TotalPings - 1
    Receiver.ReadStamp R, Ping
    If Stamp.Valid Then PingCount = PingCount + 1
Next Ping

'return ping count
Pings = PingCount

End Property
Public Property Get AllPings(R As Integer) As Long
Dim Ping As Long
Dim TotalPings As Long
Dim PingCount As Long
'count only valid pings
TotalPings = ReceiverDataInMemory(R).ePings
'return ping count
AllPings = TotalPings

End Property
Public Sub ReadGroupStamp(ByRef R() As Long, R_Bound As Long, ByVal e As Long)
'This sub reads a stamp from a receiver group

Dim iR As Long
Dim GN As Long
Dim ee As Long
Dim Total As Single
Dim Found As Boolean

Total = e
ee = e

Do
    Total = Total - CLng(ReceiverDataInMemory(R(iR)).eDetection_TTL)
    If Total <= 0 Then
        Found = True
        e = ee
    Else
        iR = iR + 1
        ee = Total
    End If
Loop Until iR > R_Bound Or Found

ReceiverTable.ReadStamp R(iR), e

End Sub
Public Property Let Detection_TTL(ByVal R As Integer, L As Long)
ReceiverDataInMemory(R).eDetection_TTL = L
End Property
Public Property Get Detection_TTL_UnGrouped(ByVal R As Integer) As Long
Detection_TTL_UnGrouped = ReceiverDataInMemory(R).eDetection_TTL
End Property
Public Property Get Detection_TTL(ByVal R As Integer) As Long

If ReceiverDataInMemory(R).eInvisible Then
    Detection_TTL = 0
Else
    Detection_TTL = ReceiverDataInMemory(R).eDetection_TTL
End If

End Property
Public Sub Detection_TTL_DOWN(ByVal R As Integer)
ReceiverDataInMemory(R).eDetection_TTL = ReceiverDataInMemory(R).eDetection_TTL - 1
End Sub

Public Property Let X(R As Integer, i As Integer)
ReceiverDataInMemory(R).eX = i
End Property
Public Property Get X(R As Integer) As Integer
X = ReceiverDataInMemory(R).eX
End Property
Public Property Let Y(R As Integer, i As Integer)
ReceiverDataInMemory(R).eY = i
End Property
Public Property Get Y(R As Integer) As Integer
Y = ReceiverDataInMemory(R).eY
End Property
Public Property Let LO(R As Integer, Si As Single)
ReceiverDataInMemory(R).eLo = Si
End Property
Public Property Get LO(R As Integer) As Single
LO = ReceiverDataInMemory(R).eLo
End Property
Public Property Let LA(R As Integer, Si As Single)
ReceiverDataInMemory(R).eLa = Si
End Property
Public Property Get LA(R As Integer) As Single
LA = ReceiverDataInMemory(R).eLa
End Property
Public Property Let Notes(R As Integer, s As String)
Dim dummy As String
dummy = Space(Len_eNotes)
LSet dummy = s
ReceiverDataInMemory(R).eNotes = dummy
End Property
Public Property Get Notes(R As Integer) As String
Notes = Trim$(ReceiverDataInMemory(R).eNotes)
End Property
Public Property Let DeployDate(R As Integer, s As String)
Dim dummy As String
dummy = Space(Len_eDeployDate)
LSet dummy = s
ReceiverDataInMemory(R).eDeployDate = dummy
End Property
Public Property Get DeployDate(R As Integer) As String
DeployDate = Trim$(ReceiverDataInMemory(R).eDeployDate)
End Property
Public Property Let RetrieveDate(R As Integer, s As String)
Dim dummy As String
dummy = Space(Len_eRetrieveDate)
LSet dummy = s
ReceiverDataInMemory(R).eRetrieveDate = dummy
End Property
Public Property Get RetrieveDate(R As Integer) As String
RetrieveDate = Trim$(ReceiverDataInMemory(R).eRetrieveDate)
End Property
Public Property Let Model(R As Integer, s As String)
Dim dummy As String
dummy = Space(Len_eModel)
LSet dummy = s
ReceiverDataInMemory(R).eModel = dummy
End Property
Public Property Get Model(R As Integer) As String
Model = ReceiverDataInMemory(R).eModel
End Property
Public Property Let Tag(R As Integer, i As Integer)
ReceiverDataInMemory(R).eTag = i
End Property
Public Property Get Tag(R As Integer) As Integer
Tag = ReceiverDataInMemory(R).eTag
End Property
Public Property Let Zone(R As Integer, i As Integer)
ReceiverDataInMemory(R).eZone = i
End Property
Public Property Get Zone(R As Integer) As Integer
Zone = ReceiverDataInMemory(R).eZone
End Property

Public Property Let Color(R As Integer, c As Long)
ReceiverDataInMemory(R).eColor = c
End Property
Public Property Get Color(R As Integer) As Long
Color = ReceiverDataInMemory(R).eColor
End Property
Public Property Let Marker(R As Integer, m As Long)
ReceiverDataInMemory(R).eMarker = m
End Property
Public Property Get Marker(R As Integer) As Long
Marker = CLng(ReceiverDataInMemory(R).eMarker)
End Property

Public Property Let Invisible(R As Integer, flag As Boolean)

'propagate status to group
If Receiver.GroupNumber(R) Then
    'only grouped receiver should be changed.  Receiver group memebers should remain invisible
    ReceiverDataInMemory(Receiver.GroupReceiverNumber(Receiver.GroupNumber(R))).eInvisible = flag
Else
    ReceiverDataInMemory(R).eInvisible = flag
End If

WriteEntry R
End Property
Public Property Get Invisible(R As Integer) As Boolean
Invisible = ReceiverDataInMemory(R).eInvisible
End Property
Private Sub WriteEntry(ReceiverNumber As Integer)
If TableOpen And Not DoNotOverwrite Then
    Put #FileNumberForAnimation, ReceiverNumber + HeaderSize, ReceiverDataInMemory(ReceiverNumber)
End If
If TotalEntries < ReceiverNumber Then TotalEntries = ReceiverNumber
End Sub
Public Property Let DataFileName(s As String)
CurrentSourceDataFile = s
End Property
Public Sub SaveData()
WriteHeader
End Sub
Private Sub WriteHeader()
Dim Name As String
If TableOpen Then
    Name = Space(Len_eID)
    LSet Name = CurrentSourceDataFile
    Header_Line.eID = Name
    Put #FileNumberForAnimation, 1, Header_Line
    LSet Name = FileDateTime(CurrentSourceDataFile)
    Header_Line.eID = Name
    Put #FileNumberForAnimation, 2, Header_Line
    Header_Line.ePings = TotalEntries
    Put #FileNumberForAnimation, 3, Header_Line
End If
End Sub
Private Sub ReadEntriesIntoMemory()
Dim ReceiverNumber As Long
'Reads all entries into memory buffer at once
For ReceiverNumber = 1 To TotalEntries
    If TableOpen Then Get #FileNumberForAnimation, ReceiverNumber + HeaderSize, ReceiverDataInMemory(ReceiverNumber)
Next ReceiverNumber
End Sub
Public Function OpenTableFile() As Boolean
'returns true if table already exists.
Dim LastSourceFullFilePath As String
Dim TimeStamp As String
Dim Buffer As String
Dim ReturnValue As Boolean

'len of single record
RecordSize = Len(NewReceiverData)

'Opens table file
Open TableFileName For Random As FN Len = RecordSize
TableOpen = True

'get header
Get #FileNumberForAnimation, 1, Header_Line
LastSourceFullFilePath = Trim$(Header_Line.eID)
Get #FileNumberForAnimation, 2, Header_Line
TimeStamp = Trim$(Header_Line.eID)
Get #FileNumberForAnimation, 3, Header_Line
TotalEntries = Header_Line.ePings

On Error GoTo FileNotFound
'check if same or new
ReturnValue = True
If (Not (LastSourceFullFilePath = CurrentSourceDataFile)) Or TimeStamp <> FileDateTime(CurrentSourceDataFile) Then
    'new
    TotalEntries = 0
    WriteHeader
    ReturnValue = False
End If

If TotalEntries > 0 Then ReadEntriesIntoMemory

OpenTableFile = ReturnValue

Exit Function

FileNotFound:
    TotalEntries = 0
    WriteHeader
End Function
Public Sub ClearPingCount()
Dim R As Integer
For R = 1 To TotalEntries
    ReceiverDataInMemory(R).ePings = 0
Next R
End Sub
Public Sub BackupTable()
'Uses memory as a backup table
'
DoNotOverwrite = True
End Sub
Public Sub RestoreTable()
'Restores table as it is in the file
ReadEntriesIntoMemory
DoNotOverwrite = False
End Sub
Private Sub InitPallette()
Dim R As Long
'Colorpal contains the 10 color pallette
ColorPal(0) = QBColor(6)
ColorPal(1) = QBColor(2)
ColorPal(2) = QBColor(9)
ColorPal(3) = QBColor(10)
ColorPal(4) = QBColor(12)
ColorPal(5) = QBColor(11)
ColorPal(6) = QBColor(13)
ColorPal(7) = QBColor(5)
ColorPal(8) = QBColor(0)
ColorPal(9) = QBColor(14)
ColorPal(10) = QBColor(4)
ColorPal(11) = QBColor(7)


End Sub
Private Sub InitMarkers()
Dim R As Long

'load colors
InitPallette

'assign a marker to each receiver
'For R = 1 To Receiver_Table.TotalReceiversLoaded
 '   AssignDefaultMarker R
'Next R

End Sub
Public Sub AssignDefaultMarker(R As Long)
Static MarkerColor As Long
Static MarkerType As Byte
Dim Color As Long
Const DefaultColor = 2 'Blue
Const DefaultMarker = 0 'Circle

'Color = ColorPal(MarkerColor)
'MarkerColor = MarkerColor + 1

'carry over to marker type
'If MarkerColor > MaxPal Then
 '   MarkerColor = 0
  '  MarkerType = MarkerType + 1
'End If

'carry over to markeropen flag
'If MarkerType > NumberofMarkerTypes Then
 '   MarkerType = 0
'End If

With ReceiverDataInMemory(R)
    .eColor = ColorPal(DefaultColor)
    .eMarker = DefaultMarker
End With

End Sub
Public Sub SetAllReceiversToINVISIBLE()
Dim R As Long
Dim G As Integer

For R = 1 To Receiver_Table.TotalReceiversLoaded
    G = Receiver.GroupNumber(R)
    If G <> 0 Then
        ReceiverDataInMemory(Receiver.GroupReceiverNumber(G)).eInvisible = True
    Else
        ReceiverDataInMemory(R).eInvisible = True
    End If
Next R

End Sub
Public Sub SetAllReceiversToVISIBLE()
Dim R As Long
Dim G As Integer

For R = 1 To Receiver_Table.TotalReceiversLoaded
    G = Receiver.GroupNumber(R)
    If G <> 0 Then
        ReceiverDataInMemory(Receiver.GroupReceiverNumber(G)).eInvisible = False
    Else
        ReceiverDataInMemory(R).eInvisible = False
    End If
Next R


End Sub
Private Sub Class_Initialize()
ResetReceivers
End Sub
Public Sub ResetReceivers()
Dim R As Long

'give marker color
InitMarkers

'make invisible and give 0 value to total and reset grouping
For R = 1 To MAX_RECEIVERS
    With ReceiverDataInMemory(R)
        .eID = ""
        .eColor = ColorPal(DefaultColor)
        .eMarker = DefaultMarker
        .eInvisible = True
        .eDetection_TTL = 0
        .ePings = 0
        .eZone = 0
        .eGroupNumber = 0
    End With
Next R

'0 receivers
TotalEntries = 0

End Sub
