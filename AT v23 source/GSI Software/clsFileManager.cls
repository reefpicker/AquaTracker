VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Enum EntryType
    FishStamp = 0
    ReceiverStamp = 1
End Enum

Const END_FILE = "[///]"
Const END_SECTION = "[---]"
Const CONTINUE_SECTION = "[...]"
Const NOP = "[.]"

Dim SaveValidStampsOnly As Boolean
Dim SaveGroupInfo As Boolean
Dim SaveMarkers As Boolean
Dim SaveRoutes As Boolean
Dim SaveSelectedOnly As Boolean
Private Function Concatenate(Index As Long, I1 As Integer, L2 As Long, I3 As Integer, L4 As Long, I5 As Integer, I6 As Boolean) As String
Dim S1 As String
Dim S2 As String
Dim S3 As String
Dim S4 As String
Dim S5 As String

'convert stamp to string

S1 = Str$(I1)
S2 = Str$(L2)
S3 = Str$(I3)
S4 = Str$(L4)
S5 = Str$(I5)
S6 = Str$(I6)

Concatenate = Str$(Index) & "," & S1 & "," & S2 & "," & S3 & "," & S4 & "," & S5 & "," & S6


End Function
Private Sub LoadStamp(BufferEntry As String, TableType As Integer)
'loads stamp from string to table
Dim Fish As String
Dim Site As String
Dim CT As String
Dim t As String
Dim d As String
Dim FishField As Long
Dim DateField As Long
Dim TimeField As Long
Dim CTimeField As Long
Dim SiteField As Long
Dim ValidField As Long
ReDim s(20) As String
Dim Index As Long
Static ed As Long

'Standard format is:
'Site,Date,Time,Ctime,Fish,VALID FLAG
'or
'Fish,Date,Time,Ctime,Site,VALID FLAG
'Depending on what is being loaded (fish or receivers stamps)

'same for everyone
DateField = 2
TimeField = 3
CTimeField = 4
ValidField = 6

'type of table determines the first and last
Select Case TableType
    Case EntryType.FishStamp
        FishField = 1: SiteField = 5
    Case EntryType.ReceiverStamp
        FishField = 5: SiteField = 1
End Select

'split incoming string
s = Split(BufferEntry, ",")

'load the stamp
With Stamp
    .Fish = CInt(s(FishField))
    .Site = CInt(s(SiteField))
    .Date = CLng(s(DateField))
    .Time = CLng(s(TimeField))
    .CTime = CLng(s(CTimeField))
    .Valid = CBool(Trim(s(ValidField)))
End With

'get index
Index = CLng(s(0))

'paste stamp to table
If TableType = EntryType.FishStamp Then
    FishTable.PasteStamp Stamp.Fish, Index
Else
    ReceiverTable.PasteStamp Stamp.Site, Index
End If

End Sub
Public Function LoadNativeFormat(InputFile As String, PBar As ProgressBar) As Long
Dim i As Long
Dim R As Integer
Dim RR As Integer
Dim F As Integer
Dim s As String
Dim Group As Long
Dim TotalNumberOfReceivers As Long
Dim TotalNumberOfFish As Long
Dim TotalNumberOfStamps As Long
Dim SaveStamp As Boolean
Dim SaveReceiver As Boolean
Dim Buffer As String
Dim LastReceiverLoaded As Boolean
Dim LastFishLoaded As Boolean
ReDim splittedstring(20) As String
Dim LA As Single
Dim LO As Single
Dim X As Long
Dim Y As Long
Dim Distance As Single
Dim GroupHasBeenDefined(MAX_GROUPS) As Boolean
Dim FakeReceiver As Integer
Dim LengthOfFile As Single
Dim icg As Long
Dim MarkerType As Long
Dim MarkerColor As Long
Dim L As Long
Dim response As Variant

Const LengthOfLine = 31

On Error GoTo ExitWithError

Form1.StatusBar.Panels(1).Text = "Opening..."

'Open file for output on buffer 1
Open InputFile For Input As #1

'get len
LengthOfFile = FileLen(InputFile)
PBar.Value = 0
If LengthOfLine = 0 Then GoTo ExitWithError
PBar.Max = CLng(LengthOfFile / LengthOfLine)


'if NOT selections only, make all receivers "visible"
If Not SaveSelectedOnly Then Receiver.MakeVisible

'First line should read:
'[Fish]

Do Until EOF(1) Or Buffer = "[Fish]"
    Line Input #1, Buffer
Loop



'load fish info
FishDatabase.LoadFish 1
If FishDatabase.TotalFishLoaded < 1 Then GoTo ExitWithError

PBar.Value = FishDatabase.TotalFishLoaded

Form1.StatusBar.Panels(1).Text = "Loading receivers stamps..."

'get receivers
'First line is a throw away marker
Line Input #1, Buffer

'Receivers stamps
Do Until EOF(1) Or LastReceiverLoaded
    Line Input #1, Buffer
    L = L + 1
    'if fish stamp table begins don't load anymore receivers.
    If Buffer = "[Stamps:Fish]" Then
        LastReceiverLoaded = True
    Else
        LoadStamp Buffer, EntryType.ReceiverStamp
        PBar.Value = PBar.Value + 1
    End If
Loop

Form1.StatusBar.Panels(1).Text = "Loading fish stamps..."
'Fish stamps (=>tracks)
Do Until EOF(1) Or LastFishLoaded
    Line Input #1, Buffer
    'if receiver desc. begins don't load anymore tracks.
    If Buffer = "[Receivers]" Then
        LastFishLoaded = True
    Else
        LoadStamp Buffer, EntryType.FishStamp
        PBar.Value = PBar.Value + 1
    End If
Loop

LastReceiverLoaded = False

Form1.StatusBar.Panels(1).Text = "Loading RDT..."
''''''''''''''''''''''''''''''''''''
'Receiver descriptor
'First line is ReceiverNumber and ID

Do Until EOF(1) Or Buffer = END_FILE Or Buffer = END_SECTION
    Line Input #1, Buffer
    If Buffer <> END_FILE And Buffer <> END_SECTION Then
        splittedstring = Split(Buffer, "=")
        R = CInt(splittedstring(0))
        s = splittedstring(1)
        
        'second line is total stamps
        Line Input #1, Buffer
        TotalNumberOfStamps = CLng(Buffer)
        
        'Third line is the coordinates
        'LA/LO
        Line Input #1, Buffer
        splittedstring = Split(Buffer, "/")
        LA = CSng(splittedstring(0))
        LO = CSng(splittedstring(1))
        
        'fourth line is x,y
        Line Input #1, Buffer
        splittedstring = Split(Buffer, "/")
        X = CSng(splittedstring(0))
        Y = CSng(splittedstring(1))
        
        'load this into a receiver
        Receiver.LoadReceiver R, TotalNumberOfStamps, s, LO, LA, X, Y
        
        'Optional: Marker information
        'If no marker saved then
        'the entry will consist of "."=> NOP
        Line Input #1, Buffer
        If Buffer <> NOP Then
            splittedstring = Split(Right(Buffer, Len(Buffer) - 1), "/")
            MarkerType = CLng(splittedstring(0))
            MarkerColor = CLng(splittedstring(1))
        Else
            MarkerType = 1
            MarkerColor = vbBlue
        End If
        
        'groups
        'if receiver is part of group this entry will read something other than 0,0
        'First part of entry will be 0 if receiver is not part of group atm
        '
        Line Input #1, Buffer
        splittedstring = Split(Buffer, ",")
        Group = CLng(splittedstring(0))
        FakeReceiver = CInt(splittedstring(1))
        If Group <> 0 Then
            If Not GroupHasBeenDefined(Group) Then
                Group = Receiver.CreateNewReceiverGroup(splittedstring(2), Group)
            End If
            GroupHasBeenDefined(Group) = True
            Receiver.AddReceiverToGroup(Group, FakeReceiver) = R
        End If
        'color and mark receiver
        Receiver.Color(R) = MarkerColor
        Receiver_Table.Marker(R) = MarkerType
        
        'End of section?
        Line Input #1, Buffer
        PBar.Value = PBar.Value + 1
        PBar.Refresh
    End If
Loop

Form1.StatusBar.Panels(1).Text = "Loading Routes and starting floating window..."
'routes too?
If Buffer = END_SECTION Then
    ''''''''''''''''''
    'Routes
    '
    'Routes that have been traced are the only ones stored
    'Thus, file can have 1 route or 10000 routes!
    '
    'Skip obligatory header
    Line Input #1, Buffer
    'process first entry
    'prefetch
    Line Input #1, Buffer
    If Buffer <> END_FILE Then
        Do
            splittedstring = Split(Buffer, ">")
            'receivers
            R = CInt(splittedstring(0))
            RR = CInt(splittedstring(1))
            'distance
            Line Input #1, Buffer
            Distance = CSng(Buffer)
            'route string
            Line Input #1, Buffer
            Receiver.LoadRoute(R, RR, Distance) = Buffer
            Line Input #1, Buffer
            PBar.Value = PBar.Value + 1
        Loop Until EOF(1) Or Buffer = END_FILE
    End If
End If

'ends here
Close #1


'transfer list

'Transfer list of fish to floating window
frmFloater.cmbFishCode.Clear
frmFloater.cmbFishCode.AddItem "ALL"
FishDatabase.TransferList frmFloater.cmbFishCode

Form1.StatusBar.Panels(StatusPanel.Map).Text = "AQN file loaded..."
'hide progress bar, its not needed anymore
PBar.Visible = False

'return ttl detections
LoadNativeFormat = L - 1

Exit Function

ExitWithError:
    If FreeFile <> 1 Then Close #1
    response = MsgBox("Error loading AQN file.  The file you are attempting to load is empty or corrupt or does not exist.", vbCritical, "Load Error")
End Function
Public Sub SaveInNativeFormat(OutputFile As String)
Dim i As Long
Dim R As Integer
Dim RR As Integer
Dim F As Integer
Dim s As String
Dim Group As Integer
Dim GroupName As String
Dim TotalNumberOfReceivers As Long
Dim TotalNumberOfFish As Long
Dim TotalNumberOfStamps As Long
Dim SaveStamp As Boolean
Dim SaveReceiver As Boolean
Dim temp As Boolean

SaveGroupInfo = True
SaveMarkers = True
SaveRoutes = True

'Open file for output on buffer 1
Open OutputFile For Output As #1

'if NOT selections only, make all receivers "visible"
If Not SaveSelectedOnly Then Receiver.MakeVisible

Print #1, "[Fish]"
FishDatabase.SaveFish 1

Print #1, "[Stamps:Receivers]"
'Get total receivers
TotalNumberOfReceivers = Receiver.TotalReceivers
'Save database info for each
For R = 1 To TotalNumberOfReceivers
    'Get total Stamps
    TotalNumberOfStamps = Receiver.Detection_TTL(R, True)
    For i = 0 To TotalNumberOfStamps - 1
        'Save
        ReceiverTable.ReadStamp R, i
        s = Concatenate(i, Stamp.Site, Stamp.Date, Stamp.Time, Stamp.CTime, Stamp.Fish, Stamp.Valid)
        Print #1, s
    Next i
Next R


Print #1, "[Stamps:Fish]"
'Get total fish
TotalNumberOfFish = FishDatabase.TotalFishLoaded

'save each fish that is visible and valid
For F = 0 To TotalNumberOfFish
    'Get total stamps
    TotalNumberOfStamps = FishDatabase.NumberOfStamps(F)
    For i = 0 To TotalNumberOfStamps - 1
        'read stamp
        FishTable.ReadStamp F, i
        'save
        s = Concatenate(i, Stamp.Fish, Stamp.Date, Stamp.Time, Stamp.CTime, Stamp.Site, Stamp.Valid)
        Print #1, s
    Next i
Next F

Print #1, "[Receivers]"
'Save each receiver's properties
For R = 1 To TotalNumberOfReceivers
    'Receiver number on database and name
    Print #1, Str$(R) & "=" & Receiver.ID(R, True)
    'total stamps, override visibility for this
    Print #1, Str$(Receiver.Detection_TTL(R, True))
    'Lat/Long
    'X/Y
    Print #1, Str$(Receiver.LA(R)) & "/" & Str$(Receiver.LO(R))
    Print #1, Str$(Receiver.X(R)) & "/"; Str$(Receiver.Y(R))
    'Shape info
    If SaveMarkers Then
        Print #1, "!" & Str$(Receiver_Table.Marker(R)) & "/" & Str$(Receiver_Table.Color(R))
    Else
        Print #1, NOP
    End If
        
    'Group info: Group in hexadecimal preceded by &H followed by , &H hex FakeReceiver
    Group = Receiver.GroupNumber(R)
    GroupName = ""
    If Not SaveGroupInfo Then
        Group = 0
    End If
    
    If Group <> 0 Then GroupName = Receiver.ID(R)
    Print #1, "&H" & Hex$(Group) & " , &H" & Hex$(Receiver.GroupReceiverNumber(Group)) & "," & GroupName
    
    'use special marker to end section or continue it
    If R <> TotalNumberOfReceivers Then
        Print #1, CONTINUE_SECTION
    End If
Next R

'
'All calculated routes
If SaveRoutes Then
    Print #1, END_SECTION
    Print #1, "[Routes]"
    For R = 1 To TotalNumberOfReceivers
        For RR = 1 To TotalNumberOfReceivers
            Receiver.SaveRoute R, RR, 1
        Next RR
    Next R
End If

'eof
Print #1, END_FILE

'mark end of file with one more line of copyrights
Print #1, "End Of AT Native format file"
Print #1, "-End of .aqn => AquaTracker Native File"
Print #1, "AQN Rev 1.0"
'close it
Close #1

End Sub

