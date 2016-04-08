VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "VERSION"
   ClientHeight    =   10710
   ClientLeft      =   2370
   ClientTop       =   -3000
   ClientWidth     =   8925
   Icon            =   "Main Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   714
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   595
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   11160
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Max             =   4
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   10335
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   2381
            MinWidth        =   2381
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2381
            MinWidth        =   2381
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   9360
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   12135
      Left            =   0
      MouseIcon       =   "Main Form.frx":08CA
      Picture         =   "Main Form.frx":0A1C
      ScaleHeight     =   805
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   621
      TabIndex        =   0
      Top             =   0
      Width           =   9375
      Begin VB.Label lblTrackNotVisible 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Track Excluded"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   12135
      Left            =   0
      MouseIcon       =   "Main Form.frx":16EFC6
      Picture         =   "Main Form.frx":16F118
      ScaleHeight     =   805
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   621
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   9375
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLoadFile 
         Caption         =   "Load data"
      End
      Begin VB.Menu mnuSaveAQN 
         Caption         =   "Save as .AQN..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFile_Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "Export Data Analysis"
         Begin VB.Menu mnuExportFishTrackData 
            Caption         =   "Fish Track Analysis"
         End
         Begin VB.Menu mnuResidenceAllFish 
            Caption         =   "Residence Analysis"
         End
         Begin VB.Menu mnuExportStrings 
            Caption         =   "Track String"
         End
         Begin VB.Menu mnuSaveExcursions 
            Caption         =   "Excursion Analysis"
         End
         Begin VB.Menu mnuExportDensity 
            Caption         =   "Density Histogram"
         End
      End
      Begin VB.Menu mnuFile_Sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImportSunriseData 
         Caption         =   "Import Sunrise/Sunset table"
      End
      Begin VB.Menu mnuFile_Sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCopyFirstLevel 
         Caption         =   "Copy"
         Begin VB.Menu mnuCopyCanvas 
            Caption         =   "Canvas Image"
         End
         Begin VB.Menu mnuCopyData 
            Caption         =   "Data"
         End
      End
      Begin VB.Menu mnuMakeVisible 
         Caption         =   "Exclude from analysis"
         Begin VB.Menu mnuOptions_Exclude_Tracks 
            Caption         =   "Fish"
         End
         Begin VB.Menu mnuOptions_Exclude_Receivers 
            Caption         =   "Receivers"
         End
         Begin VB.Menu mnuOptions_Dates 
            Caption         =   "Dates"
         End
      End
   End
   Begin VB.Menu mnuMap 
      Caption         =   "&Map"
      Begin VB.Menu mnuLoadMap 
         Caption         =   "Load from BMP"
      End
      Begin VB.Menu mnuNewCanvas 
         Caption         =   "New Blank Canvas"
      End
      Begin VB.Menu mnuMap_Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPickWaterColor 
         Caption         =   "Pick color of water"
      End
   End
   Begin VB.Menu mnuReceivers 
      Caption         =   "&Receivers"
      Begin VB.Menu mnuShowAllReceivers 
         Caption         =   "Show All"
      End
      Begin VB.Menu mnuShowDensity 
         Caption         =   "Density Plot"
         Begin VB.Menu mnuDensityByPings 
            Caption         =   "Ping density"
         End
         Begin VB.Menu mnuPercentFish 
            Caption         =   "Fish density"
         End
         Begin VB.Menu mnuShowDDbyDielCycle 
            Caption         =   "Fish density by diel cycle"
            Begin VB.Menu mnuNight 
               Caption         =   "Night"
            End
            Begin VB.Menu mnuDayAM 
               Caption         =   "Morning"
            End
            Begin VB.Menu mnuDayPM 
               Caption         =   "Evening"
            End
            Begin VB.Menu mnuDawn 
               Caption         =   "Dawn Twilight"
            End
            Begin VB.Menu mnuDusk 
               Caption         =   "Dusk Twilight"
            End
         End
      End
      Begin VB.Menu mnuReceiverGroupWindow 
         Caption         =   "Groups"
      End
   End
   Begin VB.Menu mnuAnalysis 
      Caption         =   "&Analyze"
      Begin VB.Menu mnuFindOverlaps 
         Caption         =   "Overlaps"
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "Singletons"
      End
      Begin VB.Menu mnuIdentifyFishGroups 
         Caption         =   "Fish in Groups"
      End
      Begin VB.Menu mnuDetectionPlot 
         Caption         =   "Scatter Plot"
      End
      Begin VB.Menu mnuMarkov 
         Caption         =   "Markov chains"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuActions 
         Caption         =   "Actions window"
      End
      Begin VB.Menu mnuNavStringShow 
         Caption         =   "Track String"
      End
      Begin VB.Menu mnuDielCycle 
         Caption         =   "Diel Histogram"
      End
      Begin VB.Menu mnuReceiverResidence 
         Caption         =   "Residence Time"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuAnimation 
         Caption         =   "Animation"
         Visible         =   0   'False
         Begin VB.Menu mnuSummarizeDensity 
            Caption         =   "Plot by accumulating total"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuDoNotSummarize 
            Caption         =   "Plot each day independently"
         End
      End
      Begin VB.Menu mnuReceiver 
         Caption         =   "Receiver"
         Visible         =   0   'False
         Begin VB.Menu mnuReceiverInformation 
            Caption         =   "Receiver Information"
         End
         Begin VB.Menu mnuReceiver_Sep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDefineGeoZone 
            Caption         =   "Define Geographic Zone"
         End
         Begin VB.Menu mnuShowExcursions 
            Caption         =   "Show Excursions"
         End
         Begin VB.Menu mnuReceiver_Sep2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuShowStamps 
            Caption         =   "Show stamps"
            Begin VB.Menu mnuShowDetectorInfo 
               Caption         =   "All stamps in receiver"
            End
            Begin VB.Menu mnuShowTrackDetails 
               Caption         =   "Track-specific stamps"
            End
         End
         Begin VB.Menu mnuDielPatternPings 
            Caption         =   "Diel Pattern of stamps"
         End
         Begin VB.Menu mnuSeparator 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAddToGroup 
            Caption         =   "Add to Group"
         End
      End
      Begin VB.Menu mnuTrackAnimation 
         Caption         =   "Animation"
         Begin VB.Menu mnuTrackAnimationStyle 
            Caption         =   "In Scaled Time"
            Index           =   0
         End
         Begin VB.Menu mnuSeparatorTA 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuCompressionFactor 
            Caption         =   "Choose Scaling Factor"
         End
         Begin VB.Menu mnuLagFactor 
            Caption         =   "Screen erase"
            Begin VB.Menu LagFactor 
               Caption         =   "Never"
               Checked         =   -1  'True
               Index           =   0
            End
            Begin VB.Menu LagFactor 
               Caption         =   "Always"
               Index           =   1
            End
            Begin VB.Menu LagFactor 
               Caption         =   "Every other frame"
               Index           =   2
            End
            Begin VB.Menu LagFactor 
               Caption         =   "Every two frames"
               Index           =   3
            End
            Begin VB.Menu LagFactor 
               Caption         =   "Every three frames"
               Index           =   4
            End
         End
      End
      Begin VB.Menu mnuOptions_Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHighLightColorChange 
         Caption         =   "Change highlight color"
      End
      Begin VB.Menu mnuOptions_Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChkAvoidLand 
         Caption         =   "Land Avoidance"
         Enabled         =   0   'False
         Begin VB.Menu mnuLandAvoidanceChoice 
            Caption         =   "Not Active"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuLandAvoidanceChoice 
            Caption         =   "Shoreline (Worst)"
            Index           =   1
         End
         Begin VB.Menu mnuLandAvoidanceChoice 
            Caption         =   "Corridors (Forced track)"
            Index           =   2
         End
         Begin VB.Menu mnuLandAvoidanceChoice 
            Caption         =   "Random Walk (Best)"
            Index           =   3
         End
      End
      Begin VB.Menu mnuLandAvoidanceOptions 
         Caption         =   "Change land avoidance parameters"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuOptions_Sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectFields 
         Caption         =   "Select fields to show/export"
      End
      Begin VB.Menu mnuOptions_Sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStayTime 
         Caption         =   "Residency Threshold"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpTopics 
         Caption         =   "Help topics"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About this program..."
      End
   End
   Begin VB.Menu mnuTrack 
      Caption         =   "Track"
      Visible         =   0   'False
      Begin VB.Menu mnuChangeTrackColor 
         Caption         =   "Change track color"
      End
      Begin VB.Menu mnuSetAsRefTrack 
         Caption         =   "Set as reference track"
      End
      Begin VB.Menu mnuTrack_Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrack_ReDoTrack 
         Caption         =   "ReDo Track"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FileName As String
Dim SubStrings() As String
Dim RN(MAX_ENTRIES) As Long
Dim FN(MAX_ENTRIES) As Long
Dim FishCodeColumn As Long
Dim ReceiverNameColumn As Long
Dim DateTimeColumn As Long
Dim DateColumn As Long
Dim TimeColumn As Long
Dim GroupColumn As Long
Dim LatColumn As Long
Dim LongColumn As Long
Dim TagColumn As Long
Dim DirectionColumn As Long
Dim FixedReceiverColumn As Long
Dim AbioticColumn As Long
Dim FishArray(MAX_FISH) As Long
Dim ZOOMING As Boolean
Dim Density(MAX_RECEIVERS) As Long
Dim AnchorPointNotValid As Boolean
Dim Disable_MouseOver_Event As Boolean

'API calls needed to synchronize external shell call to Converter
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Enum flag
    OK
    CancelFile
    ConvertFile
    OK_But_Check
End Enum

Private Enum Diel
    Night = 1
    Dawn = 2
    AM = 3
    PM = 4
    Dusk = 5
End Enum

Private Enum FileType
    Native = 1
    CSV = 2
End Enum

Private Enum Style
    ScaledTime = 0
End Enum

Private Enum Side
    Adjacent
    Opposite
End Enum

Private Type Coordinates
    X As Long
    Y As Long
    Distance As Single
End Type

Private Type ZoneBox
    X1 As Long
    X2 As Long
    Y1 As Long
    Y2 As Long
End Type

Private Type FishCorridorNode
    X As Integer
    Y As Integer
    ConnectTo As Integer
End Type
Const MAX_ANCHORS = 100
Dim FishCorridor_AnchorPoints(MAX_ANCHORS) As FishCorridorNode
Dim Anchors As Integer
Dim AnchorMap(MAX_X, MAX_Y) As Integer
Dim ANCHORING As Boolean
Dim CurrentAnchorNumber As Integer

Dim NEW_CONTROL_POINTS As Boolean

Dim ZoneCoordinates As ZoneBox

'File entry buffer (need to access 2x)
Dim MemBuffer(MAX_ENTRIES) As String
Dim MemBuffer_Last As Long


Dim Proximity As Double


Dim Dates_To_Show(365) As String
Dim Dates_Index As Long

Dim Dark_Time As Long

Dim SX As Single
Dim SY As Single

Dim Begining_X As Single
Dim Begining_Y As Single
Dim End_X As Single
Dim End_Y As Single

Dim Track(MAX_RECEIVERS) As Boolean

Dim LastField As Long
Dim ClickedReceiver As Integer

Dim MovingReceiver As Boolean
Const COMMA = ","
Const APPEND_STRING = "Converted_"
Const NUMBER_OF_LINES_IN_HEADER = 1

Public Sub DisableCanvas()
Disable_MouseOver_Event = True
End Sub
Public Sub EnableCanvas()
Disable_MouseOver_Event = False
End Sub
Public Sub ShowDetectors(Optional d As Boolean = False, Optional SHOWBOX As Boolean = False)
'Show all entries
'Also can show PING DENSITY as intensity/heat map
Dim ReceiverNumber As Integer
Dim Max As Long


'x is longitude
'y is lattitude
'Set Scales


'Clear Screen(s)
ClearScreen

'indicate
If DielCycleFormIsLoaded Then ImageProcessingEngine.ClearDielCycle

If SHOWBOX Then Picture1.Line (Begining_X, Begining_Y)-(End_X, End_Y), vbBlack, B

If Not d Then
    StatusBar.Panels(StatusPanel.Map) = "Showing receivers"
    For ReceiverNumber = 1 To Receiver.TotalReceivers
    'show it
        Receiver.DrawReceiver Picture1, ReceiverNumber
    Next ReceiverNumber
Else
    StatusBar.Panels(StatusPanel.Map) = "Showing receiver density"
    Max = 0
    'get max # of pings
    For ReceiverNumber = 1 To Receiver.TotalReceivers
        Density(ReceiverNumber) = Receiver.Pings(ReceiverNumber)
        If Density(ReceiverNumber) > Max Then Max = Density(ReceiverNumber)
    Next ReceiverNumber
   
   'draw receiver
   If Max > 0 Then ImageProcessingEngine.DrawDensityPlotForReceivers Form1.Picture1, Max, Density, LARGE_MARKER
End If
'Now receivers are visible
Receivers_Are_Visible = True

'if diel cycle window is loaded, show track (if any selected) and all detections from
'visible receivers in track(s)
If DielCycleFormIsLoaded Then frmDayLightCycle.UpdateHistogram

End Sub
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
Private Function TimeDifference(T1 As String, T2 As String) As Long
'Parse time decimal to find time difference in minutes
'
Dim T1_TTL As Long
Dim T2_TTL As Long

T1_TTL = hour(T1) * 60 + minute(T1)
T2_TTL = hour(T2) * 60 + minute(T2)

TimeDifference = Abs(T1_TTL - T2_TTL)


End Function

Private Sub cmdShowHide_Click()


End Sub

Private Sub ExportTo(FileName As String)
'Exports to file FileName
'all fish in List
Dim FishNumber As Long
Dim i As Long

'Open path to file
DeviceBuffer.Select_Device Device_Type.File, , FileName

'Progressbar
ProgressBar.Max = frmFloater.cmbFishCode.ListCount

'Get index to code number
For FishNumber = 0 To FishDatabase.TotalFishLoaded
    
    'show progress
    ProgressBar.Value = FishNumber
    DoEvents
    'new fish number, reset accumulators!!
    TrackCalculator.Clear
    FishDatabase.Fish = FishNumber
       'Sort date and time
       'Sort FishNumber

       For i = 0 To FishDatabase.NumberOfStamps - 1
            'get stamp
            FishTable.ReadStamp FishNumber, i
            If Stamp.Valid Then
               'load site into track calculator
               TrackCalculator.Site = Stamp.Site
                       
               'load time data into calculator
               TrackCalculator.Day = Stamp.Date
               TrackCalculator.Time = Stamp.Time
               'obstacle present?
               'If Modify_Navigation(.Site(i), .Site(i - 1)) Then
        
               'calculate accumulators
               TrackCalculator.Calculate
            End If
       Next i
    
    'Write fish info to file
    WriteTrackInformation FishNumber
    
Next FishNumber

'close
DeviceBuffer.CloseDevice
        
End Sub

Private Sub Set_Area_In_Map(OX As Integer, OY As Integer, c As Long, R As Integer)
'Sets an area in memory-based map to a specific value to mark a waypoint, or other value
Dim i As Long
Dim ii As Long
Dim X As Integer
Dim Y As Integer


For i = 0 To R
    For ii = 0 To R
        X = OX + i
        Y = OY + ii
        If X > MAX_X Then X = MAX_X
        If Y > MAX_Y Then Y = MAX_Y
        
        MapImage(X, Y) = c
        X = OX - i
        Y = OY - ii
        If X < 0 Then X = 0
        If Y < 0 Then Y = 0
        MapImage(X, Y) = c
    Next ii
Next i

End Sub
Private Sub Search_For_Water()

Dim Direction_X As Long
Dim Direction_Y As Long
Dim N As Long
Dim X As Single
Dim Y As Single
Dim p As Long

Dim Success As Boolean

'BOUNDARY NOT CONSIDERED!!
'only 2/4 directions considered in a switch-type loop

Const Search_Radius = 20

Direction_X = -1
Direction_Y = -1

Success = False
N = 0
X = SX
Y = SY

'Picture1.Cls

'check if already in water?
p = MapImage(X, Y)
If p = 0 Then
    Exit Sub
End If

Do
    Do
        X = X + Direction_X
        Y = Y + Direction_Y
        N = N + 1
        p = MapImage(X, Y)
        If p = 0 Then
            Success = True
            Exit Do
        End If
    Loop While N <= Search_Radius
    
    If Success Then Exit Do
    N = 0
    
    If Direction_X = 1 And Direction_Y = 1 Then
        N = -1
    End If
    
    If Direction_X = -1 And Direction_Y = -1 Then
        Direction_X = 1
        Direction_Y = 1
        X = SX
        Y = SY
    End If

    
Loop Until N = -1


If Success Then
    SX = X
    SY = Y
End If

End Sub

Public Sub ScanMap()
'Scans map into a map image array (B&W)

Dim X As Long
Dim Y As Long
Dim p As Long
Dim c As Long
Dim i As Long
Dim Maximum As Long
Dim result As Variant

'set canvas size
X = Form1.Picture1.ScaleWidth
Y = Form1.Picture1.ScaleHeight


ImageProcessingEngine.SetMapSize X, Y

'load this picture into the dib section
MapDib.CreateFromPicture Form1.Picture1
DoEvents
ProgressBar.Value = 2

StatusBar.Panels(StatusPanel.Map) = "Scanning Map..."
'set buffer
MapDib.InitializeMatrix
DoEvents

'get edges
ImageProcessingEngine.Edge
DoEvents

ProgressBar.Value = 0
ProgressBar.Visible = True

StatusBar.Panels(StatusPanel.Coordinates) = "Map Accepted."


MapWasScanned = True
'if map scanned make CV stuff available
mnuChkAvoidLand.Enabled = True
mnuLandAvoidanceOptions.Enabled = True

StatusBar.Panels(StatusPanel.Map) = ""

Exit Sub

ExitWithErrors:

result = MsgBox("Error loading BMP:!", vbOKOnly, "Error")
StatusBar.Panels(StatusPanel.Map) = "Unable to read map"
End Sub


Private Sub Command3_Click()
ScanMap

End Sub

Private Sub Form_Load()
Dim i As Long
Dim ii As Long
Dim SplashTime As Long
Dim response As Variant

Me.Hide
'Reset ALL values
ResetAll

'Load version
Me.Caption = "AquaTracker " & Version & " by Jose J. Reyes-Tomassini"
'load values from registry, if values do not exist, write default values
Scale_Y = 0.001474938
Scale_X = 0.001501479
Origin_Lat = 48.26254
Origin_Long = 123.3934
ResidenceThreshold = 60
Begining_X = -1: Begining_Y = -1

'Load days per month for time calculations
LoadDaysOfTheMonthForCalculations

StatusBar.Panels(1) = "Unknown Map"
Factor = 3
Proximity = 0.01

AM_TH = 480
PM_TH = 1200

'load fields to calculate
DeviceBuffer.Assign_Field("FID") = "Fish ID"
DeviceBuffer.Assign_Field("Distance") = "TTL Distance (km)"
DeviceBuffer.Assign_Field("Time") = "Total Time (D)"
DeviceBuffer.Assign_Field("Speed") = "Travel Rate (m/h)"
DeviceBuffer.Assign_Field("Range") = "Range (km)"
DeviceBuffer.Assign_Field("T") = "Linearity"
DeviceBuffer.Assign_Field("RI") = "Meandering"
DeviceBuffer.Assign_Field("SS") = "***Stay Site***"
DeviceBuffer.Assign_Field("ST") = "Stay Time (hrs)"
DeviceBuffer.Assign_Field("La") = "Avg (La)"
DeviceBuffer.Assign_Field("Lo") = "Avg (Lo)"
DeviceBuffer.Assign_Field("RA") = "%R. Active"
DeviceBuffer.Assign_Field("RS") = "Release Site"

'land avoidance defaults to not active
Land_Avoidance = AvoidanceMode.NotActive
Nav_Segment = 5
SEGMENT_SIZE_THRESHOLD = 50
SEARCHRADIUS = 5000#
Persistance_TH = 0.6

'Time compression
TimeCompressionFactor = 60


'load values from registry, if values do not exist, write default values
With ProgramSettings
    'get
    .Origin_Lat = GetSetting(APPLICATION, REGISTRY_SECTION, "Origin_Lat", Str$(Origin_Lat))
    .Origin_Long = GetSetting(APPLICATION, REGISTRY_SECTION, "Origin_Long", Str$(Origin_Long))
    .Scale_X = GetSetting(APPLICATION, REGISTRY_SECTION, "Scale_X", Str$(Scale_X))
    .Scale_Y = GetSetting(APPLICATION, REGISTRY_SECTION, "Scale_Y", Str$(Scale_Y))
    .MapFile = GetSetting(APPLICATION, REGISTRY_SECTION, "MapFile", "")
    .MoviePath = GetSetting(APPLICATION, REGISTRY_SECTION, "Movie_Folder", App.PATH)
    .SplashTime = GetSetting(APPLICATION, REGISTRY_SECTION, "SplashTime", Str$(SplashTime))
    .ColorOfWater = GetSetting(APPLICATION, REGISTRY_SECTION, "WATER", Str$(vbWhite))
    .SizeX = GetSetting(APPLICATION, REGISTRY_SECTION, "Size_Map_Width", "500")
    .SizeY = GetSetting(APPLICATION, REGISTRY_SECTION, "Size_Map_Heght", "500")
    .LastFileLoaded = GetSetting(APPLICATION, REGISTRY_SECTION, "Data_File", "")
    'set
    Origin_Lat = CDbl(.Origin_Lat)
    Origin_Long = CDbl(.Origin_Long)
    Scale_X = CSng(.Scale_X)
    Scale_Y = CSng(.Scale_Y)
    SplashTime = CLng(.SplashTime)
    WaterColor = CLng(.ColorOfWater)
    'first time you run is 5 seconds, next is 1.5
    If SplashTime = 0 Then
        SplashTime = 5000
        SaveSetting APPLICATION, REGISTRY_SECTION, "SplashTime", "2000"
    End If
    
    'load map
    If .MapFile <> "" Then
        LoadFromMap .MapFile
    Else
        'No map
        Picture1.Width = Val(.SizeX)
        Picture2.Width = Val(.SizeX)
        Picture1.Height = Val(.SizeY)
        Picture2.Height = Val(.SizeY)
        AdjustWindow 0, 0
        'set map to no map
        Set Picture1.Picture = Nothing
        Set Picture2.Picture = Nothing
    End If
End With

FishCorridor_Color = DEFAULT_FISHCORRIDOR_COLOR
HighLightColor = vbRed

NEW_CONTROL_POINTS = False

'set mapimage to allwater
'will change once map is read
For i = 0 To MAX_X
    For ii = 0 To MAX_Y
        MapImage(i, ii) = 1
    Next ii
Next i


'load and show splash window
Load frmSplash
frmSplash.Timer1.Interval = SplashTime
frmSplash.Timer1.Enabled = True
frmSplash.Show

End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer
Dim result As Variant

'save settings and quit all forms


With ProgramSettings
'always save last file
    SaveSetting APPLICATION, REGISTRY_SECTION, "Data_File", .LastFileLoaded
    'Process settings to be saved:
    .Origin_Lat = Str$(Origin_Lat)
    .Origin_Long = Str$(Origin_Long)
    .Scale_X = Str$(Scale_X)
    .Scale_Y = Str$(Scale_Y)

    'save?
    If NEW_CONTROL_POINTS Then
        result = MsgBox("Save new control points for next time?", vbYesNoCancel, "Quit Program and Save Work")
        If result = vbCancel Then
            Cancel = True
            Exit Sub
        End If
        
        If result = vbYes Then
            SaveSetting APPLICATION, REGISTRY_SECTION, "Origin_Lat", .Origin_Lat
            SaveSetting APPLICATION, REGISTRY_SECTION, "Origin_Long", .Origin_Long
            SaveSetting APPLICATION, REGISTRY_SECTION, "Scale_X", .Scale_X
            SaveSetting APPLICATION, REGISTRY_SECTION, "Scale_Y", .Scale_Y
        End If
    End If
End With

'loop thru all forms.  Unload all.
For i = Forms.count - 1 To 0 Step -1
    Unload Forms(i)
Next i
End Sub


Private Sub LagFactor_Click(Index As Integer)
Dim i As Long
For i = 0 To LagFactor.UBound
    LagFactor(i).Checked = False
Next i

LagFactor(Index).Checked = True
Lap_MAX = Index
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuActions_Click()
frmFloater.Show vbModeless, Form1


End Sub

Private Sub mnuAddToGroup_Click()
Load frmAssignReceiverToGroup
frmAssignReceiverToGroup.NewGroup
frmAssignReceiverToGroup.Show
End Sub
Private Sub mnuChangeTrackColor_Click()
frmFloater.Show
frmFloater.ChangeTrackColor
End Sub

Private Sub mnuClose_Click()
Unload Me
End Sub


Private Sub mnuCompressionFactor_Click()
'choose compression factor
Dim result As String
Dim i As Long

result = InputBox("Enter a time scaling factor in milliseconds. This is the time elapsed in real time when 30 minutes track time has elapsed.  Default is set to 60 (e.g. 1hr=0.12s).  Set to 0 for no compression.", "Set Time Compression Factor", Str$(TimeCompressionFactor))

If result <> "" And IsNumeric(result) Then
    TimeCompressionFactor = CLng(result)
    mnuTrackAnimationStyle(Style.ScaledTime).Checked = True
End If
 
End Sub

Private Sub mnuDay_Click()
End Sub

Private Sub mnuCopyCanvas_Click()
CopyCanvasImage
End Sub
Private Sub CopyCanvasImage()
Clipboard.Clear
Clipboard.SetData Picture1.Image
End Sub

Private Sub mnuCopyData_Click()
'export what's being displayed
Select Case WhatToShowOnCanvas
    Case ShowOnCanvas.FishTrack
        ExportFishTrackData
    Case ShowOnCanvas.Receivers
        ExportReceiverData
    Case ShowOnCanvas.Residence
        frmReceiversResidenceTime.Show
        frmReceiversResidenceTime.CopyTableToClipBoard
    Case Else
        ImageProcessingEngine.ExportDensityPlotDataToClipBoard Density
End Select
End Sub
Private Sub ExportFishTrackData()
'if all is selected, exports only a list of fish
'else output all valid stamps
Dim s As Long
Dim f As Integer
Dim ClipBoardOutput As String
Dim Header As String
Dim TableSize As Long

If frmFloater.cmbFishCode.Text = "ALL" Or frmFloater.cmbFishCode.Text = "" Then
    For f = 0 To FishDatabase.TotalFishLoaded
        If FishDatabase.IsVisible(f) Then ClipBoardOutput = ClipBoardOutput & FishDatabase.Code(f) & Chr$(13) & Chr$(10)
    Next f
Else
    Header = "Code,Date,Time,Site" & Chr$(13) & Chr$(10)
    ClipBoardOutput = Header
    f = FishDatabase.GetFishNumber(frmFloater.cmbFishCode.Text)
    If FishDatabase.IsVisible(f) Then
        TableSize = FishDatabase.NumberOfStamps(f)
        If ConfirmSize(TableSize) Then
            For s = 0 To TableSize - 1
                FishTable.ReadStamp f, s
                If Stamp.Valid Then ClipBoardOutput = ClipBoardOutput & FishDatabase.Code(f) & "," & Str$(Stamp.Date) & "," & Str$(Stamp.Time) & "," & Receiver.ID(Stamp.Site) & Chr$(13) & Chr$(10)
            Next s
        End If
    End If
End If
Clipboard.Clear
ClipBoardOutput = ClipBoardOutput
Clipboard.SetText ClipBoardOutput

End Sub
Private Sub ExportReceiverData()
'if all is selected, exports only a list of receivers
'else output all valid stamps
Dim s As Long
Dim R As Integer
Dim ClipBoardOutput As String
Dim Header As String
Dim f As Integer
Dim TableSize As Long

If frmFloater.cmbFishCode.Text = "ALL" Or frmFloater.cmbFishCode.Text = "" Then
    For R = 1 To Receiver.TotalReceivers
        If Not Receiver_Table.Invisible(R) Then ClipBoardOutput = ClipBoardOutput & Receiver.ID(R) & Chr$(13) & Chr$(10)
    Next R
Else
    Header = "Receiver,Date,Time,Fish" & Chr$(13) & Chr$(10)
    ClipBoardOutput = Header
    f = FishDatabase.GetFishNumber(frmFloater.cmbFishCode.Text)
    For R = 1 To Receiver.TotalReceivers
        If Not Receiver_Table.Invisible(R) Then
            TableSize = FishDatabase.NumberOfStamps(f)
            If ConfirmSize(TableSize) Then
                For s = 0 To TableSize - 1
                    ReceiverTable.ReadStamp R, s
                    If Stamp.Valid And Stamp.Fish = f Then ClipBoardOutput = ClipBoardOutput & Receiver.ID(R) & "," & Str$(Stamp.Date) & "," & Str$(Stamp.Time) & "," & FishDatabase.Code(Stamp.Fish) & Chr$(13) & Chr$(10)
                Next s
            End If
        End If
    Next R
End If

Clipboard.Clear
Clipboard.SetText ClipBoardOutput

End Sub
Private Function ConfirmSize(s As Long) As Boolean
Dim response As Variant
Dim Accepted As Boolean

If s > SUGGESTED_MAX_EXPORT_TABLE Then
    response = MsgBox("Warning: This is a very large track with " & s & " detections.  Copying this large amount of data could halt or crash the program.  Are you sure you want to proceed with the clipboard export?", vbYesNo, "Warning")
    If response = vbYes Then Accepted = True Else Accepted = False
Else
    Accepted = True
End If

ConfirmSize = Accepted

End Function
Private Sub mnuDawn_Click()
WhatToShowOnCanvas = ShowOnCanvas.DawnDensity
ShowDensity_DielPhase Diel.Dawn
End Sub

Private Sub mnuDayAM_Click()
WhatToShowOnCanvas = ShowOnCanvas.MorningDensity
ShowDensity_DielPhase Diel.AM
End Sub

Private Sub mnuDayPM_Click()
WhatToShowOnCanvas = ShowOnCanvas.EveningDensity
ShowDensity_DielPhase Diel.PM
End Sub

Private Sub mnuDefineGeoZone_Click()
'define the geographic zone in geo map
'
Dim DescriptiveName As String
Dim i As Long
Dim ii As Long
Dim ZoneTagNumber As Long


ZoneTagNumber = Receiver.NumberOfTags + 1

DescriptiveName = InputBox("Enter descriptive name for this zone", "Define Zone with Tag #" & Str$(ZoneTagNumber))

If DescriptiveName = "" Then Exit Sub

'Tag column will show the value with a % sign and two decimals.
'Therefore, name should be at least: 123.56%8
'8 characters long (with a space just in case!)
'Pad string if < 8
If Len(DescriptiveName) < 8 Then DescriptiveName = DescriptiveName & Space(8 - Len(DescriptiveName))

'now define tag
Receiver.NewReceiverTag = DescriptiveName
'and zone
With ZoneCoordinates
    Receiver.DefineZone ZoneTagNumber, .X1, .Y1, .X2, .Y2
End With
MultiSelect = False
ShowDetectors
End Sub




Private Sub mnuDensityByPings_Click()
Const SHOW_DENSITY = True
ColorScale.AutoSet = True
WhatToShowOnCanvas = ShowOnCanvas.PingDensity
ShowDetectors SHOW_DENSITY
If DateListIsLoaded Then frmDates.UpdateList

End Sub
Private Sub mnuDetectionPlot_Click()
'show hour glass
Form1.MousePointer = vbHourglass
StatusBar.Panels(StatusPanel.Map) = "Analyzing receiver and fish databases..."
frmJPlot.Show
End Sub

Private Sub mnuDielCycle_Click()
'when this option is selected
'diel cycle analysis is also shown on the analysis window and on export
'

If Receiver.TotalReceivers = 0 Then Exit Sub

'setup calculator to do it
TrackCalculator.AnalyzeDielMoves = True

frmDayLightCycle.Show vbModeless, Form1
'show on title bar
If CURRENT_FISH = -1 Then
    frmDayLightCycle.Caption = "Histogram for all included fish in array"
Else
    frmDayLightCycle.Caption = "Histogram for " & FishDatabase.Code(CInt(CURRENT_FISH))
End If

End Sub

Private Sub mnuDielPatternPings_Click()
Dim ReceiverNumber As Integer
Dim FishNumber As Integer
Dim StampNumber As Long
Dim Title As String

'get station number
ReceiverNumber = Receiver.CurrentStation_Number

'open window&clear it
frmDayLightCycle.Show
ImageProcessingEngine.ClearDielCycle
FishNumber = CInt(CURRENT_FISH)
If FishNumber = -1 Then
    Title = "Histogram for Receiver " & Receiver.ID(ReceiverNumber)
Else
    Title = "Histogram for Fish " & FishDatabase.Code(FishNumber) & " @ Receiver " & Receiver.ID(ReceiverNumber)
End If

'read first ping
StampNumber = Receiver.ReadFishStampOnReceiver(ReceiverNumber, FishNumber, 0)
Do Until StampNumber = -1
    ImageProcessingEngine.DrawDielCycle
    StampNumber = Receiver.ReadFishStampOnReceiver(ReceiverNumber, FishNumber, StampNumber)
Loop

frmDayLightCycle.Caption = Title

End Sub

Private Sub mnuDusk_Click()
WhatToShowOnCanvas = ShowOnCanvas.DuskDensity
ShowDensity_DielPhase Diel.Dusk
End Sub

Private Sub mnuExportDensity_Click()
Dim HistogramFile As New clsGenericIO
Dim ReceiverNumber As Integer
Dim ttl As Long
Dim FishInReceiver(MAX_RECEIVERS) As Long
Dim N As Long
Dim RScale As Single
Dim Max As Long

Dim X As Single
Dim Y As Single

Dim R As Integer
Dim Radius As Long
Dim c As Long

Dim Cycle As Long

Dim Site As Integer
Dim FileName As String

'Export histogram density plot data to .CSV file

With CommonDialog
    .FileName = ""
    .DialogTitle = "Export to a CSV File"
    .CancelError = False
    'ToDo: set the flags and attributes of the common dialog control
    .Filter = "CSV Files (*.csv)|*.csv"
    .ShowSave
End With

If Len(CommonDialog.FileName) = 0 Then
    Exit Sub
End If

FileName = CommonDialog.FileName

'gather receiver information
For ReceiverNumber = 1 To Receiver.TotalReceivers
    N = Receiver.CountUniqueEntries(ReceiverNumber, FieldNames.Fish)
    FishInReceiver(ReceiverNumber) = N
    ttl = ttl + N
    If N > Max Then
        Max = N
    End If
Next ReceiverNumber

'init device and select device
HistogramFile.Select_Device Device_Type.File, , FileName
HistogramFile.WriteALineToFile "Detection Volume/Density Plot"
HistogramFile.WriteALineToFile "Total Fish:" & Str$(ttl)


'define fields to export
With HistogramFile
    .Assign_Field("Receiver") = "Receiver ID"
    .Assign_Field("Fish") = "TTL Fish"
    .Assign_Field("Pings") = "TTL Pings"
    .Assign_Field("Percent") = "%Fish/TTL"
End With

For ReceiverNumber = 1 To Receiver.TotalReceivers
   HistogramFile.WriteField("Receiver") = Receiver.ID(ReceiverNumber)
   HistogramFile.WriteField("Fish") = Str$(FishInReceiver(ReceiverNumber))
   HistogramFile.WriteField("Pings") = Str$(Receiver.Pings(ReceiverNumber))
   HistogramFile.WriteField("Percent") = Str$((FishInReceiver(ReceiverNumber) / ttl) * 100)
   'write to file
   HistogramFile.WriteLine
Next ReceiverNumber

'close file
HistogramFile.CloseDevice

End Sub

Private Sub mnuExportFishTrackData_Click()
'Export track data to .CSV file
Dim FileName As String

With CommonDialog
    .FileName = ""
    .DialogTitle = "Export to a CSV File"
    .CancelError = False
    'ToDo: set the flags and attributes of the common dialog control
    .Filter = "CSV Files (*.csv)|*.csv"
    .ShowSave
End With
If Len(CommonDialog.FileName) = 0 Then
    Exit Sub
End If

FileName = CommonDialog.FileName

ExportTo FileName

End Sub
Private Sub AutoScale()
'Auto-Scales to fit Canvas Size to Coverage Area of Receiver
'
Dim ULa As Single 'upper
Dim ULo As Single '
Dim LLa As Single 'lower
Dim LLo As Single '

Dim CornerX As Single
Dim CornerY As Single

Dim result As Variant

'calculate
Receiver.CalculateCoverageArea
'load coverage area upper and lower coordinates
ULa = Receiver.CoverageArea_Upper_La
ULo = Receiver.CoverageArea_Upper_Lo
LLa = Receiver.CoverageArea_Lower_La
LLo = Receiver.CoverageArea_Lower_Lo

'screen corner
CornerX = Picture1.Width
CornerY = Picture1.Height

'set this coordinates as control points
'set so that the control points are located at +20,+20 and -20,-20 from origin and end corner of screen respectively
SetControlPoints ULa, ULo, LLa, LLo, 20, 20, CornerX - 20, CornerY - 20

'Redraw w/ new scale
Receiver.ReDraw
'refresh
'Picture1.Refresh


End Sub

Private Sub ReadFile()
'This sub reads file
'it also populates a pull down menu style combo list
'so that user can select fish name/ID and track just that one fish

'First Entry: HEADER
'Fields: Code, Site, Date/time, location, lat (dec), log (dec)
'
Dim Entry As String
Dim FishNumber As Long
Dim ReceiverName As String
Dim ReceiverNumber As Long
Dim Code As String
Dim This_Time As Long
Dim This_Day As Single
Dim i As Long
Dim FishIndex(MAX_FISH) As Long
Dim ReceiverIndex(MAX_RECEIVERS) As Long
ReDim a(5) As String
Dim Previous_ReceiverName As String
Dim Previous_FishCode As String

'Scan and locate entries belonging to each fish
'
On Error GoTo ErrorTrap

ProgressBar.Value = 0
ProgressBar.Max = MemBuffer_Last
Do
    Entry = MemBuffer(i)
    ProgressBar.Value = i
    row = i
        'extract from CSV
        SplitEntry Entry
        
        'get number
        FishNumber = FN(i)
        'and receiver
        ReceiverNumber = RN(i)
        
        'update database
        If DateColumn = TimeColumn Then
            a = Split(SubStrings(DateTimeColumn))
            This_Day = DayNumber(a(0))
            This_Time = ConvertTime(a(1))
        Else
            This_Day = DayNumber(SubStrings(DateColumn))
            This_Time = ConvertTime(SubStrings(TimeColumn))
        End If
        
        With Stamp
            .Date = This_Day
            .Time = This_Time
            .Fish = FishNumber
            .Site = ReceiverNumber
            If This_Day < EPOCH Then .Valid = False
        End With
                      
            'write into table the stamp
            FishTable.WriteStamp FishNumber, FishIndex(FishNumber)
            ReceiverTable.WriteStamp ReceiverNumber, ReceiverIndex(ReceiverNumber)
            'accumulators
            FishIndex(FishNumber) = FishIndex(FishNumber) + 1
            ReceiverIndex(ReceiverNumber) = ReceiverIndex(ReceiverNumber) + 1
    i = i + 1
Loop While i < MemBuffer_Last
Close #1
ProgressBar.Value = 0: ProgressBar.Visible = False


'transfer date list
TransferDatesToList

Exit Sub

ErrorTrap:
    'if error invalidate stamp
    Stamp.Valid = False
    Resume Next
End Sub

Private Function ReadHeader(FileNumber As Long) As Long
'Read Header
'Change Constant here to reflect header size
Dim i As Long
Dim Header As String
Dim Columns() As String
Dim p As Long
Dim response As Variant
Dim WrongFormat As Boolean
Dim FormatOK As Boolean
Dim ReturnFlag As Long

On Error GoTo Its_A_Trap
'last line in header, should be:
'[Fish_Code]     ==> Fish code from VEMCO tag
'[Receiver_Name] ==> Receiver name or site name
'[Date/Time]     ==> Date time field from VEMCO's acoustic receiver
'[Group]         ==> Group / Release site
'[Lat]           ==> Receiver Lat Position
'[Long]          ==> Receiver Long Position
'OPTIONAL:
'[Abiotic]       ==> Abiotic Factor
'[Tag]           ==> Tag for receiver, used for analyzing receivers by tag group
'
'When line is not as such, it will default to the order above

'Default order
FishCodeColumn = 0
ReceiverNameColumn = 1
DateTimeColumn = 2
DateColumn = 2
TimeColumn = 2
GroupColumn = 3
LatColumn = 4
LongColumn = 5
TagColumn = -1
DirectionColumn = -1
FixedReceiverColumn = -1
AbioticColumn = -1


'skip some lines before leader header
For i = 1 To NUMBER_OF_LINES_IN_HEADER
    If EOF(1) Then GoTo Its_A_Trap
    Line Input #1, Header
Next i

'read header

'check for delimeter
p = InStr(1, Header, ",")
FormatOK = True
If p Then
    Columns = Split(Header, ",")
    'Column by column, check for tags
    For i = 0 To UBound(Columns)
        Select Case UCase(Columns(i))
            Case "[FISH_CODE]"
                FishCodeColumn = i
            Case "[RECEIVER_NAME]"
                ReceiverNameColumn = i
            Case "[DATE/TIME]"
                DateTimeColumn = i: DateColumn = i: TimeColumn = i
            Case "[DATE]"
                DateColumn = i
            Case "[TIME]"
                TimeColumn = i
            Case "[GROUP]"
                GroupColumn = i
            Case "[LAT]"
                LatColumn = i
            Case "[LONG]"
                LongColumn = i
            Case "[TAG]"
                TagColumn = i
            Case "[DIRECTION]"
                DirectionColumn = i
            Case Else
                WrongFormat = True
        End Select
    Next i
Else
    WrongFormat = True
End If

'assume all is ok
ReturnFlag = flag.OK
Its_A_Trap:
If WrongFormat Then
    response = MsgBox("The file you are importing is missing one or more column tags. This means it may not be properly formatted for AT.  Would you like to run the converter?", vbYesNoCancel)
    Select Case response
        Case vbYes
            ReturnFlag = flag.ConvertFile
        Case vbCancel
            ReturnFlag = flag.CancelFile
        Case vbNo
            ReturnFlag = flag.OK_But_Check
    End Select
End If

If Not FormatOK Then
    response = MsgBox("Error importing file: Incorrect format", vbOKOnly, "File Import Error")
    ReturnFlag = flag.CancelFile
End If

ReadHeader = ReturnFlag

End Function
Private Function OpenHeader() As Boolean
Dim Success As Boolean
Dim ReturnFlag As Long
Dim result As Variant
Dim NewFileName As String
Dim H As Long
Dim pid As Long
Dim PassedTest As Boolean
Dim FileOpened As Boolean

On Error GoTo ErrorHandeler
'Open read file
Open FileName For Input As #1

'Read Header
ReturnFlag = ReadHeader(1)

'set defaults
PassedTest = True

Select Case ReturnFlag
    Case flag.CancelFile
        Close #1
        Exit Function
    Case flag.ConvertFile
        Close #1
        NewFileName = AppendToFileName(FileName)
        pid = Shell(App.PATH & "\Converter " & FileName & " > " & NewFileName)
        H = OpenProcess(&H100000, True, pid)
        WaitForSingleObject H, -1
        CloseHandle H
    
        'now open file
        FileName = NewFileName
        'try again
    
        'Open read file
        Open FileName For Input As #1
        FileOpened = True
        'Read Header
        ReturnFlag = ReadHeader(1)
    Case flag.OK_But_Check
        Success = True
        PassedTest = TryToCastValues(1)
        FileOpened = True
End Select


'this has to be outside select because you read the header again and check for success when using converter (see last line inside select case ConvertFile)
If ReturnFlag = flag.OK Then Success = True

ErrorHandeler:
If Not Success Then
    result = MsgBox("Unable to open file.  Converter not found or access was denied. Make sure CONVERTER.EXE is in " & App.PATH, vbExclamation, "Import Error")
End If

If Not PassedTest Then
    result = MsgBox("Unable to open file.  This file is not formatted properly.  Try running the converter next time! ", vbExclamation, "Import Error")
    Success = False
End If

If FileOpened And Not Success Then
    Close #1
End If

OpenHeader = Success
End Function
Private Function TryToCastValues(BuffNumber As Long) As Boolean
Dim i As Long
Dim Header As String
Dim Success As Boolean
ReDim s(10) As String
Dim B As String

Dim FC As Long
ReDim DTC(2) As String
Dim d As Single
Dim t As Long

Dim LA As Single
Dim LO As Single
Dim R As String
Dim RG As String


Success = False

On Error GoTo ExitWithError


Line Input #BuffNumber, B

'split
s = Split(B, ",")

'FishCodeColumn
FC = CInt(CLng(s(FishCodeColumn)))
If FC <= 0 Then GoTo ExitWithError

'ReceiverNameColumn
R = s(ReceiverNameColumn)

'update database
If DateColumn = TimeColumn Then
    DTC = Split(s(DateTimeColumn))
    d = DayNumber(DTC(0))
    t = ConvertTime(DTC(1))
Else
    d = DayNumber(s(DateColumn))
    t = ConvertTime(s(TimeColumn))
End If

'GroupColumn
RG = s(GroupColumn)


'LatColumn = 4
'LongColumn = 5
'certain rules apply
LA = CSng(s(LatColumn))
LO = CSng(s(LongColumn))
If Abs(LO) > 180 Or Abs(LA) > 90 Then GoTo ExitWithError

Close #BuffNumber
'Its good to go.  Open it again

Open FileName For Input As BuffNumber

'skip some lines before leader header
For i = 1 To NUMBER_OF_LINES_IN_HEADER
    If EOF(BuffNumber) Then GoTo ExitWithError
    Line Input #BuffNumber, Header
Next i

'success!
Success = True

ExitWithError:

TryToCastValues = Success

End Function
Private Function AppendToFileName(s As String) As String
Dim p As Long
Dim c As String

p = InStrRev(s, "\")
If p <> 0 Then
    c = Left(s, p) & APPEND_STRING & Right(s, Len(s) - p)
Else
    c = s
End If

'return string
AppendToFileName = c

End Function
Private Sub PreLoadFile()
'This sub reads file in order to partition the two tables used by the main databases
'**It populates the receiver and the fish tables but not the databases**
Dim Entry As String
Dim i As Long
Dim Longitude As Single
Dim Latitude As Single
Dim Code As String
Dim Last_Time As Long
Dim This_Time As Long
Dim This_Day As Single
Dim Progress As Long
Dim FileLength As Long
Dim TotalLines As Long
Dim ReceiverName As String
Dim result As Variant
Dim FishNumber As Long
Dim FileNameWithoutPath
Dim Previous_ReceiverName As String
Dim Previous_FishCode As String
Dim OutputRow(MAX_ENTRIES) As String
Dim Rejected As Long
Dim FileOpened As Boolean
Dim ReturnFlag As Long

On Error GoTo ErrorHandeler

'clear table
For i = 0 To MAX_RECEIVERS
    SizeReceiverData(i) = 0
Next i
For i = 0 To MAX_FISH
    SizeFishData(i) = 0
Next i

'Prepare progress bar
FileLength = FileLen(FileName)
ProgressBar.Value = 0
ProgressBar.Refresh

'Scan and locate entries belonging to each fish
'load fish table in order
FileNameWithoutPath = ExtractFileName(FileName)
StatusBar.Panels(StatusPanel.Map) = "Extracting stamps... "

i = 0

Do
    Line Input #1, Entry
    If i = 0 Then
        TotalLines = CLng(FileLength / Len(Entry))
        ProgressBar.Max = TotalLines
    End If
    
    'save for loader
    row = i
    MemBuffer(i) = Entry
    SplitEntry Entry
    'Precond
    
    'get fishcode
    Code = SubStrings(FishCodeColumn)
    FishNumber = AssignFishNumber(Code)
    
    'enter in database
    FishDatabase.Code(FishNumber) = Code
    'set fish code and grouping variable (here, release site)
    FishDatabase.Release_Site = SubStrings(GroupColumn)
    StoreFishData FishNumber
    
    FN(i) = FishNumber
    Longitude = Abs(Val(SubStrings(LongColumn)))
    Latitude = Abs(Val(SubStrings(LatColumn)))

    'get receivername
    ReceiverName = SubStrings(ReceiverNameColumn)
    
    'validate
    RN(i) = StoreFixedReceiver(Longitude, Latitude, ReceiverName, FishNumber)
    
    'show progress
    If Progress < TotalLines Then Progress = Progress + 1
    ProgressBar.Value = Progress

AtLoopEnd:
    DoEvents 'do not halt the computer
    'next!
    i = i + 1
    
    'show progress
    If Progress < TotalLines Then Progress = Progress + 1
    ProgressBar.Value = Progress
    DoEvents
        
Loop While Not EOF(1)
Close #1

'store last entry
MemBuffer_Last = i


'show errors/warnings
If Rejected Then
    'open a rejected file just in case
    Open FileName & "_REJECTED" For Output As #2
    For i = 0 To Rejected
        Print #2, OutputRow(i)
    Next i
    Close #2
    result = MsgBox("Warning: Loaded with " & Str$(Rejected) & " bad long/lats AND/OR unconforming data points.  These data points have been rejected.  Rejected rows are in _REJECTION file on same directory.  Consult the User Manual for more information on how to fix this issue.", vbExclamation, "WARNING")
End If

'Transfer list of fish to floating window
FishDatabase.TransferList frmFloater.cmbFishCode


StatusBar.Panels(StatusPanel.Map) = "Partitioning tables..."
StatusBar.Panels(StatusPanel.Alert) = ""
'Finally, partition and name tables before loading any data into them
FishTable.PartitionTable SizeFishData()
ReceiverTable.PartitionTable SizeReceiverData()
Exit Sub

ErrorHandeler:
If i < 1 Then
    result = MsgBox("Incorrect File Format", vbExclamation, "Import Error")
    Close #1
Else
    OutputRow(Rejected) = Entry
    Rejected = Rejected + 1
    Resume AtLoopEnd
End If

End Sub
Private Sub SplitEntry(Entry As String)
Dim a() As String
Dim B() As String
Dim i As Long
Dim ii As Long
Dim temp As String
Dim Concatenate As String
Dim L As Long

'some databases output info with ""
'we need to take those out
If InStr(1, Entry, Chr$(34)) Then
    a = Split(Entry, ",")
    If a(0) = "" Then Exit Sub
    For i = 0 To UBound(a)
        B = Split(a(i), Chr$(34))
        temp = ""
        For ii = 0 To UBound(B)
            If Len(B(ii)) > 0 Then temp = B(ii)
        Next ii
        If temp = "" Then temp = a(i)
        Concatenate = Concatenate & "," & temp
    Next i
    L = Len(Concatenate)
    If L > 0 Then
        Entry = Right(Concatenate, L - 1)
    End If
End If

'here its the same
SubStrings = Split(Entry, ",")

End Sub
Private Sub StoreFishData(FishNumber As Long)

SizeFishData(FishNumber) = SizeFishData(FishNumber) + 1
FishDatabase.NumberOfStampsUP CInt(FishNumber) 'add stamp to ttl

End Sub
Private Function StoreFixedReceiver(Longitude As Single, Latitude As Single, ReceiverName As String, Code As Long) As Integer
'find a unique number for receiver, and store general info
Dim ReceiverNumber As Integer
Static RN As Integer
Static s As String

'if you already did this, no need to ask to lookup again!
If s = ReceiverName Then
    ReceiverNumber = Receiver.Store(Longitude, Latitude, ReceiverName, Code, RN)
Else
    ReceiverNumber = Receiver.Store(Longitude, Latitude, ReceiverName, Code)
    s = ReceiverName
    RN = ReceiverNumber
End If

'prep for partitioning stamp table
SizeReceiverData(ReceiverNumber) = SizeReceiverData(ReceiverNumber) + 1

StoreFixedReceiver = ReceiverNumber

End Function
Public Sub TransferDatesToList()
'Transfers list of dates in so user can see and select them
'
Dim i As Long

With frmFloater
    For i = 0 To .lstDates.ListCount - 1
        .lstDates_Std.AddItem .lstDates.List(i)
    Next i
End With

End Sub

Private Function Transform_To_Number(s As String) As Long
'Take a time stamp and creates a number that can be sorted, compared, etc.
ReDim Digits(5) As String
Dim Concatenated As String
Dim i As Long


'error trap
If s = "" Then
    MsgBox ("Error #002")
    Exit Function
End If

Digits = Split(s, ":")

For i = 0 To 1
    Concatenated = Concatenated & Digits(i)
Next i

Concatenated = Trim(Concatenated)

Transform_To_Number = CLng(Concatenated)



End Function
Private Function AssignFishNumber(Code As String) As Integer
AssignFishNumber = FishDatabase.GetFishNumber(Code)
End Function

Private Function InMinutes(t As String) As Long
'one liner
InMinutes = hour(t) * 60 + minute(t)

End Function
Private Function Degrees(d As Single) As String
'Decimal degrees to Classical Degrees/Minutes
'Entry: Decimal
'Exit: "DD'MM"

Dim Decimal_Portion As Single
Dim Degrees_Portion As Single

Degrees_Portion = Int(d)
Decimal_Portion = d - Int(d)

Degrees = Str$(Degrees_Portion) & "'" & Str$((Decimal_Portion * 60))

End Function

Private Sub mnuExportStrings_Click()
'Export track data to .CSV file
Dim FileName As String

With CommonDialog
    .FileName = ""
    .DialogTitle = "Export to a CSV File"
    .CancelError = False
    'ToDo: set the flags and attributes of the common dialog control
    .Filter = "CSV Files (*.csv)|*.csv"
    .ShowSave
End With
If Len(CommonDialog.FileName) = 0 Then
    Exit Sub
End If

FileName = CommonDialog.FileName

ExportToNavStringData FileName
End Sub
Private Sub ExportToNavStringData(FileName As String)
'Exports the nav strings of a group of fish
Dim BufferNumber As Integer
Dim FishNumber As Long

If frmFloater.cmbFishCode.ListCount = 0 Then Exit Sub

BufferNumber = FreeFile
Open FileName For Output As #BufferNumber
'copyright notices
Print #BufferNumber, "AQUATRACKER Navigation String Export"
Print #BufferNumber, "by Jose J. Reyes-Tomassini / NOAA Fisheries"
Print #BufferNumber, ""
Print #BufferNumber, "Fish Code, String"
'Get index to code number
For FishNumber = 0 To frmFloater.cmbFishCode.ListCount - 1
     Print #BufferNumber, FishDatabase.Code(FishNumber) & "," & TrackCalculator.CreateVerboseTrackString(FishNumber)
Next FishNumber

Close #BufferNumber


End Sub

Private Sub mnuFilter_Click()
frmFilter.Show

End Sub

Private Sub mnuFindOverlaps_Click()
'overlap analysis
Dim Threshold As Long
Dim result As Variant

'show all receivers
ShowDetectors

'get threshold
result = InputBox("Enter threshold for minimum time between receiver detection at different sites", "Threshold", "1")
If Val(result) > 0 Then
    Threshold = CLng(result)
    frmOverlaps.Show
    TrackCalculator.AnalyzeOverlaps Threshold, frmFloater.cmbFishCode.ListCount - 1, frmOverlaps.lstOverlaps, frmOverlaps.lstOverlapParsed, frmOverlaps.lstFishes
End If



End Sub



Private Sub mnuHighLightColorChange_Click()
On Error GoTo ExitWithError
With CommonDialog
    .CancelError = True
    .ShowColor
    HighLightColor = .Color
End With
ExitWithError:
'Nop
End Sub

Private Sub mnuIdentifyFishGroups_Click()
'IDs fish groups
Dim Threshold As Long
Dim result As Variant

'get threshold
result = InputBox("Enter threshold for minimum time between co-detections", "Co-detection Threshold", "1")
If Val(result) > 0 Then
    Threshold = CLng(result)


    StatusBar.Panels(StatusPanel.Map) = "Calculating fish groups..."
    
    Receiver.FindFishGroups Threshold
    
    'load form
    Load frmFishGroups
    
    'set slider
    frmFishGroups.sldrThreshold.Value = Threshold
    
    'show new value
    frmFishGroups.lblThreshold.Caption = Str$(Threshold) & " Min"


    StatusBar.Panels(StatusPanel.Map) = ""

    'show form
    frmFishGroups.Show
End If

End Sub

Private Sub mnuImport_Click()
'call with no arguments to choose selection
OpenFile ""
End Sub

Private Sub mnuImportSunriseData_Click()
LoadSunSetData
End Sub

Private Sub mnuLandAvoidanceChoice_Click(Index As Integer)
Dim Warning As Boolean
Dim i As Long
Dim temp As Long
Dim response As Variant
Static AlreadySelected As Boolean
Static PastSelection As Long

'select avoidance mode
For i = 0 To mnuLandAvoidanceChoice.UBound
    If i <> 0 And mnuLandAvoidanceChoice(i).Checked = True Then
        Warning = True
        temp = i
    End If
    mnuLandAvoidanceChoice(i).Checked = False
Next i

If (Index > 0 And Warning And Index <> PastSelection) Or (Index > 0 And AlreadySelected And Index <> PastSelection) Then
    'warn of issues when switching between two different types of LA
    response = MsgBox("Warning: If you switch between land avoidance after the routes have been defined, program could crash or give plot inaccurate tracks.  YOU HAVE BEEN WARNED! I suggest you cancel this action...", vbOKCancel, "Warning")
    If response = vbCancel Then
        mnuLandAvoidanceChoice(i).Checked = True
        Exit Sub
    End If
End If

'store and show
Land_Avoidance = Index
AlreadySelected = True
If Index <> 0 Then PastSelection = Index
mnuLandAvoidanceChoice(Index).Checked = True
frmFloater.RefreshCanvas
End Sub

Private Sub mnuLandAvoidanceOptions_Click()
frmLandAvoidanceOptions.Show
End Sub

Private Sub ResetAll()
'reset all loaded classes and tables
'And Clear hashtables

'Reset internal counters and tables for database
FishTable.ResetTable
ReceiverTable.ResetTable

'Reset receiver and fish objects
'Set all groups to nothing and all routes to nothing
Receiver.Reset
FishDatabase.Clear

'init receiver assigments colors
Receiver_Table.ResetReceivers

'Done!
'Show floating window
Load frmFloater
With frmFloater.cmbFishCode
    .Clear
    .AddItem "ALL"
End With

End Sub
Private Sub mnuLoadFile_Click()
Dim FileName As String

'Reset ALL values
ResetAll

'load file

'Open file
With CommonDialog
    .FileName = ""
    .DialogTitle = "Load data file"
    .CancelError = False
    .Filter = "AquaTracker Native File Format (*.aqn)|*.aqn|Comma Separated Values CSV(*.csv)|*.csv"
    .FilterIndex = 1
    .ShowOpen
    If Len(.FileName) = 0 Then
        Exit Sub
    End If
    'load file
    FileName = CommonDialog.FileName
    'select type
    Select Case .FilterIndex
        Case FileType.Native
            LoadAQN FileName
        Case FileType.CSV
            OpenFile FileName
    End Select
    mnuFindOverlaps.Enabled = True
End With


End Sub
Public Sub LoadAQN(FileName As String)


Dim FileManager As New clsFileManager
Dim t As Single
On Error GoTo ExitWithError
t = Timer


'Map

'show hour glass
Form1.MousePointer = vbHourglass

If ProgramSettings.MapFile <> "" Then
    ScanMap
End If

'load aqn file
MemBuffer_Last = FileManager.LoadNativeFormat(FileName, Form1.ProgressBar)

'store in registry
ProgramSettings.LastFileLoaded = FileName

mnuSaveAQN.Enabled = True

'get tracks and floater on display!
LoadFloater FileName, t

Exit Sub

ExitWithError:
StatusBar.Panels(StatusPanel.Map) = "Unable to open AQN file."
End Sub
Private Sub mnuLoadMap_Click()
Dim results As Variant
Dim SizeX As Long
Dim SizeY As Long
Dim CurrentMapName As String

SizeX = Picture1.Width
SizeY = Picture1.Height

'store
CurrentMapName = MapFileName

'call sub to get map loaded
LoadFromMap

'load into memory
If MapFileName <> "" And CurrentMapName <> MapFileName Then
    'Ask if current Map is OK
    results = MsgBox("Accept map?", vbYesNo, "Map")
    If results = vbYes Then ScanMap
    'ask if change scale now?
    results = MsgBox("Auto-scale both axes to fit plot area?", vbYesNo, "Auto-scale")
    If results = vbYes Then AutoScale
    
    'last question, is this going to be the map?
    results = MsgBox("Make this your map on startup next time?", vbYesNo, "Map")
    If results = vbYes Then
        'set map
        SaveSetting APPLICATION, REGISTRY_SECTION, "MapFile", ProgramSettings.MapFile
    End If
    
    frmFloater.Show
    If MapWasScanned Then
        Receiver.ReDraw 'update receivers
    End If
    
    'refresh canvas
    RePaint
End If

End Sub
Private Sub LoadFromMap(Optional MapFile As String = "")
Dim results As Variant
Dim SizeX As Long
Dim SizeY As Long


On Error GoTo Unable_To_Load_Map


'load from registry filename or query filename?
If MapFile = "" Then
    With CommonDialog
        .DialogTitle = "Open BMP File"
        .CancelError = False
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "BMP Files (*.bmp)|*.bmp"
        .ShowOpen
        If Len(CommonDialog.FileName) = 0 Then
            Exit Sub
        End If
        MapFile = .FileName
    End With
    
    'clear map
    ImageProcessingEngine.ClearMap
    MapWasScanned = False
End If

MapFileName = MapFile
ProgramSettings.MapFile = MapFile

'get old size
SizeX = Picture1.Width
SizeY = Picture1.Height


'get picture
Picture1 = LoadPicture(MapFileName)
Picture2 = LoadPicture(MapFileName)

AdjustWindow SizeX, SizeY

Exit Sub

Unable_To_Load_Map:
    Form1.Picture1.Cls
    results = MsgBox("Unable to load map or default map no longer valid!", vbOKOnly)
    
End Sub
Private Sub mnuMarkov_Click()
Dim result As Variant
Dim R As Integer
Dim AskQuestion As Boolean

'are all receivers visible
For R = 1 To Receiver.TotalReceivers
    If Receiver_Table.Invisible(R) And Receiver.GroupNumber(R) = 0 Then AskQuestion = True
Next R

If AskQuestion Then
    result = MsgBox("Some receivers are set as EXCLUDED.  Do you want to include these receivers in the Markov Chain analysis?", vbYesNo, "Markov Chain")
    If result = vbYes Then
        Receiver.MakeVisible
    End If
End If
   
frmMarkov.Show , Me


End Sub

Private Sub mnuNavStringShow_Click()
frmNavigationString.Show vbModeless, Form1
End Sub

Private Sub mnuNewCanvas_Click()
'Replaces canvas with a blank.  Changes canvas size
Dim Canvas_Width As Long
Dim Canvas_Height As Long
Dim results As String
Dim SizeX As Long
Dim SizeY As Long

'get old size
SizeX = Picture1.Width
SizeY = Picture1.Height

'set map to no map
Set Picture1.Picture = Nothing
Set Picture2.Picture = Nothing

'disable stuff
frmFloater.Show

MapFileName = ""
MapWasScanned = False


'clear picture
ImageProcessingEngine.ClearMap

'no map on settings
ProgramSettings.MapFile = ""

'new dimmensions

'HEIGHT
results = InputBox("Canvas Height", "Height", Str$(Picture1.Height))

If results <> "" Then
    Canvas_Height = Val(results)
Else
    Exit Sub
End If

'WIDTH
results = InputBox("Canvas Width", "Width", Str$(Picture1.Width))

If results <> "" Then
    Canvas_Width = Val(results)
Else
    Exit Sub
End If

If Canvas_Height > 0 And Canvas_Height <= MAX_Y Then
    Picture1.Height = Canvas_Height
    Picture2.Height = Canvas_Height
End If

If Canvas_Width > 0 And Canvas_Width <= MAX_X Then
    Picture1.Width = Canvas_Width
    Picture2.Width = Canvas_Width
End If

AdjustWindow SizeX, SizeY


'change scale now?
If FishDatabase.TotalFishLoaded >= 1 Then
    results = MsgBox("Auto-scale both axis to fit plot area?", vbYesNo, "Auto-scale")
    If results = vbYes Then AutoScale
End If

'last question, is this going to be the map?
results = MsgBox("Use this blank canvas on startup next time?", vbYesNo, "Canvas")
If results = vbYes Then
    'set map to blank in registry
    SaveSetting APPLICATION, REGISTRY_SECTION, "MapFile", ""
    'write size
    SaveSetting APPLICATION, REGISTRY_SECTION, "Size_Map_Width", Str$(Canvas_Width)
    SaveSetting APPLICATION, REGISTRY_SECTION, "Size_Map_Height", Str$(Canvas_Height)
End If

'refresh canvas
frmFloater.RefreshCanvas
End Sub
Private Sub AdjustWindow(SX As Long, SY As Long)
Dim w As Long
Dim H As Long
Dim StatusBarWidth As Single
Dim i As Long

'find stat bar min needed size
For i = 1 To StatusBar.Panels.count
    StatusBarWidth = StatusBarWidth + StatusBar.Panels(i).Width
Next i
StatusBarWidth = StatusBarWidth * TwipsPerPixelX()
If Picture1.Width > StatusBarWidth Then StatusBarWidth = Picture1.Width

'progress bar
ProgressBar.Top = Picture1.Height
ProgressBar.Width = StatusBar.Width

'auto adjust form to fit picture
'pad in Y for status bar and in X for neatness
Form1.Width = StatusBarWidth
Form1.Height = (Picture1.Height + ProgressBar.Height + StatusBar.Height + 75) * TwipsPerPixelY()


End Sub
Private Function ExtractFileName(PATH As String) As String
Dim L As Long
Dim Name As String
Dim p As Long

'gets name of file from a given path separated by "\"

'invalid path?
If PATH = "" Then Exit Function

L = Len(PATH)
p = InStrRev(PATH, "\")
Name = Right(PATH, L - p)

ExtractFileName = Name

End Function

Private Sub mnuNight_Click()
WhatToShowOnCanvas = ShowOnCanvas.NightDensity
ColorScale.AutoSet = True
ShowDensity_DielPhase Diel.Night
End Sub
Private Sub ShowDensity_DielPhase(Phase As Long)
Dim ReceiverNumber As Integer
Dim N As Long
Dim Max As Long
'
StatusBar.Panels(StatusPanel.Map) = "Showing density by diel phase"
'gather receiver information
For ReceiverNumber = 1 To Receiver.TotalReceivers
    N = Receiver.CountUniqueEntriesInReceiver_ByDielPhase(ReceiverNumber, Phase)
    Density(ReceiverNumber) = N
    If N > Max Then
        Max = N
    End If
Next ReceiverNumber

ImageProcessingEngine.DrawDensityPlotForReceivers Form1.Picture1, Max, Density, LARGE_MARKER

'Now receivers are visible
Receivers_Are_Visible = True
If DateListIsLoaded Then frmDates.UpdateList
End Sub
Public Sub OpenFile(s As String)
Dim results As Variant
Dim i As Long
Dim Do_Not_ReDraw As Boolean
Dim t As Single
Dim FileOpened As Boolean

Dim NumberOfFish As Long

'no redraw for receivers (they will be redrawn later!)

Do_Not_ReDraw = False

FileName = s
'change scale now?
If ProgramSettings.MapFile = "" Then
    results = MsgBox("Auto-scale both axis to fit plot area?", vbYesNo, "Auto-scale")
    If results = vbYes Then AutoScale
End If

ProgramSettings.LastFileLoaded = ""

'read map b4 reading file

If ProgramSettings.MapFile <> "" Then
    ScanMap
End If


'show hour glass
Form1.MousePointer = vbHourglass


'Success
FileOpened = OpenHeader

If Not FileOpened Then
    StatusBar.Panels(StatusPanel.Map) = "Unable to open file."
    'show hour glass
    Form1.MousePointer = vbArrow
    Exit Sub
End If

t = Timer

'Read File
PreLoadFile

'empty file? there was an error
If MemBuffer_Last = 0 Then
    StatusBar.Panels(StatusPanel.Map) = "Failed to load stamps."
    'show hour glass
    Form1.MousePointer = vbArrow
    Exit Sub
End If

'save new table
StatusBar.Panels(StatusPanel.Map) = "Preparing and sorting stamps..."

ReadFile

'Backup receiver data
Receiver.Backup

mnuSaveAQN.Enabled = True

LoadFloater FileName, t
End Sub
Private Sub LoadFloater(f As String, t As Single)
Dim i As Long
Dim FileNameWithoutPath As String
Dim NumberOfFish As Long
Dim LoadSpeed As Single

FileNameWithoutPath = ExtractFileName(f)

'show hour glass
Form1.MousePointer = vbArrow

'update actions window
frmFloater.Caption = "Actions on " & FileNameWithoutPath

'Enable WALK Buttons
mnuReceivers.Enabled = True
frmFloater.cmdWalk.Enabled = True

NumberOfFish = FishDatabase.TotalFishLoaded
StatusBar.Panels(StatusPanel.Map) = Str$(Receiver.TotalReceivers) & " receivers loaded to memory. " & Str$(NumberOfFish) & " tracks loaded."

'clear summary
TrackCalculator.Reset

'draw tracks in memory
Form1.Picture2.Visible = True
Form1.Picture1.Visible = False
'show all tracks
For i = 0 To frmFloater.cmbFishCode.ListCount - 1
    ShowTrack i, Form1.Picture1
    TrackCalculator.ComputeSummary
    CURRENT_FISH = i
Next i

'get averages
TrackCalculator.Average

'make visible
Form1.Picture2.Visible = False
Form1.Picture1.Visible = True

CURRENT_FISH = -1

frmFloater.cmbFishCode.Text = "ALL"
'now enable selection
frmFloater.cmbFishCode.Enabled = True
frmFloater.Show vbModeless, Form1

LoadSpeed = Timer - t
StatusBar.Panels(StatusPanel.Alert) = "Loaded " & Format(MemBuffer_Last, "###,###,###") & " stamps in " & Format((LoadSpeed), "0.0") & "s"
End Sub
Private Sub mnuProximityRadius_Click()
'Enter Proximity Radius
Dim result As Variant
result = InputBox("Range Radius in km:", "Range Radius", Proximity)

If result <> "" And result > 0 Then
    Proximity = result
End If

End Sub

Private Sub mnuOptions_Dates_Click()
frmDates.Show
End Sub

Private Sub mnuOptions_Exclude_Receivers_Click()
frmSelectReceivers.Show
End Sub

Private Sub mnuOptions_Exclude_Tracks_Click()
frmFishTrackWindow.Show
End Sub

Private Sub mnuPercentFish_Click()
WhatToShowOnCanvas = ShowOnCanvas.FishDensity
ShowFishDensity
End Sub
Private Sub ShowFishDensity()

Dim ReceiverNumber As Integer
Dim N As Long
Dim Max As Long
Dim R As Integer
Dim G As Integer
Dim FirstReceiverInGroup(MAX_GROUPS) As Integer
Dim GroupSeen(MAX_GROUPS) As Boolean

WhatToShowOnCanvas = ShowOnCanvas.FishDensity

'gather receiver information
For ReceiverNumber = 1 To Receiver.TotalReceivers
    N = Receiver.CountUniqueEntries(ReceiverNumber, FieldNames.Fish)
    Density(ReceiverNumber) = N
    If N > Max Then
        Max = N
    End If
Next ReceiverNumber


ImageProcessingEngine.DrawDensityPlotForReceivers Form1.Picture1, Max, Density, LARGE_MARKER

'Now receivers are visible
Receivers_Are_Visible = True
If DateListIsLoaded Then frmDates.UpdateList

End Sub
Private Sub mnuPickWaterColor_Click()
ChooseWaterColor = True
frmPickColorOfWater.Show

End Sub



Private Sub mnuReceiverGroupWindow_Click()
frmAssignReceiverToGroup.Show
End Sub

Private Sub mnuReceiverInformation_Click()
'Show Receiver Information Window
frmReceiverInformation.Show

End Sub

Private Sub mnuReceiverResidence_Click()
If CURRENT_FISH <> -1 Then
    frmReceiversResidenceTime.Show vbModeless, Form1
End If

End Sub

Private Sub mnuResidenceAllFish_Click()
Dim FileName As String
Dim FishIndex As Integer
Const FileAccessNumber = 1
Dim Pointer As Variant

'choose file

With CommonDialog
    .FileName = ""
    .DialogTitle = "Export to a CSV File"
    .CancelError = False
    .Filter = "CSV Files (*.csv)|*.csv"
    .ShowSave
End With

If Len(CommonDialog.FileName) = 0 Then
    Exit Sub
End If

FileName = CommonDialog.FileName
'advise of updating
Form1.MousePointer = vbHourglass
Pointer = Form1.Picture1.MousePointer
Form1.Picture1.MousePointer = vbHourglass


'open file and write header
Open FileName For Output As #FileAccessNumber
WriteHeaderForResidenceInformationFile FileAccessNumber

'Calculates residence for all fish
'From first fish to last fish
'Will not calculate it if fish is not visible
For FishIndex = 0 To FishDatabase.TotalFishLoaded
    If FishDatabase.IsVisible(FishIndex) Then
        WriteResidenceInformation FishIndex, FileAccessNumber
    End If
Next FishIndex

Close #FileAccessNumber


'go back to normal
Form1.MousePointer = vbArrow
Form1.Picture1.MousePointer = Pointer
End Sub
Private Sub WriteHeaderForResidenceInformationFile(FileAccessNumber As Long)
'write the header
'header is all the receivers
Dim R As Integer
Dim s As String

'get all the names
For R = 1 To Receiver.TotalReceivers
    s = s & "," & Receiver.ID(R)
Next R

Print #FileAccessNumber, s

End Sub
Private Sub WriteResidenceInformation(FishIndex As Integer, FileAccessNumber As Long)
Dim ResidenceTime(MAX_RECEIVERS) As Long
Dim R As Integer
Dim s As String

'get residence for fish
TrackCalculator.ComputeResidence ResidenceTime, FishIndex

'Fish
s = FishDatabase.Code(FishIndex)

For R = 1 To Receiver.TotalReceivers
    s = s & "," & Str$(ResidenceTime(R))
Next R

Print #FileAccessNumber, s


End Sub


Private Sub mnuSaveAQN_Click()
'Export or append information to file
'as of now, no append
Dim FileName As String
Dim FileManager As New clsFileManager
Dim L As Long
Dim response As Variant

'Open file
With CommonDialog
    .FileName = ""
    .DialogTitle = "Save as AquaTracker Native file..."
    .CancelError = False
    .Filter = "AquaTracker Native File Format (*.aqn)|*.aqn"
    .ShowSave
End With
If Len(CommonDialog.FileName) = 0 Then
    Exit Sub
End If

FileName = CommonDialog.FileName
On Error Resume Next
L = FileLen(FileName)
If L > 0 Then
    response = MsgBox("Our records indicate you are attempting to overwrite an existing file.  Are you sure you want to proceed?", vbYesNoCancel, "File already exists")
    If response <> vbYes Then Exit Sub
End If

'Save file
FileManager.SaveInNativeFormat (FileName)

'store in registry
ProgramSettings.LastFileLoaded = FileName

'show hour glass
Form1.MousePointer = vbArrow

'Show on title bar
'update actions window
frmFloater.Caption = "Actions on " & ExtractFileName(FileName)
End Sub

Private Sub mnuSaveExcursions_Click()
Dim FileName As String
Dim result As Long
With CommonDialog
    .FileName = ""
    .DialogTitle = "Save as... CSV File"
    .CancelError = False
    'ToDo: set the flags and attributes of the common dialog control
    .Filter = "CSV Files (*.csv)|*.csv"
    .ShowSave
End With

If Len(CommonDialog.FileName) = 0 Then
    Exit Sub
End If

FileName = CommonDialog.FileName


'length min of excursion
result = CLng(InputBox("Enter minimum excursion time in minutes:", "Excursion Threshold", "10"))

If result >= 0 Then
    Excursion.Threshold_Excursion = result
Else
    Exit Sub
End If

'min time to count as residence
result = CLng(InputBox("Enter minimum residence time in minutes:", "Residence Threshold", "10"))

If result >= 0 Then
    Excursion.Threshold_Residence = result
Else
    Exit Sub
End If

Excursion.ExportExcursion 0, FileName

End Sub




Private Sub mnuSelectFields_Click()
frmSelectAnalysis.Show

End Sub

Private Sub mnuSetAsRefTrack_Click()
'Set as ref track, the current track on display, then start UDT window
frmUserDefinedTrack.Show
End Sub

Private Sub mnuShowAllReceivers_Click()
ShowAllReceivers
End Sub
Public Sub ShowAllReceivers()
WhatToShowOnCanvas = ShowOnCanvas.Receivers
ShowDetectors
End Sub


Private Sub mnuShowDetectorInfo_Click()
Dim Station As Integer
Dim NextPageStart As Long
Me.Tag = ""

frmShowPings.Show
Station = Receiver.CurrentStation_Number
Load frmShowPings
frmShowPings.Caption = Receiver.ID(Station)
NextPageStart = Receiver.ShowAllPings(Station, frmShowPings.picPings)


End Sub

Private Sub mnuShowExcursions_Click()

Dim Excursion_TH As Long
Dim Residence_TH As Long
Dim response As Variant

'length min of excursion
Excursion_TH = CLng(InputBox("Enter minimum excursion time in minutes:", "Excursion Threshold", "30"))

If Excursion_TH >= 0 Then
    Excursion.Threshold_Excursion = Excursion_TH
End If

'min time to count as residence
Residence_TH = CLng(InputBox("Enter minimum residence time in minutes:", "Residence Threshold", "12"))

If Residence_TH >= 0 Then
    Excursion.Threshold_Residence = Residence_TH
End If

If Residence_TH >= Excursion_TH Then
    response = MsgBox("Warning: The threshold for residence is equal to or exceeds the threshold for excursions.  The analysis may have unpredictable results.  Are you sure you want to proceed?", vbYesNo, "Excursion Analysis Warning")
    If response = vbNo Then Exit Sub
End If

frmExcursions.Show

End Sub
Private Sub mnuShowTrackDetails_Click()
Dim Station As Integer

Dim NextPageStart As Long
Dim FishNumber As Long

'get fish number
FishNumber = frmFloater.cmbFishCode.ListIndex - 1


Me.Tag = Str$(FishNumber)
frmShowPings.Show
Station = Receiver.CurrentStation_Number
Load frmShowPings
frmShowPings.Caption = Receiver.ID(CLng(Station))
NextPageStart = Receiver.ShowAllPings(Station, frmShowPings.picPings, FishNumber)

End Sub

Private Sub mnuStayTime_Click()
'change stay time th
Dim TH As Integer
Dim s As String
s = InputBox("Enter a threshold setting for residency time (in mins.).  If fish stays in same station less than this value, it is counted as a stay.", "Residency Time Threshold", ResidenceThreshold)
If IsNumeric(s) Then
    TH = CInt(s)
    If TH >= 0 Then
        ResidenceThreshold = TH
        frmFloater.RefreshCanvas
    End If
End If
    
End Sub

Private Sub LoadSunSetData()
Dim results As Variant
Dim FileName As String

results = MsgBox("Warning: Aquatracker can only accepts certain text files! Read manual before using this feature!", vbOKOnly)

With CommonDialog
    .DialogTitle = "Open Sunrise/Sunset txt File"
    .CancelError = False
    'ToDo: set the flags and attributes of the common dialog control
    .Filter = "Photoperiod File (*.txt)|*.txt"
    .ShowOpen
End With

If Len(CommonDialog.FileName) = 0 Then
    Exit Sub
End If

FileName = CommonDialog.FileName

PhotoPeriodCalculator.ReadFile FileName

End Sub

Private Sub mnuTrack_ReDoTrack_Click()
Dim FishNumber As Integer

'get fish number
With frmFloater.cmbFishCode
    If .Text <> "" And .Text <> "ALL" Then
        'get fish #
        FishNumber = .ListIndex - 1
    End If
End With

'ignore routes
'and then retrack
Receiver.IgnorePreviousTrackRoute = True
'cls
ClearScreen
ShowTrack FishNumber, Form1.Picture1
'return to normal
Receiver.IgnorePreviousTrackRoute = False
End Sub

Private Sub mnuTrackAnimationStyle_Click(Index As Integer)
Dim i As Long
Dim Selected As Long
mnuTrackAnimationStyle(Index).Checked = Not mnuTrackAnimationStyle(Index).Checked
End Sub



Private Sub Picture1_KeyPress(KeyAscii As Integer)
'if ctrl-c then copy canvas
If KeyAscii = 3 Then CopyCanvasImage
ABORT_PROCESS = True
End Sub

Private Sub Picture1_LostFocus()
If Tool_Number = ToolBox.Calibrate_Tool Then Picture1.Cls
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Station As Long

'''''''''''''''''''''''''''''''''''''''''''''
'From here on, if set to disable, quit
If Disable_MouseOver_Event Then Exit Sub


If ChooseWaterColor Then frmPickColorOfWater.shpWater.FillColor = Picture1.Point(X, Y): Exit Sub


ClickedReceiver = 0
If Shift Then ShowAllReceivers

Select Case Button
    Case vbLeftButton
        'If left button is clicked
        'activate tool from toolbox
        CallTool Tool_Number, X, Y
    Case vbRightButton
        If MultiSelect Then
            mnuReceiverInformation.Visible = False
            mnuShowExcursions.Visible = False
            mnuShowDetectorInfo.Visible = False
            mnuDielPatternPings.Visible = False
            mnuShowStamps.Visible = False
            
            PopupMenu mnuReceiver
            
            mnuReceiverInformation.Visible = True
            mnuShowExcursions.Visible = True
            mnuShowDetectorInfo.Visible = True
            mnuDielPatternPings.Visible = True
            mnuShowStamps.Visible = True
            MultiSelect = False
            ClearScreen
            RePaint
        Else
        
            'ID Current Station
            Station = IDStation(X, Y)
            If Station > 1 Then
                'Receiver Context Menu
                PopupMenu mnuReceiver
            Else
                'track menu
                If CURRENT_FISH = -1 Then
                    mnuChangeTrackColor.Enabled = False
                    mnuSetAsRefTrack.Enabled = False
                Else
                    mnuChangeTrackColor.Enabled = True
                    mnuSetAsRefTrack.Enabled = True
                End If
                PopupMenu mnuTrack
            End If
        End If
End Select
    
End Sub
Private Sub Plot_Path(X As Long, Y As Long)
Dim Index As Long
Dim FromSite As Long
Dim Site As Long

'ID current station
'If on a station then use this as path point, otherwise ignore
If IDStation(X, Y) > 1 Then
        Site = Receiver.CurrentStation_Number
        'add to route
        Receiver.AddReceiverToUserDefinedTrack Site
        'add to list on screen
        frmUserDefinedTrack.lstReceivers.AddItem Receiver.ID(CInt(Site))
        frmUserDefinedTrack.cmdCalculate.Enabled = True
End If



'add tracks and draw
FromSite = Receiver.UserDefinedTrack_Receiver(0)

'calculate UD track params
Do While Receiver.UserDefinedTrack_Receiver(Index) <> -1
    Site = Receiver.UserDefinedTrack_Receiver(Index)
    Receiver.DrawRoute Site, FromSite, Form1.Picture1, vbRed
    Index = Index + 1
    FromSite = Site
Loop

End Sub
Private Sub CallTool(Tool As Integer, X As Single, Y As Single)
'Selects tool from toolbox
Select Case Tool
    Case ToolBox.Select_Tool
        SelectTool X, Y
    Case ToolBox.Calibrate_Tool
        CalibrateTool X, Y
    Case ToolBox.Measure_Tool
        Measure_Tool X, Y
    Case ToolBox.Plot_Tool
        'Picture1.Refresh
        Plot_Path CLng(X), CLng(Y)
    Case ToolBox.DrawFishCorridor
        DrawFishCorridor X, Y
    Case ToolBox.Zoom
        Zoom X, Y
End Select
End Sub
Private Sub DrawFishCorridor(X As Single, Y As Single)
Dim CurrentAnchor As Integer
Dim NextAnchorPoint_X As Integer
Dim NextAnchorPoint_Y As Integer
    
'Get current x,y and anchor
CurrentAnchor = AnchorMap(X, Y)

If ANCHORING Then
    If AnchorPointNotValid Then
        With FishCorridor_AnchorPoints(CurrentAnchorNumber)
            .ConnectTo = -1
        End With
        DrawAnchors X, Y
        'end anchoring
        ANCHORING = False
    Else
        If CurrentAnchor = 0 Then
            Anchors = Anchors + 1
            If Anchors > MAX_ANCHORS Then Anchors = MAX_ANCHORS
            With FishCorridor_AnchorPoints(CurrentAnchorNumber)
                .ConnectTo = Anchors
            End With
            With FishCorridor_AnchorPoints(Anchors)
                .X = CInt(X)
                .Y = CInt(Y)
                .ConnectTo = 0
            End With
            MarkAnchorInMemory X, Y, Anchors
            CurrentAnchorNumber = Anchors
        Else
            FishCorridor_AnchorPoints(CurrentAnchorNumber).ConnectTo = CurrentAnchor
        End If
    End If
Else
    'Set flag
    ANCHORING = True
    
    'Change cursor to crosshairs
    Form1.Picture1.MousePointer = vbCross
    
    'if no anchor, deploy as point and allow to continue drawing
    If CurrentAnchor = 0 Then
        'validate point
        If MapImage(X, Y) <> 0 Then
            Anchors = Anchors + 1
            If Anchors > MAX_ANCHORS Then Anchors = MAX_ANCHORS
            CurrentAnchorNumber = Anchors
            With FishCorridor_AnchorPoints(Anchors)
                .X = CInt(X)
                .Y = CInt(Y)
                .ConnectTo = 0
            End With
            MarkAnchorInMemory X, Y, CurrentAnchor
        Else
            ANCHORING = False
        End If
    Else
        CurrentAnchorNumber = CurrentAnchor
        FishCorridor_AnchorPoints(CurrentAnchor).ConnectTo = 0
    End If
End If

'on exit, if anchoring is now off, set the newly drawn corridor in memory
If Not ANCHORING Then
    Form1.Picture1.MousePointer = vbArrow
    'redraw with no connectors showing and no receivers
    DrawAnchors X, Y, 0
    ImageProcessingEngine.StoreFishCorridor
    DrawAnchors X, Y
    ImageProcessingEngine.ConnectReceiversToCorridor
    ImageProcessingEngine.DrawReceiverConnectionsToCorridor
    'warn if not connected
    Load frmUnconnectedReceivers
    ImageProcessingEngine.ShowUnConnectedReceiverTable frmUnconnectedReceivers.lstUnconnectedReceivers
    If frmUnconnectedReceivers.lstUnconnectedReceivers.ListCount > 0 Then
        frmUnconnectedReceivers.Show
    Else
        Unload frmUnconnectedReceivers
    End If
End If

End Sub
Private Sub MarkAnchorInMemory(X As Single, Y As Single, AnchorNumber As Integer)
Dim i As Long
Dim ii As Long
Dim LX As Long
Dim LY As Long

'draw in map
LX = CLng(X) - 1
LY = CLng(Y) - 1
For i = 0 To 1
    LX = LX + i
    For ii = 0 To 1
        LY = LY + ii
        'make sure its well marked!
        If (LX <= MAX_X And LX >= 0) And (LY <= MAX_Y And LY >= 0) Then AnchorMap(LX, LY) = AnchorNumber
    Next ii
Next i
End Sub

Private Sub Measure_Tool(X As Single, Y As Single)
Static Stored As Boolean
Static Old_X As Single
Static Old_Y As Single
Static Old_La As Double
Static Old_Lo As Double


Dim LA As Double
Dim LO As Double


Dim DistanceCalculated As Double
Dim results As Variant


Const COLOR_SECOND_CONTROLPOINT = vbBlue
Const COLOR_FIRST_CONTROLPOINT = vbRed


If Stored Then
    'Mark SECOND Control Point
    Picture1.Circle (X, Y), 3, COLOR_SECOND_CONTROLPOINT
    Picture1.Line (Old_X, Old_Y)-(X, Y), vbBlack
    
    With ZoomRegion
        'convert coordinates
        Old_La = Origin_Lat - (Old_Y / .ScaleY + .OriginY) * Scale_Y
        Old_Lo = Origin_Long - (Old_X / .ScaleX + .OriginX) * Scale_X
        LA = Origin_Lat - (Y / .ScaleY + .OriginY) * Scale_Y
        LO = Origin_Long - (X / .ScaleX + .OriginX) * Scale_X
    End With
        
    'get distance
    DistanceCalculated = TrackCalculator.Calculate_Distance(Old_Lo, LO, Old_La, LA)
    results = MsgBox("Distance = " & Str$(DistanceCalculated) & " km", vbOKOnly, "Distance Tool")
    'if run again, make sure you query 1st control point again!
    Stored = False
    RePaint
Else
    'Mark FIRST Control Point
    Picture1.Circle (X, Y), 3, COLOR_FIRST_CONTROLPOINT
    Stored = True
    Old_X = X
    Old_Y = Y
End If


End Sub

Private Sub CalibrateTool(X As Single, Y As Single)
Static Stored As Boolean
Static Old_X As Single
Static Old_Y As Single
Static Old_La As Single
Static Old_Lo As Single


Dim LA As Single
Dim LO As Single


Dim Coordinates As String

Dim Max_Lat As Single
Dim Max_Long As Single
Dim Min_X As Single
Dim Min_Y As Single

Const COLOR_SECOND_CONTROLPOINT = vbBlue
Const COLOR_FIRST_CONTROLPOINT = vbRed
Const Longitude = 1
Const Latitude = 2


If Stored Then
    'Mark SECOND Control Point
    Picture1.Circle (X, Y), 3, COLOR_SECOND_CONTROLPOINT
    'Query for the Lat and Long
    Coordinates = InputBox("Enter Long(X-axis), Lat (Y-axis) in the format: Deg'Min, Deg'Min", "Control Point 2/2", "123'10,48'10")
    If Coordinates = "" Then Exit Sub
            
    'Get Lats and Longs
    LA = ExtractCoordinates(ByVal Coordinates, Latitude)
    LO = ExtractCoordinates(ByVal Coordinates, Longitude)
    
    'Set Control points
    SetControlPoints LA, LO, Old_La, Old_Lo, X, Y, Old_X, Old_Y
    'if run again, make sure you query 1st control point again!
    Stored = False
    'Clear the screen
    Picture1.Line (X, Y)-(Old_X, Old_Y), vbRed
    'refresh canvas and switch tool off
    frmFloater.SelectTool ToolBox.Select_Tool
    Picture1.Cls
    RePaint
Else
    'Mark FIRST Control Point
    Picture1.Circle (X, Y), 3, COLOR_FIRST_CONTROLPOINT
    Stored = True
    Old_X = X
    Old_Y = Y
    'Query for the Lat and Long
    Coordinates = InputBox("Enter Long(X-axis), Lat (Y-axis) in the format: Deg'Min, Deg'Min", "Control Point 1/2", "123'10,48'10")
    If Coordinates = "" Then Exit Sub
    Old_La = ExtractCoordinates(ByVal Coordinates, Latitude)
    Old_Lo = ExtractCoordinates(ByVal Coordinates, Longitude)
End If



End Sub
Private Sub SetControlPoints(LA As Single, LO As Single, Old_La As Single, Old_Lo As Single, X As Single, Y As Single, Old_X As Single, Old_Y As Single, Optional ChangeControlPoints As Boolean = True)
Dim Distance_Pixels_X As Single
Dim Distance_Pixels_Y As Single

Dim Distance_Lat As Single
Dim Distance_Long As Single


Dim Max_Lat As Single
Dim Max_Long As Single
Dim Min_X As Single
Dim Min_Y As Single

'Calculate Distance
Distance_Lat = Abs(Old_La - LA)
Distance_Long = Abs(Old_Lo - LO)

'Calibrate using this coordinates and previous coordinates
Distance_Pixels_X = Abs(Old_X - X)
Distance_Pixels_Y = Abs(Old_Y - Y)

'exit condition set when no distance b/w points
If Distance_Pixels_X = 0 And Distance_Pixels_Y = 0 Then Exit Sub
       
'Scale Ratio
Scale_X = Distance_Long / Distance_Pixels_X
Scale_Y = Distance_Lat / Distance_Pixels_Y

'Find Origin
'Note this map is oriented so that Log and Lat increase TOWARDS the origin (NORTH OF THE EQUATOR)

'Find max first
If Old_La > LA Then
    Max_Lat = Old_La
    Min_Y = Old_Y
Else
    Max_Lat = LA
    Min_Y = Y
End If

If Old_Lo > LO Then
    Max_Long = Old_Lo
    Min_X = Old_X
Else
    Max_Long = LO
    Min_X = X
End If

'get origins for lat and long
Origin_Lat = Max_Lat + (Min_Y * Scale_Y)
Origin_Long = Max_Long + (Min_X * Scale_X)

'New control points!
NEW_CONTROL_POINTS = ChangeControlPoints


'Redraw w/ new scale
Receiver.ReDraw
End Sub
Private Sub SelectTool(X As Single, Y As Single)
ClickedReceiver = IDStation(X, Y) - 1
If ClickedReceiver = -1 Then ClickedReceiver = 0
End Sub

Private Function IDStation(ByVal X As Single, ByVal Y As Single) As Long
Dim p As Long

p = Receiver.LookUpMapImage(X, Y)
If p > 1 Then
    With Receiver
        .CurrentStation_Number = CInt(p - 1)
        .DrawReceiver Picture1, p - 1
        'highlight
        .DrawReceiver Picture1, p - 1, 1, QBColor(1)
    End With
End If

IDStation = p

End Function

Private Function ExtractCoordinates(ByVal Coordinates As String, CoordinateType As Long) As Single
'Get degrees to decimal from string

Dim Minutes As Single
Dim Decimal_L As Single

Dim p As Long
Dim L As Long
Dim s As String

Dim ss_len As Long

Const Longitude = 1
Const Latitude = 2


L = Len(Coordinates)
p = InStr(Coordinates, ",")

'validate
If p >= 1 Then
    'First to the LEFT of the comma or SECOND (RIGHT OF COMMA)
    If CoordinateType = Longitude Then
        'LEFT
        ss_len = p - 1
        p = 0
    Else
        'RIGHT
        ss_len = L - p
    End If
    
    'Extract
    s = Mid$(Coordinates, p + 1, ss_len)

    'Now extract degrees and minutes
    'minutes are to the RIGHT of the ' character
    'degrees to the LEFT of the ' character
    'so this is similar to what we did above!
    
    L = Len(s)
    p = InStr(s, "'")
    
    If p >= 1 Then
        'LEFT to get the degrees
        ss_len = p - 1
        
        Decimal_L = Val(Mid$(s, 1, ss_len))
        
        'Minutes to the RIGHT
        ss_len = L - p
        Minutes = Val(Mid$(s, p + 1, ss_len))
        
        'Calculate
        Decimal_L = Decimal_L + (Minutes / 60)
    Else
        'assume decimal by default
        Decimal_L = CSng(s)
    End If
Else
    Exit Function
End If

'return value
ExtractCoordinates = Decimal_L

End Function

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim LA As Double
Dim LO As Double
Dim Coordinates As String
Dim p As Long
Dim results As Variant
Dim Color As Long
Dim R As Integer
Dim H As Single
Dim w As Single

Static Last_Receiver_Marked As Integer

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'MAP CALIBRATION
'Very Special Tool
'Creates two lines that intersect
'erases as it goes until user hits button or changes tool
If Tool_Number = ToolBox.Calibrate_Tool Then
    Picture1.Cls
    w = Picture1.Width
    H = Picture1.Height
    Form1.Picture1.Line (0, Y)-(w, Y), vbBlack
    Form1.Picture1.Line (X, 0)-(X, H), vbBlack
    Form1.Picture1.Circle (X, Y), 3, vbRed
    Exit Sub
End If

'''''''''''''''''''''''''''''''''''''''''''''
'From here on, if set to disable, quit
If Disable_MouseOver_Event Then Exit Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'DRAG AND DROP
'allow drag and drop of receiver
If Button = vbLeftButton And ClickedReceiver Then
    MovingReceiver = True
    ClearScreen
    If Receiver.LookUpMapImage(X, Y) = 1 Then
        'use arrow if valid
        Form1.Picture1.MousePointer = vbArrow
        'draw
        Receiver.DrawReceiver Form1.Picture1, ClickedReceiver, 1, vbWhite, CLng(X), CLng(Y)
    Else
        Form1.Picture1.MousePointer = vbNoDrop
    End If
    Exit Sub
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'SELECT TOOL and PLOT USER DEFINED TRACK
'this is the older part of the sub, which deals with marking receivers, detecting land, etc.
If X < 0 Or Y < 0 Then Exit Sub
'show coordinate in decimals and Degrees/Minutes
LA = Origin_Lat - (Y * Scale_Y)
LO = Origin_Long - (X * Scale_X)

Coordinates = Format$(LA, "##.000") & " / " & Format$(LO, "##.000")
StatusBar.Panels(StatusPanel.Coordinates) = Coordinates

p = Receiver.LookUpMapImage(X, Y)
If p <= 1 Then
    Select Case p
         Case 1
            StatusBar.Panels(StatusPanel.PixelDesc) = "Water"
         Case 0
            StatusBar.Panels(StatusPanel.PixelDesc) = "Land"
    End Select
    'unmark last marked one
    If Last_Receiver_Marked <> -1 Then
        If WhatToShowOnCanvas <> ShowOnCanvas.FishTrack Then Form1.RePaint
        Last_Receiver_Marked = -1
    End If
Else
    R = p - 1
    'show receiver in status bar
    If Not Receiver_Table.Invisible(R) Then
         StatusBar.Panels(StatusPanel.PixelDesc) = Receiver.ID(R)
        If (Tool_Number = ToolBox.Select_Tool Or Tool_Number = ToolBox.Plot_Tool) And WhatToShowOnCanvas <> ShowOnCanvas.FishTrack Then
            Color = Receiver.Color(R) Or HighLightColor 'make sure its marked a different color and tone
            Receiver.DrawReceiver Picture1, R, 2, Color
            Last_Receiver_Marked = R
        End If
    End If
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'MULTI-SELECT TOOL
'if using select tool and button is pressed, draw box around area
If Button = vbLeftButton And Tool_Number = ToolBox.Select_Tool And Shift Then
    MultiSelect = True
    End_X = X: End_Y = Y
    
    If Begining_X = -1 And Begining_Y = -1 Then
        Begining_X = X: Begining_Y = Y
    End If
    
    Picture1.Line (Begining_X, Begining_Y)-(X, Y), vbBlack, B
    ClearScreen
    RePaint
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'FISH CORRIDORS
If Tool_Number = ToolBox.DrawFishCorridor Then
    If AnchorMap(X, Y) > 0 Then
        Form1.Picture1.Circle (X, Y), 2, vbRed
    End If
End If
If ANCHORING Then
    DrawAnchors X, Y
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub
Private Sub DrawAnchors(X As Single, Y As Single, Optional Radius As Long = 1)
Dim i As Long
Dim X1 As Integer
Dim X2 As Integer
Dim Y1 As Integer
Dim Y2 As Integer
Dim ConnectedTo As Integer
Dim ColorOfLine As Long

'draw/erase anchor points as user draws new anchors

'if radius is set to 0, do not show connectors or receivers
If Radius = 0 Then
    ClearScreen
Else
    ShowDetectors
End If

'draw old anchors
'this is a connected graph so this code should be a bit different
For i = 1 To Anchors
    X1 = FishCorridor_AnchorPoints(i).X
    Y1 = FishCorridor_AnchorPoints(i).Y
    ConnectedTo = FishCorridor_AnchorPoints(i).ConnectTo
    If ConnectedTo > 0 Then
        X2 = FishCorridor_AnchorPoints(ConnectedTo).X
        Y2 = FishCorridor_AnchorPoints(ConnectedTo).Y
    Else
        X2 = CInt(X)
        Y2 = CInt(Y)
    End If
    'draw the line
    If ConnectedTo = 0 Then
        'if going over land, warn and draw with red
        With ImageProcessingEngine
            If Not .LandBetweenPositions(CLng(X), CLng(Y), CLng(X1), CLng(Y1)) Then
                ColorOfLine = FishCorridor_Color
                AnchorPointNotValid = False
            Else
                ColorOfLine = vbRed
                AnchorPointNotValid = True
            End If
        End With
        Form1.Picture1.Line (X1, Y1)-(X2, Y2), ColorOfLine
    Else
        ColorOfLine = FishCorridor_Color
        If Not (ConnectedTo = -1) Then Form1.Picture1.Line (X1, Y1)-(X2, Y2), ColorOfLine
    End If
Next i

End Sub


Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'if selection was in process
'multi-select stations/receivers enclosed by box

Dim i As Single
Dim ii As Single
Dim Station As Long
Dim s As Long
Dim R As Integer
Dim MinY As Long
Dim MaxY As Long
Dim MaxX As Long
Dim MinX As Long
Dim NumberOfReceivers As Long
Dim RX As Long
Dim RY As Long

If Button = vbLeftButton And MovingReceiver Then
    'validate drop location
    If Receiver.LookUpMapImage(X, Y) = 1 Then
        'delete original receiver from map (replaces it with water)
        Receiver.MoveReceiver_InMemory ClickedReceiver, X, Y
        ClickedReceiver = 0
        MovingReceiver = False
        ShowDetectors
    Else
        'no drop
        Form1.Picture1.MousePointer = vbArrow
        ClickedReceiver = 0
        MovingReceiver = False
        ShowDetectors
    End If
    Exit Sub
End If

If MultiSelect And Button = vbLeftButton And (Begining_X >= 0 And Begining_Y >= 0) And (Begining_X <> X And Begining_Y <> Y) Then
    'clear station list for multiselect
    NumberOfReceivers = Receiver.TotalReceivers
    For s = 1 To NumberOfReceivers
        Receiver_Selected(s) = False
    Next s
     
    If Y > Begining_Y Then
        MaxY = Y: MinY = Begining_Y
    Else
        MaxY = Begining_Y: MinY = Y
    End If
    If X > Begining_X Then
        MaxX = X: MinX = Begining_X
    Else
        MaxX = Begining_X: MinX = X
    End If
    
    For R = 1 To NumberOfReceivers
        RX = Receiver.X(R)
        RY = Receiver.Y(R)
        If RX >= MinX And RX <= MaxX And RY >= MinY And RY <= MaxY Then
            Station = IDStation(RX, RY)
            Receiver_Selected(R) = True
        End If
    Next R
    
    'store coordinates
    With ZoneCoordinates
        .X1 = Begining_X
        .Y1 = Begining_Y
        .X2 = X
        .Y2 = Y
    End With
    'clear for next
    Begining_X = -1: Begining_Y = -1
End If

    
End Sub
Public Sub RePaint()
Dim i As Long
Dim c As Long

'if markov chains are loaded, then control of window goes to MK
If MKWindowLoaded Then
    ClearScreen
    frmMarkov.RefreshChains
    Exit Sub
End If


'refreshes map
If Receivers_Are_Visible And Not REDRAWING Then
    'redrawing now
    REDRAWING = True
    'if coords for box are undefined, not in multiselect
    If Begining_X = -1 And Begining_Y = -1 Then
        Select Case WhatToShowOnCanvas
            Case ShowOnCanvas.FishDensity
                ShowFishDensity
            Case ShowOnCanvas.PingDensity
                ShowDetectors True
            Case ShowOnCanvas.MorningDensity
                ShowDensity_DielPhase Diel.AM
            Case ShowOnCanvas.NightDensity
                ShowDensity_DielPhase Diel.Night
            Case ShowOnCanvas.EveningDensity
                ShowDensity_DielPhase Diel.PM
            Case ShowOnCanvas.DawnDensity
                ShowDensity_DielPhase Diel.Dawn
            Case ShowOnCanvas.DuskDensity
                ShowDensity_DielPhase Diel.Dusk
            Case Else
                ShowDetectors
        End Select
    Else
        'if box is defined, its multiselect: show receivers only
        ShowDetectors False, True
    End If
    'done
    REDRAWING = False
Else
    If Track_Visible And Not REDRAWING Then
        'redrawing now
        REDRAWING = True
        ClearScreen
        frmFloater.RefreshCanvas
        'done
        REDRAWING = False
    End If
    
End If

If Tool_Number = ToolBox.Plot_Tool Then ShowReferenceTrack
        

End Sub

Private Sub ShowReferenceTrack()
'draw all points in reference tack
Dim i As Long
Dim CurrentSite As Long
Dim LastSite As Long

LastSite = -1
REDRAWING = True
Do While Receiver.UserDefinedTrack_Receiver(i) <> -1
    CurrentSite = Receiver.UserDefinedTrack_Receiver(i)
    If LastSite <> -1 Then
        'draw route
        Receiver.DrawRoute CurrentSite, LastSite, Picture1, vbRed
    End If
    LastSite = CurrentSite
    i = i + 1
Loop
REDRAWING = False
End Sub
Private Sub Zoom(X As Single, Y As Single)
'Zooms map to show receivers more closely
'Uses the auto-scale function to redraw receivers

Dim ULa As Single 'upper
Dim ULo As Single '
Dim LLa As Single 'lower
Dim LLo As Single '

Dim CornerX As Single
Dim CornerY As Single

Dim result As Variant

Dim temp As Single

Static OLA As Double
Static OLO As Double
Static SY As Single
Static SX As Single

Const ZOOM_DIMENSION_X = 50
Const ZOOM_DIMENSION_Y = 50

If ZoomRegion.Zoomed Then
    'back to normal!
    With ZoomRegion
        'notify rest of prg with flag raised
        .Zoomed = False
        .OriginX = 0
        .OriginY = 0
        .ScaleX = 1
        .ScaleY = 1
    End With
    Picture1.Cls
Else
   
    If X < ZOOM_DIMENSION_X Then X = 0
    If Y < ZOOM_DIMENSION_Y Then Y = 0
    
    '
    
    'screen corner
    CornerX = Picture1.ScaleWidth
    CornerY = Picture1.ScaleHeight
    
    'Pass to drawing routine the details of this zoom window
    With ZoomRegion
        'notify rest of prg with flag raised
        .Zoomed = True
        .OriginX = X - ZOOM_DIMENSION_X
        .OriginY = Y - ZOOM_DIMENSION_Y
        .SizeX = ZOOM_DIMENSION_X * 2
        .SizeY = ZOOM_DIMENSION_Y * 2
        .ScaleX = CornerX / .SizeX
        .ScaleY = CornerY / .SizeY
    End With
    
    'draw map
    DrawZoomMap X, Y, ZOOM_DIMENSION_X, ZOOM_DIMENSION_Y
    
    
End If

'redraw
'Receiver.ReDraw
'refresh
RePaint


End Sub
Private Sub DrawZoomMap(Optional X As Single = 0, Optional Y As Single = 0, Optional ZX As Integer = 0, Optional ZY As Integer = 0)
Dim ReadX As Integer
Dim ReadY As Integer

Dim SizeX As Integer
Dim SizeY As Integer

Dim i As Integer
Dim ii As Integer

Dim p As Long
Dim AnimateZoom As Long

Static SavedX As Integer
Static SavedY As Integer
Static SavedZX As Integer
Static SavedZY As Integer
Dim ZZX As Integer
Dim ZZY As Integer
Dim BeginStep As Long



If ZX = 0 Or ZY = 0 Then
    X = SavedX
    Y = SavedY
    ZX = SavedZX
    ZY = SavedZY
Else
    SavedX = X
    SavedY = Y
    SavedZX = ZX
    SavedZY = ZY
End If


With ZoomRegion
    If Int(.ScaleX) <> .ScaleX Then SizeX = Int(.ScaleX) + 1 Else SizeX = Int(.ScaleX)
    If Int(.ScaleY) <> .ScaleY Then SizeY = Int(.ScaleY) + 1 Else SizeY = Int(.ScaleY)
End With

For ReadX = X - ZX To X + ZX
    For ReadY = Y - ZY To Y + ZY
        p = Picture2.Point(ReadX, ReadY)
        Picture1.Line (i, ii)-(i + SizeX, ii + SizeY), p, BF
        ii = ii + SizeY
    Next ReadY
    i = i + SizeX
    ii = 0
Next ReadX


End Sub
Public Sub ClearScreen()
'clears screen

'normal clear if not zoomed
If ZoomRegion.Zoomed Then
    DrawZoomMap
Else
    Picture1.Cls
End If

'Assume track is visible
'make sure this is set here too
frmFloater.WarningTrackNotVisible False

End Sub

Public Sub AskIfUserWantsToDeletePreviousCorridor()
'ask if user needs to delete previous
Dim response As Variant
Dim X As Long
Dim Y As Long
Dim a As Long

'if anchors are previously defined, ask if need to delete previous corridor
If Anchors Then
    response = MsgBox("Delete previous corridor? If you answer NO, you can expand on previous corridor or add to it", vbYesNo, "Fish Corridor")
    If response = vbYes Then
        'delete all anchor points
        Anchors = 0
        For X = 0 To MAX_X
            For Y = 0 To MAX_Y
                AnchorMap(X, Y) = 0
            Next Y
        Next X
        
        'delete connected graph
        For a = 0 To MAX_ANCHORS
            FishCorridor_AnchorPoints(a).ConnectTo = -1
        Next a
    Else
        'DrawFishCorridor X, Y
    End If
End If

End Sub

