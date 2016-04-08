Attribute VB_Name = "mdlMain"
Option Explicit
'vid capture

Global Creator As New cAVICreator
Global ImageToMovie As New cBmp
Global VH As New cVideoHandler
Global VHS As New cVideoHandlers
Public Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

'canvas
Public Type CanvasRegion
    Zoomed As Boolean
    OriginX As Single
    OriginY As Single
    ScaleX As Single
    ScaleY As Single
    SizeX As Integer
    SizeY As Integer
End Type
Global ZoomRegion As CanvasRegion

'''''''
Global IGNORE_ALL As Boolean
Global Receivers_Are_Visible As Boolean
Global Track_Visible As Boolean

Global DateListIsLoaded As Boolean
Global MKWindowLoaded As Boolean
Global TipWindowLoaded As Boolean
Global DielCycleFormIsLoaded As Boolean
Global JPlotIsLoaded As Boolean
Global ResidenceWindowIsLoaded As Boolean
Global TrackStringWindowIsLoaded As Boolean

Global MapWasScanned As Boolean

Global Scale_X As Single
Global Scale_Y As Single
Global Origin_Lat As Double
Global Origin_Long As Double
Global Topics(20) As String
Global CAPTURE As Boolean

'File IO globals for error trapping and notification
Global row As Long
Global column As Long
Global CURRENT_FISH As Long

'Show on canvas
Global WhatToShowOnCanvas As Long
Enum ShowOnCanvas
    FishTrack
    Receivers
    Residence
    StampList
    ResidenceAndMoves
    PingDensity
    FishDensity
    NightDensity
    MorningDensity
    EveningDensity
    DawnDensity
    DuskDensity
End Enum

'Color of water
Global WaterColor As Long

'Objects
Global Receiver As New clsReceiver
Global DeviceBuffer As New clsGenericIO
Global TrackCalculator As New Calculator
Global FishDatabase As New clsFishEntry
Global FishTable As New clsTable
Global ReceiverTable As New clsTable
Global GroupTable As New clsTable
Global Receiver_Table As New clsReceiverEntry
Global ImageProcessingEngine As New clsImageProcessor
Global BufferDib As New cDIBSection
Global MapDib As New cDIBSection
Global Excursion As New clsExcursionAnalyzer
Global PhotoPeriodCalculator As New clsPhotoPeriod
Global AstroMech As New clsAstronomicalCalculator
Global ColorScale As New clsScale

Global FactorMove As New clsGenericIO
Global FactorStay As New clsGenericIO
'Global Classifier As New clsClassifier
Global RND_MT As New clsMersenneTwister64


Global MapFileName As String
Global Tool_Number As Integer
Global Factor As Single
Global MultiSelect As Boolean
Global AM_TH As Integer
Global PM_TH As Integer
Global ResidenceThreshold As Long

'Registry
Global Const APPLICATION = "AquaTracker"
Global Const REGISTRY_SECTION = "Map"

Global Const Version = "2.41"

'Max entries program can handle (2^23= 8M)
Global Const MAX_ENTRIES = 2 ^ 23 'Most important limit!

'Grouping of receivers limits
Global Const MAX_RECEIVERS_PERGROUP = 75
Global Const MAX_GROUPS = 50

'Misc limits
Global Const MAX_TRACK_STRING = 30000
Global Const SUGGESTED_MAX_EXPORT_TABLE = 10000
'Databases
Global Const MAX_FISH = 500
Global Const MAX_RECEIVERS = 255
Global Const MAX_DETECTIONS_PER_RECEIVER = 2 ^ 21 'about 2 million detections per receiver
Global Const MAX_DETECTIONS_PER_FISH = 2 ^ 21     'and fish

'Image boundaries
Global Const MAX_X = 1024
Global Const MAX_Y = 1024

'Const related to navigation and land-awareness
Global SEGMENT_SIZE_THRESHOLD As Long  '= 50
Global Const MAX_NODES = 200
Global Const CONTOUR_SIZE = 4000
Global Nav_Segment As Long 'this is now under user control
Global Persistance_TH As Single
Global SEARCHRADIUS '= 600#
Global ABORT_PROCESS As Boolean
Global FishCorridor_Color As Long
Global Const DEFAULT_FISHCORRIDOR_COLOR = vbBlue
'Const related to the auto-grouping of Receivers
Global Const AUTOMATIC = -1
Global Const THRESHOLD_TIMEBIN = 2

'Const related to Grouping of fish
Global Const MAX_SUBGROUPS = 500

'receiver tags and receiver zones
Global Const MAX_TAGS = 50
Global Const MAX_ZONES = 50

Global Land_Avoidance As Long

'Astromech related constants
Global Const ALWAYSup = "ALL_DAY"
Global Const ALWAYSdown = "ALL_NIGHT"
Global Const NoEVENT = "N/A"
Global Const TimeBinsForAnimation = 30

'Constants related to marker color, size, and type
Global Const LARGE_MARKER = 2
Global Const MaxPal = 11
Global ColorPal(MaxPal) As Long
Global HighLightColor As Long

Global MapImage(MAX_X, MAX_Y) As Long
Global MapEdge(MAX_X, MAX_Y) As Long
Global FishCorridorMap(MAX_X, MAX_Y) As Long
Global GeographicZones(MAX_X, MAX_Y) As Long
Global TimeCompressionFactor As Long
Global TotalTime As Long
Global Min As TrackParameters
Global Max As TrackParameters
Global Average_For_All_Tracks As TrackParameters

Global FileNameWithoutPath As String

Global Receiver_Selected(MAX_RECEIVERS) As Boolean
Global Lap_MAX As Long
'Globals related to program functioning and user modes

Global REDRAWING As Boolean
Global EXCLUDE_SINGLETONS As Boolean

Global Const EPOCH = 726896#
Global SizeReceiverData(MAX_RECEIVERS) As Long
Global SizeFishData(MAX_FISH) As Long


Public Type ExternalStamp
    Fish As Integer
    Site As Integer
    Date As Long
    Time As Integer
    CTime As Long
    Valid As Boolean
End Type

'program settings
Public Type ProgramSettingsType
    Scale_X As String
    Scale_Y As String
    Origin_Long As String
    Origin_Lat As String
    MapFile As String
    MoviePath As String
    SplashTime As String
    ColorOfWater As Long
    SizeX As String
    SizeY As String
    LastFileLoaded As String
End Type

Global ProgramSettings As ProgramSettingsType
Global ChooseWaterColor As Boolean
Global Stamp As ExternalStamp

Public Type TrackParameters
    LA As Single
    LO As Single
    Linearity As Single
    Meandering As Single
    Distance As Single
End Type

Enum Device_Type
    Printer = -1
    File = 1
    Window = 99
End Enum

Enum StatusPanel
    Map = 1
    Coordinates = 2
    PixelDesc = 3
    Alert = 4
End Enum

Enum AvoidanceMode
    NotActive = 0
    Shoreline = 1
    FishCorridor = 2
    RandomWalk = 3
End Enum

Enum ToolBox
    Select_Tool = 0
    Zoom = 1
    Calibrate_Tool = 2
    Measure_Tool = 3
    Plot_Tool = 4
    DrawFishCorridor = 5
    RecordFrames = 6
End Enum

Enum FieldNames
    Fish = 1
    Site = 2
    Date = 3
    Time = 4
    CTime = 5
    All = 6
    MJD = 7
End Enum

Enum AstralFunction
    Sunrise = 1
    Begin_Civil_Twilight = 2
    Begin_Nautical_Twilight = 3
    Begin_Astro_Twilight = 4
    Sunset = -1
    End_Civil_Twilight = -2
    End_Nautical_Twilight = -3
    End_Astro_Twilight = -4
End Enum

'this constant is related to the number of diel phases, two pair of twilights and one midday.  Base 0.
Global Const TotalDielPhases = 4

Enum DielPhase
    Dawn = 0
    Sunrise = 1
    Midpoint = 2
    Sunset = 3
    Dusk = 4
End Enum


Enum MarkerType
    Circular
    Triangle
    Rectangle
    InvertedTriangle
    Diamond
End Enum


Dim Days_Per_Month(12) As Long

Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
  ByVal hDC As Long) As Long
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, _
  ByVal nIndex As Long) As Long
Public Sub LoadDaysOfTheMonthForCalculations()
Days_Per_Month(0) = 0
Days_Per_Month(1) = 31
Days_Per_Month(2) = 59
Days_Per_Month(3) = 90
Days_Per_Month(4) = 120
Days_Per_Month(5) = 151
Days_Per_Month(6) = 181
Days_Per_Month(7) = 212
Days_Per_Month(8) = 243
Days_Per_Month(9) = 273
Days_Per_Month(10) = 304
Days_Per_Month(11) = 334
Days_Per_Month(12) = 365
End Sub
Public Function ConvertTime(t As String) As Long
'Converts a string time variable to -number of minutes past-midnight-
'assumes military (24hr) time

Dim number_of_colons As Long
Dim Hours As Long
Dim Minutes As Long
Dim Total As Long

Dim position_of_separator As Long

Dim L As Long
Dim note As String

On err GoTo ExitSub

'validate hours
position_of_separator = InStr(1, t, ":")
If position_of_separator > 1 Then
    Hours = Left(t, position_of_separator - 1)
    If Val(Hours) > 23 Then
        Hours = 23
        Exit Function
    End If
End If

'validate minutes
If position_of_separator > 1 And Len(t) > 3 Then
    Minutes = Mid(t, position_of_separator + 1, 2)
    If Val(Minutes) > 59 Then
        Minutes = 59
    End If
End If

If Len(t) <= 3 Or position_of_separator <= 1 Then
    Hours = 23
    Minutes = 59
End If


'Split incoming string using SOP if not outside range
'If Hours < 23 Then Hours = hour(T)
'If Minutes < 59 Then Minutes = minute(T)

'is there a PM/AM notation at the end?
L = Len(t)

If L > 2 Then
    note = Right(t, 2)
    If UCase(note) = "PM" Then Hours = Hours + 12
End If

Total = (Hours * 60) + Minutes

'can't exceed the number of minutes in a day!!
If Total > 1439 Then Total = 1440


ExitSub:
ConvertTime = Total

End Function
Public Function DayNumber(s As String) As Long
'Converts a date in format MM/DD/YYYY to modified Julian day number
ReDim a(3) As String
Dim d As Long
Dim m As Long
Dim Yr As Long
Dim result As Variant

Dim i As Long
Dim DaysSinceStartOfMJD As Double
Dim NumberOfLeapYears As Double
Dim LeapDay As Long
'
'   takes the year, month and day as a Gregorian calendar date
'   and returns the modified julian day number
'
On Error GoTo RaiseError

m = month(s)
d = Day(s)
Yr = Year(s)

'Find if leap year?
If (Yr Mod 4 = 0 And Yr Mod 100 <> 0) Or Yr Mod 400 = 0 Then
    'add extra day
    If m > 2 Then
        LeapDay = 1
    End If
End If

d = d + Days_Per_Month(m - 1) + LeapDay

'calculate number of leapyears since 0 year and add to total (#leap years = number of days that need to be added for correction)
NumberOfLeapYears = Fix(Yr / 400#) - Fix(Yr / 100#) + Fix(Yr / 4#)
DaysSinceStartOfMJD = 365# * Yr - EPOCH
DayNumber = DaysSinceStartOfMJD + NumberOfLeapYears + d

Exit Function

RaiseError:
'if set to ignore all, then change date
If IGNORE_ALL Then
    DayNumber = 0
Else
    'if no flag set, ask user what to do
    result = MsgBox("Error at " & row & ": Incorrect day format.", vbAbortRetryIgnore, "Import Error")
    If result = vbRetry Then DayNumber = DayNumber(s)
    If result = vbAbort Then Unload Form1: Stop
    'if ignore, ask if ignore all
    If result = vbIgnore Then
        DayNumber = 0
        result = MsgBox("Ignore all?", vbYesNo, "Import Error")
        If result = vbYes Then
            'inform user of consequences
            result = MsgBox("Warning: Any invalid points will be flagged with the Epoch date (1991).", vbOKOnly, "Import Error")
            'set flag
            IGNORE_ALL = True
        End If
    End If
End If
End Function
Public Function Convert_DayNumber(ByVal N As Long) As String
'Converts from an MJD date to a standard DD.
Dim Year As Single
Dim DayNumber As Integer
Dim i As Long
Dim DayNumberString As String
Dim NumberOfLeapYearDays As Long
Dim DN As Long
Dim LeapDay As Long
Dim correction_for_leap_year As Long

Const AVG_DAYS_YR = 365.25

'correct for epoch
N = N + EPOCH
'get year by first calculating leap year days
NumberOfLeapYearDays = Fix(N / (400# * AVG_DAYS_YR)) - Fix(N / (100# * AVG_DAYS_YR)) + Fix(N / (4# * AVG_DAYS_YR))
'get residual days
N = N - NumberOfLeapYearDays

'year
Year = Fix(N / 365#)
'residual is number of days in year
DN = N - (Year * 365)

'store in astromech (?No idea why this is here?)
AstroMech.SetYear = Year

'Find if leap year?
If (Year Mod 4 = 0 And Year Mod 100 <> 0) Or Year Mod 400 = 0 Then
    LeapDay = 1
End If

'find month
i = 0
Do
    i = i + 1
    If i >= 2 Then correction_for_leap_year = LeapDay
Loop Until DN <= Days_Per_Month(i) + correction_for_leap_year Or i = 12

'establish day of month
If i <= 2 Then correction_for_leap_year = 0
DN = DN - (Days_Per_Month(i - 1) + correction_for_leap_year)

'convert to a string with a leading zero
DayNumberString = Trim$(Str$(DN))
If Len(DayNumberString) = 1 Then
    DayNumberString = "0" & DayNumberString
End If

'Now assemble all the pieces together
If Year = 0 Then
    Convert_DayNumber = Str$(i) & "/" & DayNumberString
Else
    Convert_DayNumber = Str$(i) & "/" & DayNumberString & "/" & Str$(Year)
End If
End Function
Public Function Convert_ToStandardTime(t As Integer) As String
Dim hour As Long
Dim minute As Long
Dim AM_PM As String
Dim H As String
Dim m As String



hour = Int(t / 60)
minute = t Mod 60

If hour >= 12 Then
    If hour > 12 Then hour = hour - 12
    AM_PM = "PM"
Else
    AM_PM = "AM"
End If

'Follow very strict rules
'a 0 to preceed any single digit numbers in the hour section
H = Trim(Str$(hour))
If Len(H) = 1 Then H = "0" & H

'a 0 to preceed any single digit number in the minute section
m = Trim(Str$(minute))
If Len(m) = 1 Then m = "0" & m

Convert_ToStandardTime = H & ":" & m & " " & AM_PM

End Function
Public Function ConsolidateTime(d As Long, t As Integer) As Long
'Consolidates the TIME and DATE array so that entries can be sorted by date AND time
'It those this by calculating the estimate of the year as minutes since year 0
'year 0 should be set to something less than 50 yrs or so from NOW
'Set to 1990
Dim response As Variant

On Error GoTo SkipError

Const MINUTES_PER_DAY = 1440#

ConsolidateTime = (d * MINUTES_PER_DAY + t)

Exit Function

SkipError:
response = MsgBox("Internal Error: Bad Data Stamp/Unable to sort.", vbAbortRetryIgnore, "Internal Error")
If response = vbAbort Then Unload Form1: Stop
If response = vbRetry Then ConsolidateTime d, t
End Function
Public Function IsWithin(ByVal Time_Begins As Long, ByVal Time_Ends As Long, ByVal Time As Long) As Boolean
Dim results As Boolean

'Checks if time TIME is within boundaries of Begin and End or if the boundaries are within the THRESHOLD
'

'Check boundaries first

If Abs(Time_Begins - Time) <= THRESHOLD_TIMEBIN Then results = True
If Abs(Time_Ends - Time) <= THRESHOLD_TIMEBIN Then results = True

'now check if time is within boundaries
If Time >= Time_Begins And Time <= Time_Ends Then results = True


IsWithin = results

End Function
Public Sub ShowTrack(ByVal FishNumber As Long, DrawingBox As PictureBox)

'Shows fish track using DrawRoute, DrawReceiver, etc.
'Options: Show Move Only, Show Diel Window, etc.

Dim X As Integer
Dim Y As Integer
Static Counter As Double
Dim ReceiverDrawn(MAX_RECEIVERS) As Boolean
Dim TrackColor As Long

Dim i As Long

Dim OldX As Single
Dim OldY As Single
Dim OldSite As Integer
Dim Frame As Long
Dim RouteDrawn(MAX_RECEIVERS, MAX_RECEIVERS) As Boolean

Dim Concatenated As String

Dim Site_ID As String

Dim Latitude As Single
Dim Longitude As Single

Dim Site As Integer
Dim Line_Color As Long
Dim TE As Long

Dim DateDetected As Long
Dim t As Integer
Dim ValidDate As Boolean


Static LastFishNumber As Long

If FishNumber = -1 Then FishNumber = LastFishNumber

'Clear Screen
Track_Visible = False
Receivers_Are_Visible = False
Receiver.MakeInvisible


'Get index to code number
If FishNumber < 0 Then Exit Sub

'Sort date and time
'Sort FishNumber

'clear
TrackCalculator.Clear

FishDatabase.Fish = FishNumber

If Not FishDatabase.IsVisible Then Exit Sub

TrackColor = FishDatabase.Color

OldSite = -1
For i = 0 To FishDatabase.NumberOfStamps - 1
    FishTable.ReadStamp FishNumber, i
    Site = Stamp.Site
   
    ValidDate = Stamp.Valid
    'check date list
    If DateListIsLoaded Then
        'if it exist and is unchecked, this function returns a False
        ValidDate = ValidDate And frmDates.AddDate(Stamp.Date)
    End If
    
    
    'only calculate if included
    If ValidDate Then
        With TrackCalculator
            .Site = Stamp.Site
            .Day = Stamp.Date
            .Time = Stamp.Time
            .Calculate
        End With
        If OldSite <> -1 And OldSite <> Stamp.Site Then
            If frmFloater.chkMove Then
                If Not RouteDrawn(OldSite, Site) Then
                    RouteDrawn(OldSite, Site) = True
                    Receiver.DrawRoute Site, OldSite, DrawingBox, TrackColor
                End If
            End If
            
            If frmFloater.chkStay Then
                If Not ReceiverDrawn(Site) Then
                    Receiver.MakeVisible Site
                    ReceiverDrawn(Site) = True
                    If frmFloater.chkStay Then Receiver.DrawReceiver DrawingBox, Site
                End If
            End If
        End If
        
        If OldSite = -1 Then
            Receiver.MakeVisible Site
            ReceiverDrawn(Site) = True
            If frmFloater.chkStay Then Receiver.DrawReceiver DrawingBox, Site
        End If
        
        OldSite = Site
    End If
Next i

'save fish number
LastFishNumber = FishNumber

'Track is visible now
Track_Visible = True

'display a summary for track
'if "ALL" is selected, then display average
With frmFloater
    If .cmbFishCode <> "ALL" Then
        .txtLinearity = Format(TrackCalculator.Linearity, "##.###")
        .txtAvgLocation = Format(TrackCalculator.Path_Similarity_Index_La, "##.##") & "," & Format(TrackCalculator.Path_Similarity_Index_Lo, "##.##")
        .txtTTLDistance = Format(TrackCalculator.Total_Displacement, "##.###")
        .txtMeandering = Format(TrackCalculator.Meandering_Index, "##.###")
    End If
End With

'display histogram for fish
If DielCycleFormIsLoaded Then
    frmDayLightCycle.DrawHistogramForFish FishNumber
    ImageProcessingEngine.DrawDielCycleNOW
End If
End Sub
'--------------------------------------------------
Public Function TwipsPerPixelX() As Single
Const HWND_DESKTOP As Long = 0
Const LOGPIXELSX As Long = 88
Const LOGPIXELSY As Long = 90
'--------------------------------------------------
'Returns the width of a pixel, in twips.
'--------------------------------------------------
  Dim lngDC As Long
  lngDC = GetDC(HWND_DESKTOP)
  TwipsPerPixelX = 1440& / GetDeviceCaps(lngDC, LOGPIXELSX)
  ReleaseDC HWND_DESKTOP, lngDC
End Function

'--------------------------------------------------
Public Function TwipsPerPixelY() As Single
Const HWND_DESKTOP As Long = 0
Const LOGPIXELSX As Long = 88
Const LOGPIXELSY As Long = 90
'--------------------------------------------------
'Returns the height of a pixel, in twips.
'--------------------------------------------------
  Dim lngDC As Long
  lngDC = GetDC(HWND_DESKTOP)
  TwipsPerPixelY = 1440& / GetDeviceCaps(lngDC, LOGPIXELSY)
  ReleaseDC HWND_DESKTOP, lngDC
End Function
Public Sub WriteTrackInformation(FishNumber As Long)
'Writes track stats to a window or file for export
Dim i As Long
Dim ii As Long

ReDim K(4) As Variant

'key and names for phases
K = Array("N", "DW", "AM", "PM", "DK")

'Write all the columns
DeviceBuffer.WriteField("fid") = FishDatabase.Code(FishNumber)
DeviceBuffer.WriteField("Distance") = Format(TrackCalculator.Total_Displacement, "##.###")
DeviceBuffer.WriteField("Time") = Str$(TrackCalculator.Total_Time_Traveled)
DeviceBuffer.WriteField("Speed") = Format(TrackCalculator.Displacement_Rate, "##.###")
DeviceBuffer.WriteField("Range") = Format(TrackCalculator.Range_Bounding_Box, "##.###")
DeviceBuffer.WriteField("T") = Format(TrackCalculator.Linearity, "##.###")
DeviceBuffer.WriteField("RI") = Format(TrackCalculator.Meandering_Index, "##.###")
DeviceBuffer.WriteField("SS") = TrackCalculator.Longest_Stay_Station
DeviceBuffer.WriteField("ST") = Str$(TrackCalculator.Longest_Stay_Duration)
DeviceBuffer.WriteField("LA") = Format(TrackCalculator.Path_Similarity_Index_La, "##.###")
DeviceBuffer.WriteField("LO") = Format(TrackCalculator.Path_Similarity_Index_Lo, "##.###")
DeviceBuffer.WriteField("RA") = Format(TrackCalculator.Percent_Receivers_Active, "##.##%")
DeviceBuffer.WriteField("RS") = FishDatabase.Release_Site

'if any tags exist, prepare to calculate them
If Receiver.NumberOfTags > 0 Then
    For i = 1 To Receiver.NumberOfTags
        DeviceBuffer.WriteField("TAG" & Str$(i)) = Format(TrackCalculator.PercentVisitsToTag(i), "##.##%") & "%"
    Next i
End If

If TrackCalculator.AnalyzeDielMoves Then
    FactorStay.WriteField("fid") = FishDatabase.Code(FishNumber)
    FactorMove.WriteField("fid") = FishDatabase.Code(FishNumber)
    For i = 0 To 4
        FactorStay.WriteField(K(i)) = Format(TrackCalculator.PercentResidenceByCategory(i), "##.##") & "%"
        For ii = 0 To 4
            FactorMove.WriteField(K(i) & "->" & K(ii)) = Format(TrackCalculator.PercentTravelByCategory(i, ii), "##.##") & "%"
        Next ii
    Next i
    FactorStay.WriteLine
    FactorMove.WriteLine
End If

DeviceBuffer.WriteLine
End Sub
Public Sub LoadLastDataFile()
Dim response As Variant

On Error GoTo Exception

Form1.Show

If ProgramSettings.LastFileLoaded <> "" Then
    response = MsgBox("Load last data file analyzed: " & ProgramSettings.LastFileLoaded & " ?", vbYesNo)
    If response = vbYes Then
        Form1.LoadAQN ProgramSettings.LastFileLoaded
    Else
        ProgramSettings.LastFileLoaded = ""
    End If
End If
 

Exit Sub

Exception:

'throw an exception
response = MsgBox("Invalid path to data file.  Select another data file.", vbCritical)

End Sub
