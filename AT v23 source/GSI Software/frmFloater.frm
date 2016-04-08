VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{50A7E9B0-70EF-11D1-B75A-00A0C90564FE}#1.0#0"; "shell32.dll"
Begin VB.Form frmFloater 
   Caption         =   "Actions"
   ClientHeight    =   5232
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   6060
   Icon            =   "frmFloater.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5232
   ScaleWidth      =   6060
   Begin VB.Timer tmrBackgroundAnimation 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   2760
      Top             =   1800
   End
   Begin VB.Timer tmrBWDays 
      Interval        =   500
      Left            =   3960
      Top             =   4680
   End
   Begin VB.Frame Frame3 
      Caption         =   "Animate/show"
      Height          =   975
      Left            =   4440
      TabIndex        =   30
      Top             =   1080
      Width           =   1335
      Begin VB.CheckBox chkStay 
         Caption         =   "Stays"
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   480
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkMove 
         Caption         =   "Moves"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Show:"
      Height          =   1935
      Left            =   3600
      TabIndex        =   25
      Top             =   2280
      Width           =   2175
      Begin VB.OptionButton optShow 
         Caption         =   "Stamp list (verbose)"
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   29
         Top             =   1440
         Width           =   1935
      End
      Begin VB.OptionButton optShow 
         Caption         =   "Receiver residence"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   28
         Top             =   1080
         Width           =   1815
      End
      Begin VB.OptionButton optShow 
         Caption         =   "Receivers visited"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton optShow 
         Caption         =   "Fish track"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.PictureBox picChangeTrackColor 
      AutoSize        =   -1  'True
      Height          =   432
      Left            =   5160
      Picture         =   "frmFloater.frx":0442
      ScaleHeight     =   384
      ScaleWidth      =   384
      TabIndex        =   24
      ToolTipText     =   "Change track color"
      Top             =   240
      Width           =   432
   End
   Begin VB.PictureBox picTool 
      AutoSize        =   -1  'True
      Height          =   432
      Index           =   5
      Left            =   3600
      Picture         =   "frmFloater.frx":110C
      ScaleHeight     =   384
      ScaleWidth      =   384
      TabIndex        =   23
      ToolTipText     =   "Fish corridors"
      Top             =   240
      Width           =   432
   End
   Begin VB.PictureBox picTool 
      AutoSize        =   -1  'True
      Height          =   432
      Index           =   6
      Left            =   4320
      Picture         =   "frmFloater.frx":19D6
      ScaleHeight     =   384
      ScaleWidth      =   384
      TabIndex        =   22
      ToolTipText     =   "Record canvas"
      Top             =   240
      Width           =   432
   End
   Begin VB.Timer tmrFrameCapture 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1320
      Top             =   4680
   End
   Begin VB.PictureBox picStop 
      Height          =   495
      Left            =   2880
      Picture         =   "frmFloater.frx":1CE0
      ScaleHeight     =   444
      ScaleWidth      =   444
      TabIndex        =   11
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ListBox lstAllDates 
      Height          =   240
      Left            =   480
      Sorted          =   -1  'True
      TabIndex        =   10
      Top             =   4320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer tmrCompressedTime 
      Enabled         =   0   'False
      Interval        =   120
      Left            =   1800
      Top             =   4680
   End
   Begin VB.PictureBox picPlay 
      Height          =   495
      Left            =   2880
      Picture         =   "frmFloater.frx":2122
      ScaleHeight     =   444
      ScaleWidth      =   444
      TabIndex        =   0
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ListBox lstDates_Std 
      Height          =   3504
      Left            =   600
      TabIndex        =   4
      Top             =   5280
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.CommandButton cmdWalk 
      Caption         =   "Animate"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      Picture         =   "frmFloater.frx":2564
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.ComboBox cmbFishCode 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Timer tmrAnimation 
      Enabled         =   0   'False
      Interval        =   120
      Left            =   840
      Top             =   4680
   End
   Begin VB.ListBox lstDates 
      Height          =   3288
      Left            =   600
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   5280
      Visible         =   0   'False
      Width           =   4935
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   360
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frmTools 
      Caption         =   "Tools"
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   5655
      Begin VB.PictureBox picTool 
         AutoSize        =   -1  'True
         Height          =   432
         Index           =   1
         Left            =   720
         Picture         =   "frmFloater.frx":2CA6
         ScaleHeight     =   384
         ScaleWidth      =   384
         TabIndex        =   33
         ToolTipText     =   "Zoom tool"
         Top             =   240
         Width           =   432
      End
      Begin VB.PictureBox picTool 
         AutoSize        =   -1  'True
         Height          =   432
         Index           =   4
         Left            =   2880
         Picture         =   "frmFloater.frx":2FB0
         ScaleHeight     =   384
         ScaleWidth      =   384
         TabIndex        =   12
         ToolTipText     =   "Draw reference track"
         Top             =   240
         Width           =   432
      End
      Begin VB.PictureBox picTool 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         Height          =   432
         Index           =   0
         Left            =   120
         Picture         =   "frmFloater.frx":3BF2
         ScaleHeight     =   384
         ScaleWidth      =   384
         TabIndex        =   9
         Tag             =   "Active"
         ToolTipText     =   "Select Receiver(s)"
         Top             =   240
         Width           =   432
      End
      Begin VB.PictureBox picTool 
         AutoSize        =   -1  'True
         Height          =   432
         Index           =   2
         Left            =   1560
         Picture         =   "frmFloater.frx":4034
         ScaleHeight     =   384
         ScaleWidth      =   384
         TabIndex        =   8
         ToolTipText     =   "Georeference map"
         Top             =   240
         Width           =   432
      End
      Begin VB.PictureBox picTool 
         AutoSize        =   -1  'True
         Height          =   432
         Index           =   3
         Left            =   2160
         Picture         =   "frmFloater.frx":4CFE
         ScaleHeight     =   384
         ScaleWidth      =   384
         TabIndex        =   7
         ToolTipText     =   "Measure distance between two points in map"
         Top             =   240
         Width           =   432
      End
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "STOP"
      Height          =   375
      Left            =   2760
      TabIndex        =   34
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Summary"
      Height          =   1935
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   2895
      Begin VB.TextBox txtTTLDistance 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtAvgLocation 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtMeandering 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtLinearity 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "TTL Distance--------->"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Avg La/Lo-------------->"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Meandering------------>"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Linearity------------------>"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1935
      End
   End
   Begin Shell32Ctl.ShellFolderViewOC Shell32 
      Left            =   480
      OleObjectBlob   =   "frmFloater.frx":55C8
      Top             =   5280
   End
   Begin VB.Label Label1 
      Caption         =   "Fish :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   855
   End
   Begin VB.Menu mnuAnimation 
      Caption         =   "Animations"
      Visible         =   0   'False
      Begin VB.Menu mnuAccumulateDates 
         Caption         =   "Accumulate detections"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuDoNotAccumulate 
         Caption         =   "Show each date"
      End
   End
End
Attribute VB_Name = "frmFloater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StopNow As Boolean
Dim PlotIndepedently As Boolean
Dim Animate As Boolean
Dim WAITING As Boolean
Dim RealTime As Integer
Dim FLAG_TIME_POINT As Boolean
Dim TimeScale As Single
Dim TICK As Long
Dim STOP_ANIMATION As Boolean
Const Button_Highlight = &H80000014
Const Button_Face = &H8000000F
Const EXTENDED_HEIGHT = 9270
Const SHORT_HEIGHT = 5175
Dim ExportedFrame As New cDIBSection
Dim FirstFrameCaptured As Boolean
Dim TimerValue As Single


Private Enum Style
    ScaledTime = 0
End Enum

'This is for selecting folders

Private shlShell As Shell32Ctl.Shell
Private shlFolder2 As Shell32Ctl.Folder2
Private objFolderItem As Shell32Ctl.FolderItem
Const BIF_RETURNONLYFSDIRS = &H1

Private Sub chkMove_Click()
RefreshCanvas
End Sub

Private Sub chkStay_Click()
RefreshCanvas
End Sub


Private Sub cmbFishCode_Click()
Dim i As Long

'this here is needed to be consintent with what the user is seeing on the canvas
If optShow(ShowOnCanvas.FishTrack).Value = True Then WhatToShowOnCanvas = ShowOnCanvas.FishTrack

'show dates active
If DateListIsLoaded Then
    Unload frmDates
    Load frmDates
    RefreshCanvas
    frmDates.Show vbModeless, frmFloater
Else
    RefreshCanvas
End If

End Sub
Public Sub RefreshCanvas()
Dim i As Long
Dim result As Variant
Dim ShowInfo As String
Dim FishNumber As Long

If cmbFishCode.Text = "ALL" Or cmbFishCode.Text = "" Then
        WhatToShowOnCanvas = ShowOnCanvas.FishTrack
        optShow(ShowOnCanvas.FishTrack).Value = True
        'unload nav string (not needed )
        Unload frmNavigationString
        ImageProcessingEngine.ClearDielCycle

    'clear screen
    Form1.ClearScreen
    'show all tracks
    WarningTrackNotVisible False
    Form1.StatusBar.Panels(StatusPanel.Map) = "Showing all tracks"
    TrackCalculator.Reset
    For i = 0 To cmbFishCode.ListCount - 2
        CURRENT_FISH = i
        ShowTrack i, Form1.Picture1
        TrackCalculator.ComputeSummary
    Next i
    CURRENT_FISH = -1
    'and make all receiversvisible
    Receiver.MakeVisible
    TrackCalculator.Average
Else
    'get fish #
    FishNumber = cmbFishCode.ListIndex - 1
    If JPlotIsLoaded Then frmJPlot.HighlightFishNumber CInt(FishNumber)
    
    If WhatToShowOnCanvas = ShowOnCanvas.StampList Then
        frmVerboseStampList.Show vbModeless, frmFloater
        If FishNumber <> CURRENT_FISH Then frmVerboseStampList.LoadStamps CInt(FishNumber)
    End If
    CURRENT_FISH = FishNumber
    Form1.StatusBar.Panels(StatusPanel.Map) = "Showing track " & cmbFishCode.Text
    If FishDatabase.IsVisible(FishNumber) Then
        WarningTrackNotVisible False
        PopulateTrackWindow FishNumber
        'show track
        Form1.ClearScreen
        ImageProcessingEngine.ClearDielCycle
        ShowTrack FishNumber, Form1.Picture1
        If WhatToShowOnCanvas = ShowOnCanvas.Residence Then
            If Not ResidenceWindowIsLoaded Then frmReceiversResidenceTime.Show Else frmReceiversResidenceTime.LoadData
        End If
    Else
        Form1.ClearScreen
        WarningTrackNotVisible True
    End If
    
End If

End Sub
Private Sub PopulateTrackWindow(FishNumber As Long)
'show track string window
If Not TrackStringWindowIsLoaded Then
    frmNavigationString.Show vbModeless, frmFloater
End If
        
'create track string
frmNavigationString.txtNavigationString.Text = TrackCalculator.CreateVerboseTrackString(FishNumber)
frmNavigationString.Caption = "Fish " & cmbFishCode.Text & " with " & Format(FishDatabase.NumberOfStamps(CInt(FishNumber)), "###,###,###,###") & " detections."

End Sub
Public Sub WarningTrackNotVisible(Status As Boolean)
Dim i As Long

If Status = True Then
    With Form1.lblTrackNotVisible
        .Left = (Form1.Picture1.ScaleWidth - .Width) / 2
        .Top = (Form1.Picture1.ScaleHeight - .Height) / 2
        .Enabled = False
        .Visible = True
    End With
    With frmFloater
        For i = 0 To .optShow.UBound
            .optShow(i).Enabled = False
        Next i
        .cmdWalk.Enabled = False
        .chkMove.Enabled = False
        .chkStay.Enabled = False
        .Frame1.Enabled = False
        .Frame2.Enabled = False
        .Frame3.Enabled = False
        .Label1.Enabled = False
        .Label2.Enabled = False
        .Label3.Enabled = False
        .Label4.Enabled = False
        .Label5.Enabled = False
        .txtAvgLocation.Enabled = False
        .txtLinearity.Enabled = False
        .txtMeandering.Enabled = False
        .txtTTLDistance.Enabled = False
    End With
    frmNavigationString.Hide
    frmNavigationString.txtNavigationString.Text = ""
    Form1.DisableCanvas
Else
    Form1.EnableCanvas
    Form1.lblTrackNotVisible.Visible = False
    With frmFloater
        For i = 0 To .optShow.UBound
            .optShow(i).Enabled = True
        Next i
        .cmdWalk.Enabled = True
        .chkMove.Enabled = True
        .chkStay.Enabled = True
        .Frame1.Enabled = True
        .Frame2.Enabled = True
        .Frame3.Enabled = True
        .Label1.Enabled = True
        .Label2.Enabled = True
        .Label3.Enabled = True
        .Label4.Enabled = True
        .Label5.Enabled = True
        .txtAvgLocation.Enabled = True
        .txtLinearity.Enabled = True
        .txtMeandering.Enabled = True
        .txtTTLDistance.Enabled = True
    End With
End If

End Sub
Private Sub WaitAMinute(Optional FirstRun As Boolean = False)
'Waits one minute according to timer
'and also animates clock
Static X As Long

Const MidNightTime = 1440
Const Y2 = 8

If FirstRun Then X = 0

WAITING = True

'user can also select to "fly" by the compression
If TimeCompressionFactor > 0 Then
    tmrCompressedTime.Enabled = True
    Do While WAITING
        DoEvents
    Loop
    tmrCompressedTime.Enabled = False
End If


'update ticker
X = X + 1
X = Round(TICK * TimeScale, 0)
frmClock.picTimeLine.Cls
ImageProcessingEngine.DrawLineWithArrow frmClock.picTimeLine, X, X, 0, Y2


'and clock face
RealTime = RealTime + TimeBinsForAnimation

If RealTime >= 1440 Then
    RealTime = TimeBinsForAnimation
End If
frmClock.txtTime.Text = Convert_ToStandardTime(RealTime)


'flag for time point?
If FLAG_TIME_POINT Then
    'set time point marker!
    frmClock.picThis.Line (X, 0)-(X, 20), vbBlack
End If

End Sub

Private Sub WalkInTime()
'Shows progression of track in realtime

Const FirstRun = True

Dim X As Integer
Dim Y As Integer


Dim FishNumber As Long
Dim i As Long

Dim OldX As Single
Dim OldY As Single

Dim Frame As Long

Dim Site_ID As String

Dim Latitude As Single
Dim Longitude As Single

Dim PM As Boolean
Dim Prev_PM As Boolean
Dim Line_Color As Long


Dim ReceiverNumber As Integer
Dim Previous_RecieverNumber As Integer

Dim DateDetected As Long
Dim t As Integer

Dim NumberOfStamps As Long
Dim StartTime As Long
Dim StopTime As Long
Dim TimeAfterMidnight As Long
Dim StartStamp As Long
Dim cStamp As Long
Dim Moved As Boolean

Static PlotNumber As Long

'Clear Screen
Form1.ClearScreen

'Receivers no longer visible
Track_Visible = False
Receivers_Are_Visible = False
Receiver.MakeInvisible


'Load Display Form
frmDisplayDetails.Show

'Load and prepare driver
DeviceBuffer.Select_Device Device_Type.Window, frmDisplayDetails.picShowInfo

'Show scroll bar if needed and adjust its max value depending on buffer
If DeviceBuffer.EndOfPage Then frmDisplayDetails.VScroll.Visible = True
frmDisplayDetails.VScroll.Max = DeviceBuffer.LastEntryInBuffer

'reset and clear track calculations
TrackCalculator.Clear
'and diel stuff
ImageProcessingEngine.ClearDielCycle


'Get index to code number
FishNumber = cmbFishCode.ListIndex - 1
If FishNumber < 0 Then Exit Sub


'change color of pen
PlotNumber = PlotNumber + 1
If PlotNumber > 15 Then PlotNumber = 0

'Load stamps
FishDatabase.Fish = FishNumber
'get number of stamps
NumberOfStamps = FishDatabase.NumberOfStamps

'get start time
i = 0
Do
    FishTable.ReadStamp FishNumber, i
    i = i + 1
Loop Until Stamp.Valid Or i = NumberOfStamps

StartStamp = i - 1
TimeAfterMidnight = Stamp.Time

'start at midnight of that day
StartTime = Stamp.CTime - TimeAfterMidnight
RealTime = 0

'and stop time
FishTable.ReadStamp FishNumber, NumberOfStamps - 1
StopTime = Stamp.CTime + TimeBinsForAnimation + 1
'total transit time
TotalTime = (StopTime - StartTime)
 
 
'each increment is 30 minutes
'so wait 30 minutes of compressed time b/w animation, then follow to next
If TimeCompressionFactor > 0 Then tmrCompressedTime.Interval = TimeCompressionFactor

'user can also select to "fly" by the compression
If TimeCompressionFactor = 0 Then tmrCompressedTime.Interval = 0

'load clock window
Load frmClock

'calculate time scale
TimeScale = (700 / TotalTime)

'show time
frmClock.txtTime.Text = Convert_ToStandardTime(RealTime)
frmClock.LoadCurrentFish
frmClock.Show

FishTable.ReadStamp FishNumber, StartStamp

'preconds
i = StartTime
cStamp = StartStamp
ReceiverNumber = Stamp.Site
TICK = 0

'draw
Receiver.DrawReceiver Form1.Picture1, ReceiverNumber, 2
WaitAMinute FirstRun
tmrBackgroundAnimation.Enabled = True
Do

    'no flag time point yet until you find a stamp that is valid for this time point
    FLAG_TIME_POINT = False
    TICK = TICK + TimeBinsForAnimation
    Do While (i = Stamp.CTime Or i > Stamp.CTime) And cStamp <= NumberOfStamps
        If StopNow Then Exit Do
        'mark this point in timeline
        FLAG_TIME_POINT = True
        'store previous coordinates
        Previous_RecieverNumber = ReceiverNumber
        'make "visible"
        ReceiverNumber = Stamp.Site
        Receiver.MakeVisible ReceiverNumber
        
        'load site into track calculator
        TrackCalculator.Site = ReceiverNumber
     
        'load coordinates
        
        Site_ID = Receiver.ID(Stamp.Site)
        frmDisplayDetails.txtReceiver.Text = Site_ID
     
        'date and time information
        DateDetected = Stamp.Date
        frmDisplayDetails.txtDate.Text = Convert_DayNumber(DateDetected)
        frmDisplayDetails.txtTime.Text = Convert_ToStandardTime(Stamp.Time)
     
        'load into calculator
        t = Stamp.Time
        TrackCalculator.Day = DateDetected
        TrackCalculator.Time = t
        'calculate accumulators
        TrackCalculator.Calculate
        
        If Previous_RecieverNumber <> ReceiverNumber Then Moved = True
                
        'draw from last to this one if move, or draw a stay
        'draw lines if position changed
        If Moved Then
            'draw line
            Receiver.DrawRoute ReceiverNumber, Previous_RecieverNumber, Form1.Picture1, FishDatabase.Color(Stamp.Fish)
            Receiver.DrawReceiver Form1.Picture1, ReceiverNumber
            Receiver.DrawReceiver Form1.Picture1, Previous_RecieverNumber
        End If
        
        
        'check if form for keeping track of diel cycle
        If DielCycleFormIsLoaded Then
           'clear
           'frmDayLightCycle.picGraph.Cls
           'for day&fish
           ImageProcessingEngine.DrawDielCycle ' DateDetected, T, CDbl(Longitude), CDbl(Latitude)
        End If
                
        'validate
        If cStamp < NumberOfStamps Then
            Do
                'advance to next valid stamp
                FishTable.ReadStamp FishNumber, cStamp
                cStamp = cStamp + 1
            Loop Until Stamp.Valid Or cStamp = NumberOfStamps Or StopNow
        Else
            Exit Do
        End If
   Loop
   
   'increment
   i = i + TimeBinsForAnimation
   WaitAMinute
Loop Until i >= StopTime Or StopNow

tmrBackgroundAnimation.Enabled = False
Receiver.DrawReceiver Form1.Picture1, ReceiverNumber
'Write fish number to window
WriteTrackInformation FishNumber

'track is visible after animation is done
Track_Visible = True
cmdWalk.Visible = True
cmdStop.Visible = False
'unload clock window
'Unload frmClock

End Sub
Private Sub CaptureFrame()
'captures frame into dib and saves it into file for creating an AVI
'*need to insert/change code*
Dim Success As Boolean
Dim lngDC As Long

With ExportedFrame
    lngDC = GetDC(Form1.Picture1.hwnd)
    Success = .CreateFromImage(Form1.Picture1, lngDC)
    Success = .SavePicture(ProgramSettings.MoviePath & "bmp")
    ReleaseDC Form1.Picture1.hwnd, lngDC
End With

ImageToMovie.Load ProgramSettings.MoviePath & "bmp"

If FirstFrameCaptured Then
    FirstFrameCaptured = False
    Creator.StreamCreate ImageToMovie
Else
    Creator.StreamAdd ImageToMovie
End If

End Sub
Private Sub HighlightVisitedReceivers(R() As Boolean)
'erases screeens and shows/highlight receivers visited
Dim i As Integer


For i = 0 To Receiver.TotalReceivers
    If R(i) = True Then Receiver.DrawReceiver Form1.Picture1, i
Next i

End Sub
Private Sub CreateDibSectionForExport()
'creates the dib used to place the image for export
'the dib is based on the map image
Dim Success As Boolean
Dim L As Long
With ExportedFrame
    .ClearUp
    .UseDrawDib = True
    Success = .Create(Form1.Picture1.ScaleWidth, Form1.Picture1.ScaleHeight)
End With
End Sub
Private Sub WriteDate(d As Long)
Dim s As String

s = "Day: " & Convert_DayNumber(d)
Form1.Picture1.Line (0, 0)-(120, 12), vbWhite, BF
Form1.Picture1.CurrentX = 4
Form1.Picture1.CurrentY = 2
Form1.Picture1.Print s

End Sub
Private Sub AnimateAllFish(ShowMoves As Boolean, ShowStays As Boolean, Optional ShowReceiverOnly As Boolean = False)

Dim i As Long
Dim Max As Long
Dim NextDayToPlot(MAX_FISH) As Long
Dim StampNumber(MAX_FISH) As Long

Dim X As Integer
Dim Y As Integer
Dim FirstDay As Long

Dim FishNumber As Integer

Dim Frame As Long

Dim Site_ID As String
Dim LastReceiverLoaded As Integer
Dim Latitude As Single
Dim Longitude As Single
Dim DayPlotted As Long
Dim FishColor As Long
Dim LastFish As Long
Dim LastDay As Long
Dim s As Long
Dim PrevRes(MAX_FISH) As Integer
Dim ReceiverNumber As Integer
Dim Lap As Long
Dim Valid As Boolean
Dim FoundValidFishPlotDate As Boolean
Dim Visited(MAX_RECEIVERS) As Boolean

'Clear Screen
ImageProcessingEngine.ClearDielCycle

'hide canvas
Form1.Picture2.Visible = True
Form1.Picture1.Visible = False

'load dates (if not loaded)
frmDates.Show

Form1.ClearScreen
'grab first day
With frmDates.lstDates
    If .ListCount = 0 Then Exit Sub
    FirstDay = .List(0)
    LastDay = .List(.ListCount - 1)
End With
If FirstDay = 0 Then Exit Sub

LastFish = FishDatabase.TotalFishLoaded
WriteDate FirstDay

'Get index to code number
DayPlotted = FirstDay

'show it again!
Form1.Picture1.Visible = True
Form1.Picture2.Visible = False


Do Until DayPlotted > LastDay Or StopNow
    FishNumber = 0
    FoundValidFishPlotDate = False
    Do Until FishNumber > LastFish Or StopNow
        'read fish stamp
         If DayPlotted = NextDayToPlot(FishNumber) Or NextDayToPlot(FishNumber) = 0 Then
            'make sure its plot day
            'and stamp is valid
            s = StampNumber(FishNumber) + 1
            If s >= FishDatabase.NumberOfStamps(FishNumber) Then
                Valid = False
            Else
                FishTable.ReadStamp FishNumber, s - 1
                
                'if not store and move on
                If DayPlotted <> Stamp.Date Then
                    NextDayToPlot(FishNumber) = Stamp.Date
                    Valid = False
                Else
                    If Stamp.Valid Then
                        Valid = True
                    Else
                        Valid = False
                        StampNumber(FishNumber) = s
                    End If
                End If
            End If
            If Valid And Stamp.Site = PrevRes(FishNumber) Then
                If ShowStays Then Receiver.DrawReceiver Form1.Picture1, Stamp.Site, 2
                Valid = False
                StampNumber(FishNumber) = s
                FoundValidFishPlotDate = True
            End If
            
            'if valid, animate
            If Valid Then
                'store stamp
                StampNumber(FishNumber) = s
                FoundValidFishPlotDate = True
                'this is for animation purposes
                Frame = Frame + 1
                If Frame > 3 Then
                    Frame = 0
                End If
             
                FishColor = FishDatabase.Color(FishNumber)
        
                ReceiverNumber = Stamp.Site
                Visited(ReceiverNumber) = True
                Receiver.MakeVisible ReceiverNumber
              
                'load coordinates
                Latitude = Receiver.LA(ReceiverNumber)
                Longitude = Receiver.LO(ReceiverNumber)
                X = Receiver.X(ReceiverNumber)
                Y = Receiver.Y(ReceiverNumber)
                
                'get site id
                Site_ID = Receiver.ID(Stamp.Site)
                'MOVE
                If ShowStays Then
                    Receiver.DrawReceiver Form1.Picture1, ReceiverNumber
                End If
                If PrevRes(FishNumber) <> 0 Then
                    If ShowMoves Then
                        Receiver.DrawReceiver Form1.Picture1, ReceiverNumber
                        Receiver.DrawRoute ReceiverNumber, PrevRes(FishNumber), Form1.Picture1, FishColor
                    End If
                End If
                PrevRes(FishNumber) = ReceiverNumber
                'check if form for keeping track of diel cycle
                DrawDielCycle
            End If
        End If
        FishNumber = FishNumber + 1
    Loop
    If Not FoundValidFishPlotDate Then
        DayPlotted = DayPlotted + 1
        WaitBetweenDays
        'update lap counter for every day
        Lap = Lap + 1
        If Lap = Lap_MAX Then
            Lap = 0
            Form1.ClearScreen
        End If
        WriteDate DayPlotted
    End If
Loop

'track is visible after animation is done
Track_Visible = True
cmdWalk.Visible = True
cmdStop.Visible = False
tmrBWDays.Enabled = False
Form1.ClearScreen
End Sub
Private Sub AnimateTrack(ShowMoves As Boolean, ShowStays As Boolean, Optional ShowReceiverOnly As Boolean = False)

'Show all entries first
Dim TimeAtReceiver(MAX_RECEIVERS) As Long
Dim Visited(MAX_RECEIVERS) As Boolean

Dim i As Long
Dim Max As Long

Dim X As Integer
Dim Y As Integer


Dim FishNumber As Integer

Dim Frame As Long

Dim Site_ID As String

Dim Latitude As Single
Dim Longitude As Single

Dim PM As Boolean
Dim Prev_PM As Boolean

Dim ReceiverNumber As Integer
Dim Previous_RecieverNumber As Integer
Dim Residence(MAX_RECEIVERS) As Long
Dim ThisTime(MAX_RECEIVERS) As Long
Dim LastTime(MAX_RECEIVERS) As Long
Dim FishColor As Long
Dim ReceiverCount As Long

Dim Lap As Long

Const Threshold = 60

'Clear Screen
Form1.ClearScreen

'Get index to code number
FishNumber = cmbFishCode.ListIndex - 1
If FishNumber < 0 Then AnimateAllFish ShowMoves, ShowStays, ShowReceiverOnly: Exit Sub
FishColor = FishDatabase.Color(FishNumber)


If ShowMoves Or ShowStays Then
    'Load Display Form
    frmDisplayDetails.Show
    
    'Load and prepare driver
    DeviceBuffer.Select_Device Device_Type.Window, frmDisplayDetails.picShowInfo
    
    'Show scroll bar if needed and adjust its max value depending on buffer
    If DeviceBuffer.EndOfPage Then frmDisplayDetails.VScroll.Visible = True
    frmDisplayDetails.VScroll.Max = DeviceBuffer.LastEntryInBuffer
End If

'reset and clear track calculations
TrackCalculator.Clear
ImageProcessingEngine.ClearDielCycle


'move this out
If CAPTURE Then CreateDibSectionForExport

'fish number
FishDatabase.Fish = FishNumber

'residence
If ResidenceWindowIsLoaded Then
    'make all receivers visible to the program to compute residence if needed
    Receiver.MakeVisible
    TrackCalculator.ComputeResidence TimeAtReceiver, FishNumber
    For i = 0 To MAX_RECEIVERS
        If TimeAtReceiver(i) > Max Then Max = TimeAtReceiver(i)
    Next i
    With ColorScale
        .Max = Max
        .Min = 0
    End With
    frmScale.Show
End If


'Receivers no longer visible
Track_Visible = False
Receivers_Are_Visible = False
Receiver.MakeInvisible

 For i = 0 To FishDatabase.NumberOfStamps - 1
      'this is for animation purposes
     Frame = Frame + 1
     If Frame > 3 Then
         Frame = 0
     End If
     
     'read fish stamp
     FishTable.ReadStamp FishNumber, i
     If Stamp.Valid Then
        Previous_RecieverNumber = ReceiverNumber
        ReceiverNumber = Stamp.Site
        Visited(ReceiverNumber) = True
     
        'Load stamp into track calculator
        LoadStampToTrackCalculator
        Receiver.MakeVisible ReceiverNumber
     
          
        'load coordinates
        Latitude = Receiver.LA(ReceiverNumber)
        Longitude = Receiver.LO(ReceiverNumber)
        X = Receiver.X(ReceiverNumber)
        Y = Receiver.Y(ReceiverNumber)
     
     
        If ShowReceiverOnly Then
            Receiver.DrawReceiver Form1.Picture1, ReceiverNumber
        End If
        
        'get site id
        Site_ID = Receiver.ID(Stamp.Site)
     
        If ShowStays Or ShowMoves Then
            'site, date and time information
            With frmDisplayDetails
                .txtReceiver.Text = Site_ID
                .txtDate.Text = Convert_DayNumber(Stamp.Date)
                .txtTime.Text = Convert_ToStandardTime(Stamp.Time)
            End With
        End If
     
        'Show location and make sure we know its active even if fish has not moved
        'draw lines if position changed
        If ReceiverNumber = Previous_RecieverNumber Then
            'STAY
            If ShowStays Then
                DrawStay ReceiverNumber, Frame
                AnimateWait
            End If
        
            If ResidenceWindowIsLoaded Then
                LastTime(ReceiverNumber) = ThisTime(ReceiverNumber)
                ThisTime(ReceiverNumber) = Stamp.CTime
                If (ThisTime(ReceiverNumber) - LastTime(ReceiverNumber)) <= Threshold Then
                    Residence(ReceiverNumber) = Residence(ReceiverNumber) + (ThisTime(ReceiverNumber) - LastTime(ReceiverNumber))
                Else
                    LastTime(ReceiverNumber) = 0
                End If
            
                If LastTime(ReceiverNumber) <> 0 Then
                    ImageProcessingEngine.DrawReceiverWithDensity ReceiverNumber, Residence(ReceiverNumber)
                End If
                
                AnimateWait
            End If
        Else
            'MOVE
            AnimateWait
            'update lap counter
            Lap = Lap + 1
            If Lap = Lap_MAX Then
                Lap = 0
                Form1.ClearScreen
            End If
        
            'if not first, draw!
            If Previous_RecieverNumber <> 0 Then
                If ShowMoves Then DrawMove ReceiverNumber, Previous_RecieverNumber, Form1.Picture1, FishColor, Lap, Visited
                Residence(ReceiverNumber) = 0
                ThisTime(ReceiverNumber) = 0
            End If
        End If
    
        'check if form for keeping track of diel cycle
        DrawDielCycle
    End If
    If StopNow Then Exit Sub
Next i

If ShowMoves Or ShowStays Then
    'Write fish number to window
    WriteTrackInformation CLng(FishNumber)
End If

'track is visible after animation is done
Track_Visible = True
cmdWalk.Visible = True
cmdStop.Visible = False
End Sub
Private Sub AnimateWait()
'turn animation on
Animate = False
tmrAnimation.Enabled = True
Do While Not Animate
    DoEvents
Loop
End Sub
Private Sub WaitBetweenDays()
'turn animation on
Animate = False
tmrBWDays.Enabled = True
Do While Not Animate
    DoEvents
Loop
End Sub
Private Sub DrawDielCycle()
If DielCycleFormIsLoaded Then
   'clear
  ' frmDayLightCycle.picGraph.Cls
   'for day&fish
   ImageProcessingEngine.DrawDielCycle
End If

End Sub

Private Sub DrawMove(ReceiverNumber As Integer, Previous_ReceiverNumber As Integer, Canvas As PictureBox, FishColor As Long, Lap As Long, Visited() As Boolean)
Receiver.DrawReceiver Canvas, ReceiverNumber
Receiver.DrawRoute ReceiverNumber, Previous_ReceiverNumber, Canvas, FishColor
If Lap = 0 Then HighlightVisitedReceivers Visited
End Sub
Private Sub DrawStay(R As Integer, F As Long)
Dim X As Long
Dim Y As Long
X = Receiver.X(R)
Y = Receiver.Y(R)
Form1.Picture1.Circle (X, Y), 8, QBColor(1 + F)
Form1.Picture1.Circle (X, Y), 5, Receiver.Color(R)
End Sub
Private Sub LoadStampToTrackCalculator()
'load stamp to track calculator
With TrackCalculator
    .Site = Stamp.Site
    .Time = Stamp.Time
    .Day = Stamp.Date
    .Calculate
End With
End Sub

Private Sub JumpTrack()

'Show all entries first


Dim X As Integer
Dim Y As Integer


Dim FishNumber As Long
Dim i As Long


Dim Latitude As Single
Dim Longitude As Single

Dim ReceiverNumber As Integer
Dim Previous_RecieverNumber As Integer

Dim DateDetected As Long
Dim t As Integer

Static PlotNumber As Long

'Clear Screen
Form1.ClearScreen


'Receivers no longer visible
Track_Visible = False
Receivers_Are_Visible = False
Receiver.MakeInvisible

'Load Display Form
frmDisplayDetails.Show

'Load and prepare driver
DeviceBuffer.Select_Device Device_Type.Window, frmDisplayDetails.picShowInfo

'Show scroll bar if needed and adjust its max value depending on buffer
If DeviceBuffer.EndOfPage Then frmDisplayDetails.VScroll.Visible = True
frmDisplayDetails.VScroll.Max = DeviceBuffer.LastEntryInBuffer


'reset and clear track calculations
TrackCalculator.Clear
ImageProcessingEngine.ClearDielCycle

'Get index to code number
FishNumber = cmbFishCode.ListIndex - 1
If FishNumber < 0 Then Exit Sub

FishDatabase.Fish = FishNumber

 For i = 0 To FishDatabase.NumberOfStamps - 1
     FishTable.ReadStamp FishNumber, i
     If Stamp.Valid Then
        Previous_RecieverNumber = ReceiverNumber
        ReceiverNumber = Stamp.Site
        'make "visible"
        Receiver.MakeVisible ReceiverNumber
     
         'load site into track calculator
         TrackCalculator.Site = ReceiverNumber
         
         'load coordinates
         Latitude = Receiver.LA(ReceiverNumber)
         Longitude = Receiver.LO(ReceiverNumber)
         X = Receiver.X(ReceiverNumber)
         Y = Receiver.Y(ReceiverNumber)
             
         frmDisplayDetails.txtReceiver.Text = Receiver.ID(Stamp.Site)
         
         'load into calculator
         t = Stamp.Time
         TrackCalculator.Day = DateDetected
         TrackCalculator.Time = t
         'calculate accumulators
         TrackCalculator.Calculate
         
         If ReceiverNumber <> Previous_RecieverNumber And Previous_RecieverNumber <> 0 Then
            'date and time information
            DateDetected = Stamp.Date
            frmDisplayDetails.txtDate.Text = Convert_DayNumber(DateDetected)
            frmDisplayDetails.txtTime.Text = Convert_ToStandardTime(Stamp.Time)
    
            Form1.Picture1.Circle (X, Y), 5, vbBlue
            Form1.Picture1.Circle (X, Y), 1, vbRed
            'draw line
            Receiver.DrawRoute ReceiverNumber, Previous_RecieverNumber, Form1.Picture1, FishDatabase.Color(Stamp.Fish)
            
            'check if form for keeping track of diel cycle
            If DielCycleFormIsLoaded Then
               'clear
               frmDayLightCycle.picGraph.Cls
               'for day&fish
               ImageProcessingEngine.DrawDielCycle ' DateDetected, T, CDbl(Longitude), CDbl(Latitude)
            End If
            
            If StopNow Then Exit Sub
       
            'turn animation on
            Animate = False
            tmrAnimation.Enabled = True
            Do While Not Animate
                DoEvents
            Loop
        End If
    End If
Next i

'output to window

'Write fish number to window
WriteTrackInformation FishNumber

'track is visible after animation is done
Track_Visible = True

End Sub
Private Sub cmdStop_Click()
cmdWalk.Visible = True
cmdStop.Visible = False
StopNow = True
End Sub

Private Sub cmdWalk_Click()
Dim Moves As Boolean
Dim Stays As Boolean

StopNow = False
cmdWalk.Visible = False
cmdStop.Visible = True

Const ShowOnlyTheReceivers = True

If chkMove.Value = vbChecked Then Moves = True
If chkStay.Value = vbChecked Then Stays = True

If ResidenceWindowIsLoaded Then Unload frmReceiversResidenceTime
'select how to animate...
If Form1.mnuTrackAnimationStyle(Style.ScaledTime).Checked = True Then
    WalkInTime
Else
    Select Case WhatToShowOnCanvas
        Case ShowOnCanvas.FishTrack
            AnimateTrack Moves, Stays
        Case ShowOnCanvas.Receivers
            AnimateTrack Moves, Stays, ShowOnlyTheReceivers
        Case ShowOnCanvas.Residence
            frmReceiversResidenceTime.Show
            AnimateTrack Moves, Stays
        Case Else
            'cancel
            cmdWalk.Visible = True
            cmdStop.Visible = False
    End Select
End If

End Sub

Private Sub Form_GotFocus()
cmbFishCode.SetFocus
End Sub

Private Sub Form_Load()
CURRENT_FISH = -1
If Not TipWindowLoaded Then frmTip.Show vbModeless, Form1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
picPlay.Visible = False
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'if unloaded by user
'go to hiding

If UnloadMode = vbFormControlMenu Then
    Cancel = True
    Me.Hide
End If

End Sub
Public Sub ChangeTrackColor()
Dim FishNumber As Integer
On Error GoTo ExitWithError
If cmbFishCode.Text <> "" And cmbFishCode <> "ALL" Then
    'get fish #
    FishNumber = cmbFishCode.ListIndex - 1
    With CommonDialog
        .CancelError = True
        .ShowColor
        FishDatabase.Color(FishNumber) = .Color
    End With
    ShowTrack FishNumber, Form1.Picture1
End If

ExitWithError:
'nop
End Sub
Private Sub optShow_Click(Index As Integer)
WhatToShowOnCanvas = Index
If WhatToShowOnCanvas = ShowOnCanvas.Receivers Then chkStay.Value = vbChecked: chkMove.Value = vbUnchecked
If WhatToShowOnCanvas = ShowOnCanvas.FishTrack Then chkMove.Value = vbChecked
If WhatToShowOnCanvas = ShowOnCanvas.StampList Then chkMove.Value = vbUnchecked: chkStay.Value = vbChecked
RefreshCanvas
End Sub
Private Sub picChangeTrackColor_Click()
If TipWindowLoaded Then frmTip.lblToolName.Caption = Topics(7)
ChangeTrackColor
End Sub
Private Sub picPlay_Click()
Dim i As Long
Dim d As Single
Dim R As Integer
Dim L As Long
Dim Success As Boolean

Receiver.Clear
STOP_ANIMATION = False
picStop.Visible = True
If PlotIndepedently Then Factor = 0.25 Else Factor = 0.5
tmrAnimation.Interval = 240

'Add dates to Array
For i = 1 To lstDates.ListCount - 1
    Form1.StatusBar.Panels(StatusPanel.Map) = "Plotting day: " & lstDates.List(i)
    d = DayNumber(lstDates.List(i))
    R = 1
    If PlotIndepedently Then Receiver.Clear
    If STOP_ANIMATION Then Exit Sub
    Do
        Success = False
        L = 0
        Do
            Receiver.ReadStamp R, L
            If Stamp.Date = d And Stamp.Valid Then
                Receiver.PingUp (R)
                Success = True
                Exit Do
            End If
            L = L + 1
        Loop Until (L > Receiver.Detection_Total(R)) Or (Success)
        R = R + 1
    Loop Until R > Receiver.TotalReceivers
    'show density
    Form1.ShowDetectors True
    'turn animation on
    Animate = False
    tmrAnimation.Enabled = True
    Do
        DoEvents
    Loop While Not Animate
Next i

'restore table
Receiver.Restore
tmrAnimation.Enabled = False
tmrAnimation.Interval = 120
Factor = 3
End Sub

Private Sub picStop_Click()
tmrAnimation.Enabled = False
STOP_ANIMATION = True
picStop.Visible = False
End Sub

Private Sub picTool_Click(Index As Integer)

If TipWindowLoaded Then frmTip.lblToolName.Caption = Topics(Index)

If Tool_Number = ToolBox.Calibrate_Tool Then Form1.Picture1.Cls

'do nothing if selecting fish corridor tool with map not scanned,
'but do show in tooltips
If MapWasScanned = False And Index = ToolBox.DrawFishCorridor Then
    'nop
Else
    SelectTool Index
End If
End Sub
Public Sub SelectTool(Index As Integer)
Dim i As Long
Dim ReceiverNumber As Long
Dim response As Variant

'show tip

'Activate Button
picTool(Index).BackColor = Button_Highlight

'Deselect and deactivate previous, unless previous is same
If Index <> Tool_Number Then picTool(Tool_Number).BackColor = Button_Face

'select
Tool_Number = Index

'use the right cursor to denote tool selection
With Form1.Picture1
    Select Case Tool_Number
        Case ToolBox.Calibrate_Tool
            Form1.StatusBar.Panels(StatusPanel.Map) = "Map Calibration"
            .MousePointer = vbCrosshair
        Case ToolBox.Zoom
            Form1.StatusBar.Panels(StatusPanel.Map) = "Zoom"
            .MousePointer = vbCustom
        Case ToolBox.Measure_Tool
            Form1.StatusBar.Panels(StatusPanel.Map) = "Measure tool"
            .MousePointer = vbCrosshair
        Case ToolBox.DrawFishCorridor
            'if anchors are present, delete previous?
            Form1.AskIfUserWantsToDeletePreviousCorridor
        Case Else
            .MousePointer = vbArrow
    End Select
End With
        

If Tool_Number = ToolBox.Plot_Tool Then
    Form1.ClearScreen
    'make all receivers visible to use in drawing the track
    Receiver.MakeVisible
    For ReceiverNumber = 1 To Receiver.TotalReceivers
        'show it
        Receiver.DrawReceiver Form1.Picture1, CInt(ReceiverNumber), 2
    Next ReceiverNumber
    
    'tell UserDefined Track that it should allow user to define track
    CURRENT_FISH = -1
    frmUserDefinedTrack.Show
    Form1.StatusBar.Panels(StatusPanel.Map) = "User-defined track"
End If

If Tool_Number = ToolBox.RecordFrames Then
    'Start capture
    CaptureVideo
End If

End Sub
Private Sub CaptureVideo()
On Error GoTo ExitWithError
Dim response As Variant

'Tell user
Form1.StatusBar.Panels(StatusPanel.Map) = "Animation capture mode"

'Ask user
With CommonDialog
    .CancelError = True 'will return error on cancel
    .DefaultExt = "avi"
    .Filter = "AVI Video (*.avi)|*.avi"
    .FilterIndex = 0
    .DialogTitle = "Save AVI stream"
    .ShowSave
    If Len(.FileName) = 0 Then GoTo ExitWithError
    ProgramSettings.MoviePath = .FileName
End With

'Create a container
With Creator
    CreateDibSectionForExport
    .bitsPerPixel = 24
    .FileName = ProgramSettings.MoviePath
    If SetFourCC Then
        .FrameDuration = 333
        .Name = "AquaTracker Movie File"
        CAPTURE = True
        FirstFrameCaptured = True
        'set timer value
        TimerValue = Timer
        tmrFrameCapture.Enabled = True
        frmCaptureNow.Show vbModeless, frmFloater
    End If
End With
Exit Sub

ExitWithError:
response = MsgBox("AVI Stream Not Opened: User Canceled or Stream couldn't be opened!", vbOKOnly, "Capture")
Form1.StatusBar.Panels(StatusPanel.Map) = ""
End Sub
Private Function SetFourCC() As Boolean
'Depth
Dim FourCCCode As Long
Dim Success As Boolean
Dim response As Variant
On Error GoTo ExitWithError

Const Depth = 24
FourCCCode = VHS.SuggestedVideoHandlerFourCC(Depth)
If FourCCCode <> 0 Then
    Creator.VideoHandlerFourCC = FourCCCode
    Success = True
Else
    Success = False
    response = MsgBox("Unable to find a 24-bit color codec.  AquaTracker can't export unless you have a valid codec.", vbOKOnly, "Error")
End If

SetFourCC = Success

Exit Function

ExitWithError:
'Nop
response = MsgBox("Unable to find a 24-bit color codec.  AquaTracker can't export unless you have a valid codec.", vbOKOnly, "Error")
End Function

Private Function OpenSelectFolder() As String
Dim result As Variant
Dim Folder As String

'Ugly code follows. Thanks microsoft for your confusing nomenclature for shell objects!!
Set shlShell = New Shell32Ctl.Shell
Set shlFolder2 = shlShell.BrowseForFolder(Me.hwnd, "Select a Folder", BIF_RETURNONLYFSDIRS)

If Not (shlFolder2 Is Nothing) Then
    Set objFolderItem = shlFolder2.Self
    Folder = objFolderItem.PATH
    'release resource
    Set objFolderItem = Nothing
End If

'release resource
Set shlShell = Nothing
Set shlFolder2 = Nothing

OpenSelectFolder = Folder

End Function
Private Sub tmrAnimation_Timer()
Animate = True
End Sub
Private Sub TransferDatesToList(List As ListBox)
'transfer contents of list alldates to a list
'(optional, yet to implement: filter by fish! Perhaps using an array as source?)
'
Dim i As Long

If lstAllDates.ListCount = 0 Then Exit Sub

For i = 0 To lstAllDates.ListCount - 1
    List.AddItem Convert_DayNumber(Val(lstAllDates.List(i)))
Next i
    
End Sub

Private Sub tmrBackgroundAnimation_Timer()
Static Q As Integer
Dim B As Long
Dim R As Integer

Q = Q + 1
If Q > 15 Then Q = 0

B = QBColor(Q)

R = Stamp.Site
Receiver.DrawReceiver Form1.Picture1, R, 1, B


End Sub

Private Sub tmrBWDays_Timer()
Animate = True
End Sub

Private Sub tmrCompressedTime_Timer()

WAITING = False

End Sub

Private Sub tmrFrameCapture_Timer()


'store frame or close stream
If CAPTURE Then
    'timer check
    'ensures 3.0fps capture rate
    If Abs(Timer - TimerValue) >= 0.33 Then
        TimerValue = Timer
        CaptureFrame
    End If
Else
    Creator.StreamClose
    tmrFrameCapture.Enabled = False
End If

'update timer

End Sub
