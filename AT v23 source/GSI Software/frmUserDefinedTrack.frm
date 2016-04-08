VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmUserDefinedTrack 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reference track"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4185
   Icon            =   "frmUserDefinedTrack.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   4185
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstFishNumber 
      Height          =   255
      Left            =   1800
      TabIndex        =   18
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   2520
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdColorCluster 
      Caption         =   "Assign color to cluster"
      Height          =   495
      Left            =   120
      TabIndex        =   17
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton cmdCopyToClipBoard 
      Caption         =   "Export to clipboard..."
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   4800
      Width           =   1815
   End
   Begin VB.PictureBox picColorScale 
      AutoRedraw      =   -1  'True
      Height          =   255
      Left            =   120
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   117
      TabIndex        =   14
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CommandButton cmdColorize 
      Caption         =   "Heat-map all tracks"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox txtThreshold 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1200
      TabIndex        =   12
      Text            =   "0.05"
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtp 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3480
      TabIndex        =   10
      Text            =   "2.0"
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox txtWeight 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3480
      TabIndex        =   9
      Text            =   "1.0"
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox picFormula 
      Height          =   975
      Left            =   2280
      Picture         =   "frmUserDefinedTrack.frx":030A
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   117
      TabIndex        =   6
      Top             =   1680
      Width           =   1815
   End
   Begin VB.ListBox lstParams 
      Height          =   1230
      Left            =   2280
      TabIndex        =   4
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Find matching tracks"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1815
   End
   Begin VB.ListBox lstFish 
      Height          =   1425
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   1695
   End
   Begin VB.ListBox lstReceivers 
      Height          =   255
      Left            =   2400
      TabIndex        =   0
      Top             =   4080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   2160
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line2 
      BorderWidth     =   4
      X1              =   0
      X2              =   2160
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label lblSimilar 
      Caption         =   "Less similar------------------>"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Threshold:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Minkowski's exponent p:"
      Height          =   495
      Left            =   2400
      TabIndex        =   8
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label lblW 
      Caption         =   "Weight (w):"
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Parameters (V):"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      X1              =   2160
      X2              =   2160
      Y1              =   0
      Y2              =   6240
   End
   Begin VB.Label Label1 
      Caption         =   "Fish with similar tracks:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Clipboard"
      Visible         =   0   'False
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
      End
   End
End
Attribute VB_Name = "frmUserDefinedTrack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FishTrack(MAX_FISH) As TrackParameters
Dim FishStart_Receiver_LA(MAX_FISH) As Single
Dim FishStart_Receiver_LO(MAX_FISH) As Single
Dim SelectedParameter As Integer
Dim Threshold As Single
Dim w(6) As Single
Dim p(6) As Single
Dim UnderProgramControl As Boolean

Dim SelectedFishTrack As Long
Dim Start_Lattitude As Single
Dim Start_Longitude As Single
Dim ReferenceValue As TrackParameters
Dim SumArray(MAX_FISH) As Single
Dim FishArray(MAX_FISH) As Long
Dim Step As Integer

Dim UserTrackCreated As Boolean


Private Type TrackParameters
    LA As Single
    LO As Single
    Linearity As Single
    Meandering As Single
    Distance As Single
End Type
Private Enum TrackParametersWeight
    PSI_La
    PSI_Lo
    Linearity
    MI
    Start_La
    Start_Lo
    TotalDistance
End Enum
Private Sub CalculateValuesForUserDefinedTrack()
'calculate UD track params
Dim Index As Long

Do While Receiver.UserDefinedTrack_Receiver(Index) <> -1
    'add site
    TrackCalculator.Site = Receiver.UserDefinedTrack_Receiver(Index)
    TrackCalculator.Calculate
    'advance
    Index = Index + 1
Loop

'get new params
With ReferenceValue
    .LA = TrackCalculator.Path_Similarity_Index_La
    .LO = TrackCalculator.Path_Similarity_Index_Lo
    .Meandering = TrackCalculator.Meandering_Index
    .Linearity = TrackCalculator.Linearity
    .Distance = TrackCalculator.Total_Displacement
End With

Start_Lattitude = Receiver.LA(Receiver.UserDefinedTrack_Receiver(0))
Start_Longitude = Receiver.LO(Receiver.UserDefinedTrack_Receiver(0))


End Sub
Private Sub CalculateValuesForUserSelectedTrack()
Dim R As Integer
Dim c As Long
Dim FirstStamp As Long

'save color
c = FishDatabase.Color(SelectedFishTrack)
'show as red now
FishDatabase.Color(SelectedFishTrack) = vbRed
ShowTrack SelectedFishTrack, Form1.Picture1

'get all the details
With ReferenceValue
    .LA = TrackCalculator.Path_Similarity_Index_La
    .LO = TrackCalculator.Path_Similarity_Index_Lo
    .Meandering = TrackCalculator.Meandering_Index
    .Linearity = TrackCalculator.Linearity
    .Distance = TrackCalculator.Total_Displacement
End With

Do
    FishTable.ReadStamp SelectedFishTrack, FirstStamp
    FirstStamp = FirstStamp + 1
Loop Until Stamp.Valid Or FirstStamp = FishDatabase.NumberOfStamps

R = Stamp.Site
Start_Lattitude = Receiver.LA(R)
Start_Longitude = Receiver.LO(R)

'revert color
FishDatabase.Color(SelectedFishTrack) = c

End Sub
Private Sub cmdCalculate_Click()
Static CalculatePSIOnce As Boolean
Dim Index As Long
Dim FishPath As Boolean
Dim TrackPath As Boolean
Dim Sum As Single
Dim NormalizedValue As TrackParameters
Dim NormalizedReferenceValue As TrackParameters
Dim Distance As Double
Dim FishNumber As Long
Dim response As Variant

On Error GoTo RaiseError

'disable if no track selected and not defined yet
If SelectedFishTrack = -1 And Receiver.UserDefinedTrack_NumberOfReceiversInTrack = 0 Then Exit Sub


'calculate fish psi's
If Not CalculatePSIOnce Then
    CalculatePSIOnce = True
    CalculatePSI
End If
'clear all accumulators
TrackCalculator.Clear

'get proper track info into ReferenceValue
If SelectedFishTrack = -1 Then
    CalculateValuesForUserDefinedTrack
Else
    CalculateValuesForUserSelectedTrack
End If

NormalizedReferenceValue = Normalize(ReferenceValue)

'clear list
lstFish.Clear
lstFishNumber.Clear

'calculate similarity and display
'Get index to code number
For FishNumber = 0 To FishDatabase.TotalFishLoaded
    'distance
    Sum = 0
    
    'normalize value
    NormalizedValue = Normalize(FishTrack(FishNumber))
    Sum = Sum + w(TrackParametersWeight.PSI_La) ^ p(TrackParametersWeight.PSI_La) * Abs(NormalizedReferenceValue.LA - NormalizedValue.LA) ^ p(TrackParametersWeight.PSI_La)
    Sum = Sum + w(TrackParametersWeight.PSI_Lo) ^ p(TrackParametersWeight.PSI_Lo) * Abs(NormalizedReferenceValue.LO - NormalizedValue.LO) ^ p(TrackParametersWeight.PSI_Lo)
    Sum = Sum + w(TrackParametersWeight.Linearity) ^ p(TrackParametersWeight.Linearity) * Abs(NormalizedReferenceValue.Linearity - NormalizedValue.Linearity) ^ p(TrackParametersWeight.Linearity)
    Sum = Sum + w(TrackParametersWeight.MI) ^ p(TrackParametersWeight.MI) * Abs(NormalizedReferenceValue.Meandering - NormalizedValue.Meandering) ^ p(TrackParametersWeight.MI)
    Sum = Sum + w(TrackParametersWeight.Start_La) ^ p(TrackParametersWeight.Start_La) * Abs(FishStart_Receiver_LA(FishNumber) - Start_Lattitude) ^ p(TrackParametersWeight.Start_La)
    Sum = Sum + w(TrackParametersWeight.Start_Lo) ^ p(TrackParametersWeight.Start_Lo) * Abs(FishStart_Receiver_LO(FishNumber) - Start_Longitude) ^ p(TrackParametersWeight.Start_Lo)
    Sum = Sum + w(TrackParametersWeight.TotalDistance) ^ p(TrackParametersWeight.TotalDistance) * Abs(NormalizedReferenceValue.Distance - NormalizedValue.Distance) ^ p(TrackParametersWeight.TotalDistance)
    SumArray(FishNumber) = Sum
    FishArray(FishNumber) = FishNumber
   
    
    If Sum <= Threshold And FishNumber <> SelectedFishTrack Then
        FishDatabase.Fish = FishNumber
        lstFish.AddItem FishDatabase.Code
        lstFishNumber.AddItem FishNumber
    End If
Next FishNumber

'enable coloring of tracks by distance Sum
cmdColorize.Enabled = True

Exit Sub


RaiseError:
response = MsgBox("Error calculating parameters.  Use another track as reference or draw a different track.  If the problem persist, contact the program author.", vbOKOnly)

End Sub

Friend Function Normalize(ByRef t As TrackParameters) As TrackParameters
Dim response As Variant

On Error GoTo ErrorRaised

'normalize parameter
Dim N As TrackParameters

N.LA = (t.LA - Average_For_All_Tracks.LA) / (Max.LA - Min.LA)

N.LO = (t.LO - Average_For_All_Tracks.LO) / (Max.LO - Min.LO)

N.Linearity = (t.Linearity - Average_For_All_Tracks.Linearity) / (Max.Linearity - Min.Linearity)
N.Meandering = (t.Meandering - Average_For_All_Tracks.Meandering) / (Max.Meandering - Min.Meandering)

N.Distance = (t.Distance - Average_For_All_Tracks.Distance) / (Max.Distance - Min.Distance)



Normalize = N
Exit Function

ErrorRaised:
'division /0
response = MsgBox("One or more parameters have range of 0.  Unable to normalize all parameters.", vbOKOnly)
Resume Next
End Function

Private Sub CalculatePSI()
'Calculates PSI for fish in list
Dim FishNumber As Long
Dim i As Long
Dim FirstStamp As Long

'Get index to code number
For FishNumber = 1 To frmFloater.cmbFishCode.ListCount - 1
    'new fish number, reset accumulators!!
    TrackCalculator.Clear
    FishDatabase.Fish = FishNumber
    FirstStamp = 0
       For i = 0 To FishDatabase.NumberOfStamps - 1
           'get stamp
           FishTable.ReadStamp FishNumber, i
           If Stamp.Valid Then
                If FirstStamp = 0 Then FirstStamp = i
                'load site into track calculator
                TrackCalculator.Site = Stamp.Site
                        
                'load time data into calculator
                TrackCalculator.Day = Stamp.Date
                TrackCalculator.Time = Stamp.Time
                'calculate accumulators
                TrackCalculator.Calculate
            End If
       Next i
       
        With FishTrack(FishNumber)
            .LA = TrackCalculator.Path_Similarity_Index_La
            .LO = TrackCalculator.Path_Similarity_Index_Lo
            .Meandering = TrackCalculator.Meandering_Index
            .Linearity = TrackCalculator.Linearity
            .Distance = TrackCalculator.Total_Displacement
        End With
        
        'get start receiver pos
        FishTable.ReadStamp FishNumber, FirstStamp
        FishStart_Receiver_LA(FishNumber) = Receiver.LA(Stamp.Site)
        FishStart_Receiver_LO(FishNumber) = Receiver.LO(Stamp.Site)
Next FishNumber

End Sub

Private Sub cmdColorCluster_Click()
Dim c As Long
Dim i As Long
Dim f As Integer
On Error GoTo ExitWithError

'pick color for cluster tracks
With CommonDialog
    .CancelError = True
    .ShowColor
    c = .Color
End With

For i = 0 To lstFish.ListCount - 1
    f = Int(lstFishNumber.List(i))
    FishDatabase.Color(f) = c
Next i

'color fish used for comparison
If SelectedFishTrack <> -1 And Receiver.UserDefinedTrack_NumberOfReceiversInTrack = 0 Then FishDatabase.Color(SelectedFishTrack) = c

frmFloater.cmbFishCode.ListIndex = 0
Unload Me


ExitWithError:
'NOP
End Sub

Private Sub cmdColorize_Click()
Dim i As Long
ColorFishTracks SumArray(), FishArray()
For i = 0 To frmFloater.cmbFishCode.ListCount - 1
    SelectedFishTrack = i
    ShowTrack i, Form1.Picture1
Next i
End Sub

Private Sub cmdCopyToClipBoard_Click()
Dim Table As String
Dim Fish As Integer
Table = "Fish," & Space(20) & "Minkowski's distance," & Chr$(13) & Chr$(10)
For Fish = 0 To FishDatabase.TotalFishLoaded
    Table = Table & Truncated(20, FishDatabase.Code(Fish)) & "," & Format(SumArray(Fish), "00.00000") & Chr$(13) & Chr$(10)
Next Fish

Clipboard.Clear
Clipboard.SetText (Table)

End Sub
Private Function Truncated(N As Long, s As String) As String
Dim L As Long
Dim ReturnString As String

L = Len(s)
If L >= N Then
    ReturnString = Left(s, N)
Else
    ReturnString = s & Space(N - L)
End If
Truncated = ReturnString

End Function
Private Sub Form_Load()
Dim i As Integer
Dim N As Long

UnderProgramControl = True

'assume that fish on floater window is selected fish.  If no ref track drawn, use this track
SelectedFishTrack = CURRENT_FISH

If Receiver.UserDefinedTrack_NumberOfReceiversInTrack > 0 Then
    response = MsgBox("Do you want to erase previous reference track?", vbYesNo, "Reference Track")
    If response = vbYes Then
        Receiver.UserDefinedTrack_Reset
    End If
End If

Receiver.DrawUserDefinedTrack Form1.Picture1
Threshold = 0.005
  
'parameters:
    'PSI_La
    'PSI_Lo
    'Linearity
    'MI
    'Start_La
    'Start_Lo
lstParams.AddItem "Average La"
lstParams.AddItem "Average Lo"
lstParams.AddItem "Linearity"
lstParams.AddItem "Meandering"
lstParams.AddItem "First Receiver La"
lstParams.AddItem "First Receiver Lo"
lstParams.AddItem "Total Distance"

w(0) = 1
w(1) = 1
w(2) = 0.5
w(3) = 0.5
w(4) = 0.5
w(5) = 0.5
w(6) = 0

p(0) = 2
p(1) = 2
p(2) = 3
p(3) = 3
p(4) = 2
p(5) = 2
p(6) = 0.3

'show values associated with parameters
txtp.Text = Str$(p(0))
txtWeight.Text = Str$(w(0))
txtThreshold.Text = Str$(Threshold)
UnderProgramControl = False

'determine color stepping
'based on # of valid tracks
For i = 0 To FishDatabase.TotalFishLoaded
    If FishDatabase.IsVisible(i) Then
        N = N + 1
    End If
Next i

If N > 0 Then Step = CInt(255 / N)
If Step < 5 Then Step = 5
If Step > 100 Then Step = 100

'draw scale on colorscale box
ImageProcessingEngine.DrawScale picColorScale, Step


End Sub

Private Sub Form_Unload(Cancel As Integer)
frmFloater.SelectTool ToolBox.Select_Tool
End Sub
Private Sub DrawReferenceTrack()
Dim Index As Long
Dim FromSite As Long
Dim Site As Long

FromSite = Receiver.UserDefinedTrack_Receiver(0)

'calculate UD track params
Do While Receiver.UserDefinedTrack_Receiver(Index) <> -1
    Site = Receiver.UserDefinedTrack_Receiver(Index)
    Receiver.DrawRoute Site, FromSite, Form1.Picture1, vbRed
    Index = Index + 1
    FromSite = Site
Loop

End Sub
Private Sub DrawFishTrack()
Dim R As Integer
Dim c As Long

'save color
c = FishDatabase.Color(SelectedFishTrack)
'show as red now
FishDatabase.Color(SelectedFishTrack) = vbRed
ShowTrack SelectedFishTrack, Form1.Picture1

'revert color
FishDatabase.Color(SelectedFishTrack) = c

End Sub

Private Sub lstFish_Click()
Dim i As Integer
Dim FishNumber As Long

For i = 0 To lstFish.ListCount - 1
    If lstFish.Selected(i) Then
        FishNumber = FishDatabase.GetFishNumber(lstFish.List(i))
        Form1.ClearScreen
        If SelectedFishTrack = -1 Then DrawReferenceTrack Else DrawFishTrack
        frmFloater.cmbFishCode.ListIndex = FishNumber + 1
    End If
Next i
End Sub

Private Sub lstFish_DblClick()
ExcludeTracks
End Sub
Private Sub ExcludeTracks()
Dim i As Integer
Dim f As Integer

Dim FishNumber As Long
Dim response As Variant

If lstFish.ListCount > 0 Then
    response = MsgBox("Exclude all except these tracks and show on canvas?", vbYesNo, "Reference track")
    If response = vbYes Then
        For f = 0 To FishDatabase.TotalFishLoaded - 1
            FishDatabase.IsVisible(f) = False
        Next f
        
        For i = 0 To lstFish.ListCount - 1
            If lstFish.Selected(i) Then
                FishNumber = FishDatabase.GetFishNumber(lstFish.List(i))
                FishDatabase.IsVisible(FishNumber) = True
            End If
        Next i
        If SelectedFishTrack <> -1 Then FishDatabase.IsVisible(SelectedFishTrack) = True
    End If
    
    'show all
    frmFloater.cmbFishCode.ListIndex = 0
End If
End Sub
Private Sub lstFish_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'show "export to clipboard" menu
If Button = vbRightButton Then
    PopupMenu mnuPopUp
End If

End Sub

Private Sub lstParams_Click()
Dim Index As Integer


'get selected
For Index = 0 To lstParams.ListCount - 1
    If lstParams.Selected(Index) Then
        SelectedParameter = Index
    End If
Next Index

'show values associated with parameters
UnderProgramControl = True
txtp.Text = Str$(p(SelectedParameter))
txtWeight.Text = Str$(w(SelectedParameter))
UnderProgramControl = False

End Sub

Private Sub mnuCopy_Click()
Dim i As Long
Dim OutputToClipBoard As String

'gather all the tracks!
For i = 0 To lstFish.ListCount - 1
    OutputToClipBoard = OutputToClipBoard & lstFish.List(i) & Chr$(13) & Chr$(10)
Next i

'add current fish
OutputToClipBoard = OutputToClipBoard & FishDatabase.Code(SelectedFishTrack) & Chr$(13) & Chr$(10)

'output to clipboard
Clipboard.Clear
Clipboard.SetText (OutputToClipBoard)


End Sub

Private Sub txtp_Change()
Dim temp As Single
'does not allow exponents smaller than 1
'as p<1 =>not a normed vector space so distance is meaningless

If IsNumeric(txtp.Text) And Not UnderProgramControl Then
    temp = CSng(txtp.Text)
    If temp < 1 Then UnderProgramControl = True: txtp.Text = "1": UnderProgramControl = False
    p(SelectedParameter) = temp
End If

End Sub

Private Sub txtThreshold_Change()
Dim temp As Single

If IsNumeric(txtThreshold.Text) And Not UnderProgramControl Then
    temp = CSng(txtThreshold.Text)
    Threshold = temp
End If

End Sub

Private Sub txtWeight_Change()
Dim temp As Single

If IsNumeric(txtWeight.Text) And Not UnderProgramControl Then
    temp = CSng(txtWeight.Text)
    w(SelectedParameter) = temp
End If

End Sub
Private Sub ColorFishTracks(ByRef ArrayOfValues() As Single, ByRef ArrayOfFish() As Long)

Dim Intensity_Value As Integer
Dim Color As Long
Dim Fish As Integer

Dim Value As Single
Dim Previous_Value As Single
'Sort Array
QuickSort ArrayOfValues(), ArrayOfFish(), 0, FishDatabase.TotalFishLoaded

'-1 is a n/a value
Previous_Value = -1

'start from first entry
For Fish = 0 To FishDatabase.TotalFishLoaded
    'get value of entry.  This is a 0 base array but the indexing is 1 base
    Value = ArrayOfValues(ArrayOfFish(Fish))
    'if same as previous do not increment
    If Previous_Value <> Value Then
        Intensity_Value = Intensity_Value + Step
    End If
    
    'calculate color
    Color = ImageProcessingEngine.Colorize(Intensity_Value)
    'color track
    FishDatabase.Color(ArrayOfFish(Fish) + 1) = Color
    
    Previous_Value = Value
    
Next Fish
End Sub

Private Sub QuickSort(ByRef pvarArray() As Single, ByRef Index() As Long, Optional ByVal plngLeft As Long, Optional ByVal plngRight As Long)
'
'QuickSort for link lists.  First argument is the sorted by value.  Second is the link to first.  It sorts the second list using the first list as the argument.
'Sort is bounded by last two arguments which are optional if list is not partitioned
'Otherwise the last two arguments are used internally during recursion.
Dim lngFirst As Long
Dim lngLast As Long
Dim varMid As Single
Dim varSwap As Long

If plngRight = 0 Then
    plngLeft = LBound(pvarArray)
    plngRight = UBound(pvarArray)
End If
lngFirst = plngLeft
lngLast = plngRight
varMid = pvarArray(Index((plngLeft + plngRight) \ 2))

Do
    Do While pvarArray(Index(lngFirst)) < varMid And lngFirst < plngRight
        lngFirst = lngFirst + 1
    Loop
    Do While varMid < pvarArray(Index(lngLast)) And lngLast > plngLeft
        lngLast = lngLast - 1
    Loop
    If lngFirst <= lngLast Then
        varSwap = Index(lngFirst)
        Index(lngFirst) = Index(lngLast)
        Index(lngLast) = varSwap
        lngFirst = lngFirst + 1
        lngLast = lngLast - 1
    End If
Loop Until lngFirst > lngLast
If plngLeft < lngLast Then QuickSort pvarArray, Index, plngLeft, lngLast
If lngFirst < plngRight Then QuickSort pvarArray, Index, lngFirst, plngRight
End Sub

