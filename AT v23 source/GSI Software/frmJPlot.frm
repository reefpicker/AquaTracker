VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmJPlot 
   Caption         =   "Daily Detection Plot"
   ClientHeight    =   9390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14625
   Icon            =   "frmJPlot.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   626
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   975
   StartUpPosition =   1  'CenterOwner
   Begin VB.HScrollBar HScroll 
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   8880
      Visible         =   0   'False
      Width           =   14175
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   9015
      Width           =   14625
      _ExtentX        =   25797
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   4304
            MinWidth        =   4304
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   10583
            MinWidth        =   10583
         EndProperty
      EndProperty
   End
   Begin VB.VScrollBar VScroll 
      Height          =   8895
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      Height          =   8775
      Left            =   360
      ScaleHeight     =   581
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   941
      TabIndex        =   2
      Top             =   0
      Width           =   14175
   End
   Begin VB.ListBox lstDates 
      Height          =   2010
      Left            =   3600
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.ListBox lstFish 
      Height          =   1815
      Left            =   1560
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
   Begin VB.Menu mnuClipBoard 
      Caption         =   "Clipboard"
      Visible         =   0   'False
      Begin VB.Menu mnuGrid 
         Caption         =   "Grid"
         Begin VB.Menu mnuShowGrid 
            Caption         =   "Hide grid"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuCopyGraphToClipBoard 
         Caption         =   "Copy graph"
      End
      Begin VB.Menu mnuCopyData 
         Caption         =   "Copy data"
      End
   End
End
Attribute VB_Name = "frmJPlot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ReceiverList(MAX_RECEIVERS) As Long
Dim NumberOfReceiversInList As Long
Dim Box_width As Long
Dim Box_height As Long
Dim ListOnX As ListBox
Dim ListOnY As ListBox
Dim WindowWidth As Long
Dim WindowHeight As Long
Dim Max_Ticks_X As Long
Dim Max_Ticks_Y As Long
Dim TicksPerPage_X As Long
Dim TicksPerPage_Y As Long
Dim Loading_Window As Boolean

'10 years worth of data can be plotted
Const MAX_DAYS_PLOT = 3650
Dim Plot(MAX_DAYS_PLOT, MAX_FISH) As Integer '10 years worth of data can be plotted
Dim AltPlot(MAX_DAYS_PLOT, MAX_FISH) As Integer '10 years worth of data can be plotted
Const MIN_BOX_WIDTH = 2
Const MIN_BOX_HEIGHT = 2
Const CHAR_WIDTH = 8
Const CHAR_HEIGHT = 14
Const SEPARATION = 3
Const Max_NumberOfDayLabelsToPrintOnAxis = 6
Dim NumberOfDayLabelsToPrintOnAxis As Long
Dim StartPage_X As Long
Dim StartPage_Y As Long
Dim SHOW_Labels As Boolean
Dim SHOW_Grid As Boolean
Dim Form_Width As Long
Dim Form_Height As Long
Dim First_Receiver As Integer
Dim Last_Receiver As Integer

Private Sub PrepareList()
'Calculates all the necesary values before plotting JPlot
'A JPlot is a plot of fish vs dates detected
'

Dim ReceiverNumber As Long
Dim i As Long
Dim FirstDay As Long
Dim LastDay As Long
Dim FirstDetection As Long
Dim LastDetection As Long
Dim FirstStamp As Long
Dim LastStamp As Long
Dim response As Variant

Dim FishInList(MAX_FISH) As Boolean
If NumberOfReceiversInList = 0 Then Exit Sub

'preconditions
FirstDay = 2 ^ 30
LastDay = 0

'scan list to create the date list
'and fish list
Do
    ReceiverNumber = ReceiverList(i)
    'get first and last date of plot
    LastStamp = Receiver.Detection_TTL(CInt(ReceiverNumber))
    
    FirstStamp = 0
    If LastStamp > 0 Then
        'first valid detection
        Do
            Receiver.ReadStamp ReceiverNumber, FirstStamp
            FirstStamp = FirstStamp + 1
        Loop Until Stamp.Valid Or FirstStamp = LastStamp
        
        'if no valid stamp, do not proceed
        If Stamp.Valid Then FirstDetection = Stamp.Date Else FirstDetection = -1
        
        'to last (valid) detection
        Do
            LastStamp = LastStamp - 1
            Receiver.ReadStamp ReceiverNumber, LastStamp
        Loop Until Stamp.Valid Or LastStamp = 0
        
        'if no valid stamp, do not proceed
        If Stamp.Valid Then LastDetection = Stamp.Date Else LastDetection = -1
        
        If LastDetection > LastDay And LastDetection <> -1 Then LastDay = LastDetection
        If FirstDetection < FirstDay And FirstDetection <> -1 Then FirstDay = FirstDetection
        'get fish
        Receiver.TransferUniqueFishEntriesToList ReceiverNumber, FishInList()
    End If
    i = i + 1
Loop Until i = NumberOfReceiversInList


'clear list to prepare it
lstDates.Clear

'verify
If LastDay - FirstDay > MAX_DAYS_PLOT Then
    'OOPS... Cant do that... Only 10 or less years
    response = MsgBox("You have more than 10 years of data to plot or there was an internal error.  Split your file or use the Date window to eliminate some dates.  No more than " & Str$(MAX_DAYS_PLOT) & " days between first and last detection is allowed!", vbExclamation, "Limit exceeded")
    'show arrow, computer ready
    Form1.MousePointer = vbArrow
    Exit Sub
Else
    'populate
    For i = FirstDay To LastDay
        lstDates.AddItem Str$(i)
    Next i
    'fish... fill in the list
    For i = 0 To MAX_FISH
        If FishInList(i) Then lstFish.AddItem FishDatabase.Code(i)
    Next i
End If
End Sub

Private Sub SetValues()
Dim BW As Single
Dim BH As Single
Dim d As Single

'axis
SHOW_Labels = True
'bounds
WindowWidth = picCanvas.ScaleWidth
WindowHeight = picCanvas.ScaleHeight - (CHAR_HEIGHT + SEPARATION)

'validate
If ListOnX.ListCount = 0 Or ListOnY.ListCount = 0 Then Exit Sub
NumberOfDayLabelsToPrintOnAxis = ListOnX.ListCount
If NumberOfDayLabelsToPrintOnAxis > Max_NumberOfDayLabelsToPrintOnAxis Then NumberOfDayLabelsToPrintOnAxis = 6

'ticks per page
'and box/cell size
d = (WindowWidth - CHAR_WIDTH * Len(ListOnX.List(0)))
BW = Fix(d / ListOnX.ListCount)
BH = Fix(WindowHeight / ListOnY.ListCount)

Box_width = Fix(BW)
Box_height = Fix(BH)


'is divisible by 2
Box_width = Box_width - (Box_width Mod 2)

If Box_width < MIN_BOX_WIDTH Then
    Box_width = MIN_BOX_WIDTH
    SHOW_Grid = False
    SHOW_Labels = False
End If

If Box_height < MIN_BOX_HEIGHT Then
    Box_height = MIN_BOX_HEIGHT
    SHOW_Grid = False
    SHOW_Labels = False
End If

'snap to window grid.  Make window smaller if needed
WindowWidth = WindowWidth - (WindowWidth Mod Box_width)
WindowHeight = WindowHeight - (WindowHeight Mod Box_height)
'find max ticks that can be placed in page
Max_Ticks_X = WindowWidth / Box_width
Max_Ticks_Y = WindowHeight / Box_height
If Max_Ticks_X < 1 Then
    Max_Ticks_X = 1
End If

If Max_Ticks_Y < 1 Then
    Max_Ticks_Y = 1
End If

'default to no scroll bars
HScroll.Visible = False
VScroll.Visible = False

'set ticks per page
TicksPerPage_X = ListOnX.ListCount 'Max_Ticks_X
TicksPerPage_Y = ListOnY.ListCount 'Max_Ticks_Y
If TicksPerPage_X >= Max_Ticks_X Then
    HScroll.Max = Fix(TicksPerPage_X / Max_Ticks_X)
    'If Fix(TicksPerPage_X / Max_Ticks_X) <> (TicksPerPage_X / Max_Ticks_X) Then HScroll.Max = HScroll.Max + 1
    TicksPerPage_X = Max_Ticks_X
    HScroll.SmallChange = 1
    HScroll.LargeChange = 1
    HScroll.Visible = True
End If

If ListOnY.ListCount >= Max_Ticks_Y Then
    VScroll.Max = Fix(TicksPerPage_Y / Max_Ticks_Y)
    'If Fix(TicksPerPage_Y / Max_Ticks_Y) <> (TicksPerPage_Y / Max_Ticks_Y) Then VScroll.Max = VScroll.Max + 1
    TicksPerPage_Y = Max_Ticks_Y
    VScroll.LargeChange = 1
    VScroll.SmallChange = 1
    VScroll.Visible = True
End If


End Sub

Private Sub Command1_Click()

ShowPlot

End Sub

Private Sub Form_Load()
LoadValues
End Sub
Private Sub LoadValues()
Dim R As Integer
Dim i As Long
Dim ii As Long
Dim G As Integer

'loaded status
JPlotIsLoaded = True
Loading_Window = True

'standard values for axis:
'TO REVERSE DEFAULT AXIS change this values! and also flip box width/height
Set ListOnX = lstDates
Set ListOnY = lstFish

'clear
ListOnX.Clear
ListOnY.Clear


'get highlighted receivers
For R = 1 To Receiver.TotalReceivers
    'use the visible property of fake receiver if in group
    G = Receiver.GroupNumber(R)
    If G Then
        If Not Receiver_Table.Invisible(Receiver.GroupReceiverNumber(G)) Then
            ReceiverList(i) = R
            i = i + 1
        End If
    Else
    
        If Not Receiver_Table.Invisible(R) Then
            ReceiverList(i) = R
            i = i + 1
        End If
    End If
Next R
StartPage_X = 0
StartPage_Y = 0
NumberOfReceiversInList = i

'store window size to be used in rescaling drawing area
Form_Height = frmJPlot.ScaleHeight
Form_Width = frmJPlot.ScaleWidth

'white background
picCanvas.BackColor = vbWhite
'Prep jplot vars
PrepareList
'
'verify lists are loaded and ready for analysis, otherwise NOP
If lstDates.ListCount <> 0 And lstFish.ListCount <> 0 Then
    'Set values for dynamic Constants and others
    SetValues
    
    'Get plot info for page
    GetPlot
    
    DrawPlot
    
    'show arrow, computer ready
    Form1.MousePointer = vbArrow
End If

'once all is done, reset form1 status bar
Form1.StatusBar.Panels(1) = ""

Loading_Window = False

End Sub
Public Sub DrawPlot()
'clear canvas
picCanvas.Cls

'Draw Grid
If SHOW_Grid Then DrawGrid

'Print Axis
'If SHOW_Labels Then
PrintAxis

'Show JPlot for this page
ShowPlot

End Sub
Private Sub ShowPlot()
'show plot for current plot area
Dim i As Long
Dim ii As Long
Dim X1 As Long
Dim X2 As Long
Dim Y1 As Long
Dim Y2 As Long
Dim Bottom As Long
Dim StartOfX As Long
Dim c As Long
Dim G As Integer
Dim R As Integer
Dim a As Integer


StartOfX = Len(ListOnX.List(0)) * CHAR_WIDTH
For i = 0 To TicksPerPage_X - 1
    For ii = 0 To TicksPerPage_Y - 1
        R = Plot(i + StartPage_X, ii + StartPage_Y)
        'adjust for groups
        G = Receiver.GroupNumber(R)
        a = CInt(AltPlot(i + StartPage_X, ii + StartPage_Y))
        If Receiver.GroupNumber(a) <> 0 Then a = Receiver.GroupReceiverNumber(Receiver.GroupNumber(a))
        If G <> 0 Then
            R = Receiver.GroupReceiverNumber(G)
        End If
        
        If R <> 0 Or a <> 0 Then
            X1 = ((Box_width) * i) + StartOfX
            Y1 = (Box_height * (ii))
            X2 = (Box_width) * (i + 1) + StartOfX
            Y2 = Box_height * (ii + 1)
            c = Receiver.Color(R)
            picCanvas.Line (X2 - Box_width / 2, Y2)-(X1, Y1), c, BF
            c = Receiver.Color(a)
            picCanvas.Line (X2, Y2)-(X1 + Box_width / 2, Y1), c, BF
        End If
    Next ii
Next i
End Sub
Private Sub GetPlot()
'gets info about fish to plot it...
'reads list from start page to end of page
Dim i As Long
Dim ii As Long
Dim R As Integer
Dim c As Long
Dim Bottom As Long
Dim X As Long
Dim Y As Long
Dim FishNumber As Long
Dim count_progress As Long
Dim CurrentDate As Long

'bottom of graph is top of plot
Bottom = ListOnY.ListCount


'clear plot
For i = 0 To ListOnX.ListCount - 1
    For ii = 0 To ListOnY.ListCount - 1
        Plot(i, ii) = 0
    Next ii
Next i


'read list

For ii = 0 To ListOnY.ListCount - 1
    FishNumber = FishDatabase.GetFishNumber(lstFish.List(ii))
    For i = 0 To ListOnX.ListCount - 1
        CurrentDate = CLng(lstDates.List(i))
        FindFish FishNumber, CurrentDate
        Plot(i, ii) = First_Receiver
        AltPlot(i, ii) = Last_Receiver
    Next i
Next ii

End Sub
Private Sub FindFish(Fish As Long, d As Long)
Dim s As Long
Dim R As Long

First_Receiver = 0
Last_Receiver = 0
For s = 0 To FishDatabase.NumberOfStamps(CInt(Fish)) - 1
    FishTable.ReadStamp Fish, s
    If Stamp.Date = d And Stamp.Valid Then
        If First_Receiver = 0 Then
            First_Receiver = Stamp.Site
        End If
        Last_Receiver = Stamp.Site
    End If
Next s

End Sub
Private Sub PrintAxis()
Dim i As Long
Dim count As Long
Dim L As Long
Dim CX As Long
Dim CY As Long
Dim CW As Long
Dim CH As Long
Dim LX1 As Long
Dim LX2 As Long
Dim LY1 As Long
Dim LY2 As Long
Dim Start_Y As Long
Dim Start_X As Long
Dim d As String
Dim O As Long


'Print y axis
CH = CHAR_HEIGHT
CX = 0
count = 0
Start_Y = CH

For i = StartPage_Y To StartPage_Y + TicksPerPage_Y - 1
    With picCanvas
        LY2 = (Box_height * count)
        CY = LY2 - CH + (SEPARATION)
        .CurrentY = Start_Y + CY
        .CurrentX = CX
        picCanvas.Print ListOnY.List(i)
        O = Len(ListOnY.List(i)) * CHAR_WIDTH
        count = count + 1
    End With
Next i

'Print X axis
'print begining date, mid date, and last date on axis
CY = WindowHeight
Start_X = (CHAR_WIDTH * Len(ListOnX.List(0)))
count = 0
If TicksPerPage_X <= 1 Or TicksPerPage_Y <= 1 Then Exit Sub
For i = StartPage_X To StartPage_X + TicksPerPage_X Step (TicksPerPage_X - 1) / NumberOfDayLabelsToPrintOnAxis
    If i < lstDates.ListCount - 1 Then
        d = Convert_DayNumber(Val(lstDates.List(i)))
    Else
        d = ""
    End If
    
    With picCanvas
        LX1 = (Box_width) * (i - StartPage_X)
        L = Len(Trim$(d))
        CW = L * CHAR_WIDTH
        CX = LX1 - (CW / 2)
        .CurrentY = CY
        .CurrentX = (O / 2) + CX + Start_X
        picCanvas.Print Trim$(d)
    End With
Next i

End Sub
Private Sub DrawGrid()
'draws gridline
Dim i As Long
Dim CX As Long
Dim CY As Long

'validate
If Box_height = 0 Or Box_width = 0 Then Exit Sub

'Y axis
For i = 0 To WindowHeight Step Box_height
    picCanvas.Line (0, i)-(WindowWidth, i), vbBlack
Next i

'X axis
For i = Len(ListOnX.List(0)) * CHAR_WIDTH To WindowWidth Step Box_width
    picCanvas.Line (i, 0)-(i, WindowHeight), vbBlack
Next i

End Sub


Private Sub Form_Resize()
Dim X As Long
Dim Y As Long
Dim new_width As Long
Dim new_height As Long
Dim DeltaX As Long
Dim DeltaY As Long
If Loading_Window Then Exit Sub
'check if loaded
If Form_Width = 0 Or Form_Height = 0 Then Exit Sub

X = frmJPlot.ScaleWidth
Y = frmJPlot.ScaleHeight

If X = 0 Or Y = 0 Then Exit Sub

DeltaX = (X - Form_Width) '* TwipsPerPixelX())
DeltaY = (Y - Form_Height) ' * TwipsPerPixelY())
new_width = picCanvas.Width + DeltaX
new_height = picCanvas.Height + DeltaY

'validate
If new_width <= 0 Then
    new_width = 1
Else
    HScroll.Width = HScroll.Width + DeltaX
End If

If new_height <= 0 Then
    new_height = 1
Else
    'change position of scroll bars and their height/width
    VScroll.Height = VScroll.Height + DeltaY
    HScroll.Top = HScroll.Top + DeltaY
End If

picCanvas.Height = new_height
picCanvas.Width = new_width

Form_Width = frmJPlot.ScaleWidth
Form_Height = frmJPlot.ScaleHeight

'refresh
SetValues
DrawPlot
End Sub

Private Sub Form_Unload(Cancel As Integer)
'unloaded
JPlotIsLoaded = False
End Sub

Private Sub HScroll_Change()
StartPage_X = (HScroll.Value) * TicksPerPage_X
DrawPlot
End Sub

Private Sub mnuCopyData_Click()
Dim FishCode As String
Dim ReceiverName As String
Dim i As Long
Dim ii As Long
Dim Concatenated As String
Dim CurrentDate As String
Dim ReceiverNumber As Integer

'clear clipboard
Clipboard.Clear

'header
Concatenated = "Date," & Space(20 - Len("Date")) & "Fish Code," & Space(20 - Len("Fish Code")) & "Receiver"

'read list
    For i = 0 To ListOnX.ListCount - 1
        CurrentDate = Convert_DayNumber(CLng(lstDates.List(i)))
        For ii = 0 To ListOnY.ListCount - 1
            FishCode = lstFish.List(ii)
            ReceiverName = ""
            ReceiverNumber = Plot(i, ii)
            If ReceiverNumber > 0 Then ReceiverName = Receiver.ID(CInt(ReceiverNumber))
            If AltPlot(i, ii) > 0 And ReceiverNumber <> AltPlot(i, ii) Then
                ReceiverName = ReceiverName & " / " & Receiver.ID(CInt(AltPlot(i, ii)))
            End If
            If ReceiverName = "" Then FishCode = ""
            Concatenated = Concatenated & Chr$(13) & Chr$(10) & CurrentDate & "," & Space(20 - Len(CurrentDate)) & FishCode & "," & Space(20 - Len(FishCode)) & ReceiverName
        Next ii
Next i
'copy to clipboard
Clipboard.SetText Concatenated

End Sub

Private Sub mnuCopyGraphToClipBoard_Click()
Clipboard.Clear
Clipboard.SetData picCanvas.Image
End Sub

Private Sub mnuShowGrid_Click()
If mnuShowGrid.Checked = True Then
    SHOW_Grid = True
    mnuShowGrid.Checked = False
Else
    SHOW_Grid = False
    mnuShowGrid.Checked = True
End If
DrawPlot
End Sub

Private Sub picCanvas_DblClick()
If Receiver.CurrentStation_Number <> 0 Then frmReceiverInformation.Show
End Sub

Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'opens up popup menu
If Button = vbRightButton Then PopupMenu mnuClipboard
End Sub

Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'changes picCanvas.ToolTipText to show date
'
Dim TempColor As Long
Dim Date_Axis_X As Long
Dim d As Long
Dim Fractional As Single
Dim StartOfX As Long
Dim FishNumber As Long
Dim f As Long
Dim ReceiverNumber As Integer
Dim Concatenated As String
Dim UseAltPlotValue As Boolean

'exit if no data loaded
If lstDates.ListCount <> 0 Then
    'get X axis position in relation to gridline
    StartOfX = Len(ListOnX.List(0)) * CHAR_WIDTH
    Date_Axis_X = Fix((X - StartOfX) / (Box_width))
    Fractional = (X - StartOfX) / (Box_width)
    
    If Fractional <> Date_Axis_X Then Date_Axis_X = Date_Axis_X + 1
    
    If Fractional - Fix(Fractional) >= 0.5 Then UseAltPlotValue = True
    
    d = Date_Axis_X + StartPage_X - 1
    
    If d >= lstDates.ListCount Then Exit Sub
    
    If d > 0 Then StatusBar.Panels(1).Text = "Day: " & Convert_DayNumber(Val(lstDates.List(d)))
    'get fish number (Y axis)
    
    FishNumber = Fix(Y / (Box_height))
    
    f = FishNumber + StartPage_Y
    
    If f >= lstFish.ListCount Or f < 0 Then Exit Sub
    Concatenated = "Fish: " & lstFish.List(f)
    
    If d < 0 Then
        HighlightFishNumber CInt(f)
    Else
        If UseAltPlotValue Then
            ReceiverNumber = AltPlot(d, f)
        Else
            ReceiverNumber = Plot(d, f)
        End If
        
        Receiver.CurrentStation_Number = ReceiverNumber
        
        If ReceiverNumber Then
            Concatenated = Concatenated & " @ " & Receiver.ID(ReceiverNumber)
            If Receiver.GroupNumber(ReceiverNumber) <> 0 Then Concatenated = Concatenated & " (" & Receiver.ID(ReceiverNumber, True) & ")"
            TempColor = Receiver.Color(ReceiverNumber)
            Receiver.Color(ReceiverNumber) = HighLightColor Or &H808080
            ShowPlot
            Receiver.Color(ReceiverNumber) = TempColor
            'Draw Grid
            If SHOW_Grid Then DrawGrid
        Else
            ShowPlot
            'Draw Grid
            If SHOW_Grid Then DrawGrid
        End If
    End If
    StatusBar.Panels(2).Text = Concatenated
End If


End Sub
Public Sub HighlightFishNumber(FishNumber As Integer)
Dim i As Integer
Dim ID As String
Dim f As Integer

    
Y = (FishNumber - StartPage_Y) * Box_height
If Y >= 0 Then
    DrawPlot
    picCanvas.Line (1, Y)-(WindowWidth - 1, Y + Box_height), HighLightColor, B
    picCanvas.Line (1, Y + 1)-(WindowWidth - 1, Y + Box_height + 1), HighLightColor, B
End If

End Sub


Private Sub VScroll_Change()
StartPage_Y = (VScroll.Value) * TicksPerPage_Y
DrawPlot
End Sub

