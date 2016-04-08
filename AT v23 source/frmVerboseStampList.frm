VERSION 5.00
Begin VB.Form frmVerboseStampList 
   Caption         =   "Verbose stamp list"
   ClientHeight    =   3270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9045
   Icon            =   "frmVerboseStampList.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3270
   ScaleWidth      =   9045
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstStamps 
      Height          =   2790
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Double click on item to explore more..."
      Top             =   240
      Width           =   8775
   End
End
Attribute VB_Name = "frmVerboseStampList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Max_List = 32766
Dim EntryPoint As Long
Dim FirstTime As String
Dim FirstDate As String
Dim Counts As Long
Dim FirstStampNumber(Max_List) As Long
Dim LastStampNumber(Max_List) As Long
Dim Fish As Integer
Dim LastReceiver As Integer
Dim L As Long
Private Sub ShowStamps()
Dim i As Long
Dim t As Long
Dim s As String

For i = 0 To L - 1
    t = LastStampNumber(i) - FirstStampNumber(i)
    FishTable.ReadStamp Fish, FirstStampNumber(i)
    
    If t = 0 Then
        s = Receiver.ID(Stamp.Site) & " on " & Convert_DayNumber(Stamp.Date) & " at " & Convert_ToStandardTime(Stamp.Time)
    Else
        s = Receiver.ID(Stamp.Site) & " from " & Convert_DayNumber(Stamp.Date) & " at " & Convert_ToStandardTime(Stamp.Time)
        FishTable.ReadStamp Fish, LastStampNumber(i)
        s = s & " to " & Convert_DayNumber(Stamp.Date) & " at " & Convert_ToStandardTime(Stamp.Time)
        s = s & " (" & Str$(t + 1) & ")"
    End If
    
    'add to list
    lstStamps.AddItem s
Next i
     
'reveal
Me.Show

End Sub
Private Sub AddStamp(i As Long)
Dim R As Integer
Dim G As Integer

'adds current stamp to list
'
'use STAMP to pass stamp
R = Stamp.Site
'groups
G = Receiver.GroupNumber(R)
If G <> 0 Then R = Receiver.GroupReceiverNumber(G)

'overflow
If L >= Max_List Then Exit Sub

If LastReceiver = -1 Or LastReceiver <> R Then
    FirstStampNumber(L) = i
    LastStampNumber(L) = i
    L = L + 1
Else
    LastStampNumber(L - 1) = i
End If
LastReceiver = R
End Sub
Private Sub ClearAll()
lstStamps.Clear
Me.Hide
L = 0
LastReceiver = -1
End Sub

Private Sub Form_Load()
If CURRENT_FISH = -1 Then Exit Sub
LoadStamps
End Sub
Public Sub LoadStamps(Optional F As Integer = -1)
Dim s As Long
'fish #

frmFloater.MousePointer = vbHourglass

If F = -1 Then
    F = CURRENT_FISH
End If

Fish = F
If Fish = -1 Then Exit Sub
'Clear
ClearAll

'load stamps
For s = 0 To FishDatabase.NumberOfStamps(F) - 1
    FishTable.ReadStamp Fish, s
    If Stamp.Valid Then
        AddStamp s
    End If
Next s

'show
ShowStamps

frmFloater.MousePointer = vbArrow
End Sub
Private Sub lstStamps_Click()
Dim i As Long
Dim R As Integer
Dim s As Long

For i = 0 To lstStamps.ListCount - 1
    If lstStamps.Selected(i) = True Then
        s = FirstStampNumber(i)
        FishTable.ReadStamp Fish, s
        R = Stamp.Site
        Form1.ClearScreen
        Receiver.Show R, Form1.Picture1
    End If
Next i

        
End Sub

Private Sub lstStamps_DblClick()
'directly accesses the stamps from fish and displays them onto a dialogue with a short verbose description
'
'
Dim s As Long
Dim c As String
Dim ReceiverStampNumber As Long
Dim misc As String
Dim R As Integer
Dim FishInfo As String
Dim MJD As Long
Dim IntegerOverflow As Boolean

For i = 0 To lstStamps.ListCount - 1
    If lstStamps.Selected(i) = True Then
        'load dialog
        Load dlgStamps
        FishTable.ReadStamp Fish, FirstStampNumber(i)
        dlgStamps.Caption = Receiver.ID(Stamp.Site) & ": " & FishDatabase.Code(Fish)
        EntryPoint = -1
        For s = FirstStampNumber(i) To LastStampNumber(i)
            FishTable.ReadStamp Fish, s
            R = Stamp.Site
            MJD = Stamp.CTime
            c = "Stamp #" & Str$(s) & " on " & Convert_DayNumber(Stamp.Date) & " at " & Convert_ToStandardTime(Stamp.Time)
            dlgStamps.lstStamps.AddItem c
            FishInfo = "Fish" & Str$(FishDatabase.Code(Fish)) & " (" & Trim$(Str$(Stamp.Fish)) & ")"
            ReceiverStampNumber = SeekStamp(Fish, MJD, R)
            misc = "Stamp #" & Trim(Str$(s)) & " on " & FishInfo & ". Stamp #" & Trim$(Str$(ReceiverStampNumber)) & " on receiver " & Receiver.ID(R, True)
            With dlgStamps.lstMsc
                'integer overflow check
                If .ListCount = Max_List Then
                    IntegerOverflow = True
                    Exit For
                End If
                .AddItem misc
            End With
        Next s
    End If
Next i

dlgStamps.Show vbModal, Me

End Sub
Private Function SeekStamp(F As Integer, t As Long, R As Integer) As Long
'Finds stamp belonging to fish on receiver table
Dim i As Long
Dim Found As Boolean
Dim LastStampOnReceiver As Long

Me.MousePointer = vbHourglass

LastStampOnReceiver = Receiver.Detection_TTL(R)

Do
    EntryPoint = EntryPoint + 1
    Receiver.ReadStamp R, EntryPoint
    If Stamp.CTime = t And Stamp.Fish = F Then
        Found = True
    End If
Loop Until Found Or EntryPoint >= LastStampOnReceiver

Me.MousePointer = vbArrow

SeekStamp = EntryPoint


End Function
