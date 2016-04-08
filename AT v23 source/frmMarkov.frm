VERSION 5.00
Begin VB.Form frmMarkov 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "1st Order Markov chain analysis"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4545
   Icon            =   "frmMarkov.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   4545
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtThreshold 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2280
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton cmdTrackBack 
      Caption         =   "Backtrack"
      Height          =   315
      Left            =   360
      TabIndex        =   6
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton cmdTrackForward 
      Caption         =   "Track forward"
      Height          =   315
      Left            =   2760
      TabIndex        =   5
      Top             =   2880
      Width           =   1455
   End
   Begin VB.ListBox lstFrom 
      Height          =   2010
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1935
   End
   Begin VB.ListBox lstTo 
      Height          =   2010
      Left            =   2520
      TabIndex        =   1
      Top             =   720
      Width           =   1935
   End
   Begin VB.ComboBox lstReceivers 
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Threshold for tracking:"
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
      Left            =   240
      TabIndex        =   7
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Coming from:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Going to:"
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "frmMarkov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Chain
    Name As String
    GoingTo(MAX_RECEIVERS) As Long
    ComingFrom(MAX_RECEIVERS) As Long
    TotalMoves As Long
    Forward_P(MAX_RECEIVERS) As Single
    Back_P(MAX_RECEIVERS) As Single
    Top_Forward As Integer
    Top_Backward As Integer
End Type
Dim Threshold As Single
Dim MarkovChain(MAX_RECEIVERS) As Chain
Dim Track(MAX_RECEIVERS) As Integer
Dim TrackIndex As Integer


Private Sub cmdTrackBack_Click()
Dim i As Long
Dim R As Integer
Dim ID As String
Dim NextR As Long
Dim CheckAsDone(MAX_RECEIVERS, MAX_RECEIVERS) As Boolean

'track forward
Form1.ClearScreen

'get receiver
ID = lstReceivers.Text
R = Receiver.ReceiverNumber(ID)
Receiver.Show R, Form1.Picture1
TrackIndex = 0
Track(0) = R

'go on the chain, link by link
Do
    i = MarkovChain(R).Top_Backward
    If MarkovChain(R).Back_P(i) > Threshold And CheckAsDone(R, i) = False Then
        Receiver.DrawRoute R, i, Form1.Picture1, vbBlack
        CheckAsDone(R, i) = True
        R = i
        TrackIndex = TrackIndex + 1
        Track(TrackIndex) = R
    Else
        R = -1
    End If
Loop Until R = -1
End Sub
Private Sub cmdTrackForward_Click()
Dim i As Long
Dim R As Integer
Dim ID As String
Dim NextR As Long
Dim CheckAsDone(MAX_RECEIVERS, MAX_RECEIVERS) As Boolean

'track forward
Form1.ClearScreen

'get receiver
ID = lstReceivers.Text
R = Receiver.ReceiverNumber(ID)
TrackIndex = 0
Track(0) = R
Receiver.Show R, Form1.Picture1
'go on the chain, link by link
Do
    i = MarkovChain(R).Top_Forward
    If MarkovChain(R).Forward_P(i) > Threshold And CheckAsDone(R, i) = False Then
        Receiver.DrawRoute i, R, Form1.Picture1, vbBlack
        CheckAsDone(R, i) = True
        R = i
        TrackIndex = TrackIndex + 1
        Track(TrackIndex) = R
    Else
        R = -1
    End If
Loop Until R = -1
End Sub
Private Sub ClearChains()
'clears all chains
Dim R As Long
Dim i As Long

For R = 1 To Receiver.TotalReceivers
    With MarkovChain(R)
        .Top_Backward = 0
        .Top_Forward = 0
        .TotalMoves = 0
        For i = 0 To Receiver.TotalReceivers
            .ComingFrom(i) = 0
            .Back_P(i) = 0
            .Forward_P(i) = 0
            .GoingTo(i) = 0
        Next i
    End With
Next R

End Sub
Public Sub RefreshChains()
Dim Selected As String
Dim i As Long
Dim R As Integer

'To refresh chains
'Clear
ClearChains

'& recalculate
CalculateMarkovChains

'remember selection
'and reselect
Selected = lstReceivers.Text
lstReceivers.Clear
LoadReceivers
lstReceivers.Text = Selected
SelectReceiver

'refresh canvas
'Chain
For i = 1 To TrackIndex
    Receiver.DrawRoute Track(i - 1), Track(i), Form1.Picture1, vbBlack
Next i
'or just a single receiver
If TrackIndex = 0 Then
    If lstReceivers.Text <> "" Then
        R = Receiver.ReceiverNumber(lstReceivers.Text)
        Receiver.DrawReceiver Form1.Picture1, R, 1, vbBlack
    End If
End If

End Sub
Private Sub Form_Load()
Dim R As Integer
Dim Found As Boolean

'calculate markov chains upon loading
CalculateMarkovChains
'get receivers into lists
LoadReceivers
'get first receiver in list that is "visible"
Do
    R = R + 1
    If Not Receiver_Table.Invisible(R) Then
        lstReceivers.Text = Receiver.ID(R)
        Found = True
    End If
Loop Until R > Receiver.TotalReceivers Or Found
'Get receiver selected into the list
SelectReceiver
'enable updates to this data
MKWindowLoaded = True
'Threshold is determined by total number of receivers
If Receiver.TotalReceivers > 0 Then Threshold = 1 / Receiver.TotalReceivers
txtThreshold.Text = Format(Threshold, "##.####")
End Sub
Private Sub LoadReceivers()
Dim R As Integer
Dim ID As String

'''''''''''''''''''''''''''''''''''
'NOTE: Need to consider GROUPS
'''''''''''''''''''''''''''''''''''

For R = 1 To Receiver.TotalReceivers
    ID = Receiver.ID(R)
    If Receiver_Table.Invisible(R) = False Then
        lstReceivers.AddItem ID
        MarkovChain(R).Name = ID
    End If
Next R

End Sub
Private Sub CalculateMarkovChains()
Dim Fish As Integer
Dim R As Integer
Dim PrevR As Integer
Dim s As Long
Dim i As Long
Dim MaxF As Single
Dim MaxB As Single

'''''''''''''''''''''''''''''''''''
'NOTE: Need to consider GROUPS
'''''''''''''''''''''''''''''''''''

'Get the MK's from the tracks
'track all fish
'get all P's calculated

'First do the tracking
For Fish = 0 To FishDatabase.TotalFishLoaded
    'For each fish
    For s = 0 To FishDatabase.NumberOfStamps(Fish) - 1
        FishTable.ReadStamp Fish, s
        If Stamp.Valid And Receiver_Table.Invisible(Stamp.Site) = False Then
            PrevR = R
            R = Stamp.Site
            If R <> PrevR Then
                'Get the To
                With MarkovChain(R)
                    .ComingFrom(PrevR) = .ComingFrom(PrevR) + 1
                    .TotalMoves = .TotalMoves + 1
                End With
                'The from
                With MarkovChain(PrevR)
                    .GoingTo(R) = .GoingTo(R) + 1
                    .TotalMoves = .TotalMoves + 1
                End With
            End If
        End If
    Next s
Next Fish



'Then do the P's
For R = 1 To Receiver.TotalReceivers
    With MarkovChain(R)
        If .TotalMoves > 0 Then
            MaxF = 0
            MaxB = 0
            For i = 1 To Receiver.TotalReceivers
                .Forward_P(i) = .GoingTo(i) / .TotalMoves
                .Back_P(i) = .ComingFrom(i) / .TotalMoves
                If .Forward_P(i) > MaxF Then
                    .Top_Forward = i
                    MaxF = .Forward_P(i)
                End If
                If .Back_P(i) > MaxB Then
                    .Top_Backward = i
                    MaxB = .Back_P(i)
                End If
            Next i
        End If
    End With
Next R
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
MKWindowLoaded = False
End Sub

Private Sub lstReceivers_Click()
TrackIndex = 0
SelectReceiver
End Sub
Private Sub SelectReceiver()
Dim R As Integer
Dim ID As String
Dim i As Long
lstTo.Clear
lstFrom.Clear

ID = lstReceivers.Text

R = Receiver.ReceiverNumber(ID)

With MarkovChain(R)
    For i = 1 To Receiver.TotalReceivers
        If .TotalMoves > 0 Then
            If .Forward_P(i) > 0 Then AddToList lstTo, MarkovChain(i).Name, .Forward_P(i)
            If .Back_P(i) > 0 Then AddToList lstFrom, MarkovChain(i).Name, .Back_P(i)
        End If
    Next i
End With

End Sub
Private Sub AddToList(List As ListBox, Name As String, p As Single)
Dim Lp As Long
Dim Ln As Long
Dim N As String
Dim ttl As Long
Dim s As Long

Const FieldLength = 25

Name = Trim(Name)
Ln = Len(Name)
N = Format(p, "##.##")
Lp = Len(N)

If Ln + Lp > FieldLength Then
    Name = Left(Name, FieldLength - Lp)
    Ln = Len(Name)
End If

'format
s = FieldLength - Ln - Lp

List.AddItem Name & String$(s, "-") & N


End Sub
Private Sub txtThreshold_Change()
If txtThreshold <> "" Then
    If IsNumeric(txtThreshold) Then Threshold = CSng(txtThreshold)
End If
End Sub
