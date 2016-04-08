VERSION 5.00
Begin VB.Form frmNavigationString 
   Caption         =   "Navigation String"
   ClientHeight    =   2760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6585
   Icon            =   "frmNavigationString.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2760
   ScaleWidth      =   6585
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNavigationString 
      Height          =   2775
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   6615
   End
End
Attribute VB_Name = "frmNavigationString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ReDrawBefore As Boolean
Dim ReceiversSelected(MAX_TRACK_STRING) As Integer
Dim Movements(MAX_TRACK_STRING) As Integer
Dim PreviousString As String
Dim Form_Height As Single
Dim Form_Width As Single
Dim Adjusting_Window As Boolean
Private Sub Form_Load()
Adjusting_Window = True
'store window size to be used in rescaling drawing area
Form_Height = frmNavigationString.Height
Form_Width = frmNavigationString.Width
ReDrawBefore = True
Adjusting_Window = False

'me is loaded!
TrackStringWindowIsLoaded = True

'focus back on window
frmFloater.SetFocus

End Sub


Private Function GetReceiversSelected(p As Integer, L As Integer) As Integer
Dim i As Integer
Dim R As Integer
Dim Previous As Integer
Dim count As Integer
Dim N As Long

'track arrows = -1 or -2
'
Previous = -1

For i = p To p + L - 1
    R = TrackCalculator.ReceiverInTrackString(i)
    If R = -1 Or R = -2 Then
        Movements(N) = R
        N = N + 1
    Else
        If R <> Previous Then
            ReceiversSelected(count) = R
            count = count + 1
        End If
        Previous = R
    End If
Next i

GetReceiversSelected = count
End Function
Private Sub ShowSelection()
'translates selected text into a track to be highlighted
'highlight color default is => RED


Dim Segment As String
Dim TrackString As String
Dim NumberOfReceiversSelected As Long
Dim FirstReceiver As Long
Dim TranslatedString As String
Dim PositionInString As Integer
Dim Separator As String
Dim Selection As String
Dim SelectionLen As Integer
Dim R As Integer
Dim i As Long


Static NumberOfReceiversBefore As Long


'get string
TrackString = txtNavigationString.Text
PositionInString = txtNavigationString.SelStart
Selection = txtNavigationString.SelText
SelectionLen = Len(Selection)

'validate and
'Get relative position in this string
If PositionInString <= Len(TrackString) And Len(Selection) > 1 Then
    'and get how many receivers are selected
    NumberOfReceiversSelected = GetReceiversSelected(PositionInString, SelectionLen)
    
    'no receivers, exit and paint if before a receiver was selected
    If NumberOfReceiversSelected = 0 Then
        If NumberOfReceiversBefore <> 0 Then frmFloater.RefreshCanvas
        Exit Sub
    End If
    'if only 1 receiver, visualize it
    If NumberOfReceiversSelected = 1 Then
        If ReDrawBefore Then frmFloater.RefreshCanvas
        Receiver.DrawReceiver Form1.Picture1, ReceiversSelected(0), 2, HighLightColor
    Else
        'draw the track segment
        If ReDrawBefore Then frmFloater.RefreshCanvas
        For R = 1 To NumberOfReceiversSelected - 1
            Receiver.DrawRoute ReceiversSelected(R), ReceiversSelected(R - 1), Form1.Picture1, HighLightColor
            'bidirectional arrow?
            If Movements(R - 1) = -2 Then Receiver.DrawRoute ReceiversSelected(R - 1), ReceiversSelected(R), Form1.Picture1, HighLightColor
        Next R
    End If
End If

NumberOfReceiversBefore = NumberOfReceiversSelected

End Sub
Private Sub CheckIfNew()
Dim p As Long
Dim Selection As String

Selection = txtNavigationString.SelText

If PreviousString = "" Or Selection = "" Then
    ReDrawBefore = True
Else
    If InStr(1, Selection, PreviousString) And Len(Selection) >= Len(PreviousString) Then
        ReDrawBefore = False
    Else
        ReDrawBefore = True
    End If
End If

PreviousString = Selection

End Sub

Private Sub Form_Resize()
Dim new_width As Single
Dim new_height As Single

Dim DeltaX As Single
Dim DeltaY As Single


'check if loaded
If Adjusting_Window Then Exit Sub


DeltaX = (frmNavigationString.Width - Form_Width)
DeltaY = (frmNavigationString.Height - Form_Height)


new_width = frmNavigationString.txtNavigationString.Width + DeltaX
new_height = frmNavigationString.txtNavigationString.Height + DeltaY

'validate
If new_width <= 0 Then
    new_width = 1
End If

If new_height <= 0 Then
    new_height = 1
End If

frmNavigationString.txtNavigationString.Width = new_width
frmNavigationString.txtNavigationString.Height = new_height
Form_Width = frmNavigationString.Width
Form_Height = frmNavigationString.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
TrackStringWindowIsLoaded = False
End Sub

Private Sub txtNavigationString_Click()
If txtNavigationString.SelText = "" Then frmFloater.RefreshCanvas
End Sub
Private Sub txtNavigationString_KeyUp(KeyCode As Integer, Shift As Integer)
CheckIfNew
ShowSelection
End Sub
Private Sub txtNavigationString_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    CheckIfNew
    ShowSelection
End If
End Sub
