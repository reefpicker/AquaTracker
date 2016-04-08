VERSION 5.00
Begin VB.Form frmReceiversResidenceTime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Residence"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5010
   Icon            =   "frmReceiverVisitedTable.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   5010
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstReceiverNumbers 
      Height          =   840
      Left            =   5040
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox lstTable 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3960
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5055
   End
   Begin VB.Label Label1 
      Caption         =   "Fish:"
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
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   495
   End
   Begin VB.Menu mnuCopytoclipboard 
      Caption         =   "Copy"
      Visible         =   0   'False
      Begin VB.Menu mnuCopyData 
         Caption         =   "Copy Data"
      End
   End
End
Attribute VB_Name = "frmReceiversResidenceTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ResidenceTimeAtReceiver(MAX_RECEIVERS) As Long
Public Sub ShowResultsInTable(ReceiversVisited() As Long)
Dim i As Integer
Dim s As String
Dim t As Long
Dim L As Long
Dim p As Long
Dim Item As Long

'cls
lstTable.Clear
lstReceiverNumbers.Clear

'scan
For i = 0 To MAX_RECEIVERS
    ResidenceTimeAtReceiver(i) = ReceiversVisited(i)
    If ReceiversVisited(i) > 0 Then
        s = Receiver.ID(CInt(i))
        t = ReceiversVisited(i)
        L = Len(s)
        If L <= 20 Then
            p = 20 - L
        Else
            p = 1
        End If
        lstTable.AddItem s & Space(p) & t & " mins"
        lstReceiverNumbers.AddItem i
    End If
Next i
End Sub

Private Sub Form_Load()
ResidenceWindowIsLoaded = True
LoadData
End Sub
Public Sub LoadData()
Dim ReceiversVisited(MAX_RECEIVERS) As Long
'clear
lstTable.Clear

'compute new residences
TrackCalculator.ComputeResidence ReceiversVisited, CURRENT_FISH

DrawResidence ReceiversVisited
End Sub
Private Sub Form_Unload(Cancel As Integer)
ResidenceWindowIsLoaded = False
End Sub
Private Sub DrawResidence(RV() As Long)
Dim Max As Long
Dim i As Long

ShowResultsInTable RV

For i = 0 To MAX_RECEIVERS
    If RV(i) > Max Then Max = RV(i)
Next i

ImageProcessingEngine.DrawDensityPlotForReceivers Form1.Picture1, Max, RV, LARGE_MARKER, False

End Sub
Public Sub CopyTableToClipBoard()
Dim i As Long
Dim s As String
s = "Receiver," & Space(20) & "Total residence time" & Chr$(13) & Chr$(10)
For i = 0 To Me.lstTable.ListCount - 1
    s = s & lstTable.List(i) & Chr$(13) & Chr$(10)
Next i

Clipboard.Clear
Clipboard.SetText s

End Sub
Private Sub lstTable_DblClick()
'if user double clicks on an item
'the program will
'take it off from the heat map analysis

Dim i As Long
Dim response As Variant

For i = 0 To lstTable.ListCount - 1
    If lstTable.Selected(i) = True Then
        response = MsgBox("Delete from list and map?", vbOKCancel, "Delete")
        If response = vbOK Then
            ResidenceTimeAtReceiver(CInt(lstReceiverNumbers.List(i))) = 0
            DrawResidence ResidenceTimeAtReceiver
            Unload frmScale
            ColorScale.ShowScale
            Exit Sub
        End If
    End If
Next i

End Sub

Private Sub lstTable_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    PopupMenu mnuCopytoclipboard
End If

End Sub

Private Sub mnuCopyData_Click()
CopyTableToClipBoard
End Sub

