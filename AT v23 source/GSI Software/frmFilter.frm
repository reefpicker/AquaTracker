VERSION 5.00
Begin VB.Form frmFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Singleton detections"
   ClientHeight    =   4560
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5100
   Icon            =   "frmFilter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDeleteSingleStamp 
      Caption         =   "Delete selected"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdPurgeAll 
      Caption         =   "Purge all..."
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   4080
      Width           =   1335
   End
   Begin VB.ListBox lstReceiversToPurge 
      Height          =   4350
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ReceiverList(MAX_ENTRIES) As Integer
Dim EntryList(MAX_ENTRIES) As Long
Dim FishList(MAX_ENTRIES) As Integer
Dim TrackStringList(MAX_FISH) As String

Option Explicit

Private Sub CancelButton_Click()
Unload Me
End Sub
Private Sub cmdDeleteSingleStamp_Click()
'this will mark stamp as not valid
'warn user first

Dim response As Variant
Dim StampVerbose As String
Dim i As Integer
Dim ChangedStrings As String

GetTrackStrings

With lstReceiversToPurge
    For i = 0 To .ListCount - 1
        If .Selected(i) = True Then
            StampVerbose = .List(i)
            response = MsgBox("Are you sure you want to mark this stamp as invalid (deleted)?", vbYesNo, StampVerbose)
            If response = vbYes Then
                DeleteSingleton i
            End If
            'now redo list
            .Clear
            Receiver.ShowSingletons 1440, lstReceiversToPurge, ReceiverList(), EntryList(), FishList()
            Exit For
        End If
    Next i
End With
    
ChangedStrings = CheckTrackStrings
If ChangedStrings <> "" Then
    response = MsgBox("The following tracks have changed because of the deletion of singletons: " & ChangedStrings, vbOKOnly, "Warning")
End If

End Sub
Private Sub DeleteSingleton(SingletonStampNumber As Integer)
Dim t As Long
Dim s As Long
Dim Found As Boolean
'read stamp to get the time
Receiver.ReadStamp ReceiverList(SingletonStampNumber), EntryList(SingletonStampNumber)
t = Stamp.CTime
'delete receiver stamp first, its easy
ReceiverTable.DeleteStamp ReceiverList(SingletonStampNumber), EntryList(SingletonStampNumber)
'then delete stamp in fish list
Do
    FishTable.ReadStamp FishList(SingletonStampNumber), s
    If Stamp.CTime = t Then
        FishTable.DeleteStamp FishList(SingletonStampNumber), s
        Found = True
    End If
    s = s + 1
Loop Until Found Or s >= FishDatabase.NumberOfStamps(FishList(SingletonStampNumber))
End Sub
Private Sub cmdPurgeAll_Click()
Dim i As Integer
Dim ChangedStrings As String
Dim response As Variant

response = MsgBox("Are you sure you want to eliminate ALL of the singletons in the list?", vbYesNo, "Delete all singletons")

If response = vbYes Then
    GetTrackStrings
    'deletes all singletons in list
    With lstReceiversToPurge
        For i = 0 To .ListCount - 1
            DeleteSingleton i
        Next i
    End With
    lstReceiversToPurge.Clear
    ChangedStrings = CheckTrackStrings
    If ChangedStrings <> "" Then
        response = MsgBox("The following tracks have changed because of the deletion of singletons: " & ChangedStrings, vbExclamation, "Warning")
    End If
    
    'unload
    Unload Me
End If

End Sub
Private Function CheckTrackStrings() As String
Dim ReturnString As String
Dim F As Long

For F = 0 To FishDatabase.TotalFishLoaded
    If TrackStringList(F) <> TrackCalculator.CreateVerboseTrackString(F) Then
        ReturnString = ReturnString & FishDatabase.Code(F) & ","
    End If
Next F

CheckTrackStrings = ReturnString

End Function
Private Sub GetTrackStrings()
Dim F As Long

For F = 0 To FishDatabase.TotalFishLoaded
    TrackStringList(F) = TrackCalculator.CreateVerboseTrackString(F)
Next F
End Sub
Private Sub Form_Load()
Receiver.ShowSingletons 1440, lstReceiversToPurge, ReceiverList(), EntryList(), FishList()
End Sub

Private Sub lstReceiversToPurge_Click()
Dim i As Integer

For i = 0 To lstReceiversToPurge.ListCount - 1
    If lstReceiversToPurge.Selected(i) = True Then
        Form1.ClearScreen
        ShowTrack FishList(i), Form1.Picture1
        Receiver.DrawReceiver Form1.Picture1, ReceiverList(i), 2, HighLightColor
    End If
Next i

    
End Sub

Private Sub OKButton_Click()
Unload Me
End Sub
