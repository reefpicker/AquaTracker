VERSION 5.00
Begin VB.Form frmFishTrackWindow 
   Caption         =   "Fish tracks"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4770
   Icon            =   "frmFishTrackWindow.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6690
   ScaleWidth      =   4770
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmTreatments 
      Caption         =   "By treatments:"
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      Top             =   4800
      Width           =   3135
      Begin VB.ListBox lstTreatments 
         Height          =   1410
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdDeselect 
      Caption         =   "Exclude ALL"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Include ALL"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdInvert 
      Caption         =   "Invert selections"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin VB.ListBox lstFish 
      Height          =   4560
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmFishTrackWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UnderProgramControl As Boolean

Private Sub cmdDeselect_Click()
Dim i As Integer

UnderProgramControl = True

For i = 0 To lstFish.ListCount - 1
    lstFish.Selected(i) = False
    FishDatabase.IsVisible(i) = False
Next i

UnderProgramControl = False
RefreshTracks
End Sub

Private Sub cmdInvert_Click()
Dim i As Integer

UnderProgramControl = True

For i = 0 To lstFish.ListCount - 1
    lstFish.Selected(i) = Not lstFish.Selected(i)
    FishDatabase.IsVisible(i) = lstFish.Selected(i)
Next i

UnderProgramControl = False
RefreshTracks
End Sub

Private Sub cmdSelect_Click()
Dim i As Integer

UnderProgramControl = True

For i = 0 To lstFish.ListCount - 1
    lstFish.Selected(i) = True
    FishDatabase.IsVisible(i) = True
Next i

UnderProgramControl = False
RefreshTracks
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim Treatment As String
Dim N As Long

UnderProgramControl = True
frmFloater.cmbFishCode.ListIndex = 0
RefreshTracks
For i = 1 To frmFloater.cmbFishCode.ListCount - 1
    'get treatment
    Treatment = FishDatabase.Release_Site(i - 1)
    If Trim$(Treatment) <> "" Then StoreTreatmentString Treatment
    'get and display fish code
    lstFish.AddItem frmFloater.cmbFishCode.List(i) & "      " & Treatment
    lstFish.Selected(i - 1) = FishDatabase.IsVisible(i - 1)
Next i
UnderProgramControl = False
If lstTreatments.ListCount = 0 Then frmTreatments.Enabled = False
End Sub
Private Sub StoreTreatmentString(TreatmentString As String)
Dim i As Integer
Dim Found As Boolean

'scan list and if found, don't add
For i = 0 To lstTreatments.ListCount - 1
    If TreatmentString = lstTreatments.List(i) Then
        Found = True
        Exit Sub 'this makes it a bit faster
    End If
Next i

'if found, add!
'store in list as "selected"
If Not Found Then
    lstTreatments.AddItem TreatmentString
    lstTreatments.Selected(lstTreatments.ListCount - 1) = True
End If

End Sub
Private Sub RefreshTracks()
frmFloater.RefreshCanvas
If JPlotIsLoaded Then
    Unload frmJPlot
    'show hour glass
    Form1.MousePointer = vbHourglass
    Form1.StatusBar.Panels(StatusPanel.Map) = "Analyzing receiver and fish databases..."
    frmJPlot.Show
End If
Me.Show
End Sub
Private Sub ClickedOnList()

Dim i As Integer

If Not UnderProgramControl Then
    For i = 0 To lstFish.ListCount - 1
         FishDatabase.IsVisible(i) = lstFish.Selected(i)
    Next i
    RefreshTracks
End If

End Sub

Private Sub lstFish_Click()
ClickedOnList
End Sub

Private Sub lstTreatments_Click()

Dim i As Integer
Dim FishIndex As Integer
Dim Treatment As String

If Not UnderProgramControl Then
    UnderProgramControl = True
    For i = 0 To lstTreatments.ListCount - 1
        Treatment = lstTreatments.List(i)
        For FishIndex = 0 To lstFish.ListCount - 1
            If FishDatabase.Release_Site(FishIndex) = Treatment Then
                FishDatabase.IsVisible(FishIndex) = lstTreatments.Selected(i)
                lstFish.Selected(FishIndex) = lstTreatments.Selected(i)
            End If
        Next FishIndex
    Next i
    frmFloater.cmbFishCode.ListIndex = 0
    RefreshTracks
    UnderProgramControl = False
End If

End Sub
