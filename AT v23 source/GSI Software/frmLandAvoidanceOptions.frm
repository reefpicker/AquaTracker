VERSION 5.00
Begin VB.Form frmLandAvoidanceOptions 
   Caption         =   "Land Avoidance Options"
   ClientHeight    =   2625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3795
   Icon            =   "frmLandAvoidanceOptions.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2625
   ScaleWidth      =   3795
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRndWlkP 
      Height          =   285
      Left            =   2640
      TabIndex        =   9
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txtNav_Segment 
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtSEARCHRADIUS 
      Height          =   285
      Left            =   2640
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtSEGMENT_SIZE_THRESHOLD 
      Height          =   285
      Left            =   2640
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Random walk persistance:"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "Connecting segment length:"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Route max. in pixels:"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Minimum length of Valid Segment:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "frmLandAvoidanceOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub
Private Sub cmdOK_Click()
SaveAndRefresh
End Sub
Private Sub SaveAndRefresh()
If SaveVariables Then
    Me.Hide
    frmFloater.RefreshCanvas
    Unload Me
End If
End Sub
Private Function SaveVariables() As Boolean

Dim R As Variant
Dim CancelSave As Boolean
Dim ReturnValue As Boolean

'default, all ok
ReturnValue = True

Nav_Segment = CLng(txtNav_Segment.Text)
If Nav_Segment < 1 Then
    SEGMENT_SIZE_THRESHOLD = 1
    R = MsgBox("Illegal value entered.  Connecting Segment size should not be less than 1 pixels.", vbOKOnly, "Warning")
    CancelSave = True
End If


SEGMENT_SIZE_THRESHOLD = CLng(txtSEGMENT_SIZE_THRESHOLD.Text)
If SEGMENT_SIZE_THRESHOLD < 5 Then
    SEGMENT_SIZE_THRESHOLD = 5
    R = MsgBox("Illegal value entered.  Valid Segment size should not be less than 5 pixels.", vbOKOnly, "Warning")
    CancelSave = True
End If


SEARCHRADIUS = CLng(txtSEARCHRADIUS.Text)
Persistance_TH = CSng(txtRndWlkP.Text)
If Persistance_TH >= 1 Then
    Persistance_TH = 0.9
    R = MsgBox("Illegal value entered.  Persistance is a number between 1 and 0, but can't be 1 or 0.", vbOKOnly, "Warning")
    CancelSave = True
End If

If Persistance_TH <= 0 Then
    R = MsgBox("Illegal value entered.  Persistance is a number between 1 and 0, but can't be 1 or 0.", vbOKOnly, "Warning")
    Persistance_TH = 0.1
    CancelSave = True
End If

If CancelSave Then ReturnValue = False

SaveVariables = ReturnValue

End Function
Private Sub Form_Load()
txtNav_Segment.Text = Str$(Nav_Segment)
txtSEGMENT_SIZE_THRESHOLD.Text = Str$(SEGMENT_SIZE_THRESHOLD)
txtSEARCHRADIUS.Text = Str$(SEARCHRADIUS)
txtRndWlkP.Text = Format(Persistance_TH, "0.#")
End Sub

