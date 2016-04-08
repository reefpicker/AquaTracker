VERSION 5.00
Begin VB.Form frmMovementAnalysis 
   Caption         =   "Movement and Residence by factor"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14580
   Icon            =   "frmMovementAnalysis.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4815
   ScaleWidth      =   14580
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Define residence"
      Height          =   855
      Left            =   9840
      TabIndex        =   16
      Top             =   1440
      Width           =   4215
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2640
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Fish stays in receiver vicinity for:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Define move"
      Height          =   975
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   8775
      Begin VB.CheckBox Check4 
         Caption         =   "Departure/arrival conditions are commutative"
         Height          =   255
         Left            =   4920
         TabIndex        =   19
         Top             =   600
         Width           =   3615
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Arrival and departure has to occur on same day"
         Height          =   375
         Left            =   4920
         TabIndex        =   15
         Top             =   240
         Width           =   3735
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Fish resides at receiver after arrival"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   3135
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Fish resided at receiver before departing"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.CommandButton cmdPrevious_Residence 
      Caption         =   "<"
      Height          =   255
      Left            =   9840
      TabIndex        =   11
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton cmdNext_Residence 
      Caption         =   ">"
      Height          =   255
      Left            =   13680
      TabIndex        =   10
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton cmdNext_Movement 
      Caption         =   ">"
      Height          =   255
      Left            =   8640
      TabIndex        =   9
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton cmdPrevious_Movement 
      Caption         =   "<"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   4440
      Width           =   495
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   1695
      Left            =   14160
      TabIndex        =   7
      Top             =   2760
      Width           =   255
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1695
      Left            =   9120
      TabIndex        =   6
      Top             =   2760
      Width           =   255
   End
   Begin VB.PictureBox picResidency 
      AutoRedraw      =   -1  'True
      Height          =   1695
      Left            =   9840
      ScaleHeight     =   1635
      ScaleWidth      =   4275
      TabIndex        =   5
      Top             =   2760
      Width           =   4335
   End
   Begin VB.PictureBox picMove 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1695
      Left            =   240
      ScaleHeight     =   1635
      ScaleWidth      =   8835
      TabIndex        =   2
      Top             =   2760
      Width           =   8895
   End
   Begin VB.ListBox lstFactor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   8280
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Residence:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   4
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Movement:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Factor:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmMovementAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdNext_Movement_Click()
FactorMove.NextPage
End Sub

Private Sub cmdNext_Residence_Click()
FactorStay.NextPage
End Sub

Private Sub cmdPrevious_Movement_Click()
FactorMove.PreviousPage
End Sub

Private Sub cmdPrevious_Residence_Click()
FactorStay.PreviousPage
End Sub

Private Sub Form_Load()
'load into list the available factors
'right now two: tides and photoperiod

lstFactor.AddItem "Photoperiod"
lstFactor.AddItem "Tides"

'set windows to white
picMove.BackColor = vbWhite
picResidency.BackColor = vbWhite

End Sub

Private Sub lstFactor_Click()
'0=diel
'1=tides
Dim i As Long
Dim Factor As Long

For i = 0 To lstFactor.ListCount - 1
    If lstFactor.Selected(i) Then
        Factor = i
    End If
Next i

'based on factor choosen, open windows and get analysis ready
If Factor = 0 Then
    'set devices
    FactorMove.Select_Device Device_Type.Window, picMove
    FactorStay.Select_Device Device_Type.Window, picResidency
    'setup calculator to do it
    TrackCalculator.AnalyzeDielMoves = True
    'load diel window
    frmDayLightCycle.Show
End If

End Sub

Private Sub VScroll1_Change()
Dim S As Long
S = VScroll1.Value
If S > 0 Then FactorMove.Scroll VScroll1.Value
End Sub

Private Sub VScroll2_Change()
Dim S As Long
S = VScroll2.Value
If S > 0 Then FactorStay.Scroll VScroll2.Value
End Sub
