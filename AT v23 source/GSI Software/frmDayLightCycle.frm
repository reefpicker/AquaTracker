VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmDayLightCycle 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diel Cycle Histogram"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12090
   Icon            =   "frmDayLightCycle.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   273
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   806
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.HScrollBar HScroll_Bin_Size 
      Height          =   135
      LargeChange     =   15
      Left            =   4920
      Max             =   240
      Min             =   5
      SmallChange     =   5
      TabIndex        =   10
      Top             =   600
      Value           =   10
      Width           =   1935
   End
   Begin VB.TextBox txtBinSize 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "10"
      Top             =   240
      Width           =   495
   End
   Begin VB.PictureBox picGraph 
      AutoRedraw      =   -1  'True
      Height          =   2775
      Left            =   720
      ScaleHeight     =   181
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   725
      TabIndex        =   0
      Top             =   1080
      Width           =   10935
   End
   Begin VB.Label Label1 
      Caption         =   "minutes"
      Height          =   255
      Left            =   6240
      TabIndex        =   11
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "Bin Size:"
      Height          =   255
      Left            =   4920
      TabIndex        =   8
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Time of Day (h)"
      Height          =   255
      Left            =   4440
      TabIndex        =   7
      Top             =   3840
      Width           =   3375
   End
   Begin VB.Label lblMid 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape shpColor 
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   4
      Left            =   5640
      Top             =   840
      Width           =   135
   End
   Begin VB.Shape shpColor 
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   3
      Left            =   4440
      Top             =   840
      Width           =   135
   End
   Begin VB.Shape shpColor 
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   2880
      Top             =   840
      Width           =   135
   End
   Begin VB.Shape shpColor 
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   720
      Top             =   840
      Width           =   135
   End
   Begin VB.Label lblMax 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label lblColorPallette 
      Alignment       =   1  'Right Justify
      Caption         =   "PM"
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
      Index           =   4
      Left            =   5640
      TabIndex        =   4
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblColorPallette 
      Alignment       =   1  'Right Justify
      Caption         =   "AM"
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
      Index           =   3
      Left            =   4440
      TabIndex        =   3
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblColorPallette 
      Alignment       =   1  'Right Justify
      Caption         =   "Dawn&&Dusk"
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
      Index           =   2
      Left            =   720
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblColorPallette 
      Alignment       =   1  'Right Justify
      Caption         =   "Night"
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
      Index           =   1
      Left            =   2760
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Visible         =   0   'False
      Begin VB.Menu mnuCopyGraph 
         Caption         =   "Copy Graph"
      End
      Begin VB.Menu mnuCopyData 
         Caption         =   "Copy Data"
      End
      Begin VB.Menu mnuTypeofGraph 
         Caption         =   "Graph type"
         Visible         =   0   'False
         Begin VB.Menu mnuHistogram 
            Caption         =   "Histogram"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuACircle 
            Caption         =   "Circular"
         End
      End
   End
End
Attribute VB_Name = "frmDayLightCycle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const V_Axis_Label = "Detections"
Private Enum ShowAndAnalyze
    Moves
    Stays
    Both
End Enum





Private Sub Form_Load()

DielCycleFormIsLoaded = True
'load proper colors
ImageProcessingEngine.LoadColorPalletteDielWindow
HScroll_Bin_Size.Value = ImageProcessingEngine.DielCycleBinSize
picGraph.BackColor = vbWhite
UpdateHistogram


End Sub
Private Sub Form_Unload(Cancel As Integer)
DielCycleFormIsLoaded = False
End Sub
Private Sub HScroll_Bin_Size_Change()
ImageProcessingEngine.DielCycleBinSize = HScroll_Bin_Size.Value
txtBinSize.Text = Str$(HScroll_Bin_Size.Value)
UpdateHistogram
End Sub
Private Sub lblColorPallette_Click(Index As Integer)
Dim Color_Choosen As Long
On Error GoTo ExitWithError
With CommonDialog
    .CancelError = True
    .ShowColor
    Color_Choosen = .Color
End With

ImageProcessingEngine.ChangeColorPallette(Index) = Color_Choosen
shpColor(Index).FillColor = Color_Choosen
UpdateHistogram

ExitWithError:
'Nop
End Sub

Public Sub UpdateHistogram()
Dim Fish As Long
Dim Pointer As Variant

'clear previous
ImageProcessingEngine.ClearDielCycle

'advise of updating
Form1.MousePointer = vbHourglass
Pointer = Form1.Picture1.MousePointer
Form1.Picture1.MousePointer = vbHourglass
Me.MousePointer = vbHourglass

'process current fish or all fish
If CURRENT_FISH <> -1 Then
    DrawHistogramForFish CURRENT_FISH
    ImageProcessingEngine.DrawDielCycleNOW
Else
    For Fish = 0 To FishDatabase.TotalFishLoaded
        DrawHistogramForFish Fish
    Next Fish
        ImageProcessingEngine.DrawDielCycleNOW
End If

'go back to normal
Form1.MousePointer = vbArrow
Me.MousePointer = vbArrow
Form1.Picture1.MousePointer = Pointer

End Sub
Public Sub DrawHistogramForFish(FishNumber As Long)
Dim i As Long
Dim StaysOnly As Boolean
Dim MovesOnly As Boolean
Dim OldSite As Integer
Dim NewSite As Integer
Dim WhatToAnalyze As Long
Dim DrawThisStamp As Boolean
Dim e As Long

If frmFloater.chkMove = vbChecked Then MovesOnly = True
If frmFloater.chkStay = vbChecked Then StaysOnly = True
If MovesOnly And StaysOnly Then WhatToAnalyze = ShowAndAnalyze.Both
If MovesOnly And Not StaysOnly Then WhatToAnalyze = ShowAndAnalyze.Moves
If StaysOnly And Not MovesOnly Then WhatToAnalyze = ShowAndAnalyze.Stays

'fish number
FishDatabase.Fish = FishNumber

If Not FishDatabase.IsVisible Then Exit Sub

For i = 0 To FishDatabase.NumberOfStamps - 1
    FishTable.ReadStamp FishNumber, i
    If Stamp.Valid Then
        NewSite = Stamp.Site
        DrawThisStamp = False
        Select Case WhatToAnalyze
            Case ShowAndAnalyze.Both
                DrawThisStamp = True
            Case ShowAndAnalyze.Moves
                If OldSite <> NewSite Then DrawThisStamp = True
            Case ShowAndAnalyze.Stays
                If OldSite = NewSite Then DrawThisStamp = True
        End Select
        
        If DrawThisStamp Then
            ImageProcessingEngine.DrawDielCycle False
        End If
        OldSite = NewSite
    End If
Next i

End Sub

Private Sub mnuCopyData_Click()
ImageProcessingEngine.CopyDielCycleToCilpBoard
End Sub

Private Sub mnuCopyGraph_Click()
'copy to clipboard
Clipboard.Clear
Clipboard.SetData picGraph.Image
End Sub

Private Sub picGraph_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'copy menu pops when the right mouse button is pressed
If Button = vbRightButton Then
    PopupMenu mnuEdit
End If

End Sub

Private Sub picGraph_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim H As Integer
Dim m As Integer
Dim s As Integer
Dim N As Long

'get number/time slice
N = ImageProcessingEngine.ReturnTimeOfBin(X, picGraph.Width)

'decompose
H = Fix(N / 60)
m = N Mod 60
picGraph.ToolTipText = TimeSerial(H, m, 0)

End Sub
