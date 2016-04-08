VERSION 5.00
Begin VB.Form frmTip 
   Caption         =   "Tool Information"
   ClientHeight    =   2910
   ClientLeft      =   2370
   ClientTop       =   2400
   ClientWidth     =   3960
   Icon            =   "frmTip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   3960
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox picTool 
      BackColor       =   &H00FFFFFF&
      Height          =   2715
      Left            =   120
      Picture         =   "frmTip.frx":0442
      ScaleHeight     =   2655
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.Label lblToolName 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Did you know..."
         Height          =   255
         Left            =   540
         TabIndex        =   2
         Top             =   180
         Width           =   2655
      End
      Begin VB.Label lblToolText 
         BackColor       =   &H00FFFFFF&
         Height          =   1635
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Tips(20) As String

Private Sub Form_Load()
Topics(0) = "Selection tool"
Tips(0) = "The selection tool (arrow) allows you to select a receiver, move a receiver around the map or press SHIFT + LEFT Mouse Button to select multiple receivers"
Topics(1) = "Zoom tool"
Tips(1) = "Zooms in into the map or canvas to show receivers or tracks.  You can use zoom to view receivers that are too close together in the map.  To return to normal view, click on the canvas a second time.  You can display the analysis as on the regular map."
Topics(2) = "Map georeferencing tool"
Tips(2) = "Use the map georeferencing to change the reference points on the map.  Center the crosshairs at known lattitudes/longitudes in the map and click.  Once two points are referenced, the program will update all georeferenced coordinates.  For more information, see AquaTracker's user manual."
Topics(3) = "Measuring tool"
Tips(3) = "On georeference maps and auto-scaled canvases, the measuring tool gives you the distance between two points.  Click on any two points in the canvas and the program draws a line and returns the distance between the points."
Topics(4) = "Reference track"
Tips(4) = "Creates a reference track.  Click on any receiver in the map to connect the receivers into a track.  When you are done, click on the FIND MATCHING TRACKS button to find tracks with similar parameters.  You can also select any real track and make it a reference track by using the canvas context menu (RIGHT-CLICK)."
Topics(5) = "Fish corridors"
Tips(5) = "Define a fish corridor for land avoidance.  To use this tool, draw a line on the map (over water).  Try to trace the line close to the receivers and make sure it follows the path you think the fish would take.  When you are done drawing, take the line over land.  The program will attempt to connect all the receivers to the lines you drew!"
Topics(6) = "Animation capture"
Tips(6) = "Use this tool to create short clips of your animations (in AVI).  Anything on the canvas will be captured, so you can also create videos of your analysis in steps!"
Topics(7) = "Change track color"
Tips(7) = "Choose a track color for the track selected."
TipWindowLoaded = True

'start up position
frmTip.Top = 0
frmTip.Left = Screen.Width - frmTip.Width
lblToolName.Caption = Topics(0)

End Sub

Private Sub Form_Unload(Cancel As Integer)
TipWindowLoaded = False
End Sub

Private Sub lblToolName_Change()
'Grab topic and use to find tip
Dim i As Long
Dim s As Long

Do Until i > 7 Or Topics(i) = lblToolName.Caption
    i = i + 1
Loop

lblToolText.Caption = Tips(i)

If i = 7 Then
    picTool.Picture = frmFloater.picChangeTrackColor.Picture
End If

If i < 7 Then
    picTool.Picture = frmFloater.picTool(i).Picture
End If

End Sub

Private Sub Picture1_Click()

End Sub
