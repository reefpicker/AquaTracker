VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmChangeMarkerForReceiver 
   Caption         =   "Choose Marker"
   ClientHeight    =   2250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8340
   Icon            =   "frmChangeMarkerForReceiver.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2250
   ScaleWidth      =   8340
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdColorAllReceiversUsingDistanceGradient 
      Caption         =   "Distance heat-map"
      Height          =   375
      Left            =   3840
      TabIndex        =   25
      Top             =   1680
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   6480
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Color"
      Height          =   2055
      Left            =   3240
      TabIndex        =   11
      Top             =   120
      Width           =   3135
      Begin VB.CommandButton Command1 
         Caption         =   "More colors..."
         Height          =   375
         Left            =   840
         TabIndex        =   24
         Top             =   1200
         Width           =   1455
      End
      Begin VB.PictureBox picColor 
         Height          =   375
         Index           =   11
         Left            =   2520
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   23
         Top             =   720
         Width           =   375
      End
      Begin VB.PictureBox picColor 
         Height          =   375
         Index           =   10
         Left            =   2520
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   22
         Top             =   240
         Width           =   375
      End
      Begin VB.PictureBox picColor 
         Height          =   375
         Index           =   9
         Left            =   2040
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   21
         Top             =   720
         Width           =   375
      End
      Begin VB.PictureBox picColor 
         Height          =   375
         Index           =   8
         Left            =   1560
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   20
         Top             =   720
         Width           =   375
      End
      Begin VB.PictureBox picColor 
         Height          =   375
         Index           =   7
         Left            =   1080
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   19
         Top             =   720
         Width           =   375
      End
      Begin VB.PictureBox picColor 
         Height          =   375
         Index           =   6
         Left            =   600
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   18
         Top             =   720
         Width           =   375
      End
      Begin VB.PictureBox picColor 
         Height          =   375
         Index           =   5
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   17
         Top             =   720
         Width           =   375
      End
      Begin VB.PictureBox picColor 
         Height          =   375
         Index           =   4
         Left            =   2040
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   16
         Top             =   240
         Width           =   375
      End
      Begin VB.PictureBox picColor 
         Height          =   375
         Index           =   3
         Left            =   1560
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   15
         Top             =   240
         Width           =   375
      End
      Begin VB.PictureBox picColor 
         Height          =   375
         Index           =   2
         Left            =   1080
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   14
         Top             =   240
         Width           =   375
      End
      Begin VB.PictureBox picColor 
         Height          =   375
         Index           =   1
         Left            =   600
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   13
         Top             =   240
         Width           =   375
      End
      Begin VB.PictureBox picColor 
         Height          =   375
         Index           =   0
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   12
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton Command4 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7080
      TabIndex        =   10
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   7080
      TabIndex        =   9
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Apply to all..."
      Height          =   375
      Left            =   7080
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox picMarkerPreview 
      AutoRedraw      =   -1  'True
      Height          =   855
      Left            =   120
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   7
      Top             =   720
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Marker"
      Height          =   2055
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   1815
      Begin VB.OptionButton optMarker 
         Caption         =   "Diamond"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   1575
      End
      Begin VB.OptionButton optMarker 
         Caption         =   "Inverted triangle"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   1575
      End
      Begin VB.OptionButton optMarker 
         Caption         =   "Triangle"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton optMarker 
         Caption         =   "Rectangle"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton optMarker 
         Caption         =   "Circle"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Preview:"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "frmChangeMarkerForReceiver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ReceiverNumber As Integer
Dim LastBox As Integer
Dim OldColor As Long
Dim OldMarkerType As Long
Private Sub cmdColorAllReceiversUsingDistanceGradient_Click()
'colors all receivers with the same color using a gradient based on distance from THIS receiver
'DEFAULT: Darker, the farther....
'This is a Euclidean distance, for other distance measures, change formula

Dim X As Long
Dim Y As Long
Dim R As Integer

Dim GradientStep As Long
Dim range As Single
Dim Min As Single
Dim Max As Single
Dim Distance As Single
Dim R_Distance(MAX_RECEIVERS)

Dim Color As Long
Dim intensity As Long

Dim response As Variant

response = MsgBox("Are you sure you want to use " & Receiver.ID(ReceiverNumber) & " to generate the distance heat map? All visible receivers will be changed!", vbYesNo, "Generate heat map")
If response = vbNo Then Exit Sub

'location of this receiver
X = Receiver.X(ReceiverNumber)
Y = Receiver.Y(ReceiverNumber)

'clear
Max = 0
Min = 1500 'max distance for a 1024 x 1024 map/canvas

'Get euclidean distance from this receivers to all other receivers
For R = 1 To Receiver.TotalReceivers
    R_Distance(R) = Sqr((X - Receiver.X(R)) ^ 2 + (Y - Receiver.Y(R)) ^ 2)
    If R_Distance(R) > Max Then Max = R_Distance(R)
    If R_Distance(R) < Min Then Min = R_Distance(R)
Next R

'get range
range = Max - Min
If range > 255 Then range = 255

'get steps
GradientStep = 255 / range

'color receivers
For R = 1 To Receiver.TotalReceivers
    intensity = (R_Distance(R) - Min) * GradientStep
    If intensity > 255 Then intensity = 255
    intensity = 255 - intensity
    Color = RGB(intensity, intensity, intensity)
    If Receiver.Detection_TTL(R) <> 0 Then Receiver.Color(R) = Color 'only colors "visible" receivers
Next R

'Make sure receiver gets the "origin" color, which by default is BLUE
Receiver.Color(ReceiverNumber) = vbBlue

Form1.ShowAllReceivers
ColorMarker
End Sub

Private Sub Command1_Click()
On Error GoTo ExitWithError
With CommonDialog
    .CancelError = True
    .ShowColor
    Receiver.Color(ReceiverNumber) = .Color
End With

'show
'draw it
picMarkerPreview.Cls
Receiver.DrawReceiver picMarkerPreview, ReceiverNumber, 2, -1, CLng(picMarkerPreview.ScaleWidth / 2), CLng(picMarkerPreview.ScaleHeight / 2)

ExitWithError:
'NOP
End Sub

Private Sub Command2_Click()
'applies style and color of current selection to ALL RECEIVERS in database
Dim UserResponse As Variant
Dim CurrentReceiverColor As Long
Dim CurrentReceiverMarker As Long
Dim R As Integer

'warn about whats gonna happen
UserResponse = MsgBox("Are you sure you want to change the marker type and color for all receivers?", vbOKCancel, "Apply current style to all receivers")
If UserResponse <> vbCancel Then
    'get current values
    CurrentReceiverColor = Receiver.Color(ReceiverNumber)
    CurrentReceiverMarker = Receiver_Table.Marker(ReceiverNumber)
    
    'propagate to all receivers
    For R = 1 To Receiver.TotalReceivers
        Receiver.Color(R) = CurrentReceiverColor
        Receiver_Table.Marker(R) = CurrentReceiverMarker
    Next R
End If
ColorMarker
End Sub
Private Sub Command3_Click()
ColorMarker
End Sub
Private Sub ColorMarker()
On Error GoTo ExitWithError
Me.MousePointer = vbHourglass
'show new color on info window
frmReceiverInformation.picMarker.Cls
Receiver.DrawReceiver frmReceiverInformation.picMarker, ReceiverNumber, LARGE_MARKER, -1, CLng(frmReceiverInformation.picMarker.ScaleWidth / 2), CLng(frmReceiverInformation.picMarker.ScaleHeight / 2)
Receiver.DrawReceiver Form1.Picture1, ReceiverNumber
frmReceiverInformation.picRelativeLocation.Refresh
'if window not there, safely exit
ExitWithError:
Me.MousePointer = vbArrow
Unload Me
End Sub
Private Sub Command4_Click()
'reverse changes to marker

'restore
Receiver.Color(ReceiverNumber) = OldColor
Receiver_Table.Marker(ReceiverNumber) = OldMarkerType

'unload for
Unload Me

End Sub

Private Sub Form_Load()
Dim i As Integer
picMarkerPreview.BackColor = vbWhite

'load color pallette
LastBox = CInt(MaxPal)
If LastBox > picColor.UBound Then LastBox = picColor.UBound

For i = 0 To LastBox
    picColor(i).BackColor = ColorPal(i)
Next i

End Sub

Public Sub LoadReceiverMarker(R As Integer)
Dim i As Integer
'loads info for receiver
ReceiverNumber = R

'store info
OldColor = Receiver.Color(ReceiverNumber)
OldMarkerType = Receiver_Table.Marker(ReceiverNumber)

'get color (exact match)
Do
    If picColor(i).BackColor = Receiver.Color(ReceiverNumber) Then
        picColor(i).BorderStyle = 1
    Else
        picColor(i).BorderStyle = 0
    End If
    i = i + 1
Loop Until i > LastBox

'marker type
For i = 0 To optMarker.UBound
    optMarker(i).Value = False
Next i
optMarker(CInt(Receiver_Table.Marker(R))).Value = True

'draw it
picMarkerPreview.Cls
Receiver.DrawReceiver picMarkerPreview, ReceiverNumber, 2, -1, CLng(picMarkerPreview.ScaleWidth / 2), CLng(picMarkerPreview.ScaleHeight / 2)

End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.RePaint
End Sub

Private Sub optMarker_Click(Index As Integer)
'assign
Receiver_Table.Marker(ReceiverNumber) = Index

'draw it
picMarkerPreview.Cls
Receiver.DrawReceiver picMarkerPreview, ReceiverNumber, 2, -1, CLng(picMarkerPreview.ScaleWidth / 2), CLng(picMarkerPreview.ScaleHeight / 2)

End Sub

Private Sub picColor_Click(Index As Integer)
Dim i As Integer

'deselect
For i = 0 To picColor.UBound
    picColor(i).BorderStyle = 0
Next i

'select
picColor(Index).BorderStyle = 1

'assign
Receiver.Color(ReceiverNumber) = picColor(Index).BackColor


'draw it
picMarkerPreview.Cls
Receiver.DrawReceiver picMarkerPreview, ReceiverNumber, 2, -1, CLng(picMarkerPreview.ScaleWidth / 2), CLng(picMarkerPreview.ScaleHeight / 2)

End Sub
