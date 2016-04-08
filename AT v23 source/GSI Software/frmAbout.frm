VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AquaTracker: v2.00B"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10860
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   10860
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   60
      Left            =   10080
      Top             =   840
   End
   Begin VB.PictureBox picPlane 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   6840
      Left            =   120
      Picture         =   "frmAbout.frx":08CA
      ScaleHeight     =   452
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   629
      TabIndex        =   1
      Top             =   120
      Width           =   9495
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   9720
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Text_String(20) As String
Dim Bottom As Single
Dim NL As String
Private Sub cmdOK_Click()
Unload Me

End Sub

Private Sub Form_Load()
Dim i As Long
Dim t As String

frmAbout.Caption = "AquaTracker " & Version

NL = Chr$(13) & Chr$(10)

Text_String(0) = "AquaTracker " & Version & " (2009-2015)"
Text_String(1) = "--A program for the visualization and dynamic data exploration of telemetry data--"
Text_String(2) = "Programmed by Jose J. Reyes-Tomassini with video capture code and other enhancements"
Text_String(3) = "from VBAccelarator.com and other sources of public domain code..."
Text_String(4) = NL
Text_String(5) = NL
Text_String(6) = "Created by Jose Reyes-Tomassini and Megan Moore with the BE Team @ NOAA-Manchester."
Text_String(7) = NL
Text_String(8) = NL
Text_String(9) = "Special thanks to F. Goetz & J. Rhodes, my two first outside 'users'"
Text_String(10) = NL
Text_String(11) = "Thanks to the users who keep this project going."
Text_String(12) = NL & "Special thanks to J. McKinney and L. Deegan for spreading the word about AquaTracker."
Text_String(13) = NL
Text_String(14) = NL
Text_String(15) = "For additional information or to report errors contact Jose.ReyesTomassini@noaa.gov."

'picture
picPlane.ForeColor = vbWhite
Bottom = picPlane.ScaleHeight
End Sub

Private Sub picPlane_Click()
Timer1.Enabled = Not Timer1.Enabled
End Sub

Private Sub Timer1_Timer()
Dim i As Single
Dim t As String

Static Line As Long
Static Scroll As Single

picPlane.Cls
picPlane.ForeColor = vbWhite

'write and scroll text
'start at bottom

picPlane.CurrentY = Bottom - Scroll
For i = 0 To 16
    picPlane.Print Text_String(i)
Next i
Scroll = Scroll + 1

If picPlane.CurrentY = 0 Then Scroll = 0


End Sub
