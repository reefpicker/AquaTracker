VERSION 5.00
Begin VB.Form frmShowPings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List of stamps"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7215
   Icon            =   "frmShowPings.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar vscStamps 
      Height          =   4095
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox picPings 
      AutoRedraw      =   -1  'True
      FontTransparent =   0   'False
      Height          =   4095
      Left            =   360
      ScaleHeight     =   269
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   445
      TabIndex        =   0
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "frmShowPings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NextStart As Long
Dim FishNumber As Long
Dim CurrentStart As Long
Dim TotalPings As Long
Dim ScrollBarScale As Integer
Dim ping As Long
Private Function Total_Detections() As Long
'
'gets detection total including total by fish!
Dim return_value As Long
Dim R As Integer
Dim ttl As Long
Dim s As Long

R = Receiver.CurrentStation_Number

If FishNumber = -1 Then
    return_value = Receiver.Detection_Total(R)
Else
    ttl = Receiver.Detection_Total(R)
    For s = 0 To ttl
        Receiver.ReadStamp R, s
        If Stamp.Fish = FishNumber Then return_value = return_value + 1
    Next s
End If

Total_Detections = return_value
End Function
Private Sub Form_Load()
Dim LastDetection As Long
CurrentStart = 1

If Form1.Tag = "" Then
    FishNumber = -1
Else
    FishNumber = CLng(Form1.Tag)
End If

LastDetection = Total_Detections

ping = 32767
If LastDetection > ping Then LastDetection = ping

'scroll bar
With vscStamps
    .Max = LastDetection
    .Min = 1
    If .Max > 18 Then
        .LargeChange = (Fix(.Max / 18)) + 1
        .SmallChange = 1
    Else
        .LargeChange = .Max - 1
        .SmallChange = .Max - 1
    End If
End With

End Sub
Private Sub ChangePage()
Dim p As Long

'work around problem of using the vscroll with a max that is not an integer
If vscStamps.Value = 32767 Then
    ping = ping + 1
    p = ping
Else
    ping = 32767
    p = vscStamps.Value
End If

NextStart = Receiver.ShowAllPings(Receiver.CurrentStation_Number, picPings, FishNumber, p)

End Sub
Private Sub picPings_Paint()
Dim dummy As Long
dummy = Receiver.ShowAllPings(Receiver.CurrentStation_Number, picPings, FishNumber, CurrentStart)
End Sub

Private Sub vscStamps_Change()
ChangePage
End Sub
