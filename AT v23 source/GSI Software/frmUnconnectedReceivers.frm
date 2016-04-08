VERSION 5.00
Begin VB.Form frmUnconnectedReceivers 
   Caption         =   "Warning"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2805
   Icon            =   "frmUnconnectedReceivers.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4455
   ScaleWidth      =   2805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   3840
      Width           =   1455
   End
   Begin VB.ListBox lstUnconnectedReceivers 
      Height          =   2985
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "The following receivers are not connected to a corridor:"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmUnconnectedReceivers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub lstUnconnectedReceivers_Click()
Dim i As Long
Dim X As Long
Dim Y As Long
Dim R As Integer

For i = 0 To lstUnconnectedReceivers.ListCount - 1
    With lstUnconnectedReceivers
        If .Selected(i) Then
            Form1.ClearScreen
            Form1.ShowDetectors
            ImageProcessingEngine.DrawReceiverConnectionsToCorridor
            R = CInt(.ItemData(i))
            X = Receiver.X(R)
            Y = Receiver.Y(R)
            Form1.Picture1.Circle (X, Y), 2, vbRed
            Form1.Picture1.Circle (X, Y), 3, vbRed
            Form1.Picture1.Circle (X, Y), 4, vbRed
        End If
    End With
Next i

End Sub
