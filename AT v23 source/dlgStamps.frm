VERSION 5.00
Begin VB.Form dlgStamps 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stamps..."
   ClientHeight    =   3435
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4125
   Icon            =   "dlgStamps.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstMsc 
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox lstStamps 
      Height          =   2595
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   3615
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
End
Attribute VB_Name = "dlgStamps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub lstStamps_Click()
Dim i As Long
Dim s As String
Dim response As Variant

For i = 0 To lstStamps.ListCount - 1
    If lstStamps.Selected(i) = True Then
        s = lstMsc.List(i)
        response = MsgBox("Additional info: " & s, vbOKOnly, "Stamp")
    End If
Next i

    
End Sub

Private Sub OKButton_Click()
Unload Me
End Sub
