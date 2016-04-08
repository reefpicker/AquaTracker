VERSION 5.00
Begin VB.Form frmSetDayLight 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Day and Night Time"
   ClientHeight    =   2775
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "frmSetDayLight.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtPhotoperiod 
      Height          =   285
      Left            =   3000
      TabIndex        =   14
      Top             =   2370
      Width           =   1575
   End
   Begin VB.TextBox txtNightTo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3000
      TabIndex        =   13
      Text            =   "HH:MM:SS"
      Top             =   1770
      Width           =   1455
   End
   Begin VB.TextBox txtNightFrom 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3000
      TabIndex        =   10
      Text            =   "HH:MM:SS"
      Top             =   1290
      Width           =   1455
   End
   Begin VB.TextBox txtDayTo 
      Height          =   285
      Left            =   720
      TabIndex        =   9
      Text            =   "HH:MM:SS"
      Top             =   1770
      Width           =   1455
   End
   Begin VB.TextBox txtDayFrom 
      Height          =   285
      Left            =   720
      TabIndex        =   6
      Text            =   "HH:MM:SS"
      Top             =   1290
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      Height          =   615
      Left            =   2880
      Picture         =   "frmSetDayLight.frx":0442
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   120
      Picture         =   "frmSetDayLight.frx":0884
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Photoperiod Duration:"
      Height          =   375
      Left            =   1320
      TabIndex        =   15
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Line Line3 
      X1              =   4560
      X2              =   4560
      Y1              =   2160
      Y2              =   0
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   4560
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label6 
      Caption         =   "To:"
      Height          =   255
      Left            =   2400
      TabIndex        =   12
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "From:"
      Height          =   255
      Left            =   2400
      TabIndex        =   11
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "To:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "From:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   495
   End
   Begin VB.Line Line1 
      X1              =   2280
      X2              =   2280
      Y1              =   0
      Y2              =   2160
   End
   Begin VB.Label Label2 
      Caption         =   "Night Period"
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Daylight Period"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
End
Attribute VB_Name = "frmSetDayLight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Unload Me

End Sub

Private Sub cmdUpdate_Click()
Dim D As Long
Dim T As String
Dim Minutes_after_Midnight As Integer
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
T = txtDayTo.Text
txtDayTo.Text = Convert_ToStandardTime(ConvertTime(T))

'convert to minutes after midnight
Minutes_after_Midnight = ConvertTime(T)

'less one minute
Minutes_after_Midnight = Minutes_after_Midnight + 1
If Minutes_after_Midnight > 1439 Then Minutes_after_Midnight = 0

'now this is the night time end
txtNightFrom.Text = Convert_ToStandardTime(Minutes_after_Midnight)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
T = txtDayFrom.Text
txtDayFrom.Text = Convert_ToStandardTime(ConvertTime(T))
'convert to minutes after midnight
Minutes_after_Midnight = ConvertTime(T)

'less one minute
Minutes_after_Midnight = Minutes_after_Midnight - 1
If Minutes_after_Midnight < 0 Then Minutes_after_Midnight = 1439

'now this is the night time end
txtNightTo.Text = Convert_ToStandardTime(Minutes_after_Midnight)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'photoperiod
D = ConvertTime(txtDayTo.Text) - ConvertTime(txtDayFrom.Text)

txtPhotoperiod = Format((D / 60), "#.#") & " HR "

End Sub

Private Sub Form_Load()
txtDayFrom.Text = Convert_ToStandardTime(AM_TH)
txtDayTo.Text = Convert_ToStandardTime(PM_TH)

End Sub

Private Sub OKButton_Click()
AM_TH = ConvertTime(txtDayFrom.Text)
PM_TH = ConvertTime(txtDayTo.Text)

Unload Me

End Sub

Private Function ValidateKey(K As Integer, Text As String) As String
Dim L As Long
Dim S As String
Dim T As String

Dim key_pressed As String

Static Position As Long

T = Text
L = Len(T)

key_pressed = Chr$(K)


If key_pressed = "A" Or key_pressed = "a" Then
    If L > 2 Then
        Text = Left(Text, L - 2) & " AM"
    Else
        Text = "AM"
    End If
End If

If key_pressed = "P" Or key_pressed = "p" Then
    If L > 2 Then
        Text = Left(Text, L - 2) & " PM"
    Else
        Text = "PM"
    End If
End If

If key_pressed >= "0" And key_pressed <= "9" Then
    'move from left to right
    Position = Position + 1
    If Position > 5 Then Position = 1 'go back
    If Position = 3 Then Position = 4 'skip over ":"
    S = Left(T, Position - 1)
    S = S & key_pressed
    Text = S & Right(T, L - Position)
End If

ValidateKey = Text
End Function
