VERSION 5.00
Begin VB.Form frmClock 
   AutoRedraw      =   -1  'True
   Caption         =   "Stamp Timeline"
   ClientHeight    =   1410
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10560
   Icon            =   "frmClock.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   94
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   704
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTimeLine 
      AutoRedraw      =   -1  'True
      Height          =   180
      Left            =   0
      ScaleHeight     =   8
      ScaleMode       =   0  'User
      ScaleWidth      =   700
      TabIndex        =   2
      Top             =   840
      Width           =   10500
   End
   Begin VB.PictureBox picThis 
      AutoRedraw      =   -1  'True
      Height          =   300
      Left            =   0
      ScaleHeight     =   16
      ScaleMode       =   0  'User
      ScaleWidth      =   700
      TabIndex        =   1
      Top             =   1080
      Width           =   10500
   End
   Begin VB.TextBox txtTime 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "12:25PM"
      Top             =   120
      Width           =   3495
   End
   Begin VB.Menu mnuContextMenu 
      Caption         =   "Context"
      Visible         =   0   'False
      Begin VB.Menu mnuCopyPic 
         Caption         =   "Copy Timeline Image"
      End
   End
End
Attribute VB_Name = "frmClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

LoadCurrentFish

End Sub
Public Sub LoadCurrentFish()
Dim FishNumber As Long
Dim NumberOfStamps As Long
Dim i As Long
Dim Spacing As Long
Dim NumberOfDays As Long

Dim X As Long
Dim FirstDay As Long
Dim LastDay As Long

Const Length = 700

'load dates into text box

'use fish number
FishNumber = FishDatabase.Fish
NumberOfStamps = FishDatabase.NumberOfStamps

'validate
If TotalTime = 0 Then Exit Sub

'calculate numberofdays
NumberOfDays = (TotalTime / 1440)

'calculate spacing b/w days
If NumberOfDays = 0 Then NumberOfDays = 1
Spacing = Length / NumberOfDays
If Spacing <= 0 Then Spacing = 1

'clear
picThis.Cls

'draw days
For i = 1 To NumberOfDays
    X = X + Spacing
    picThis.Line (X, 0)-(X, 20), vbRed
Next i
End Sub

Private Sub mnuCopyPic_Click()
Clipboard.Clear
Clipboard.SetData picThis.Image
End Sub

Private Sub picThis_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then PopupMenu mnuContextMenu
End Sub
