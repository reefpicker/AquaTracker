VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFishGroups 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fish Groups"
   ClientHeight    =   6960
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   10290
   Icon            =   "frmFishGroups.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   10290
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstFormattedForCSV 
      Height          =   255
      Left            =   6480
      TabIndex        =   11
      Top             =   4920
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   6480
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdEport 
      Caption         =   "Export as..."
      Height          =   375
      Left            =   7800
      TabIndex        =   10
      Top             =   6480
      Width           =   1215
   End
   Begin MSComctlLib.Slider sldrThreshold 
      Height          =   495
      Left            =   1560
      TabIndex        =   7
      Top             =   6000
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   60
      Min             =   1
      Max             =   1441
      SelStart        =   3
      Value           =   3
   End
   Begin VB.ListBox lstGroupOfFish 
      Height          =   450
      Left            =   5640
      TabIndex        =   4
      Top             =   3720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ListBox lstFish 
      Height          =   450
      Left            =   7200
      TabIndex        =   3
      Top             =   3720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtFish 
      Height          =   1215
      Left            =   5640
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   2400
      Width           =   4455
   End
   Begin VB.ListBox lstGroups 
      Height          =   5325
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   5175
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   9000
      TabIndex        =   0
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label lblThreshold 
      Alignment       =   2  'Center
      Caption         =   "3 Min"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Threshold:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Fish belonging to group:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   6
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Fish Groups:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmFishGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Unload Me

End Sub

Private Sub cmdEport_Click()
'Export fish group analysis
Dim FileNumber As Long
Dim FileName As String
Dim i As Long
Dim Concatenated As String
Dim result As Variant

On Error GoTo ErrFileAccess

With CommonDialog
    .DialogTitle = "Export to a CSV File"
    .CancelError = False
    'ToDo: set the flags and attributes of the common dialog control
    .Filter = "CSV Files (*.csv)|*.csv"
    .ShowOpen
End With

If Len(CommonDialog.FileName) = 0 Then
    Exit Sub
End If

FileName = CommonDialog.FileName

'low level for this
FileNumber = FreeFile

Open FileName For Output As #FileNumber

Print #FileNumber, "Aquatracker (c) by Jose J. Reyes"
Print #FileNumber, " "
Print #FileNumber, "Fish Group Analysis"
Print #FileNumber, "Threshold: " & Str$(sldrThreshold.Value)
Print #FileNumber, " "
Print #FileNumber, "FG, Receiver(s), Number of fish, Detection time (last), Detection date (last), Fish List"

For i = 0 To lstFormattedForCSV.ListCount - 1
    Print #FileNumber, lstFormattedForCSV.List(i)
Next i

Close #FileNumber

Exit Sub

ErrFileAccess:
result = MsgBox("File Access Error", vbOKOnly)
End Sub

Private Sub Form_Load()
Receiver.OutputFishGroups lstGroups, lstFish, lstGroupOfFish, lstFormattedForCSV
End Sub

Private Sub lstGroups_Click()
'show members of group
Dim i As Long

'get selection
For i = 0 To lstGroups.ListCount - 1
    If lstGroups.Selected(i) = True Then
        txtFish.Text = lstFish.List(i)
    End If
Next i
  
End Sub

Private Sub OKButton_Click()
Unload Me

End Sub

Private Sub sldrThreshold_Click()
Dim Threshold As Long

'change threshold value
Threshold = sldrThreshold.Value

'show new value
lblThreshold.Caption = Str$(Threshold) & " Min"

'calculate using new value

Receiver.FindFishGroups Threshold

'erase old groups
lstGroups.Clear
lstFish.Clear
lstGroupOfFish.Clear
lstFormattedForCSV.Clear

'show new groups
Receiver.OutputFishGroups lstGroups, lstFish, lstGroupOfFish, lstFormattedForCSV

End Sub
