VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmDisplayDetails 
   Caption         =   "Fish Track Details"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10365
   Icon            =   "frmDisplayDetails.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5370
   ScaleWidth      =   10365
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll 
      Height          =   3015
      Left            =   10080
      TabIndex        =   12
      Top             =   1440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export as..."
      Height          =   375
      Left            =   7320
      TabIndex        =   11
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Height          =   255
      Left            =   9600
      TabIndex        =   9
      Top             =   4440
      Width           =   495
   End
   Begin VB.PictureBox picShowInfo 
      AutoRedraw      =   -1  'True
      Height          =   3015
      Left            =   120
      ScaleHeight     =   2955
      ScaleWidth      =   9915
      TabIndex        =   8
      Top             =   1440
      Width           =   9975
   End
   Begin VB.CommandButton cmdCloseMe 
      Caption         =   "Close"
      Default         =   -1  'True
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
      Left            =   8640
      TabIndex        =   6
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox txtReceiver 
      Height          =   285
      Left            =   5520
      TabIndex        =   2
      Top             =   240
      Width           =   2415
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Left            =   5520
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox txtTime 
      Height          =   285
      Left            =   5520
      TabIndex        =   0
      Top             =   960
      Width           =   2415
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   9120
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "Current Track Point"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   7
      Top             =   720
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   4
      Height          =   1215
      Left            =   4680
      Top             =   120
      Width           =   3375
   End
   Begin VB.Line Line2 
      BorderWidth     =   4
      X1              =   4680
      X2              =   3600
      Y1              =   1320
      Y2              =   840
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      X1              =   4680
      X2              =   3600
      Y1              =   120
      Y2              =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Receiver:"
      Height          =   255
      Left            =   4800
      TabIndex        =   5
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Time:"
      Height          =   255
      Left            =   4920
      TabIndex        =   4
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Date:"
      Height          =   255
      Left            =   4920
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "frmDisplayDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub tmrAnimation_Timer()

End Sub

Private Sub cmdCloseMe_Click()
Me.Hide


End Sub

Private Sub cmdExport_Click()
'Export to .CSV file
Dim FileName As String

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

DeviceBuffer.WriteBuffer_to_File FileName

picShowInfo.Cls
Unload Me

End Sub

Private Sub cmdNext_Click()
DeviceBuffer.NextPage
End Sub

Private Sub cmdPrevious_Click()
DeviceBuffer.PreviousPage
End Sub

Private Sub Form_Load()
'prep window
With picShowInfo
    .Cls
    .BackColor = vbWhite
    .ForeColor = vbBlack
    .FontBold = True
End With
End Sub

Private Sub VScroll1_Change()

End Sub

Private Sub VScroll_Change()
DeviceBuffer.Scroll VScroll.Value
End Sub
