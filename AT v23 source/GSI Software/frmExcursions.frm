VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmExcursions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Excursions from Receiver "
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8130
   Icon            =   "frmExcursions.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   8130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   6840
      TabIndex        =   3
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   6120
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   1920
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   5775
      Left            =   7800
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picWindow 
      AutoRedraw      =   -1  'True
      Height          =   5775
      Left            =   240
      ScaleHeight     =   5715
      ScaleWidth      =   7515
      TabIndex        =   0
      Top             =   120
      Width           =   7575
   End
End
Attribute VB_Name = "frmExcursions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Output As New clsGenericIO
Private Sub Picture1_Click()

End Sub

Private Sub cmdClose_Click()
Unload Me

End Sub

Private Sub cmdExport_Click()
Dim FileName As String

With CommonDialog
    .DialogTitle = "Save as... CSV File"
    .CancelError = False
    'ToDo: set the flags and attributes of the common dialog control
    .Filter = "CSV Files (*.csv)|*.csv"
    .ShowOpen
End With

If Len(CommonDialog.FileName) = 0 Then
    Exit Sub
End If

FileName = CommonDialog.FileName

Excursion.ExportExcursion Receiver.CurrentStation_Number, FileName
End Sub

Private Sub Form_Load()
Dim Station As Long
Dim LastEntry As Long

'
picWindow.Cls
picWindow.BackColor = vbWhite
picWindow.Print "Processing..."
picWindow.Cls
Excursion.SetOutputAs = Output

'show excursions for receiver
Excursion.ShowExcursions CLng(Receiver.CurrentStation_Number), picWindow
'Show scroll bar if needed and adjust its max value depending on buffer
If Output.EndOfPage Then VScroll1.Visible = True
LastEntry = Output.LastEntryInBuffer
VScroll1.Max = LastEntry
If LastEntry > 1 Then cmdExport.Enabled = True

VScroll1.LargeChange = 27
VScroll1.SmallChange = 2
''''''''''''''''''''''''''''''''''''''
Station = Receiver.CurrentStation_Number
Me.Caption = Receiver.ID(CLng(Station))

Output.Scroll 1

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Output = Nothing

End Sub

Private Sub VScroll1_Change()
Output.Scroll VScroll1.Value

End Sub
