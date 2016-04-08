VERSION 5.00
Begin VB.Form frmSelectReceivers 
   Caption         =   "Receivers"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2385
   Icon            =   "frmSelectReceivers.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5730
   ScaleWidth      =   2385
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdInvert 
      Caption         =   "Invert selections"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Include ALL"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton cmdDeselect 
      Caption         =   "Exclude ALL"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   4560
      Width           =   1335
   End
   Begin VB.ListBox lstReceivers 
      Height          =   4335
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmSelectReceivers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UnderProgramControl As Boolean
Dim ReceiverNumber(MAX_RECEIVERS) As Integer
Private Sub cmdDeselect_Click()
Dim R As Long
UnderProgramControl = True
For R = 0 To lstReceivers.ListCount - 1
    lstReceivers.Selected(R) = False
Next R
Receiver.MakeInvisible

'update screen
Form1.RePaint
UnderProgramControl = False
RefreshJPlot
Me.Show
End Sub
Private Sub RefreshJPlot()
If JPlotIsLoaded Then
    Unload frmJPlot
    'show hour glass
    Form1.MousePointer = vbHourglass
    Form1.StatusBar.Panels(StatusPanel.Map) = "Analyzing receiver and fish databases..."
    frmJPlot.Show
End If
End Sub
Private Sub cmdInvert_Click()
Dim R As Long
UnderProgramControl = True
For R = 0 To lstReceivers.ListCount - 1
    lstReceivers.Selected(R) = Not lstReceivers.Selected(R)
    Receiver_Table.Invisible(ReceiverNumber(R)) = Not Receiver_Table.Invisible(ReceiverNumber(R))
Next R

'update screen
Form1.RePaint
UnderProgramControl = False
RefreshJPlot
Me.Show
End Sub

Private Sub cmdSelect_Click()
Dim R As Long
UnderProgramControl = True
For R = 0 To lstReceivers.ListCount - 1
    lstReceivers.Selected(R) = True
Next R

Receiver.MakeVisible
'update screen
Form1.RePaint
UnderProgramControl = False
RefreshJPlot
Me.Show
End Sub

Private Sub Form_Load()
'load list in form with all receivers
Dim R As Long
Dim G As Long
Dim ShowGroup(MAX_GROUPS) As Boolean
Dim ListIndex As Long
Dim Update As Boolean

UnderProgramControl = True
For R = 1 To Receiver.TotalReceivers
    G = Receiver.GroupNumber(CInt(R))
    Update = False
        If G = 0 Then
            lstReceivers.AddItem Receiver.ID(CInt(R))
            ReceiverNumber(ListIndex) = R
            ListIndex = ListIndex + 1
            Update = True
        Else
            If Not ShowGroup(G) Then
                lstReceivers.AddItem Receiver.ID(CInt(R))
                ReceiverNumber(ListIndex) = Receiver.GroupReceiverNumber(CInt(G))
                ShowGroup(G) = True
                ListIndex = ListIndex + 1
                Update = True
            End If
        End If
    If Update Then
        lstReceivers.Selected(ListIndex - 1) = Not Receiver_Table.Invisible(ReceiverNumber(ListIndex - 1))
    End If
Next R
'update screen
Form1.ShowDetectors
UnderProgramControl = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.StatusBar.Panels(1).Text = ""
End Sub

Private Sub lstReceivers_Click()
Dim R As Long
Dim G As Integer
Dim Response As Variant
Dim Changed As Boolean

If Not UnderProgramControl Then
    For R = 0 To lstReceivers.ListCount - 1
        Receiver_Table.Invisible(ReceiverNumber(R)) = Not lstReceivers.Selected(R)
        Changed = True
    Next R
    'update screen
    If Changed Then
        Form1.RePaint
        RefreshJPlot
        Me.Show
        If WhatToShowOnCanvas = ShowOnCanvas.StampList Then
            frmVerboseStampList.Show
            frmVerboseStampList.LoadStamps
        End If
    End If
End If

End Sub

