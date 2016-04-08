VERSION 5.00
Begin VB.Form frmDates 
   Caption         =   "Dates"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2310
   Icon            =   "frmDates.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5685
   ScaleWidth      =   2310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDeselect 
      Caption         =   "Deselect all"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select all"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton cmdInvert 
      Caption         =   "Invert all"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   5280
      Width           =   1335
   End
   Begin VB.ListBox lstShowDates 
      Height          =   4335
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.ListBox lstDates 
      Height          =   4110
      Left            =   840
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "frmDates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ListHasChanged As Boolean
Dim UnderProgramControl As Boolean
Dim ListHasBeenUpdated As Boolean

Private Sub cmdDeselect_Click()
Dim d As Long
ListHasBeenUpdated = True
UnderProgramControl = True
For d = 0 To lstShowDates.ListCount - 1
    lstShowDates.Selected(d) = False
    lstDates.Selected(d) = True
Next d
UnderProgramControl = False
'update screen
Form1.RePaint


End Sub

Private Sub cmdInvert_Click()
Dim d As Long
Dim Valid As Boolean
ListHasBeenUpdated = True
UnderProgramControl = True
For d = 0 To lstShowDates.ListCount - 1
    Valid = Not lstShowDates.Selected(d)
    lstShowDates.Selected(d) = Valid
    lstDates.Selected(d) = Not Valid
Next d
UnderProgramControl = False
'update screen
Form1.RePaint

End Sub

Private Sub cmdSelect_Click()
Dim d As Long
UnderProgramControl = True
ListHasBeenUpdated = True
For d = 0 To lstShowDates.ListCount - 1
    lstShowDates.Selected(d) = True
    lstDates.Selected(d) = False
Next d
UnderProgramControl = False
'update screen
Form1.RePaint


End Sub

Private Sub Form_Load()
DateListIsLoaded = True
Form1.RePaint
UpdateList
End Sub
Public Sub UpdateList()
Dim i As Long

UnderProgramControl = True

If ListHasChanged Then
    lstShowDates.Clear
    For i = 0 To lstDates.ListCount - 1
        lstShowDates.AddItem Convert_DayNumber(CLng(lstDates.List(i)))
        lstShowDates.Selected(i) = True
    Next i
    ListHasChanged = False
End If

UnderProgramControl = False

End Sub
Public Function AddDate(NumberDate As Long) As Boolean
'Add date to list
'convert to CTime, sort, search, and add date if necesary...

Dim DetectionDate As Date
Dim Found As Boolean
Dim H As Long
Dim L As Long
Dim p As Long
Dim Code As Long
Dim TotalDates As Long
Dim List As ListBox
Dim d As Long
Dim Last As Boolean
Static LastDayNumberProcessed As Long
Static OutputOfLastIteration As Boolean

'list name
Set List = Me.lstDates
'assume true
ValidDate = True

If LastDayNumberProcessed = NumberDate And Not ListHasBeenUpdated Then
    ValidDate = OutputOfLastIteration
    GoTo ExitNow
End If
'save
LastDayNumberProcessed = NumberDate

'list parameters
TotalDates = List.ListCount
H = TotalDates - 1
L = 0
Found = False
Last = False

'special case: first time run
If TotalDates = 0 Then
    'insert first item
    'checkmarked
    List.AddItem NumberDate
    Found = True
    p = 0
Else
    'binary search
    'half-ass implementation only cuts ops in 1/2
    p = (H - L) / 2
    If NumberDate = List.List(p) Then Found = True
    If NumberDate > List.List(p) Then d = 1
    If NumberDate < List.List(p) Then d = -1
    
    Do Until Found Or Last
        p = p + d
        If p < 0 Or p > H Then
            Last = True
        Else
            If NumberDate = List.List(p) Then Found = True
        End If
    Loop
End If

If Not Found Then
    'insert new item
    'checkmarked
    List.AddItem NumberDate
    ListHasChanged = True
Else
    'check if checkmarked
    'this is an inverse operation here
    'this has to do with my refusal to write another sorting routine, as I don't have time
    'and I don't want to spend a couple of hours doing it or trying to impement code already
    'in use here elsewhere (recycle)
    If List.Selected(p) Then ValidDate = False Else ValidDate = True
End If

'under program control, only user can update, only program can change list
ListHasBeenUpdated = False

'save for next iteration
OutputOfLastIteration = ValidDate

ExitNow:

'send back value to caller
AddDate = ValidDate

End Function

Private Sub Form_Unload(Cancel As Integer)
DateListIsLoaded = False
End Sub
Private Sub lstShowDates_Click()
Dim i As Long
Dim Changed As Boolean
If Not UnderProgramControl Then
    ListHasBeenUpdated = True
    For i = 0 To lstShowDates.ListCount - 1
        lstDates.Selected(i) = Not lstShowDates.Selected(i)
        Changed = True
    Next i
    If Changed Then
        If WhatToShowOnCanvas = ShowOnCanvas.StampList Then
            frmVerboseStampList.Show
            frmVerboseStampList.LoadStamps
        End If
        Form1.RePaint
    End If
End If
End Sub

