VERSION 5.00
Begin VB.Form frmSelectAnalysis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fields to Include"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7065
   Icon            =   "frmSelectAnalysis.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstFieldNumber 
      Height          =   1035
      Left            =   6120
      TabIndex        =   9
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "<---"
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "--->"
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   1920
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Field information"
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   4440
      Width           =   5775
      Begin VB.TextBox txtDescription 
         Height          =   855
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   360
         Width           =   5535
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.ListBox lstIncluded 
      Height          =   3960
      ItemData        =   "frmSelectAnalysis.frx":0442
      Left            =   3480
      List            =   "frmSelectAnalysis.frx":0449
      TabIndex        =   3
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
   Begin VB.ListBox lstCalculatedParamters 
      Height          =   3960
      ItemData        =   "frmSelectAnalysis.frx":045A
      Left            =   120
      List            =   "frmSelectAnalysis.frx":045C
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Select fields you want to include:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmSelectAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ActiveField As Long

Private Sub cmdAdd_Click()
Dim i As Long
Dim Found As Boolean
Dim Selection As Long

'selection
For i = 0 To lstCalculatedParamters.ListCount - 1
    If lstCalculatedParamters.Selected(i) Then
        Selection = i
    End If
Next i


'scan to see if already added
For i = 0 To lstFieldNumber.ListCount - 1
    If CLng(lstFieldNumber.List(i)) = Selection Then
        Found = True
    End If
Next i

If Not Found Then
    lstFieldNumber.AddItem Selection
    lstIncluded.AddItem lstCalculatedParamters.List(Selection)
End If

    
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim i As Integer

'No show all
For i = 0 To lstCalculatedParamters.ListCount - 1
    DeviceBuffer.Field_Is_Included(i) = False
Next i

'whats included
For i = 0 To lstFieldNumber.ListCount - 1
    DeviceBuffer.Field_Is_Included(CLng(lstFieldNumber.List(i))) = True
Next i

Unload Me
  
End Sub

Private Sub cmdRemove_Click()
Dim i As Integer

For i = 0 To lstIncluded.ListCount - 1
    If lstIncluded.Selected(i) Then
        lstIncluded.RemoveItem i
        lstFieldNumber.RemoveItem i
        Exit For
    End If
Next i
End Sub

Private Sub Form_Load()
Dim i As Integer
LoadList

'Cache to local variables to use local function

End Sub
Private Sub LoadList()
Dim i As Integer
Dim FormJustLoaded As Boolean

'form loaded or just reloading list?
If lstFieldNumber.ListCount > 0 Then FormJustLoaded = False Else FormJustLoaded = True

lstCalculatedParamters.Clear
lstIncluded.Clear

ActiveField = -1
DeviceBuffer.ShowFields lstCalculatedParamters

If FormJustLoaded Then
    For i = 0 To lstCalculatedParamters.ListCount - 1
        If DeviceBuffer.Field_Is_Included(CLng(i)) Then
            lstIncluded.AddItem lstCalculatedParamters.List(i)
            lstFieldNumber.AddItem i
        End If
    Next i
Else
    For i = 0 To lstFieldNumber.ListCount - 1
        lstIncluded.AddItem lstCalculatedParamters.List(lstFieldNumber.List(i))
    Next i
End If
End Sub

Private Sub lstCalculatedParamters_Click()
Dim i As Integer
'display parameters in text box
For i = 0 To lstCalculatedParamters.ListCount - 1
    If lstCalculatedParamters.Selected(i) Then
        txtDescription.Text = DeviceBuffer.FIELD(CLng(i))
        ActiveField = i
    End If
Next i


End Sub

Private Sub lstCalculatedParamters_DblClick()
Dim result As String
Dim s As String
Dim w As Long

Dim i As Integer
'display parameters in text box
For i = 0 To lstCalculatedParamters.ListCount - 1
    If lstCalculatedParamters.Selected(i) Then
        txtDescription.Text = DeviceBuffer.FIELD(CLng(i))
        ActiveField = i
    End If
Next i

ChangeField

End Sub
Private Sub ChangeField()
Dim result As String
Dim w As Long
Dim s As String

If ActiveField < 0 Then Exit Sub

result = InputBox("New Field Name (leave blank to use same):", "Change Field Name")

If result = "" Then result = DeviceBuffer.Name(ActiveField)

s = result

result = InputBox("Field Length (enter 0 for default):", "Change Field Length")

If result = "" Or Len(result) > 4 Then result = "0"

w = CLng(Val(result))
If w <= 0 Or w > 50 Then
    w = -1 'makes it auto-select
    s = Trim$(s)
End If

DeviceBuffer.Edit_Field(ActiveField, s) = w
'reload list
LoadList

End Sub
Private Sub lstIncluded_Click()
Dim i As Integer
'display parameters in text box
For i = 0 To lstIncluded.ListCount - 1
    If lstIncluded.Selected(i) Then
        txtDescription.Text = DeviceBuffer.FIELD(CLng(lstFieldNumber.List(i)))
        ActiveField = i
    End If
Next i
End Sub

Private Sub txtDescription_Click()
ChangeField
End Sub
