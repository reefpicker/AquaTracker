VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAssignReceiverToGroup 
   Caption         =   "Assign Receiver to Group"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6705
   Icon            =   "frmAssignReceiverToGroup.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   3390
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   11360
            MinWidth        =   11360
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdJoin 
      Caption         =   "Create/Join to this group"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   3000
      Width           =   2415
   End
   Begin VB.ListBox lstReceivers 
      Height          =   255
      Left            =   5760
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox lstGroupNumber 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox lstGroupNames 
      Height          =   1035
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.ListBox lstGroupMembers 
      Height          =   2595
      Left            =   2640
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label lblGroupName 
      Caption         =   "Receivers in group :"
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Select a Group:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
End
Attribute VB_Name = "frmAssignReceiverToGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ReceiversToAdd(MAX_RECEIVERS) As Integer
Dim NumberOfReceiversToAdd As Long
Dim GroupNumber As Long
Dim UnDo As Boolean
Dim NewGroupNumber As Long
Public Sub NewGroup()
'This sub has the preconds that:
'Multiple receivers are selected: Receiver_Selected(R) is a boolean
'and if True, R is selected
'If only one receiver added: Receiver # is in CurrentStation
'
'
Dim Name As String
Dim GN As Long
Dim G As Long
'name
Name = UCase(InputBox("To create a new group, enter a name for the group, or to add receiver(s) to an existing group leave blank.", "Group Name"))

'do not allow user to click OK, unless user has clicked JOINED
cmdOK.Enabled = False

'new receivers
NumberOfReceiversToAdd = CollectReceivers

'if new name, hatch new group
'else use whatever group is there to use
If Name <> "" Then
    GN = Receiver.CreateNewReceiverGroup(Name)
    NewGroupNumber = GN
    lstGroupNumber.AddItem Str$(GN)
    lstGroupNames.AddItem Name
    G = lstGroupNumber.ListCount - 1
    SelectGroup G
    UnDo = True
Else
    SelectGroup 0
End If
   
End Sub
Private Function CollectReceivers() As Long
Dim R As Integer
Dim N As Long

'gets receivers
If MultiSelect Then
    For R = 0 To MAX_RECEIVERS
        If Receiver_Selected(R) Then
            If Not Receiver_Table.Invisible(R) Then
                ReceiversToAdd(N) = R
                N = N + 1
                'clear
                Receiver_Selected(R) = False
            End If
        End If
    Next R
Else
    R = Receiver.CurrentStation_Number
    ReceiversToAdd(0) = R
    N = 1
End If

If N > 0 Then
    cmdJoin.Enabled = True
    StatusBar.Panels(1).Text = "Click on JOIN to join these receivers into the selected group."
End If

CollectReceivers = N

End Function
Private Sub SelectGroup(G As Long)
'selected group will show on lbl box and list of members
'
Dim Name As String
Dim i As Long
Dim List(MAX_RECEIVERS) As Integer
Dim m As Long

If lstGroupNames.ListCount = 0 Then Exit Sub

'name
Name = lstGroupNames.List(G)
lblGroupName.Caption = "Receivers in group " & Name & ":"
'members
lstGroupMembers.Clear
lstReceivers.Clear
'add members
GroupNumber = CLng(lstGroupNumber.List(G))
m = Receiver.GetReceiversInGroup(GroupNumber, List())
i = 0

'receivers already in group
Do
    If List(i) <> 0 Then
        lstGroupMembers.AddItem Receiver.ID(List(i), True)
        lstReceivers.AddItem Str$(List(i))
    End If
    i = i + 1
Loop Until i >= m Or i > MAX_RECEIVERS

'new receivers if any
For i = 0 To NumberOfReceiversToAdd - 1
    lstGroupMembers.AddItem Receiver.ID(ReceiversToAdd(i))
    lstReceivers.AddItem ReceiversToAdd(i)
Next i

End Sub

Private Sub cmdCancel_Click()
If UnDo Then
    'kill group that was not joined
    Receiver.DeleteGroup NewGroupNumber
End If

Unload Me
End Sub

Private Sub cmdJoin_Click()
Dim TotalReceiversInGroup As Long
Dim L(MAX_RECEIVERS) As Integer
Dim GroupName As String
Dim R As Integer
Dim i As Long
Dim Joined As Boolean
Dim response As Variant


'no groups? Nothing to do
If lstGroupNames.ListCount = 0 Then
    Unload Me
Else
    'make group shown, group in memory
    If NumberOfReceiversToAdd > 0 Then
        If GroupNumber = NewGroupNumber Then UnDo = False
        TotalReceiversInGroup = Receiver.GetReceiversInGroup(GroupNumber, L())
        For i = TotalReceiversInGroup To lstReceivers.ListCount - TotalReceiversInGroup - 1
            R = CInt(lstReceivers.List(i))
            Receiver.AddReceiverToGroup(GroupNumber) = R
            Joined = True
        Next i
    End If
End If

If Joined Then
    Form1.RePaint
    Receiver.Show R, Form1.Picture1
    cmdJoin.Enabled = False
End If

cmdOK.Enabled = True


End Sub

Private Sub cmdOK_Click()
'refresh canvas
Form1.RePaint
Unload Me
End Sub

Private Sub Form_Load()
Dim i As Long

Dim G As Long
Dim L(MAX_GROUPS) As String
NumberOfReceiversToAdd = 0
'add group name to window
G = Receiver.GetGroupNames(L())
For i = 0 To G
    If L(i) <> "" Then
        lstGroupNames.AddItem L(i)
        lstGroupNumber.AddItem i
    End If
Next i

'Select First one for now
SelectGroup 0

'Show tips
StatusBar.Panels(1).Text = "Double click on a group or receiver to delete it..."
End Sub

Private Sub lstGroupMembers_DblClick()
'Delete receivers?
Dim i As Integer
For i = 0 To lstGroupMembers.ListCount - 1
    If lstGroupMembers.Selected(i) Then
        DeleteReceiver i
        Exit Sub
    End If
Next i

End Sub
Private Sub DeleteReceiver(i As Integer)
Dim ListInGroupAlready(MAX_RECEIVERS) As Integer
Dim AlreadyInGroup As Long
Dim R As Integer
Dim response As Variant

'get receivers in group
AlreadyInGroup = (Receiver.GetReceiversInGroup(GroupNumber, ListInGroupAlready())) - 1
If i >= AlreadyInGroup Then
    response = MsgBox("Remove receiver from list?", vbYesNo, "Remove from group")
    If response = vbYes Then
        lstGroupMembers.RemoveItem i
        lstReceivers.RemoveItem i
    End If
Else
    R = lstReceivers.List(i)
    response = MsgBox("Remove receiver from list? Note that any deletions of receivers already in a group can't be undone by the CANCEL button in this window", vbYesNo, "Remove from group")
    If response = vbYes Then
        Receiver.DeleteReceiver GroupNumber, R
        lstGroupMembers.RemoveItem i
        lstReceivers.RemoveItem i
    End If
End If

End Sub
Private Sub lstGroupNames_Click()
Dim i As Long
For i = 0 To lstGroupNames.ListCount - 1
    If lstGroupNames.Selected(i) Then
        SelectGroup i
        Exit Sub
    End If
Next i

End Sub

Private Sub lstGroupNames_DblClick()
Dim i As Integer
For i = 0 To lstGroupNames.ListCount - 1
    If lstGroupNames.Selected(i) Then
        QueryDeleteGroup i
        Exit Sub
    End If
Next i

End Sub
Private Sub QueryDeleteGroup(G As Integer)
Dim GroupNumber As Long
Dim GroupName As String

Dim response As Variant

GroupNumber = CLng(lstGroupNumber.List(G))
GroupName = lstGroupNames.List(G)

response = MsgBox("Delete group " & GroupName & "? This will remove the group permanently from memory and unlink all receivers from the group.  This action can't be undone.", vbYesNoCancel, "Delete Group")
If response = vbYes Then
    Receiver.DeleteGroup GroupNumber
    lstGroupNames.RemoveItem G
    SelectGroup 0
End If

End Sub

Private Sub lstGroupNames_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If cmdJoin.Enabled = True And lstGroupNames.ListCount > 1 Then StatusBar.Panels(1).Text = "Select a group in which to place your receivers."
End Sub
