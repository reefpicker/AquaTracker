VERSION 5.00
Begin VB.Form frmOverlaps 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Overlapping pairs of receivers"
   ClientHeight    =   4830
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "frmOverlaps.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstFishes 
      Height          =   450
      Left            =   4800
      TabIndex        =   5
      Top             =   1800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtFishes 
      Height          =   1095
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   3600
      Width           =   4455
   End
   Begin VB.ListBox lstOverlapParsed 
      Height          =   450
      Left            =   4800
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox lstOverlaps 
      Height          =   2790
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   4455
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Fish that generate overlaps:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   2175
   End
End
Attribute VB_Name = "frmOverlaps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Type Table
    Equivalencies(5000) As Long
    LastEntry As Long
End Type
Dim Equates(5000) As Table
Dim EquMax As Long

Private Sub AutoGroup()
'automatically create groups based on receiver overlaps
'
Dim i As Long
Dim S() As String
Dim R(1, MAX_FISH) As Long
Dim LastPair As Long
Dim StartPair As Long
Dim GroupMembers(MAX_GROUPS) As String
Dim LastGroup As Long
Dim Group As Long
Dim Receiver_To_Compare As Long
Dim SeedPair As Long
Dim pair As Long
Dim Found As Boolean
Dim flag As Boolean
Dim G As Long
Dim EntryNumber As Long

'once analysis is done, don't allow again
Form1.mnuFindOverlaps.Enabled = False
    
'Create List of Pairs
LastPair = lstOverlaps.ListCount - 1
For i = 0 To LastPair
    'get pairs
    S = Split(lstOverlapParsed.List(i), "-")
    R(0, i) = CLng(S(0)): R(1, i) = CLng(S(1))
Next i

'Search pairs and create table of equivalents
For i = 0 To LastPair
    'scan table
    ScanTable R(0, i), R(1, i)
Next i

'now using table, consolidate lists
ConsolidateTable
                          
With frmAssignReceiverToGroup
    'Create groups automatically
    For EntryNumber = 0 To EquMax
        If Equates(EntryNumber).LastEntry > 0 Then
            Group = Group + 1
            G = Receiver.CreateNewReceiverGroup("G" & Str$(Group))
            For i = 0 To Equates(EntryNumber).LastEntry - 1
                Receiver.AddReceiverToGroup(G) = CInt(Equates(EntryNumber).Equivalencies(i))
            Next i
        End If
    Next EntryNumber
End With

End Sub
Private Sub ConsolidateTable()
Dim entry1 As Long
Dim entry2 As Long
Dim I1 As Long
Dim i2 As Long
Dim Change As Boolean

'compare entries in table, round-robin style
Do
    Change = False
    entry1 = 0
    Do
        entry2 = entry1 + 1
        Do
          'compare all branches
              For I1 = 0 To Equates(entry1).LastEntry - 1
                  For i2 = 0 To Equates(entry2).LastEntry - 1
                      If Equates(entry1).Equivalencies(I1) = Equates(entry2).Equivalencies(i2) Then
                          JoinEntry entry1, entry2
                          'DeleteEntry entry2
                          Equates(entry2).LastEntry = -1
                          Change = True
                          Exit For
                      End If
                  Next i2
                  If Change Then Exit For
              Next I1
              entry2 = entry2 + 1
          Loop Until Change Or entry2 > EquMax
          entry1 = entry1 + 1
          
    Loop Until Change Or entry1 > EquMax
    'reset counter and do again until no changes
Loop Until Not Change
End Sub
Private Sub JoinEntry(E1 As Long, E2 As Long)
'joins entries in equivalency table
'copies E2 into E1
Dim equivalencies_to_copy As Long
Dim Entry As Long
Dim copy_position As Long
Dim Count As Long


equivalencies_to_copy = Equates(E2).LastEntry
copy_position = Equates(E1).LastEntry

'copy
For Entry = 0 To equivalencies_to_copy
    Equates(E1).Equivalencies(Entry + copy_position) = Equates(E2).Equivalencies(Entry)
Next Entry


'update last marker
Equates(E1).LastEntry = copy_position + equivalencies_to_copy


End Sub
Private Sub ScanTable(Pair1 As Long, Pair2 As Long)
Dim p(1) As Long
Dim i As Long
Dim Entry As Long
Dim pair As Long
Dim Found As Boolean
Dim L As Long

p(0) = Pair1
p(1) = Pair2

Do
    Do
        With Equates(Entry)
            'search for entry that has the equivalency
                'search down the branch
                i = 0
                Found = False
                Do
                    If .Equivalencies(i) = p(pair) Then
                        For L = 0 To .LastEntry - 1
                            If .Equivalencies(L) = p(Abs(pair - 1)) Then
                                Exit For
                                Found = True
                            End If
                         Next L
                        If Not Found Then
                            .Equivalencies(.LastEntry) = p(Abs(pair - 1))
                            .LastEntry = .LastEntry + 1
                        End If
                        Found = True
                    End If
                    i = i + 1
                Loop Until i >= .LastEntry
        End With
        Entry = Entry + 1
    Loop Until Entry > EquMax
    pair = pair + 1
    Entry = 0
Loop Until pair = 2

'if entry does not have an equivalency table entry then create using pair
If Not Found Then
    With Equates(EquMax)
        .Equivalencies(0) = p(0)
        .Equivalencies(1) = p(1)
        .LastEntry = 2
    End With
    EquMax = EquMax + 1
End If
        
End Sub


Private Sub CancelButton_Click()
Unload Me

End Sub

Private Sub lstOverlaps_Click()
'tease pairs out
'and show on map
Dim S() As String
Dim R1 As Long
Dim R2 As Long
Dim i As Long

'get selection
For i = 0 To lstOverlaps.ListCount - 1
    If lstOverlaps.Selected(i) = True Then
        'get pairs
        S = Split(lstOverlapParsed.List(i), "-")
        R1 = CLng(S(0)): R2 = CLng(S(1))
        'clean slate
        Form1.ClearScreen
        'draw pairs
        Receiver.Show R1, Form1.Picture1
        Receiver.Show R2, Form1.Picture1
        txtFishes.Text = lstFishes.List(i)
    End If
Next i
End Sub
Private Sub lstOverlaps_KeyDown(KeyCode As Integer, Shift As Integer)
Const vbDelete = 46
Dim i As Long
If KeyCode = vbDelete Then
    For i = 0 To lstOverlaps.ListCount - 1
        If lstOverlaps.Selected(i) = True Then
            lstOverlaps.RemoveItem i
            lstOverlapParsed.RemoveItem i
            lstFishes.RemoveItem i
            Exit Sub
        End If
    Next i
End If
End Sub

Private Sub OKButton_Click()
Dim results As Variant

If lstOverlaps.ListCount > 0 Then
    results = MsgBox("Do you want to use pairs to automatically create groups?", vbYesNo, "Create groups")
    If results = vbYes Then
        AutoGroup
        'then let user see, edit, or delete groups
        Me.Hide
        Load frmAssignReceiverToGroup
        With frmAssignReceiverToGroup
            .cmdCancel.Enabled = False
            .Show
        End With
    End If
End If

Unload Me

End Sub

