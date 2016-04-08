VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFishEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim HashTable(MAX_FISH, MAX_FISH) As Long
Dim CodeTable(MAX_FISH, MAX_FISH) As Long
Dim NumberOfFish As Long
Dim FishNumber As Long
Dim FCode(MAX_FISH) As Long
Dim FNumberOfStamps(MAX_FISH) As Long
Dim FRelease_Site(MAX_FISH) As String
Dim FVisible(MAX_FISH) As Boolean
Dim FColor(MAX_FISH) As Long
Public Sub SaveFish(FN As Long)
'save all fish in table
'Format
'Code, Color, Stamp Total
Dim i As Integer
For i = 0 To NumberOfFish - 1
    Print #FN, Code(i) & "," & Str$(FColor(i)) & "," & FRelease_Site(i) & "," & Str$(FNumberOfStamps(i))
Next i
'write the end section code here
Print #FN, "[---]"
End Sub
Public Sub LoadFish(FN As Long)
Dim s As String
Dim Done As Boolean
Dim F As Integer
Dim Code As String

ReDim SplitString(10) As String

On Error GoTo ExitWithError

''''''''Important condition:
''''''''
''''''''YOU NEED TO SAVE ALL FISH AND LOAD ALL FISH EVEN THOSE THAT ARE SET TO EXCLUDED
''''''''OTHERWISE THE DBASES MAY NOT LINE-UP
Do
    Line Input #FN, s
    If s = "[---]" Then
        Done = True
    Else
        SplitString = Split(s, ",")
        Code = SplitString(0)
        F = GetFishNumber(Code)
        FColor(F) = CLng(SplitString(1))
        FVisible(F) = False 'make all visible by setting this to false
        FRelease_Site(F) = SplitString(2)
        FNumberOfStamps(F) = CLng(SplitString(3))
        F = F + 1
    End If
Loop While Not EOF(1) And Not Done

'total number of fish is f
NumberOfFish = F

Exit Sub

ExitWithError:
NumberOfFish = 0

End Sub
Property Get TotalFishLoaded() As Long
TotalFishLoaded = NumberOfFish - 1
End Property
Property Let Color(Optional ByVal FN As Integer = -1, c As Long)
If FN > -1 And FN <= MAX_FISH Then FishNumber = FN
FColor(FishNumber) = c
End Property
Property Get Color(Optional ByVal FN As Integer = -1) As Long
If FN > -1 And FN <= MAX_FISH Then FishNumber = FN
Color = FColor(FishNumber)
End Property
Property Let IsVisible(Optional ByVal FN As Integer = -1, v As Boolean)
'since values default to false, this is inverted here
If FN > -1 And FN <= MAX_FISH Then FishNumber = FN
FVisible(FishNumber) = Not v
End Property
Property Get IsVisible(Optional ByVal FN As Integer = -1) As Boolean
'since the values default to false, this is inverted here
If FN > -1 And FN <= MAX_FISH Then FishNumber = FN
IsVisible = Not FVisible(FishNumber)
End Property
Property Get Code(Optional ByVal FN As Integer = -1) As String
If FN > -1 And FN <= MAX_FISH Then FishNumber = FN
'returns the string or hex for the fish code
Code = Trim$(Str$(FCode(FishNumber)))
End Property
Property Let Code(Optional ByVal FN As Integer = -1, c As String)
On Error GoTo RaiseError

If FN > -1 And FN <= MAX_FISH Then FishNumber = FN

If IsNumeric(c) Then
    FCode(FishNumber) = Val(c)
Else
    FCode(FishNumber) = Val("&H" & c)
End If

Exit Property

RaiseError:
'invalid number stored?
FCode(FishNumber) = 0
End Property
Property Get NumberOfStamps(Optional FN As Integer = -1) As Long
If FN > -1 And FN <= MAX_FISH Then FishNumber = FN
NumberOfStamps = FNumberOfStamps(FishNumber)
End Property
Public Sub NumberOfStampsUP(Optional FN As Integer = -1)
If FN > -1 And FN <= MAX_FISH Then FishNumber = FN
FNumberOfStamps(FishNumber) = FNumberOfStamps(FishNumber) + 1
End Sub

Public Sub Clear()
Dim i As Integer
Dim ii As Integer

'clears hashtable and stamps
For i = 0 To MAX_FISH
    FNumberOfStamps(i) = 0
    FCode(i) = 0
    For ii = 0 To MAX_FISH
        HashTable(i, ii) = 0
        CodeTable(i, ii) = 0
    Next ii
    'default color for tracks is blue
    FColor(i) = vbBlue
Next i

NumberOfFish = 0

End Sub

Property Get Release_Site(Optional FN As Integer = -1) As String
If FN > -1 And FN <= MAX_FISH Then FishNumber = FN
Release_Site = FRelease_Site(FishNumber)
End Property
Property Let Release_Site(Optional FN As Integer = -1, s As String)
If FN > -1 And FN <= MAX_FISH Then FishNumber = FN
FRelease_Site(FishNumber) = s
End Property
Property Get Fish() As Long
Fish = FishNumber
End Property
Property Let Fish(FN As Long)
If FN > -1 And FN <= MAX_FISH Then FishNumber = FN

End Property
Public Function GetFishNumber(CodeString As String) As Integer
'translate fish code into fish number
'PRECOND: DOES NOT ACCEPT A CODESTRING WITH HEX VALUES unless preceded by &H
Dim Found As Boolean
Dim c As Long
Dim CodeValue As Long
Dim ValueInTable As Long
Dim i As Long

Dim response As Variant

On Error GoTo ExitWithError
Static SeedValue As Long

CodeValue = Val(CodeString)
If SeedValue = 0 Then SeedValue = Val(CodeString)

c = Abs(SeedValue - Val(CodeString))

If c > MAX_FISH Then c = MAX_FISH

'search hash table
i = -1
Do
    i = i + 1
    ValueInTable = HashTable(c, i)
    If ValueInTable = CodeValue Then Found = True
Loop Until Found Or i > MAX_FISH Or ValueInTable = 0

'if not found, add to table in order found
If Not Found Then
    FCode(NumberOfFish) = CodeValue
    CodeTable(c, i) = NumberOfFish
    HashTable(c, i) = CodeValue
    '0 based
    NumberOfFish = NumberOfFish + 1
End If

GetFishNumber = CodeTable(c, i)

Exit Function

ExitWithError:
response = MsgBox("Illegal fish code found or too many fish codes.", vbOKOnly, "Error in import file")
End Function
Public Sub TransferList(List As ComboBox)
Dim i As Long

If NumberOfFish = 0 Then Exit Sub
For i = 0 To NumberOfFish - 1
    List.AddItem Str$(FCode(i))
Next i

End Sub

Private Sub Class_Initialize()
Dim i As Long
'default color for tracks is blue
For i = 0 To MAX_FISH
    FColor(i) = vbBlue
Next i

End Sub
Private Function CountReceivers(ByVal F As Long) As Long
'Search the fish field (or any other stamp-formatted table entry) for # of receivers
Dim ReceiverIsInTable(MAX_FISH) As Boolean
Dim range As Long
Dim i As Long
Dim Count As Long

For i = 0 To FishDatabase.NumberOfStamps(CInt(F)) - 1
    FishTable.ReadStamp F, i
    If (Not ReceiverIsInTable(Stamp.Site)) And Stamp.Valid Then
        ReceiverIsInTable(Stamp.Site) = True
        Count = Count + 1
    End If
Next i

'rts
CountReceivers = Count

End Function

