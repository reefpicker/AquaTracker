VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Table Class
'This class stores the stamps produced by receivers in chronological order
'

'PARTITION TABLE before writting to it. Pass an array with expected size/range of each entry.
'Access by setting entry or directly by accessing

Private Type RangeNumber
    Low As Long
    High As Long
End Type

Dim FishNumber(MAX_ENTRIES) As Integer
Dim ReceiverID(MAX_ENTRIES) As Integer
Dim DateStamp(MAX_ENTRIES) As Single
Dim TimeStamp(MAX_ENTRIES) As Integer
Dim ReferenceTable(MAX_ENTRIES) As Long
Dim CTime(MAX_ENTRIES) As Long
Dim EntryRange(MAX_ENTRIES) As RangeNumber
Dim RI As Long

Public Sub ResetTable()
Dim i As Long

'clear all entry ranges
Do
    EntryRange(i).Low = 0
    EntryRange(i).High = 0
    i = i + 1
Loop Until i >= MAX_ENTRIES

RI = 0
End Sub
Public Sub DeleteStamp(ByVal EntryNumber As Long, ByVal Index As Long)
'delete stamp
'writes blanks to all stamp fields
'Note: This object does not know what kind of table it is so delinking
'from the other tables (e.g. number of detections per fish) that point to this entry
'is the responsability of the calling sub.
Dim RI As Long

'get index
RI = ReferenceIndex(EntryNumber, Index)

'write blanks to stamp
FishNumber(RI) = 0
ReceiverID(RI) = 0
DateStamp(RI) = 0
TimeStamp(RI) = 0
CTime(RI) = 0
End Sub
Public Sub ReadStamp(ByVal EntryNumber As Long, ByVal Index As Long)
'reads stamp to table
'
Dim RI As Long

'get index
RI = ReferenceIndex(EntryNumber, Index)

'write stamp to table
Stamp.Fish = FishNumber(RI)
Stamp.Site = ReceiverID(RI)
Stamp.Date = DateStamp(RI)
Stamp.Time = TimeStamp(RI)
Stamp.CTime = CTime(RI)

'assume valid
Stamp.Valid = True

If DateListIsLoaded Then
    'if it exist and is unchecked, this function returns a False
    Stamp.Valid = frmDates.AddDate(Stamp.Date)
End If

'deleted stamp is not valid
If DateStamp(RI) = 0 And TimeStamp(RI) = 0 Then Stamp.Valid = False

End Sub
Public Sub TransferUniqueReceiverEntriesToList(F As Long, ReceiverIsInTable() As Boolean)
'Search the FISH field (or any other stamp-formatted table entry) for # of receivers
'transfers to a list/updates list
'list is  SORTED on return

Dim ReceiverNumber As Long
Dim range As Long
Dim i As Long

range = EntryRange(F).High - EntryRange(F).Low

For i = 1 To range
    ReadStamp F, i
    If Not ReceiverIsInTable(Stamp.Site) Then
        ReceiverIsInTable(Stamp.Site) = True
    End If
Next i

End Sub
Private Sub GenericTransferUniqueEntriesToList(ByVal EntryNumber As Long, QueryField As Long, List As ListBox)

'Transfer contents of field or fields to a list
'
Dim i As Long
Dim range As Long
Dim ItemToAdd As String
Dim TempList(MAX_ENTRIES) As String
Dim LastItem As Long
Dim Item As Long
Dim Found As Boolean
Dim LastItemAdded As String

'if list contains something, transfer contents
If List.ListCount >= 1 Then
    For i = 0 To List.ListCount - 1
        TempList(i) = List.List(i)
    Next i
End If

'last item in list
LastItem = i

'get range
range = EntryRange(EntryNumber).High - EntryRange(EntryNumber).Low

'step thru list
For i = 1 To range
    ReadStamp EntryNumber, i
    ItemToAdd = Choose(QueryField, FishDatabase.Code(Stamp.Fish), Receiver.ID(Stamp.Site), Convert_DayNumber(Stamp.Date), Convert_ToStandardTime(Stamp.Time), Convert_DayNumber(Stamp.Date) & "  " & Convert_ToStandardTime(Stamp.Time), Str$(Stamp.Fish) & " on " & Convert_DayNumber(Stamp.Date) & " at " & Convert_ToStandardTime(Stamp.Time), Str$(Stamp.Date))
    
    'search item in list
    Found = False
    Item = 0
    If ItemToAdd = LastItemAdded Then
        Found = True
    Else
        Do
            If ItemToAdd = TempList(Item) Then
                'if already in list, do not add
                Found = True
            End If
            Item = Item + 1
        Loop While (Not Found) And Item <= LastItem
    End If
    
    'else add
    If Not Found Then
        TempList(LastItem) = ItemToAdd
        LastItem = LastItem + 1
        'also add to list
        List.AddItem ItemToAdd
        LastItemAdded = ItemToAdd
    End If
Next i


End Sub
Public Sub PasteStamp(EntryNumber As Integer, Index As Long)
'Copy stamp and move index ahead
'Warning: ONLY WORKS IF COPYING IS DONE IN SEQUENCE
'Uses the base index from previous entries top (last) entry

Dim First As Long
Dim Last As Long
Dim TableIndex As Long
Dim RelativeIndex As Long

Dim BaseIndex As Long
Dim Base As Long


'get index
Base = EntryNumber - 1
If Base = -1 Then
    BaseIndex = 0
Else
    BaseIndex = EntryRange(Base).High
End If

ReferenceTable(RI) = RI

'write stamp to table
FishNumber(RI) = Stamp.Fish
ReceiverID(RI) = Stamp.Site
DateStamp(RI) = Stamp.Date
TimeStamp(RI) = Stamp.Time
CTime(RI) = ConsolidateTime(Stamp.Date, Stamp.Time)

RI = RI + 1


'at the end update range counters
EntryRange(EntryNumber).Low = BaseIndex         'makes sure exists for this entry
EntryRange(EntryNumber).High = RI  'moves top entry ahead

End Sub
Public Sub WriteStamp(EntryNumber As Long, Index As Long)
'writes stamp to table
'
Dim RI As Long
Dim First As Long
Dim Last As Long
Dim TableIndex As Long
Dim RelativeIndex As Long


'get index
RI = ReferenceTable(EntryRange(EntryNumber).Low + Index)
'write stamp to table
FishNumber(RI) = Stamp.Fish
ReceiverID(RI) = Stamp.Site
DateStamp(RI) = Stamp.Date
TimeStamp(RI) = Stamp.Time
CTime(RI) = ConsolidateTime(Stamp.Date, Stamp.Time)

'check to see if last value has been written in table, if so, sort!
If Index + EntryRange(EntryNumber).Low = EntryRange(EntryNumber).High Then
    QuickSort CTime(), ReferenceTable(), EntryRange(EntryNumber).Low, EntryRange(EntryNumber).High
End If

End Sub

Private Function ReferenceIndex(Entry As Long, Index As Long) As Long
Dim RI As Long
RI = ReferenceTable(EntryRange(Entry).Low + Index)

ReferenceIndex = RI
End Function
Private Sub QuickSort(ByRef pvarArray() As Long, ByRef Index() As Long, Optional ByVal plngLeft As Long, Optional ByVal plngRight As Long)
    '
    'QuickSort for link lists.  First argument is the sorted by value.  Second is the link to first.  It sorts the second list using the first list as the argument.
    'Sort is bounded by last two arguments which are optional if list is not partitioned
    'Otherwise the last two arguments are used internally during recursion.
          Dim lngFirst As Long
          Dim lngLast As Long
          Dim varMid As Long
          Dim varSwap As Long
   
          If plngRight = 0 Then
              plngLeft = LBound(pvarArray)
              plngRight = UBound(pvarArray)
          End If
          
          lngFirst = plngLeft
          lngLast = plngRight
          varMid = pvarArray(Index((plngLeft + plngRight) \ 2))
          
          Do
              Do While pvarArray(Index(lngFirst)) < varMid And lngFirst < plngRight
                  lngFirst = lngFirst + 1
              Loop
              Do While varMid < pvarArray(Index(lngLast)) And lngLast > plngLeft
                  lngLast = lngLast - 1
              Loop
              If lngFirst <= lngLast Then
                  varSwap = Index(lngFirst)
                  Index(lngFirst) = Index(lngLast)
                  Index(lngLast) = varSwap
                  lngFirst = lngFirst + 1
                  lngLast = lngLast - 1
              End If
          Loop Until lngFirst > lngLast
          If plngLeft < lngLast Then QuickSort pvarArray, Index, plngLeft, lngLast
          If lngFirst < plngRight Then QuickSort pvarArray, Index, lngFirst, plngRight
      End Sub
Private Function ConsolidateTime(d As Long, t As Integer) As Long
'Consolidates the TIME and DATE array so that entries can be sorted by date AND time
'It those this by calculating the estimate of the year as minutes since year 0
'year 0 should be set to something less than 50 yrs or so from NOW
'Set to 1990
Dim response As Variant

On Error GoTo SkipError

Const MINUTES_PER_DAY = 1440#

ConsolidateTime = (d * MINUTES_PER_DAY + t)

Exit Function

SkipError:
response = MsgBox("Internal Error: Bad Data Stamp/Unable to sort.", vbAbortRetryIgnore, "Internal Error")
If response = vbAbort Then Unload Form1: Stop
If response = vbRetry Then ConsolidateTime d, t
End Function
Public Sub PartitionTable(ByRef EntrySize() As Long)
'Partition table based on the size of each entry range
'
Dim Index As Long
Dim Entry As Long
Dim i As Long

'Set partitions
Entry = 0

Do
    'set partition bounds
    EntryRange(Entry).Low = Index
    Index = Index + EntrySize(Entry) - 1 'define upper boundary based on size+lower bound
    EntryRange(Entry).High = Index
    Entry = Entry + 1
    Index = Index + 1 'advance to next entry w/ no overlap!!
Loop Until Entry > UBound(EntrySize)

'prep table for key-based sorting
For i = 0 To EntryRange(Entry - 1).High
    ReferenceTable(i) = i
Next i

End Sub

