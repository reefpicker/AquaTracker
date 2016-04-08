VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExcursionAnalyzer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Excursion Analysis Tool by Jose J. Reyes-Tomassini
'(c) 2009-2010
'added residence analysis in 2014
Const MAX_EXCURSIONS_PER_RECEIVER = 10000
Const MAX_EXCURSIONS_PER_FISH = 10000


Private Type SAL 'Stays And Leaves
    Fish As String
    DayStart As Long
    TimeStart As Integer
    DayEnd As Long
    TimeEnd As Integer
    Duration As Long
    PrevResidence As Long
End Type

Private Type Table
    FishID As String
    LastConsolidatedTime As Long
    Residence As Long
    LastDay As Single
    LastTime As Integer
End Type
Dim Output As clsGenericIO
Dim Max_FishTable As Long
Dim FishTable(MAX_FISH) As Table
Dim Max_FishExcursion As Long
Dim FishExcursion(MAX_EXCURSIONS_PER_RECEIVER) As SAL
Dim FishResidence(MAX_EXCURSIONS_PER_RECEIVER) As SAL
Dim ResidenceCount(MAX_FISH) As Long
Dim Excursion_TH As Long
Dim Residence_TH As Long

Public Property Let SetOutputAs(O As clsGenericIO)
Set Output = O
Output.Assign_Field("FID") = "Fish ID"
Output.Assign_Field("SDay") = "Start Day"
Output.Assign_Field("STime") = "Start Time"
Output.Assign_Field("EDay") = "End Day  "
Output.Assign_Field("ETime") = "End Time"
Output.Assign_Field("D") = "Duration"
Output.Assign_Field("PR") = "Prev. Residence"

End Property
Public Property Let Threshold_Excursion(TH As Long)
Excursion_TH = TH
End Property
Public Property Let Threshold_Residence(TH As Long)
Residence_TH = TH
End Property
Public Property Get Fish(Entry As Long) As Long
'Fish code index for entry
End Property
Public Property Get TimeStart(Entry As Long) As Long
'Excursion start time
End Property
Public Property Get TimeEnd(Entry As Long) As Long
'Excursion end time
End Property
Public Property Get Duration(Entry As Long) As Long
'duration of excursion
End Property
Public Property Get PreviousResidence(Entry As Long) As Long
'residence time at site before excursion
End Property
Private Sub ClearAll()
'clear all
Dim Fish As Long
Dim Excursion As Long

For Fish = 0 To MAX_FISH
    With FishTable(Fish)
        .FishID = ""
        .LastConsolidatedTime = 0
        .Residence = 0
    End With
Next Fish

For Excursion = 0 To MAX_EXCURSIONS_PER_FISH
    With FishExcursion(Excursion)
        .Fish = ""
        .DayEnd = 0
        .DayStart = 0
        .TimeEnd = 0
        .TimeStart = 0
        .PrevResidence = 0
        .Duration = 0
    End With
Next Excursion
        

End Sub
Public Sub ExportExcursion(R As Integer, FileName As String)
'Writes excursion information to file
'R=0 : All receivers
Dim Excursion As Long

'Open device
Output.Select_Device Device_Type.File, , FileName
Output.WriteHeader
'count excursions
If R > 0 Then
    Output.WriteALineToFile "Receiver Excursion for : " & Receiver.ID(CInt(R)) & "(" & Str$(R) & ")"
    AnalyzeExcursions R
    'write information
    WriteInformation
Else
    For R = 1 To Receiver.TotalReceivers
        Output.WriteALineToFile "Receiver Excursion for : " & Receiver.ID(CInt(R)) & "(" & Str$(R) & ")"
        AnalyzeExcursions R
        WriteInformation
    Next R
End If

'close device
Output.CloseDevice
End Sub
Private Sub AnalyzeExcursions(R As Integer)
Dim Group As Long

'clear all values
ClearAll

'count excursions
'is receiver part of group?
If Receiver.GroupNumber(R) = 0 Then
    CountExcursions R, CLng(R), ReceiverTable
Else
    'create group table
    Group = CLng(Receiver.GroupNumber(R))
    Receiver.CreateGroupTable (Group)
    CountExcursions R, 0, GroupTable
End If

End Sub

Public Sub ShowExcursions(R As Integer, WindowBox As PictureBox)

'Do the analysis
AnalyzeExcursions R

'Allow for at least that first column (fish id) to be under user control
With DeviceBuffer
    Output.Edit_Field(0, .Name(0)) = .FieldLen(0)
End With

'Open path to window
Output.Select_Device Device_Type.Window, WindowBox

'print directly to window
WriteInformation

'end device communication
Output.CloseDevice

End Sub
Private Sub WriteInformation()
Dim ExcursionNumber As Long

'Write Information
Do
    With FishExcursion(ExcursionNumber)
        Output.WriteField(0) = .Fish
        Output.WriteField(1) = Convert_DayNumber(.DayStart)
        Output.WriteField(2) = Convert_ToStandardTime(.TimeStart)
        Output.WriteField(3) = Convert_DayNumber(.DayEnd)
        Output.WriteField(4) = Convert_ToStandardTime(.TimeEnd)
        Output.WriteField(5) = .Duration
        Output.WriteField(6) = .PrevResidence
    End With
    'write
    Output.WriteLine
    ExcursionNumber = ExcursionNumber + 1
Loop Until ExcursionNumber >= Max_FishExcursion


End Sub
Private Sub CountExcursions(R As Integer, TableEntry As Long, Entries As clsTable)

'Count all valid excursions coming from group that includes receiver R
'Returns Maximum entry point for ReceiverExcursion variable

Dim ReceiverEntry As Long
Dim Fish As Long
Dim TimeDifference As Long


'Reset counter
Max_FishExcursion = 0

'get
'Read all stamps
For ReceiverEntry = 0 To Receiver.Detection_TTL(R) - 1
    'retrieve stamp
    Entries.ReadStamp TableEntry, ReceiverEntry
    'validate
    If Stamp.Valid Then
        'Get ID
        Fish = Stamp.Fish
        With FishTable(Fish)
            If .FishID = "" Then
                'first entry
                TimeDifference = 0
                'Load correct fish entry #
                .FishID = FishDatabase.Code(Fish)
            Else
                TimeDifference = Stamp.CTime - FishTable(Fish).LastConsolidatedTime
            End If
        End With
        
        'has an excursion occured?
        If TimeDifference >= Excursion_TH Then
            'add excursion...
            'but only if fish has resided at least once,
            If FishTable(Fish).Residence >= Residence_TH Then
                With FishExcursion(Max_FishExcursion)
                    .Fish = FishTable(Fish).FishID
                    .DayStart = FishTable(Fish).LastDay
                    .TimeStart = FishTable(Fish).LastTime
                    .DayEnd = Stamp.Date
                    .TimeEnd = Stamp.Time
                    .PrevResidence = FishTable(Fish).Residence
                    .Duration = Stamp.CTime - FishTable(Fish).LastConsolidatedTime
                End With
                'reset residence
                 FishTable(Fish).Residence = 0
                'update index
                Max_FishExcursion = Max_FishExcursion + 1
            End If
        End If
        
        If TimeDifference <= Residence_TH Then
            'update table residence time
            FishTable(Fish).Residence = FishTable(Fish).Residence + TimeDifference
        End If
        
        'last time fish was observed:
        With FishTable(Fish)
            .LastConsolidatedTime = Stamp.CTime
            .LastDay = Stamp.Date
            .LastTime = Stamp.Time
        End With
    End If
Next ReceiverEntry

'return

End Sub
