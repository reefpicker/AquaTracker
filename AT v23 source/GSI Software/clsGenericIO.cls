VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGenericIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Unified IO class
'Formats output of -Calculator- class to print or file (CSV)
'(Also will handle IMPORT Embedded Label Format )
Option Explicit
Private Type FIELD
    Name As String
    Width As Long
    Value As String
    X As Long
    Show As Boolean
    Key As String
End Type

Private Type PageDescriptor
    Start_Column As Long
    End_Column As Long
End Type

Const MAX_FIELDS = 50
Const MAX_PAGES = 100
Const UPDATE_WINDOW = True
Const MAX_ENTRIES = 5000

Dim Columns(MAX_FIELDS) As FIELD


Dim Device As Long

Dim File_Name As String
Dim File_Number As Long
Dim Box As PictureBox
Dim Driver_Status As String

Dim MaxWidth As Long
Dim MaxHeight As Long
Dim Entries(MAX_ENTRIES, MAX_FIELDS) As String
Dim LastEntry As Long
Dim FirstEntry_OnScreen As Long
Dim LastEntry_OnScreen As Long
Dim Page As Long
Dim Page_Max As Long
Dim Page_Info(MAX_PAGES) As PageDescriptor
Dim EntriesPerPage As Long
Dim WriteColumnHeader As Boolean
Dim LastColumnAssigned As Long
Public Property Get EndOfPage() As Boolean
If EntriesPerPage > 0 Then EndOfPage = True
End Property
Public Property Get LastEntryInBuffer() As Long
LastEntryInBuffer = LastEntry
End Property
Public Sub ShowFields(FieldList As ListBox)
Dim i As Long
FieldList.Clear

For i = 0 To MAX_FIELDS
    If Columns(i).Name > "" Then FieldList.AddItem Columns(i).Name
Next i
    
End Sub
Public Property Get Field_Is_Included(ByVal FieldNumber As Long) As Boolean

Field_Is_Included = Columns(FieldNumber).Show

End Property
Public Property Let Field_Is_Included(ByVal FieldNumber As Long, s As Boolean)
Columns(FieldNumber).Show = s
End Property
Public Sub CloseDevice()
Close #File_Number

End Sub
Public Property Get FIELD(i As Long)
Dim Concatenated As String
Dim NL As String
NL = Chr$(13) & Chr$(10)
'Name of field
Concatenated = UCase(Columns(i).Name)

'field length
Concatenated = Concatenated & NL & "Field Length =" & Str$(Columns(i).Width)

FIELD = Concatenated

End Property
Public Sub Select_Device(DeviceID As Long, Optional WindowName As PictureBox = "", Optional Name As String)

Driver_Status = "Opening Device for Communication..."
Device = DeviceID

Select Case DeviceID
    Case Device_Type.Printer
        File_Number = -1
    Case Device_Type.File
        File_Number = 1
        On err GoTo Error
            File_Number = 1
            Open Name For Output As #File_Number
            WriteHeader File_Number
    Case Device_Type.Window
        Set Box = WindowName
        MaxWidth = WindowName.ScaleWidth
        MaxHeight = WindowName.ScaleHeight
End Select

'set flag
WriteColumnHeader = True


Exit Sub

Driver_Status = "Device Open"

Error:

Driver_Status = "Failure to Open device"

End Sub
Public Property Let Assign_Field(Key As String, Name As String)

Dim ColumnNumber As Long
Dim Width As Long

'first make sure it does not already exist
'rule: if an entry with the same KEY and NAME already exist, exit without raising an error
'same key is OK. same name is OK.
Do
    With Columns(ColumnNumber)
        If .Key = Key And .Name = Name Then
            Exit Property
        End If
    End With
    ColumnNumber = ColumnNumber + 1
Loop Until ColumnNumber >= LastColumnAssigned
    
'insert field at the end of column list
ColumnNumber = LastColumnAssigned

If ColumnNumber > MAX_FIELDS Then ColumnNumber = MAX_FIELDS
Width = Len(Name) + 2
With Columns(ColumnNumber)
    .Name = Formatting(Name, Width)
    .Width = Width
    .Show = True
    .Key = UCase(Key)
End With

'next column
LastColumnAssigned = LastColumnAssigned + 1

End Property
Public Property Get FieldLen(ColumnNumber As Long) As Long
FieldLen = Columns(ColumnNumber).Width
End Property
Public Property Let Edit_Field(ColumnNumber As Long, Name As String, Width As Long)
If ColumnNumber > MAX_FIELDS Then ColumnNumber = MAX_FIELDS
If Width = AUTOMATIC Then Width = Len(Name) + 2
With Columns(ColumnNumber)
    .Name = Formatting(Name, Width)
    .Width = Width
    .Show = True
End With

End Property
Public Property Get Name(ColumnNumber) As String
Name = Columns(ColumnNumber).Name
End Property
Public Property Get GetFieldNumber(Key As String) As Long
'gets column number based on occurrance of first key
Dim i As Long
Dim Found As Boolean

Do
    With Columns(i)
        If UCase(Key) = .Key Then Found = True
    End With
    i = i + 1
Loop Until Found Or i > MAX_FIELDS

GetFieldNumber = i
End Property
Public Property Get Status() As String
Status = Driver_Status
End Property

Public Property Let WriteField(Parameter As Variant, Value As String)
'Write a value to the field using the key
'or using its column/field number
Dim FieldNumber As Long
Dim Key As String
Dim i As Long

If IsNumeric(Parameter) Then
    FieldNumber = CLng(Val(Parameter))
    If FieldNumber <= MAX_FIELDS Then
        Columns(FieldNumber).Value = Formatting(Value, Columns(FieldNumber).Width)
    End If
Else
    Key = UCase(Parameter)
    Do
        With Columns(i)
            If .Key = Key Then
                .Value = Formatting(Value, .Width)
            End If
        End With
        i = i + 1
    Loop Until i >= LastColumnAssigned
End If

End Property

Private Function Formatting(InputString As String, Width As Long) As String
'low level string operations to format string to fit column in desired position
'and to add a leading zero
Dim l As Long
Dim OutputString As String
Dim d As Double
Dim s As Long
Dim p As Long

InputString = Trim(InputString)

'check if decimal point?!
If IsNumeric(InputString) Then
    d = Abs(Val(InputString))
    If InStr(1, InputString, "E") Then
        'special case if its an "E" number
        If Width >= 3 Then
            InputString = Format(InputString, "##." & String(Width - 2, "#"))
        Else
            InputString = "0.0"
        End If
    Else
        If d > 0 And d < 1 Then
            p = InStr(1, InputString, ".")
            l = Len(InputString)
            InputString = Left(InputString, p - 1) & "0" & Right(InputString, l - (p - 1))
        End If
    End If
End If

If Width > 0 Then
    l = Len(InputString)
    If Width >= l Then
        OutputString = InputString & Space(Width - l)
    Else
        OutputString = Left(InputString, Width)
    End If
End If


Formatting = OutputString

End Function
Public Sub Scroll(First As Long)
Const UPDATEWINDOW = True
FirstEntry_OnScreen = First
LastEntry_OnScreen = First + EntriesPerPage
If LastEntry_OnScreen > LastEntry Then LastEntry_OnScreen = LastEntry
WriteLine_to_Window UPDATEWINDOW
End Sub
Public Sub WriteLine()
'writes all columns to one line

Select Case Device
    Case Device_Type.File
        WriteLine_to_File
    Case Device_Type.Window
        'if on window entries, select first and last entry in display
        If EntriesPerPage > 0 Then
            LastEntry_OnScreen = FirstEntry_OnScreen + EntriesPerPage
        Else
            LastEntry_OnScreen = LastEntry
        End If
        'call sub
        WriteLine_to_Window
End Select

End Sub
Public Sub NextPage()
'window driver uses pages
'a page is a window that the user can scroll
'This is to allow multiple columns (not for multiple rows)
'(rows are efimerous)
Page = Page + 1
If Page > MAX_PAGES Or Page > Page_Max Then Page = Page_Max

WriteLine_to_Window (UPDATE_WINDOW)

End Sub
Public Sub PreviousPage()
'window driver uses pages
'a page is a window that the user can scroll
'This is to allow multiple columns (not for multiple rows)
'(rows are efimerous)
Page = Page - 1
If Page < 0 Then Page = 0

WriteLine_to_Window (UPDATE_WINDOW)
End Sub

Private Sub WriteLine_to_Window(Optional UpdateOnly As Boolean = False)
'Write Line to window
'
Dim i As Long
Dim Concatenated As String
Dim FirstColumn As Long
Dim LastColumn As Long
Dim row As Long
Dim Alternate As Boolean
Dim X1 As Long
Dim X2 As Long
Dim Y1 As Long
Dim Y2 As Long


With Page_Info(Page)
    FirstColumn = .Start_Column
    LastColumn = .End_Column
End With


If Page_Info(Page).End_Column = 0 Then
    LastColumn = MAX_FIELDS
End If

If Page_Info(Page).Start_Column = 0 Then
    FirstColumn = 0
End If

'prep window
With Box
    .Cls
    .BackColor = vbWhite
    .ForeColor = vbBlack
    .FontBold = True
End With


'write columns
For i = FirstColumn To LastColumn
    If Columns(i).Show Then
        'align field to column
        Columns(i).X = Box.CurrentX
        If Columns(i).Show Then Box.Print Columns(i).Name;
        If Box.CurrentX > MaxWidth Then
            'if boundary exceeded then allow user to hit a next or cont button
            'and prepare for a second page
            Page_Info(Page).End_Column = i
            LastColumn = i
            Page_Max = Page_Max + 1
            If Page_Max > MAX_PAGES Then Page_Max = MAX_PAGES
            Page_Info(Page_Max).Start_Column = i
            Exit For
        End If
    End If
Next i

'advance carriage
Box.Print
Box.FontBold = False

If UpdateOnly Then
    'Print previous rows
    For row = FirstEntry_OnScreen To LastEntry_OnScreen
        For i = FirstColumn To LastColumn
            Box.CurrentX = Columns(i).X
            If Columns(i).Show Then Box.Print Entries(row, i);
        Next i
        'advance carriage
        Box.Print
        If Alternate Then
            Alternate = False
            Box.ForeColor = vbBlack
        Else
            Alternate = True
            Box.ForeColor = vbBlue
        End If
    Next row
Else
    'Print previous rows
    For row = FirstEntry_OnScreen To LastEntry_OnScreen - 1
        For i = FirstColumn To LastColumn
            Box.CurrentX = Columns(i).X
            Box.Print Entries(row, i);
        Next i
        'advance carriage
        Box.Print
        If Alternate Then
            Alternate = False
            Box.ForeColor = vbBlack
        Else
            Alternate = True
            Box.ForeColor = vbBlue
        End If
    Next row
    
    'print and split this entry
    'note fields=columns!!
    For i = 0 To MAX_FIELDS
        Entries(LastEntry, i) = Columns(i).Value
        If i >= FirstColumn And i <= LastColumn And Columns(i).Show Then
            Box.CurrentX = Columns(i).X
            Box.Print Columns(i).Value;
            Entries(LastEntry, i) = Columns(i).Value
            If Box.CurrentY > MaxHeight And EntriesPerPage = 0 Then EntriesPerPage = LastEntry - 1
        End If
    Next i
    'add as new entry
    LastEntry = LastEntry + 1
End If

'draw lines
For i = FirstColumn To LastColumn
    X1 = Columns(i).X - 60
    If X1 < 0 Then X1 = 0
    X2 = X1
    Y1 = 0
    Y2 = MaxHeight
    If X1 > 0 Then
        Box.Line (X1 - 20, Y1)-(X2 + 20, Y2), vbBlack, BF
    End If
Next i


End Sub
Public Sub WriteHeader(Optional FileNumber As Long = 0)
If FileNumber = 0 Then FileNumber = File_Number
'writes header of program on file or printer at buffer number
Print #FileNumber, "AquaTracker " & Version & "by Jose J. Reyes-Tomassini"
Print #FileNumber, "Manchester Research Station"
Print #FileNumber, "National Ocean and Atmospheric Administration"
Print #FileNumber, ""
Print #FileNumber, ""
Print #FileNumber, Date & "," & Time
Print #FileNumber, ""
Print #FileNumber, "Source: " & FileNameWithoutPath

End Sub
Public Sub WriteALineToFile(s As String)
Print #File_Number, s
End Sub
Public Sub WriteBuffer_to_File(FileName As String)
Dim row As Long
Dim column As Long
Dim Concatenated As String

Open FileName For Output As #2

WriteHeader 2


'write columns
For column = 0 To MAX_FIELDS
    If Columns(column).Show Then
        'align field to column
        Concatenated = Concatenated & Columns(column).Name & ","
    End If
Next column

'write to file
Print #2, Concatenated

'save all buffer entries
For row = 0 To LastEntry
    Concatenated = ""
    For column = 0 To MAX_FIELDS
        If Columns(column).Show Then
            Concatenated = Concatenated & Entries(row, column) & ","
        End If
    Next column
    'print to file
    Print #2, Concatenated
Next row

'clear
LastEntry = 0

'close
Close #2

End Sub
Private Sub WriteLine_to_File()
'Writes line to device FILE
'
Dim i As Long
Dim Concatenated As String

Const COMMA = ","

'write columns only once
If WriteColumnHeader Then
    For i = 0 To MAX_FIELDS
        If Columns(i).Show Then
            Concatenated = Concatenated & Columns(i).Name & COMMA
        End If
    Next i
    WriteColumnHeader = False
    Print #File_Number, Concatenated
End If

'write line itself
Concatenated = ""
For i = 0 To MAX_FIELDS
    If Columns(i).Show Then
        Concatenated = Concatenated & Columns(i).Value & COMMA
    End If
Next i

Print #File_Number, Concatenated

End Sub
Private Sub Class_Initialize()
Dim i As Long

'reset
Driver_Status = "Initialized"

'init
For i = 0 To MAX_FIELDS
    Columns(i).Name = ""
    Columns(i).Width = 0
Next i

EntriesPerPage = 0
LastEntry = 0
WriteColumnHeader = False


End Sub
