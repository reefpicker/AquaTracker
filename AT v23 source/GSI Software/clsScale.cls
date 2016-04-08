VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim AutomaticallySelectValues As Boolean
Dim range As Long
Dim MaxValue As Long
Dim MinValue As Long
Dim BinSize As Single
Dim BinColor(255) As Long
Dim BinVisible(255) As Boolean
Dim NumberOfShades As Integer
Dim OneEightOfBins As Single
Dim Width As Long
Dim Height As Long
Dim BinWidth As Single
Dim DensityOfReceiver(MAX_RECEIVERS) As Long
Dim StoreMarkerSize As Long
Dim ScaleFactor As Single
Private Enum ColorSchemeSelected
    GreenToRed
    Multi
    AquaToPurple
    GreyScale
End Enum

Dim MRed As Integer
Dim MGreen As Integer
Dim MBlue As Integer
Dim Scheme As Integer
Const AboveRangeColor = vbRed
Const BelowRangeColor = vbWhite
Public Property Get WidthOfBin() As Single
WidthOfBin = BinWidth
End Property
Public Property Let ColorScheme(Value As Integer)
'Select color scheme
Scheme = Value
End Property
Public Property Get Max() As Long
Max = MaxValue
End Property
Public Property Let Max(Value As Long)
If AutomaticallySelectValues Then
    MinValue = 1
    If Value > 255 Then NumberOfShades = 255 Else NumberOfShades = Value
End If
If Value <= MinValue Then
    MaxValue = MinValue
    MinValue = Value
Else
    MaxValue = Value
End If
range = MaxValue - MinValue

AutoRange
AutoAssignColors

End Property
Private Sub AutoRange()
'get values for scale
If range < NumberOfShades Then NumberOfShades = range + 1
If NumberOfShades < 1 Then NumberOfShades = 1
ScaleFactor = range / NumberOfShades
BinSize = ScaleFactor
If ScaleFactor > 1 Then ScaleFactor = 1
End Sub
Public Property Let Min(Value As Long)
MinValue = Value
range = MaxValue - MinValue
AutoRange
End Property
Public Property Get Min() As Long
Min = MinValue
End Property
Public Property Let NumberOfBins(Value As Integer)
If NumberOfBins > 255 Then NumberOfBins = 255
If NumberOfBins <= 0 Then NumberOfBins = 1
NumberOfShades = Value
End Property

Public Property Get NumberOfBins() As Integer
NumberOfBins = NumberOfShades
End Property
Public Sub AutoAssignColors()
Dim ToneLevel As Integer
Dim i As Long
Dim intensity As Integer
Dim Color As Long
Dim result As Variant
Dim Y As Integer
Dim Residual As Integer
Dim Level As Integer
Dim BaseIntensity As Integer
Dim RedC As Integer
Dim BlueC As Integer
Dim GreenC As Integer
Dim LeftShifted As Integer
Dim RightShifted As Integer
Dim Shifted As Integer
Dim B As Long

AutoRange
OneEightOfBins = CSng(NumberOfShades * BinSize) / 8

'tones
BaseIntensity = 255


ToneLevel = Fix(BaseIntensity / NumberOfShades)
Shifted = MinValue / BinSize
LeftShifted = Shifted
RightShifted = Shifted
If Shifted + NumberOfShades > 255 Then RightShifted = 0
If Shifted < 0 Then LeftShifted = 0

'assign colors
If ScaleFactor = 0 Then
    R = MsgBox("No scale available.  Make sure valid receivers are displayed", vbOKOnly)
    Exit Sub
End If

B = 0
For i = 0 + LeftShifted To NumberOfShades + RightShifted - 1
    intensity = BaseIntensity - (ToneLevel * i / ScaleFactor)
    If intensity < 0 Then intensity = 0
    Level = Fix(intensity / 128)
    Residual = intensity Mod 128
    
    
    Select Case Scheme
        Case ColorSchemeSelected.GreenToRed
            If Level = 0 Then
                RedC = 255: GreenC = Residual * 2: BlueC = 0
            Else
                RedC = (128 - Residual) * 2: GreenC = 255: BlueC = 0
            End If
        Case ColorSchemeSelected.Multi
            If Level = 0 Then
                RedC = 0: GreenC = (256 - Residual): BlueC = Residual * 2
            Else
                RedC = 128: GreenC = Residual * 2: BlueC = (256 - Residual)
            End If
        Case ColorSchemeSelected.AquaToPurple
            If Level = 0 Then
                RedC = (128 - Residual) * 2: GreenC = 0: BlueC = 255
            Else
                RedC = 0: GreenC = Residual * 2: BlueC = 255
            End If
        Case ColorSchemeSelected.GreyScale
            RedC = intensity: BlueC = intensity: GreenC = intensity
    End Select
    Color = RGB(RedC, GreenC, BlueC)
    BinColor(B) = Color
    B = B + 1
Next i


End Sub
Public Sub ShowScale()
Dim X As Long
Dim i As Long

If NumberOfShades = 0 Then Exit Sub

'quick here, ensure form is loaded
frmScale.Show

With frmScale.picScale
    'window width and height!
    Width = .ScaleWidth
    Height = .ScaleHeight
    'clear!
    .Cls
End With


'valid values?
If Width > 0 And Height > 0 Then
    BinWidth = Width / NumberOfShades
    
    'paint the bins
    For i = 0 To NumberOfShades - 1
        X = i * BinWidth
        If BinVisible(i) Then
            frmScale.picScale.Line (X, 0)-(X + BinWidth, Height), BinColor(i), BF
        Else
            'if bin is not enabled, draw white
            frmScale.picScale.Line (X, 0)-(X + BinWidth, Height), vbWhite, BF
        End If
    Next i
    
    'make sure is all the way to the end
    'If BinVisible(NumberOfShades) Then
     '   frmScale.picScale.Line (X + BinWidth, 0)-(Width, Height), BinColor(NumberOfShades), BF
    'Else
        'if bin is not enabled, draw white
     '   frmScale.picScale.Line (X + BinWidth, 0)-(Width, Height), vbWhite, BF
    'End If
    
    'label 1/8 ticks
    For i = 0 To 8
        frmScale.lblScaleValue(i).Caption = Format((i * OneEightOfBins), "##")
    Next i
End If
End Sub
Private Function ReturnBinNumber(Value As Long) As Integer
Dim i As Long
Dim BinNumber As Single
Dim Color As Long
'can return negatives
If BinSize > 0 Then
    BinNumber = Fix((Value - MinValue) / BinSize)
End If

If BinNumber < 1 Then BinNumber = 0
If BinNumber > 255 Then BinNumber = 255
If Value > MaxValue Then BinNumber = -1
ReturnBinNumber = CInt(BinNumber)

End Function
Public Property Get ReturnColorOfBin(Value As Long) As Long
Dim Color As Long
Dim WithinRange As Boolean

'value in range or out of range?
WithinRange = True
If Value < MinValue Then
    WithinRange = False
    Color = BelowRangeColor
End If

If Value > MaxValue Then
    WithinRange = False
    Color = AboveRangeColor
End If

If WithinRange Then
    Color = BinColor(ReturnBinNumber(Value))
End If

'return value
ReturnColorOfBin = Color
End Property
Public Property Let AssignColor(BinNumber As Integer, Color As Long)
BinColor(BinNumber) = Color
End Property
Public Property Get GetBinColor(BinNumber As Integer) As Long
If BinNumber >= 0 Then GetBinColor = BinColor(BinNumber)
End Property
Public Property Get Bin(X As Single) As Integer

If BinWidth > 0 Then
    Bin = Fix(X / BinWidth)
End If

End Property
Public Sub ShowPlot(MapPicture As PictureBox, DensityData() As Long, MarkerSize As Long, Optional WIPE_CANVAS As Boolean = True)
Dim ReceiverNumber As Integer
Dim Color As Long
Dim GroupReceiver As Integer
Dim R As Integer
Dim BinNumber As Integer

StoreMarkerSize = MarkerSize

'Refresh Scale first
ShowScale

'clear screen
If WIPE_CANVAS Then Form1.ClearScreen

'draw receivers one by one
For ReceiverNumber = 1 To Receiver.TotalReceivers
    'calculate bin size
    'consider if part of group
    If Receiver.GroupNumber(ReceiverNumber) Then
        R = Receiver.GroupReceiverNumber(Receiver.GroupNumber(ReceiverNumber))
    Else
        R = ReceiverNumber
    End If

    DensityOfReceiver(R) = DensityData(R)
    Color = ReturnColorOfBin(DensityOfReceiver(R))
    
    BinNumber = ReturnBinNumber(DensityOfReceiver(R))
    
    'draw if visible
    If BinNumber >= 0 Then
        If (Not Receiver_Table.Invisible(R)) And BinVisible(BinNumber) Then
            Receiver.DrawReceiver MapPicture, R, MarkerSize, Color
        End If
    End If
Next ReceiverNumber

End Sub
Public Sub RefreshPlot()
Dim ReceiverNumber As Integer
Dim Color As Long
Dim GroupReceiver As Integer
Dim R As Integer
Dim BinNumber As Integer

'Refresh Scale first
ShowScale

'clear screen
Form1.ClearScreen

'draw receivers one by one
For ReceiverNumber = 1 To Receiver.TotalReceivers
    'calculate bin size
    'consider if part of group
    If Receiver.GroupNumber(ReceiverNumber) Then
        R = Receiver.GroupReceiverNumber(Receiver.GroupNumber(ReceiverNumber))
    Else
        R = ReceiverNumber
    End If
    Color = ReturnColorOfBin(DensityOfReceiver(R))
    BinNumber = ReturnBinNumber(DensityOfReceiver(R))
    
    'draw if visible
    If BinNumber >= 0 Then
        If Receiver.Detection_TTL(R) <> 0 And BinVisible(BinNumber) Then
            Receiver.DrawReceiver Form1.Picture1, R, StoreMarkerSize, Color
        End If
    End If
Next ReceiverNumber

End Sub
Public Property Let AutoSet(Value As Boolean)
AutomaticallySelectValues = Value
End Property
Public Property Get AutoSet() As Boolean
AutoSet = AutomaticallySelectValues
End Property
Public Function RangeString(BinNumber As Integer) As String
'gets a range
'in the form of a string "A - B"
'for the bin number
Dim s As String
Dim Low As Long
Dim High As Long
Dim t As Long

Low = (BinNumber * BinSize) + MinValue
High = (((BinNumber + 1) * BinSize) - 1) + MinValue

'check if scale is up or down
If Low > High Then
    'Swap
    t = Low
    Low = High
    High = t
End If

s = Str$(Low) & " - " & Str$(High)

RangeString = s

End Function
Public Sub HighlightBin(HighlightedBin As Integer)
Dim ReceiverNumber As Integer
Dim Color As Long
Dim GroupReceiver As Integer
Dim R As Integer
Dim BinNumber As Integer

'nothing to draw if no bins
If BinSize = 0 Then Exit Sub

'Refresh Scale first
ShowScale

'clear screen
Form1.ClearScreen

'draw receivers one by one
For ReceiverNumber = 1 To Receiver.TotalReceivers
    'calculate bin size
    'consider if part of group
    If Receiver.GroupNumber(ReceiverNumber) Then
        R = Receiver.GroupReceiverNumber(Receiver.GroupNumber(ReceiverNumber))
    Else
        R = ReceiverNumber
    End If
    'get bin number
    BinNumber = ReturnBinNumber(DensityOfReceiver(R))
    Color = ReturnColorOfBin(DensityOfReceiver(R))
    
    'compare to bin to highlight
    If HighlightedBin = BinNumber And Not Receiver_Table.Invisible(R) Then
        'special draw
        'highlight!
        Receiver.Show R, Form1.Picture1
        
    Else
        'normal draw
        'draw if visible
        If BinNumber >= 0 Then
            If (Not Receiver_Table.Invisible(R)) And BinVisible(BinNumber) Then
                Receiver.DrawReceiver Form1.Picture1, R, StoreMarkerSize, Color
            End If
        End If
    End If
Next ReceiverNumber
End Sub
Public Property Get IsVisible(BinNumber As Integer) As Boolean
IsVisible = BinVisible(BinNumber)
End Property
Public Property Let MakeVisible(BinNumber As Integer)
BinVisible(BinNumber) = True
End Property
Public Property Let MakeNotVisible(BinNumber As Integer)
BinVisible(BinNumber) = False
End Property
Private Sub Class_Initialize()
Dim i As Long
For i = 0 To 255
    BinVisible(i) = True
Next i
'30 shades of grey
NumberOfShades = 30
MinValue = 1
Scheme = ColorSchemeSelected.GreyScale
AutomaticallySelectValues = True
End Sub

