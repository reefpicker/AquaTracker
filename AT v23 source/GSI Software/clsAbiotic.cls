VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAbiotic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim MaxDepth As Single
Dim MinDepth As Single
Dim FishNumber As Long
Dim ReceiverNumber As Long
Dim ScaleReset As Boolean
Public Property Get Max() As Single
Max = MaxDepth
End Property
Public Property Get Min() As Single
Min = MinDepth
End Property
Public Property Let SelectFish(FN As Long)
'set fish  number
'NOTE: THIS PROCEDURE MAY CHANGE THE VALUE OF THE CURRENT STAMP!

FishNumber = FN
GetFishMaxMin

End Property
Public Property Let SelectReceiver(R As Long)
'set receiver number
'NOTE: THIS PROCEDURE MAY CHANGE THE VALUE OF THE CURRENT STAMP!
ReceiverNumber = R
GetReceiverMaxMin
End Property
Private Sub GetReceiverMaxMin()
Dim TotalStamps As Long
Dim i As Long
'NOTE: THIS PROCEDURE WILL CHANGE THE VALUE OF THE CURRENT STAMP!

'get number of stamps
TotalStamps = Receiver.Detection_Total(ReceiverNumber)

'get max/min
MaxDepth = 0
MinDepth = MAX_DEPTH_EVER

'find max and min among all stamps
For i = 1 To TotalStamps
    ReceiverTable.ReadStamp ReceiverNumber, i
    If MaxDepth < Stamp.AbioticFactor Then
        MaxDepth = Stamp.AbioticFactor
    End If
    
    If MinDepth > Stamp.AbioticFactor Then
        MinDepth = Stamp.AbioticFactor
    End If
Next i


End Sub
Private Sub GetFishMaxMin()
Dim i As Long
Const MAX_DEPTH_EVER = 2 ^ 30
'NOTE: THIS PROCEDURE WILL CHANGE THE VALUE OF THE CURRENT STAMP!

'get max/min
MaxDepth = 0
MinDepth = MAX_DEPTH_EVER

'find max and min among all stamps
For i = 1 To FishDatabase.NumberOfStamps
    FishTable.ReadStamp FishNumber, i
    If MaxDepth < Stamp.AbioticFactor Then
        MaxDepth = Stamp.AbioticFactor
    End If
    
    If MinDepth > Stamp.AbioticFactor Then
        MinDepth = Stamp.AbioticFactor
    End If
Next i

'reset to first stamp
FishTable.ReadStamp FishNumber, 1

'reset window scale
 ScaleReset = True


End Sub
Public Sub DrawDepthHistogram_Fish(PicBox As PictureBox)
'Draw depth histogram
'
Dim Bins As Long
Dim BinSize As Long
Dim PicBoxWidth As Long
Dim PicBoxHeight As Long
Dim i As Long
Dim range As Long

'get scale info
range = CLng(Abs(MaxDepth - MinDepth))
BinSize = PicBox.ScaleHeight / range


End Sub
Public Sub DrawCurrentDepth(StampNumber As Long)
'draws depth info on window
Dim DepthScale As Single
Dim PicBoxWidth As Long
Dim PicBoxHeight As Long
Dim Position As Long
Dim i As Long
Dim range As Long
Dim Scale_Increments As Long
Dim Scale_Value As Long


'calculate parameters
range = CLng(Abs(MaxDepth - MinDepth))
DepthScale = (frmDepth.picCurrentDepth.ScaleHeight) / range

Position = CLng((Stamp.AbioticFactor - MinDepth) * DepthScale)
PicBoxWidth = frmDepth.picCurrentDepth.ScaleWidth

'set window if not previously set
If ScaleReset Then
    'write scale
    Scale_Increments = range / frmDepth.txtScaleLabel.UBound
    Scale_Value = MinDepth
    
    For i = 0 To frmDepth.txtScaleLabel.UBound
        frmDepth.txtScaleLabel(i).Text = Str$(MinDepth + i * Scale_Increments) & "-"
    Next i
End If


'write depth info
frmDepth.picCurrentDepth.Line (1, Position)-(PicBoxWidth, Position + DepthScale), vbRed, BF

frmDepth.txtDepth.Text = Str$(Stamp.AbioticFactor)

End Sub
