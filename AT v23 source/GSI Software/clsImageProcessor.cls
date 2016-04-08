VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImageProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'limits
Const MAXNodeNUMBER = 5000
Const Shift = 20
Private Type NodePoint
    X As Integer
    Y As Integer
    Count As Integer
    direction(8) As Integer
    ContourSize As Long
    gap As Long
End Type

Private Type Star
    SelectedDirection As Integer
    X As Integer
    Y As Integer
    direction(8) As Integer
End Type

Dim ReceiverConnectionToCorridor(MAX_RECEIVERS) As Star

Dim Buffer_H(MAX_X, MAX_Y)
Dim Buffer_V(MAX_X, MAX_Y)
Dim ImageBuffer(MAX_X, MAX_Y)
Dim ContourMap(MAX_X, MAX_Y) As Long
Dim X_Increment(7) As Long
Dim Y_increment(7) As Long
Dim GAP_X As Long
Dim GAP_Y As Long

Dim FishCorridorStored As Boolean

Dim SizeOfContour As Long
Dim PictureInMemory As StdPicture
Dim CornerX As Long
Dim CornerY As Long

Dim Node(MAXNodeNUMBER) As NodePoint
Dim NodeNumber As Integer
Dim ContourIndex As Long

Dim K_X(8) As Long
Dim K_Y(8) As Long
Dim s(3) As Long

Dim Slot_Size As Long
Const SLOTS = 1440

Dim TimeSlot(SLOTS) As Long
Dim PhaseSlot(SLOTS) As Long
Dim AlternatePhaseSlot(SLOTS) As Long
Dim TimeMax As Long
Dim Color_Pallette(16) As Long


'Define for all processing subs and functions:
Const LAND = 0
Const WATER = 1
Const TARGET = &HF
Const SET_BITS = &HFFFFFFFF
Const ZERO = 0


Dim Set_Max As Long

'For drawing arrows
Const Pi = 3.14159
Const ArrowHeadLength = 10
Const ArrowHeadAngle = 20
Const Angle = ArrowHeadAngle * Pi / 180
Const Grey = &H707070
Const Ninety_In_Radians = 90 * Pi / 180

Public Property Let SetMax(Value As Long)
If Value >= 0 Then Set_Max = Value
End Property
Public Property Get SetMax() As Long
SetMax = Set_Max
End Property
Public Property Let DielCycleBinSize(s As Long)
Slot_Size = s
End Property
Public Property Get DielCycleBinSize() As Long
DielCycleBinSize = Slot_Size
End Property
Public Sub SetMapSize(X As Long, Y As Long)
CornerX = X
CornerY = Y
End Sub
Public Sub ClearMap()
Dim X As Long
Dim Y As Long

For X = 0 To MAX_X
    For Y = 0 To MAX_Y
        Buffer_H(X, Y) = 0
        Buffer_V(X, Y) = 0
        ImageBuffer(X, Y) = 0
        ContourMap(X, Y) = 0
        MapImage(X, Y) = 0
        MapEdge(X, Y) = 0
    Next Y
Next X

End Sub
Public Sub Clear()
Dim X As Long
Dim Y As Long

For X = 0 To MAX_X
    For Y = 0 To MAX_Y
        Buffer_H(X, Y) = 0
        Buffer_V(X, Y) = 0
        ImageBuffer(X, Y) = 0
        ContourMap(X, Y) = 0
    Next Y
Next X

End Sub
Public Sub DrawMapToScale(Img As PictureBox, Source As PictureBox)
'Draws map to scale on the image
'
Dim X As Long
Dim Y As Long
Dim ScanBoxWidth As Single
Dim ScanBoxHeight As Single
Dim NumberOfScanWindows_X As Long
Dim NumberOfScanWindows_Y As Long
Dim Box_X As Long
Dim Box_Y As Long
Dim p As Long
Dim DRAWLAND As Boolean
Dim TempImg(MAX_X, MAX_Y) As Long
Dim BoxCorner_X As Long
Dim BoxCorner_Y As Long
Dim ScreenCornerX As Long
Dim ScreenCornerY As Long
Dim dummy As Boolean
Dim X_Offset As Long
Dim Y_Offset As Long
Dim Color As Long


Const Grey = &H707070

'get source scale
ScreenCornerX = Source.ScaleWidth
ScreenCornerY = Source.ScaleHeight
'get scale
ScanBoxWidth = ScreenCornerX / Img.ScaleWidth
ScanBoxHeight = ScreenCornerY / Img.ScaleHeight

'get number of scans
NumberOfScanWindows_X = Img.ScaleWidth
NumberOfScanWindows_Y = Img.ScaleHeight

'if transfering to a container of same size then do a simple copy and paste
If ScanBoxWidth = 1 And ScanBoxHeight = 1 Then
    For X = X_Offset To ScreenCornerX
        For Y = Y_Offset To ScreenCornerY
            If MapEdge(X, Y) = LAND Then
                TempImg(X, Y) = vbBlack
            Else
                TempImg(X, Y) = vbWhite
            End If
        Next Y
    Next X
Else
    
    'beging resizing
    For Box_X = 0 To NumberOfScanWindows_X
        For Box_Y = 0 To NumberOfScanWindows_Y
            DRAWLAND = False
            'scan source map
            BoxCorner_X = (Box_X * ScanBoxWidth) + X_Offset
            BoxCorner_Y = (Box_Y * ScanBoxHeight) + Y_Offset
            For X = BoxCorner_X To BoxCorner_X + ScanBoxWidth
                If X <= ScreenCornerX Then
                    For Y = BoxCorner_Y To BoxCorner_Y + ScanBoxHeight
                        If Y <= ScreenCornerY Then
                            p = MapEdge(X, Y)
                            If p = LAND Then
                                DRAWLAND = True
                            End If
                        End If
                    Next Y
                End If
            Next X
            
            'draw
            If DRAWLAND Then
                Color = vbBlack
            Else
                Color = vbWhite
            End If
            TempImg(Box_X, Box_Y) = Color
            
        Next Box_Y
    Next Box_X
End If

dummy = BufferDib.Create(Img.ScaleWidth, Img.ScaleHeight)
BufferDib.WriteScaledMatrix TempImg()
BufferDib.PaintPicture Img.hDC

End Sub
Public Sub DrawZoomMapToScale(Img As PictureBox, Source As PictureBox)
'Draws map to scale on the image
'
Dim X As Long
Dim Y As Long
Dim ScanBoxWidth As Single
Dim ScanBoxHeight As Single
Dim NumberOfScanWindows_X As Long
Dim NumberOfScanWindows_Y As Long
Dim Box_X As Single
Dim Box_Y As Single
Dim p As Long
Dim DRAWLAND As Boolean
Dim TempImg(MAX_X, MAX_Y) As Long
Dim BoxCorner_X As Long
Dim BoxCorner_Y As Long
Dim ScreenCornerX As Long
Dim ScreenCornerY As Long
Dim dummy As Boolean
Dim X_Offset As Long
Dim Y_Offset As Long
Dim i As Long
Dim ii As Long
Dim Color As Long


Const Grey = &H707070

With ZoomRegion
    ScreenCornerX = .OriginX + .SizeX
    ScreenCornerY = .OriginY + .SizeY
    X_Offset = .OriginX
    Y_Offset = .OriginY
    'get scale
    ScanBoxWidth = Img.ScaleWidth / .SizeX
    ScanBoxHeight = Img.ScaleHeight / .SizeY
    'get number of scans
    NumberOfScanWindows_X = .SizeX
    NumberOfScanWindows_Y = .SizeY
End With


'beging resizing

For Y = Y_Offset To ScreenCornerY
    Box_X = 0
    Box_Y = Box_Y + ScanBoxHeight
    For X = X_Offset To ScreenCornerX
        Color = vbWhite
        p = MapEdge(X, Y)
        If p = LAND Then
            Color = vbBlack
        End If
        Box_X = Box_X + ScanBoxWidth
        For i = 0 To ScanBoxWidth
            For ii = 0 To ScanBoxHeight
                TempImg(Box_X + i, Box_Y + ii) = Color
            Next ii
        Next i
        
    Next X
Next Y
dummy = BufferDib.Create(Img.ScaleWidth, Img.ScaleHeight)
BufferDib.WriteScaledMatrix TempImg()
BufferDib.PaintPicture Img.hDC

End Sub
Public Sub ShowImageFromMatrix(Img As PictureBox, ByRef t() As Long)
'shows an image in IMG picturebox from a matrix T (long and 2d)
'
Dim dummy As Boolean

dummy = MapDib.Create(Img.ScaleWidth, Img.ScaleHeight)
MapDib.WriteMatrix t()
MapDib.PaintPicture Img.hDC

End Sub

Public Sub Edge()

SaltAndPepper
KirshED (2)
Smooth
'EdgeDetect

End Sub
Public Function Redirect(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, ByRef PX() As Long, ByRef PY() As Long, Site_To As Long) As Long
Dim Success As Boolean
Dim x_array(20, CONTOUR_SIZE) As Long
Dim y_array(20, CONTOUR_SIZE) As Long
Dim Sizes(20) As Long
Dim temp_x(CONTOUR_SIZE) As Long
Dim temp_y(CONTOUR_SIZE) As Long
Dim N As Long
Dim Min As Long
Dim Smallest As Long

Dim i As Long

'show hour glass
Form1.MousePointer = vbHourglass
Form1.StatusBar.Panels(StatusPanel.Map) = "Calculating route..."

Do
    Do
        If ABORT_PROCESS Then
            ABORT_PROCESS = False
            If CURRENT_FISH <> -1 Then Form1.StatusBar.Panels(StatusPanel.Alert).Text = "User aborted at Fish #" & Str$(FishDatabase.Code(CURRENT_FISH))
            Exit Function
        End If
        Success = FollowContour(X1, Y1, X2, Y2, temp_x(), temp_y(), Site_To)
        DoEvents
    Loop While (Not Success)
    For i = 0 To SizeOfContour
        x_array(N, i) = temp_x(i)
        y_array(N, i) = temp_y(i)
    Next i
    Sizes(N) = SizeOfContour
    N = N + 1
Loop While N <= 5

Min = CONTOUR_SIZE

For i = 0 To 5
    If Sizes(i) < Min Then
        Smallest = i
        Min = Sizes(i)
    End If
Next i

For i = 0 To Sizes(Smallest)
    PX(i) = x_array(Smallest, i)
    PY(i) = y_array(Smallest, i)
Next i

Redirect = Sizes(Smallest)
'show hour glass
Form1.MousePointer = vbArrow
Form1.StatusBar.Panels(StatusPanel.Map) = ""

End Function
Private Sub SaltAndPepper()
'salt and pepper filter
'Eliminate artifacts from dithering and scaling etc
'WATER=1 ; LAND=0
Dim X As Long
Dim Y As Long
Dim p As Long
Dim Surroundings As Long
Dim i As Long
Dim ii As Long
Dim O As Double

For X = 2 To CornerX
    For Y = 2 To CornerY
        If MapImage(X, Y) = 1 Then
            Surroundings = 0
            For i = -2 To 2
                For ii = -2 To 2
                    'weight surroundings and delete pixel if surrounded by other "land" pixels
                    If i <> 0 And ii <> 0 Then Surroundings = Surroundings + MapImage(X - i, Y - ii)
                Next ii
            Next i
            If Surroundings = 0 Then
                MapImage(X, Y) = 0
            End If
        End If
    Next Y
Next X

End Sub
Private Sub EdgeDetect()
'dilates map image
Dim X As Long
Dim Y As Long
Dim p As Long



'Horizontal
For X = 3 To CornerX Step 2
    For Y = 0 To CornerY
        EdgeHorizontal X, Y
    Next Y
Next X

'Vertical
For X = 0 To CornerX
    For Y = 3 To CornerY Step 2
        EdgeVertical X, Y
    Next Y
Next X

'Result
For X = 0 To CornerX
    For Y = 0 To CornerY
        MapEdge(X, Y) = Buffer_H(X, Y) And Buffer_V(X, Y)
    Next Y
Next X

Smooth

End Sub

Private Sub EdgeHorizontal(XP As Long, Yp As Long)
Dim X As Long
Dim Y As Long
Dim p As Long
Dim side_a As Long
Dim side_b As Long

'edge by Delta
For X = XP - 3 To XP - 2
    side_a = side_a + MapImage(X, Yp)
Next X

For X = XP - 1 To XP
    side_b = side_b + MapImage(X, Yp)
Next X

If side_a <> side_b Then
    p = 0
Else
    p = 1
End If

For X = XP - 3 To XP
    Buffer_H(X, Yp) = p
Next X



End Sub
Public Sub KirshED(TreshHoldResponse As Integer)
Const SubMatrixSize As Integer = 2

Dim K(7) As Long
Dim MASK As Long
Dim X As Long
Dim Y As Long
Dim PosX As Long
Dim PixelColor As Long
Dim i As Long
Dim j As Long
Dim CheckChannel As Byte
Dim N As Long
Dim MaskResponse(2) As Long
Dim Channel As Integer
Dim KirshMask(7, 2, 2) As Long

'define kirshmasks at runtime
'K0
KirshMask(0, 0, 0) = -3: KirshMask(0, 1, 0) = -3: KirshMask(0, 2, 0) = 5
KirshMask(0, 0, 1) = -3: KirshMask(0, 1, 1) = 0: KirshMask(0, 2, 1) = 5
KirshMask(0, 0, 2) = -3: KirshMask(0, 1, 2) = -3: KirshMask(0, 2, 2) = 5
'K1
KirshMask(1, 0, 0) = -3: KirshMask(1, 1, 0) = 5: KirshMask(1, 2, 0) = 5
KirshMask(1, 0, 1) = -3: KirshMask(1, 1, 1) = 0: KirshMask(1, 2, 1) = 5
KirshMask(1, 0, 2) = -3: KirshMask(1, 1, 2) = -3: KirshMask(1, 2, 2) = -3
'K2
KirshMask(2, 0, 0) = 5: KirshMask(2, 1, 0) = 5: KirshMask(2, 2, 0) = 5
KirshMask(2, 0, 1) = -3: KirshMask(2, 1, 1) = 0: KirshMask(2, 2, 1) = -3
KirshMask(2, 0, 2) = -3: KirshMask(2, 1, 2) = -3: KirshMask(2, 2, 2) = -3
'K3
KirshMask(3, 0, 0) = 5: KirshMask(3, 1, 0) = 5: KirshMask(2, 2, 0) = -3
KirshMask(3, 0, 1) = 5: KirshMask(3, 1, 1) = 0: KirshMask(2, 2, 1) = -3
KirshMask(3, 0, 2) = -3: KirshMask(3, 1, 2) = -3: KirshMask(2, 2, 2) = -3
'K4
KirshMask(4, 0, 0) = 5: KirshMask(4, 1, 0) = -3: KirshMask(4, 2, 0) = -3
KirshMask(4, 0, 1) = 5: KirshMask(4, 1, 1) = 0: KirshMask(4, 2, 1) = -3
KirshMask(4, 0, 2) = 5: KirshMask(4, 1, 2) = -3: KirshMask(4, 2, 2) = -3
'K5
KirshMask(5, 0, 0) = -3: KirshMask(5, 1, 0) = -3: KirshMask(5, 2, 0) = -3
KirshMask(5, 0, 1) = 5: KirshMask(5, 1, 1) = 0: KirshMask(5, 2, 1) = -3
KirshMask(5, 0, 2) = 5: KirshMask(5, 1, 2) = 5: KirshMask(5, 2, 2) = -3
'K6
KirshMask(6, 0, 0) = -3: KirshMask(6, 1, 0) = -3: KirshMask(6, 2, 0) = -3
KirshMask(6, 0, 1) = -3: KirshMask(6, 1, 1) = 0: KirshMask(6, 2, 1) = -3
KirshMask(6, 0, 2) = 5: KirshMask(6, 1, 2) = 5: KirshMask(6, 2, 2) = 5
'K7
KirshMask(7, 0, 0) = -3: KirshMask(7, 1, 0) = -3: KirshMask(7, 2, 0) = -3
KirshMask(7, 0, 1) = -3: KirshMask(7, 1, 1) = 0: KirshMask(7, 2, 1) = 5
KirshMask(7, 0, 2) = -3: KirshMask(7, 1, 2) = 5: KirshMask(7, 2, 2) = 5
'Constant: Treshold value for response to Edge detector

'THis is different here
Const BLACK = &HFFFFFFFF
Const WHITE = 0


'begin on left corner of screen
'work all the way to the bottom

For Y = 1 To CornerY
    For X = 1 To CornerX
                
        'do some matrix math for each channel
        'add side rows and top/bottom colums, compare with center
        
        'compare with all masks
        For MASK = 0 To 7
            For j = 0 To 2
                For i = 0 To 2
                        K(MASK) = K(MASK) + (MapImage((X + i), (Y + j)) * KirshMask(MASK, i, j))
                Next i
            Next j
        Next MASK
            
    'get 2 largest responses to mask
    '
    'clear values before entering loop
    MaskResponse(1) = 0
    MaskResponse(2) = 0
            
    'Loop thru all values
    For i = 0 To 7
        'Compare current max to this value
        If K(i) > MaskResponse(1) Then
            'if this number is bigger than MAX, store MAX in secondmax and make this new
            'value the MAX
            MaskResponse(2) = MaskResponse(1)
            MaskResponse(1) = K(i)
        'if not bigger than MAX, is it bigger than secondmax?
        ElseIf K(i) > MaskResponse(2) Then MaskResponse(2) = K(i)
        End If
        
        'clear after each check
        K(i) = 0
    Next i
        
    'test edge detector for response
    
    If MaskResponse(1) > TreshHoldResponse And MaskResponse(2) > TreshHoldResponse Then
    'pixel is an edge
        MapEdge(X, Y) = 0
    Else
    'pixel is not an edge
        MapEdge(X, Y) = 1
    End If

    Next X
Next Y
        
End Sub
Private Sub EdgeVertical(XP As Long, Yp As Long)
Dim X As Long
Dim Y As Long
Dim side_a As Long
Dim side_b As Long
Dim p As Long


'edge by Delta
For Y = Yp - 3 To Yp - 2
    side_a = side_a + MapImage(XP, Y)
Next Y

For Y = Yp - 1 To Yp
    side_b = side_b + MapImage(XP, Y)
Next Y

If side_a <> side_b Then
    p = 0
Else
    p = 1
End If
    
For Y = Yp - 3 To Yp
    Buffer_V(XP, Y) = p
Next Y

End Sub
Public Sub Smooth(Optional TimesToErode As Long = 0)
'Erase all non-corner pixels of objects
'

Dim X As Long
Dim Y As Long
Dim Index As Long
Dim i As Long
Dim ii As Long
Dim PixelInMatrix As Long
Dim gap As Single
Dim SEARCHRADIUS As Integer
Dim MinimumPixel As Integer
Dim ObjectNumber As Long
Dim Connectivity As Long

Dim OutputArray(MAX_X, MAX_Y) As Long
Dim Boundary As Long
Dim Count As Long
Dim Value As Long

Dim UNMODIFIED As Boolean

Dim iterations

'Start at the top left hand corner of frame
'and repeat as many times as TimesToErode
Do
    UNMODIFIED = True
    For X = 0 To CornerX
        For Y = 0 To CornerY
                If MapEdge(X, Y) = LAND Then
                    'check if pixel is connected to others
                    'by counting all water (background) pixels surrounding this pixel
                    Connectivity = 0
                    For i = 0 To 3
                        Connectivity = Connectivity + _
                        (PixelValue(X, Y, s(i)) - (PixelValue(X, Y, s(i)) * PixelValue(X, Y, s(i) + 1) * PixelValue(X, Y, s(i) + 2)))
                    Next i
                    
                    Count = 0
                    For i = 1 To 8
                        Value = PixelValue(X, Y, i)
                        Count = Count + Abs(Value - 1)
                    Next i
                    
                    If Connectivity = 1 And (Count >= 2 And Count <= 6) Then
                        If (PixelValue(X, Y, 1) = WATER Or PixelValue(X, Y, 3) = WATER Or PixelValue(X, Y, 5) = WATER) _
                        And (PixelValue(X, Y, 3) = WATER Or PixelValue(X, Y, 7) = WATER Or PixelValue(X, Y, 5) = WATER) Then
                            OutputArray(X, Y) = WATER
                            UNMODIFIED = False
                        Else
                            OutputArray(X, Y) = LAND
                        End If
                        
                    Else
                        OutputArray(X, Y) = LAND
                    End If
                Else
                    OutputArray(X, Y) = WATER
                End If
        Next Y
    Next X
    
    'now transfer copybuffer contents to original buffer and overwrite its contents
    'only the boundary pixels are kept

    'if no erosion req, simply copy buffer as is to main pixel buffer
    For X = 0 To CornerX
        For Y = 0 To CornerY
            MapEdge(X, Y) = OutputArray(X, Y)
        Next Y
    Next X
    
Exit Sub
    If UNMODIFIED Then Exit Do
    UNMODIFIED = True
    
    For X = 0 To CornerX
        For Y = 0 To CornerY
                If MapEdge(X, Y) = LAND Then
                    'check if pixel is connected to others
                    'by counting all water (background) pixels surrounding this pixel
                    Connectivity = 0
                    
                    For i = 0 To 3
                        Connectivity = Connectivity + _
                        (PixelValue(X, Y, s(i)) - (PixelValue(X, Y, s(i)) * PixelValue(X, Y, s(i) + 1) * PixelValue(X, Y, s(i) + 2)))
                    Next i
                   
                    
                    Count = 0
                    For i = 1 To 8
                        Value = PixelValue(X, Y, i)
                        Count = Count + Abs(Value - 1)
                    Next i
                    
                    If Connectivity = 1 And (Count >= 2 And Count <= 6) Then
                        If (PixelValue(X, Y, 3) = WATER Or PixelValue(X, Y, 1) = WATER Or PixelValue(X, Y, 7) = WATER) _
                        And (PixelValue(X, Y, 1) = WATER Or PixelValue(X, Y, 7) = WATER Or PixelValue(X, Y, 5) = WATER) Then
                            OutputArray(X, Y) = WATER
                            UNMODIFIED = False
                        Else
                            OutputArray(X, Y) = LAND
                        End If
                        
                    Else
                        OutputArray(X, Y) = LAND
                    End If
                Else
                    OutputArray(X, Y) = WATER
                End If
                
        Next Y
    Next X

    'now transfer copybuffer contents to original buffer and overwrite its contents
    'only the boundary pixels are kept

    'if no erosion req, simply copy buffer as is to main pixel buffer
    For X = 0 To CornerX
        For Y = 0 To CornerY
            MapEdge(X, Y) = OutputArray(X, Y)
        Next Y
    Next X
Loop Until UNMODIFIED
End Sub
Private Function PixelValue(X As Long, Y As Long, N As Long) As Long
Dim xi As Long
Dim yi As Long

If N > 8 Then
    N = N - 8
End If
xi = X + K_X(N)
yi = Y + K_Y(N)
If xi > CornerX Then xi = CornerX
If xi < 0 Then xi = 0
If yi > CornerY Then yi = CornerY
If yi < 0 Then yi = 0


PixelValue = IIf(MapEdge(xi, yi) = LAND, 0, 1)

End Function
Private Sub SnapToRoute(ByRef X_to_change As Long, ByRef Y_to_change As Long)
'Snaps starting or ending coordinates to controur (route)
'works by searching for closest coordinates from site to route track (contour of map)
'
Dim X As Long
Dim Y As Long
Dim SearchPattern As Long
Dim Search_Radius As Long

Dim Found As Boolean

Const Radius = 1


Do
    X = X_to_change
    Y = Y_to_change
    Do
        'start at origin
        X = X + X_Increment(SearchPattern)
        Y = Y + Y_increment(SearchPattern)
        If ValidContourPixel(X, Y) Then
            Found = True
        End If
        Search_Radius = Search_Radius + 1
    Loop While (Not Found) And Search_Radius <= Radius
    SearchPattern = SearchPattern + 1
    SEARCHRADIUS = 0
Loop While (Not Found) And SearchPattern <= 5

'found, get new coords
If Found Then
    X_to_change = X
    Y_to_change = Y
End If

End Sub
Private Function ValidContourPixel(X As Long, Y As Long) As Boolean
'valid pixel if at least one "water" pixel is in direct contact
Dim direction As Long
Dim i As Long
Dim Found As Boolean
Dim nx As Long
Dim ny As Long

If ContourMap(X, Y) = LAND Then
    i = 1
    Do
        nx = X + K_X(i)
        ny = Y + K_Y(i)
             'validate
             If (nx >= 0 And nx <= CornerX) And (ny >= 0 And ny <= CornerY) Then
                 If ContourMap(nx, ny) = WATER Then
                     Found = True
                 End If
             End If
         i = i + 1
    Loop Until i > 8 Or Found
End If

ValidContourPixel = Found

End Function
Private Function SearchAroundPixel(X_Center As Long, Y_Center As Long, PreviousDirection As Long, gap As Long) As Long
'search around pixel using gap factor GAP
'it applies current direction with GAP to check if anything beyond pixel exist
'if not it changes direction
'if no "LAND BORDER" exists around the pixel, within GAP distance, it returns a -1 (failed to encounter more land border)

Dim X As Long
Dim Y As Long
Dim direction As Long
Dim Found As Boolean
Dim d(7) As Integer
Dim i As Long
Dim CountValidDirections As Long
Dim DirectionToFollow As Long
Dim OX As Long
Dim OY As Long
Dim R As Integer
Dim GapCounter As Long

'use previous direction as start unless no direction existed... in that case, start at random
If PreviousDirection = 0 Then
    RND_MT.Randomize Timer
    PreviousDirection = RND_MT.Generate_Number(8) + 1
End If

direction = PreviousDirection
i = 1
DirectionToFollow = -1
'Internally, gap is a base 1 variable
GapCounter = gap + 1
Do
        X = X_Center: Y = Y_Center
        'allow for gaps
        Found = False
        'move towards new direction
        X = X + (K_X(direction) * GapCounter): Y = Y + (K_Y(direction) * GapCounter)
    
        'test pixel
        If ValidContourPixel(X, Y) Then
            If DirectionToFollow = -1 Then
                DirectionToFollow = direction
            End If
            Found = True
        End If
        'near-target test
        'if target is near, move towards it!
        If ContourMap(X, Y) = TARGET Then
            DirectionToFollow = direction
            Found = True
        End If
        
        If Found Then
            d(CountValidDirections) = direction
            CountValidDirections = CountValidDirections + 1
        End If
        
        direction = direction + 1
        i = i + 1
        'validate direction
        If direction > 8 Then direction = 1

Loop Until i > 8

'more than one branching/valid directions? (i.e. connected to more than one section of outline)
If CountValidDirections > 1 Then
    'hatch a node
    With Node(NodeNumber)
        .X = X_Center
        .Y = Y_Center
        .Count = CountValidDirections
        .ContourSize = ContourIndex
        .gap = GapCounter
        Randomize Timer
        R = CInt(Rnd(10) * (CountValidDirections - 1))
        For i = 0 To CountValidDirections - 1
            .direction(i) = d(R)
            R = R + 1
            If R >= CountValidDirections Then R = 0
        Next i
    End With
    'update pointer
    NodeNumber = NodeNumber + 1
    If NodeNumber > MAXNodeNUMBER Then Exit Function
End If


If CountValidDirections = 0 Then
    DirectionToFollow = -1
End If


SearchAroundPixel = DirectionToFollow

'''''''''''''''''''''''''
End Function
Private Sub EnlargeTarget(X As Long, Y As Long)
'enlarge/propagate pixel x,y into surrounding land/water map
'so that target is detected with 100% accuracy
'
Dim Pixel As Long
Dim xi As Long
Dim yi As Long
Dim i As Long
Dim ii As Long
Const Radius = 4

'propaget North, South, East and West by RADIUS*2 amount
'not circle but box
xi = X
For i = Radius * -1 To Radius
    xi = xi + i
    yi = Y
    For ii = Radius * -1 To Radius
        yi = yi + ii
        ContourMap(xi, yi) = TARGET
    Next ii
Next i

End Sub
Public Function FollowContour(StartPixel_X As Long, StartPixel_Y As Long, EndPixel_X As Long, EndPixel_Y As Long, ByRef PixelX() As Long, ByRef PixelY() As Long, Site_To As Long) As Boolean
'Entry:Start and End pixels of route
'      An array too put the information of which pixels make the most direct route should be provided.
'
'Follows contour to find closest route to PIXEL END from PIXEL START
'
Dim X As Long
Dim Y As Long
Dim PX As Long
Dim PY As Long
Dim direction As Long
Dim Found As Boolean
Dim ChangeDirection As Boolean
Dim i As Long
Dim NextDirection As Long
Dim SX As Long
Dim SY As Long
Dim eX As Long
Dim eY As Long
Dim DistanceToTarget As Single
Dim ExploreNode As Long
Dim exl As Long
Dim N As Long

Const GAP_LIMIT = 0
Dim gap As Long
Dim Current_Gap As Long


SX = StartPixel_X: SY = StartPixel_Y
eX = EndPixel_X: eY = EndPixel_Y
SizeOfContour = 0
ContourIndex = 0
NodeNumber = 0

'copy contour edge map to an internal array
'(speeds up things + allows to start over)
For X = 0 To CornerX
    For Y = 0 To CornerY
        ContourMap(X, Y) = MapEdge(X, Y)
    Next Y
Next X

'Search for the closest start pixel in the map
SnapToRoute SX, SY
SnapToRoute eX, eY
'make the target big enough to cover the land-water boundary and beyond
EnlargeTarget eX, eY

'define previous point
X = SX
Y = SY


'follow route until end pixel is reached
Do
    Do
        'keep moving in choosen direction
        X = X + (K_X(direction) * Current_Gap): Y = Y + (K_Y(direction) * Current_Gap)
        'If x = 190 Then Stop
        Current_Gap = 1
        'if pixel belongs to contour then,-+
        If ValidContourPixel(X, Y) Then
            'set gap value for algorithm
            'start at 0 gap
            gap = 0
            'check surrounding pixels
            Do
                NextDirection = SearchAroundPixel(X, Y, direction, gap)
                
                'widen gap
                gap = gap + 1
                Current_Gap = gap
            Loop Until gap > GAP_LIMIT Or NextDirection <> -1
    
            'if no more "land" pixels, go back, and change direction
            If NextDirection = -1 Then
                'erase pixel
                ContourMap(X, Y) = SET_BITS
                ChangeDirection = True
            Else
                'store
                PixelX(ContourIndex) = X
                PixelY(ContourIndex) = Y
                ContourIndex = ContourIndex + 1
                'erase visited pixels from map
                ContourMap(X, Y) = SET_BITS
                'Form1.Picture1.PSet (x, y), vbRed
                ChangeDirection = False
                 direction = NextDirection
            End If
        Else
            NextDirection = -1
        End If
        If Not LandBetweenPositions(X, Y, Receiver.X(Site_To), Receiver.Y(Site_To)) Then Found = True
        If ContourMap(X, Y) = TARGET Then Found = True
    Loop Until Found Or ChangeDirection Or NextDirection = -1
    
    'change direction?
    If Not Found And NextDirection = -1 Then
        'go to last node
        ExploreNode = NodeNumber - 1
        'no gap
        Current_Gap = 1
        If ExploreNode >= 0 Then
            ChangeDirection = False
            With Node(ExploreNode)
                'explore node branches
                If .Count = 0 Then Stop
                X = .X: Y = .Y
                NextDirection = .direction(.Count - 1)
                ChangeDirection = True
                .Count = .Count - 1
                ContourIndex = .ContourSize
                Current_Gap = .gap
                'delete node if this is the last branch
                If .Count = 0 Then exl = exl + 1: NodeNumber = NodeNumber - 1
            End With
        End If
    End If

    
        'if no more branches to explore and no new direction, exit... can't explore any more
        If ExploreNode = -1 And NextDirection = -1 Then Exit Do
        
    
    
    'reposition if gap
    If GAP_X > 0 And GAP_Y > 0 Then X = GAP_X: Y = GAP_Y
        
    direction = NextDirection
    
    'distance to target
    DistanceToTarget = Sqr((X - eX) ^ 2 + (Y - eY) ^ 2)

Loop Until Found
'mark end
SizeOfContour = ContourIndex

'rts
FollowContour = Found
End Function
Private Sub Class_Initialize()
Dim i As Long
For i = 0 To MAX_RECEIVERS
    With ReceiverConnectionToCorridor(i)
        .SelectedDirection = -1
    End With
Next i

X_Increment(5) = 0: Y_increment(5) = -1
X_Increment(6) = 1: Y_increment(6) = -1
X_Increment(7) = 1: Y_increment(7) = 0
X_Increment(0) = 1: Y_increment(0) = 1
X_Increment(1) = 0: Y_increment(1) = 1
X_Increment(2) = -1: Y_increment(2) = 1
X_Increment(3) = -1: Y_increment(3) = 0
X_Increment(4) = -1: Y_increment(4) = -1

'define cells around central pixel
K_X(0) = 0: K_X(1) = 1: K_X(2) = 1: K_X(3) = 0: K_X(4) = -1: K_X(5) = -1: K_X(6) = -1: K_X(7) = 0: K_X(8) = 1
K_Y(0) = 0: K_Y(1) = 0: K_Y(2) = -1: K_Y(3) = -1: K_Y(4) = -1: K_Y(5) = 0: K_Y(6) = 1: K_Y(7) = 1: K_Y(8) = 1
'define search set
s(0) = 1: s(1) = 3: s(2) = 5: s(3) = 7

'color pallette for phases
'Phase:
'1) Night: from nautical dusk to nautical dawn
'2) Dawn: from nautical dawn to sunrise
'3) Day AM: from sunrise to the midpoint between sunrise and sunset
'4) Day PM: from the midpoint between sunrise and sunset to sunset
'5) Dusk, from sunset to nautical dusk
Color_Pallette(0) = vbBlack 'not valid in this version
Color_Pallette(1) = &H303030
Color_Pallette(2) = &HBFBFBF
Color_Pallette(3) = &HAAF0AA 'light green

Color_Pallette(4) = &H505050
Color_Pallette(5) = &HBFBFBF

Slot_Size = 10 'slots for bins in diel cycle. Default 10 mins / bin
End Sub
Public Property Let ChangeColorPallette(p As Integer, c As Long)
Color_Pallette(p) = c
End Property
Public Sub LoadColorPalletteDielWindow()

Dim i As Long

For i = 1 To 4
    With frmDayLightCycle
        .shpColor(i).FillColor = Color_Pallette(i)
    End With
Next i

End Sub
Private Function ExtractVersion(s As String) As Single
Dim v As String
Dim p As Long
Dim L As Long

L = Len(s)
p = InStr(1, s, ".")

If p > 1 Then
    v = Mid$(s, p - 1, 4)
End If

ExtractVersion = CSng(v)


End Function

Public Sub ClearDielCycle()
Dim i As Long

Const MINUTES_PER_DAY = 1440

If DielCycleFormIsLoaded Then frmDayLightCycle.picGraph.Cls
For i = 0 To MINUTES_PER_DAY
    TimeSlot(i) = 0
    PhaseSlot(i) = 0
    AlternatePhaseSlot(i) = 0
Next i
TimeMax = 0

End Sub
Public Sub DrawDielCycleNOW()
Dim Phase As Long
Dim Phase_Time As Long
Dim ScaleFactor_X As Single
Dim ScaleFactor_Y As Single
Dim DayOfDetection As Long
Dim Height As Long
Dim X As Long
Dim LineColor As Long
Dim i As Long
Dim Max As Long
Dim Bin As Long
Dim WindowHeight As Long
Dim ii As Long
Dim Width As Single
Dim ChangeScale As Boolean
Dim timeofday As Long


Const TICK_HEIGHT = 3
'get day
If TimeMax = 0 Then Exit Sub

With frmDayLightCycle.picGraph
    'get scale factor
    ScaleFactor_X = (.ScaleWidth - Shift) / 1440
    WindowHeight = .ScaleHeight - Shift
    ScaleFactor_Y = WindowHeight * 0.9
End With


'redraw histogram
frmDayLightCycle.picGraph.Cls
Width = (ScaleFactor_X * (Slot_Size))
For i = 0 To (CLng(SLOTS / Slot_Size) - 1)
   
    X = Fix(i * (Width)) + Shift
    Height = CLng((TimeSlot(i) / TimeMax) * ScaleFactor_Y)
    frmDayLightCycle.picGraph.Line (X, WindowHeight)-(X + Width, WindowHeight - Height), Color_Pallette(PhaseSlot(i)), BF
    frmDayLightCycle.picGraph.Line (X, WindowHeight)-(X, WindowHeight - TICK_HEIGHT), Not Color_Pallette(PhaseSlot(i)), BF
    
    'paint line pattern if two diel phases in same slot
    If AlternatePhaseSlot(i) <> 0 Then
        For ii = 0 To Height Step 5
            frmDayLightCycle.picGraph.Line (X, WindowHeight - ii)-(X + Width, WindowHeight - ii), Color_Pallette(AlternatePhaseSlot(i))
        Next ii
    End If
Next i

'draw the axis

'vertical line
frmDayLightCycle.picGraph.Line (Shift, WindowHeight)-(Shift, 0), vbBlack
'horizontal line
frmDayLightCycle.picGraph.Line (Shift, WindowHeight)-(frmDayLightCycle.picGraph.ScaleWidth, WindowHeight), vbBlack

'change scale
DielCycleDrawAxis WindowHeight, Shift
End Sub
Private Sub DielCycleDrawAxis(WH As Long, s As Long)
Dim X As Long
Dim timeofday As Long
Dim L As Integer
Dim TimeLabel As String
Const NumberOfLabelsOnX = 8
Dim t As Long
Dim Bin As Long
Dim SX As Single
Dim i As Long


frmDayLightCycle.picGraph.CurrentX = 0
frmDayLightCycle.picGraph.CurrentY = 5
frmDayLightCycle.picGraph.Print Str$(TimeMax) & "--"
frmDayLightCycle.picGraph.CurrentX = 0
frmDayLightCycle.picGraph.CurrentY = ((WH) / 2)
frmDayLightCycle.picGraph.Print Str$(Fix(TimeMax / 2)) & "--"
SX = (frmDayLightCycle.picGraph.Width) / 1440
For i = 1 To NumberOfLabelsOnX
    timeofday = timeofday + (24 / NumberOfLabelsOnX)
    t = 60 * timeofday
    X = (t + 1) * SX
    If timeofday > 0 Then TimeLabel = Str$(timeofday): L = Len(TimeLabel)
    X = X
    frmDayLightCycle.picGraph.CurrentX = X '- (L * 5)
    frmDayLightCycle.picGraph.CurrentY = frmDayLightCycle.picGraph.ScaleHeight - s
    If TimeLabel <> "" Then frmDayLightCycle.picGraph.Print TimeLabel
Next i

End Sub
Public Function ReturnTimeOfBin(X As Single, w As Single) As Long
Dim t As Long
Dim Bin As Long
Dim SX As Single

X = X - Shift

'get scale factor
SX = (w - Shift) / 1440
Bin = X / SX
t = Bin
ReturnTimeOfBin = t
End Function
Public Sub DrawDielCycle(Optional DRAW_AS_YOU_GO As Boolean = True)
'Draws a histogram showing diel phases for detections
'Argument: none in the call but STAMP should contain the stamp used to draw the histogram.
'Need to call until all stamp drawn
Static Phase As Long
Dim Phase_Time As Long
Dim ScaleFactor_X As Single
Dim ScaleFactor_Y As Single
Dim DayOfDetection As Long
Dim Height As Long
Dim X As Long
Dim LineColor As Long
Dim i As Long
Dim Max As Long
Dim Bin As Long
Dim WindowHeight As Long
Dim ii As Long
Dim Width As Single
Dim G As Integer
Dim timeofday As Long
Dim R As Integer



'is receiver visible? If not visible, assume this stamps are not to be used for the
'histogram
'DO NOT DRAW IF RECEIVER IS NOT VISIBLE
'HOWEVER USE GROUP RECEIVER IF IN GROUP
G = Receiver.GroupNumber(Stamp.Site)
If G <> 0 Then R = Receiver.GroupReceiverNumber(G)
If Receiver_Table.Invisible(R) Then Exit Sub

'also check if valid
If Not Stamp.Valid Then Exit Sub

'get day
DayOfDetection = Stamp.Date

With frmDayLightCycle.picGraph
    'get scale factor
    ScaleFactor_X = (.ScaleWidth - Shift) / 1440
    WindowHeight = .ScaleHeight - Shift
    ScaleFactor_Y = WindowHeight * 0.9
End With

'Get bin, and accumulate
Bin = Fix(Stamp.Time / Slot_Size)
TimeSlot(Bin) = TimeSlot(Bin) + 1

If Phase = 0 Then Phase = 1
'choose pen color
Do Until PhotoPeriodCalculator.StampIsInPhase(Phase)
    Phase = Phase + 1
    If Phase > 5 Then Phase = 1
Loop

'merge colors if not the same.
If PhaseSlot(Bin) = 0 Then
    PhaseSlot(Bin) = Phase
Else
    If PhaseSlot(Bin) <> Phase Then AlternatePhaseSlot(Bin) = Phase
End If

'get max
If TimeSlot(Bin) > TimeMax Then
    TimeMax = TimeSlot(Bin)
    If DRAW_AS_YOU_GO Then DrawDielCycleNOW
Else
    If DRAW_AS_YOU_GO Then
        'Get height and position
        Width = (ScaleFactor_X * Slot_Size)
        X = Fix(Bin * (Width)) + Shift
        Height = CLng((TimeSlot(Bin) / TimeMax) * ScaleFactor_Y) 'draw a point in the histogram
        frmDayLightCycle.picGraph.Line (X, WindowHeight)-(X + Width, WindowHeight - Height), Color_Pallette(PhaseSlot(Bin)), BF
        If AlternatePhaseSlot(Bin) <> 0 Then
                For ii = 0 To Height Step 10
                    frmDayLightCycle.picGraph.Line (X, WindowHeight - ii)-(X + Width, WindowHeight - ii), Color_Pallette(AlternatePhaseSlot(Bin))
                Next ii
        End If
    End If
End If

End Sub
Public Sub CopyDielCycleToCilpBoard()
'Copies content of loaded diel cycle histogram to clipboard
'formatted for TAB DELIM
Dim Bin As Long
Dim BinValueHour As String
Dim Count As String
Dim Concatenated As String
Dim Phase As String

'clear clipboard
Clipboard.Clear

'Phase:
'1) Night: from nautical dusk to nautical dawn
'2) Dawn: from nautical dawn to sunrise
'3) Day AM: from sunrise to the midpoint between sunrise and sunset
'4) Day PM: from the midpoint between sunrise and sunset to sunset
'5) Dusk, from sunset to nautical dusk

'Header for table
Concatenated = "Time," & Space(15 - Len("time")) & "Counts," & Space(15 - Len("counts")) & "Phase" & Chr$(13) & Chr$(10)
For Bin = 0 To (CLng(SLOTS / Slot_Size) - 1)
    BinValueHour = Format((Slot_Size * Bin) / 60, "0.00")
    Count = Str$(TimeSlot(Bin))
    If PhaseSlot(Bin) <> 0 Then Phase = Choose(PhaseSlot(Bin), "Night", "Dawn", "AM", "PM", "Dusk") Else Phase = "No counts on bin."
    If AlternatePhaseSlot(Bin) <> 0 Then
        Phase = Phase & " and " & Choose(AlternatePhaseSlot(Bin), "Night", "Dawn", "AM", "PM", "Dusk")
    End If
    Concatenated = Concatenated & BinValueHour & "," & Space(15 - Len(BinValueHour)) & Count & "," & Space(15 - Len(Count)) & Phase & Chr$(13) & Chr$(10)
    Clipboard.SetText Concatenated
Next Bin


End Sub


Public Sub DrawLineWithArrow(Pic As PictureBox, X1 As Long, X2 As Long, Y1 As Long, Y2 As Long, Optional c As Long = vbBlue)
Dim Y3, Y4 As Long
Dim X3, X4 As Long
Dim LineAngle, EndAngle1, EndAngle2 As Single
Dim DeltaX, DeltaY As Long



 ' - Draw Line
Pic.Line (X1, Y1)-(X2, Y2), c

'get sign for delta x and delta y
DeltaY = Y2 - Y1
DeltaX = X2 - X1

' - Get Angle of Line
If DeltaX <> 0 Then
    LineAngle = Atn((Y2 - Y1) / (X2 - X1))
Else
    LineAngle = Ninety_In_Radians
End If

' - Get Angle for lines to make up arrow head
EndAngle1 = LineAngle + Angle
EndAngle2 = LineAngle - Angle
    

' - Get EndPoints of Arrow Head Lines
If DeltaX <> 0 Then
    Y3 = Y2 - ArrowHeadLength * Sin(EndAngle1) * Sgn(DeltaX)
    Y4 = Y2 - ArrowHeadLength * Sin(EndAngle2) * Sgn(DeltaX)
    X3 = X2 - ArrowHeadLength * Cos(EndAngle1) * Sgn(DeltaX)
    X4 = X2 - ArrowHeadLength * Cos(EndAngle2) * Sgn(DeltaX)
Else
    Y3 = Y2 - ArrowHeadLength * Sin(EndAngle1) * Sgn(DeltaY)
    Y4 = Y2 - ArrowHeadLength * Sin(EndAngle2) * Sgn(DeltaY)
    X3 = X2 + ArrowHeadLength * Cos(EndAngle1)
    X4 = X2 + ArrowHeadLength * Cos(EndAngle2)
End If

' - Draw Arrow Head
Pic.Line (X2, Y2)-(X3, Y3), c
Pic.Line (X2, Y2)-(X4, Y4), c
End Sub
Public Sub ExportDensityPlotDataToClipBoard(DensityData() As Long)
Dim Bin As Integer
Dim ReceiverNumber As Integer
Dim ReceiverName As String
Dim RN As String
Dim Counts As String
Dim Concatenated As String
Dim G As Integer
Dim GroupAlreadyCopied(MAX_GROUPS) As Boolean

'clear clipboard
Clipboard.Clear

'header
Concatenated = "Receiver Number," & Space(20 - Len("Receiver Number")) & "Receiver Name," & Space(20 - Len("Receiver Name")) & " Detections"

'receivers one by one
For ReceiverNumber = 1 To Receiver.TotalReceivers
    'copy only visible receivers
    'copy groups only once
    If Receiver.Detection_TTL(ReceiverNumber) <> 0 Then
        G = Receiver.GroupNumber(ReceiverNumber)
        If Not GroupAlreadyCopied(G) Then
            If G <> 0 Then GroupAlreadyCopied(G) = True
            RN = Str$(ReceiverNumber)
            Counts = DensityData(ReceiverNumber)
            ReceiverName = Receiver.ID(ReceiverNumber)
            If Len(ReceiverName) > 15 Then ReceiverName = Left(ReceiverName, 15)
            Concatenated = Concatenated & Chr$(13) & Chr$(10) & RN & "," & Space(20 - Len(RN)) & ReceiverName & "," & Space(20 - Len(ReceiverName)) & Counts
        End If
    End If
Next ReceiverNumber

'copy to clipboard
Clipboard.SetText Concatenated


End Sub
Public Sub DrawDensityPlotForReceivers(MapPicture As PictureBox, Max As Long, DensityData() As Long, Optional MarkerSize As Long = 1, Optional WIPE_CANVAS As Boolean = True)

'Override calculated max setting?
Set_Max = Max

If ColorScale.AutoSet Then
    Max = Set_Max
    ColorScale.Max = Max
Else
    Max = ColorScale.Max
End If

ColorScale.ShowPlot MapPicture, DensityData(), MarkerSize, WIPE_CANVAS
End Sub
Public Sub DrawReceiverWithDensity(ReceiverNumber As Integer, TotalDensity As Long)
Dim intensity As Long
Dim BinSize As Single
Dim Bin As Integer
Dim Color As Long
Dim Intensity_Value As Integer

Color = ColorScale.ReturnColorOfBin(TotalDensity)
Receiver.DrawReceiver Form1.Picture1, ReceiverNumber, LARGE_MARKER, Color

End Sub
Public Sub StoreFishCorridor()
'transfers dib section to the fish corridor matrix
Dim v As Boolean
MapDib.ClearUp
v = MapDib.Create(Form1.Picture1.ScaleWidth, Form1.Picture1.ScaleHeight)
v = MapDib.CreateFromImage(Form1.Picture1, Form1.Picture1.hDC)
MapDib.TransferToFishCorridorMatrix
FishCorridorStored = True
End Sub
Public Sub ShowUnConnectedReceiverTable(List As ListBox)
'shows ids of receivers that are not connected
'
Dim R As Integer
Dim ReceiverName As String
Dim Group As Integer
Dim GroupDone(MAX_GROUPS) As Boolean
Dim ReceiverNumber As Integer
'clear list before adding anything to it
List.Clear

For R = 1 To Receiver.TotalReceivers
    Group = Receiver.GroupNumber(R)
    ReceiverNumber = R
    'part of a group? check which receiver has the connection info
    If Group Then
        If Not GroupDone(Group) Then
            ReceiverNumber = GetCorrectReceiverNumber(Group)
            If ReceiverNumber = -1 Then
                List.AddItem Receiver.ID(R)
                List.ItemData(List.ListCount - 1) = CLng(R)
            End If
        End If
    Else
        'not part of a group?
        'check connection info
        If ReceiverConnectionToCorridor(ReceiverNumber).SelectedDirection = -1 Then
            List.AddItem Receiver.ID(R)
            List.ItemData(List.ListCount - 1) = CLng(R)
        End If
    End If
Next R
End Sub
Private Function GetCorrectReceiverNumber(G As Integer) As Integer
'gets which receiver number has the connection information
'returns -1 if no connection
Dim R As Long
Dim ReturnValue As Integer

ReturnValue = -1
For R = 1 To Receiver.TotalReceivers
    If Receiver.GroupNumber(R) = G Then
        If ReceiverConnectionToCorridor(R).SelectedDirection <> -1 Then
            ReturnValue = R
        End If
    End If
Next R
GetCorrectReceiverNumber = ReturnValue
End Function
Public Sub ConnectReceiversToCorridor()
Dim GroupDone(MAX_GROUPS) As Boolean
Dim Valid As Boolean
Dim Group As Integer
Dim R As Integer
Dim X As Integer
Dim Y As Integer
Dim SearchPatternResults(MAX_RECEIVERS) As Star
Dim DirectionPattern As Integer
Dim Distance As Integer
Dim DirectionSelected As Integer
Dim DistanceMin As Integer
Dim i As Long
Const MaximumDistanceFromCorridorAllowed = 2048


'For each receiver
For R = 1 To Receiver.TotalReceivers
    'Check if its in a group
    Group = Receiver.GroupNumber(R)
    If Group Then
        'If in group, take X/Y and skip all others in group
        If GroupDone(Group) Then
            Valid = True
        Else
            Valid = False
            GroupDone(Group) = True
            ReceiverConnectionToCorridor(R).SelectedDirection = -1
        End If
    Else
        Valid = True
    End If
    
    If Valid Then
        X = CInt(Receiver.X(R))
        Y = CInt(Receiver.Y(R))
        'Search around at a 90 angle first, go clockwise
        DirectionPattern = 1
        DirectionSelected = -1
        DistanceMin = MaximumDistanceFromCorridorAllowed
        Do
            Distance = SearchAroundReceiver(R, DirectionPattern)
            'Choose shortest path
            If Distance <> 0 Then
                If Distance < DistanceMin Then
                    DirectionSelected = DirectionPattern
                    DistanceMin = Distance
                End If
            End If
            'Save results of each search
            SearchPatternResults(R).direction(DirectionPattern) = Distance
            DirectionPattern = DirectionPattern + 1
        Loop Until DirectionPattern > 8
        'Save connection info in a separate array and store in object for use later
        With ReceiverConnectionToCorridor(R)
            .SelectedDirection = DirectionSelected
            If DirectionSelected <> -1 Then
                .X = X + (DistanceMin * K_X(DirectionSelected))
                .Y = Y + (DistanceMin * K_Y(DirectionSelected))
            End If
        End With
    End If
Next R
End Sub
Public Function LandLockedReceiver(ByVal R As Integer) As Boolean
'Given a receiver R, returns a boolean, true if receiver is landlocked, false if it isn't
'
'
Dim X As Long
Dim Y As Long
Dim ReturnValue As Boolean
Dim direction As Long
Dim i As Long
Dim Found As Boolean
Dim nx As Long
Dim ny As Long
Dim WaterPixels As Long
Dim gap As Long



'at least 4 water pixels!
Const Threshold = 4
Const GAP_LIMIT = 2

'start with gap 3 because receiver is drawn in map as a 2 pixel square
gap = 3

'get coords
X = Receiver_Table.X(R)
Y = Receiver_Table.Y(R)

'valid if surrounded by water (direct contact, no gap)
i = 1
Do
    nx = X + (K_X(i) * gap)
    ny = Y + (K_Y(i) * gap)
         'validate
         If (nx >= 0 And nx <= CornerX) And (ny >= 0 And ny <= CornerY) Then
             If MapImage(nx, ny) <> LAND Then
                 WaterPixels = WaterPixels + 1
             End If
         End If
     i = i + 1
Loop Until i > 8 Or WaterPixels >= Threshold

'min threhsold has to be exceeded, otherwise receiver is locked in land
If WaterPixels >= Threshold Then ReturnValue = False Else ReturnValue = True

LandLockedReceiver = ReturnValue

End Function

Private Function SearchAroundReceiver(R As Integer, direction As Integer) As Integer
Dim i As Long
Dim Found As Boolean
Dim X As Long
Dim Y As Long
Dim Distance As Integer

'searches in the direction DIRECTION for a connection to the fish corridor
'if land is found, aborts and returns 0
'if the edge of screen is reached and nothing is found, aborts and returns 0
'returns distance to corridor

X = Receiver.X(R)
Y = Receiver.Y(R)
Do
    i = i + 1
    X = X + K_X(direction)
    Y = Y + K_Y(direction)
    
    'test if valid
    If X >= CornerX Or X <= 0 Then
        Exit Do
    End If
    
    If Y >= CornerY Or Y <= 0 Then
        Exit Do
    End If
    
    If FishCorridorMap(X, Y) = 1 Then Found = True
Loop Until Found Or (MapImage(X, Y) = 0)

If Found Then Distance = i

SearchAroundReceiver = Distance
    
End Function
Public Sub DrawReceiverConnectionsToCorridor()
'draws receiver connection to corridor on default map
Dim R As Integer
Dim X1 As Long
Dim Y1 As Long
Dim X2 As Long
Dim Y2 As Long
For R = 1 To Receiver.TotalReceivers
    With ReceiverConnectionToCorridor(R)
        If .SelectedDirection <> -1 Then
            X1 = .X
            Y1 = .Y
            X2 = Receiver.X(R)
            Y2 = Receiver.Y(R)
            Form1.Picture1.Line (X1, Y1)-(X2, Y2), vbBlue
        End If
    End With
Next R
      
End Sub
Public Function CorridorPoint_X(ByVal ReceiverNumber As Integer) As Long
Dim G As Integer
Dim R As Integer
Dim X As Integer

'returns -1 if no connection
X = -1
'group?
G = Receiver.GroupNumber(ReceiverNumber)
R = ReceiverNumber
If G Then
    'will return -1 if no receiver in group is connected
    R = GetCorrectReceiverNumber(G)
End If


'connected
If R <> -1 Then
    If ReceiverConnectionToCorridor(R).SelectedDirection <> -1 Then
        X = ReceiverConnectionToCorridor(R).X
    End If
End If

CorridorPoint_X = CLng(X)
End Function
Public Function CorridorPoint_Y(ByVal ReceiverNumber As Long) As Long
Dim G As Integer
Dim R As Integer
Dim Y As Integer

'returns -1 if no connection
Y = -1
'group?
G = Receiver.GroupNumber(ReceiverNumber)
R = ReceiverNumber
If G Then
    'will return -1 if no receiver in group is connected
    R = GetCorrectReceiverNumber(G)
End If

'connected
If R <> -1 Then
    If ReceiverConnectionToCorridor(R).SelectedDirection <> -1 Then
        Y = ReceiverConnectionToCorridor(R).Y
    End If
End If

CorridorPoint_Y = CLng(Y)
End Function

Public Function LandBetweenPositions(ByVal New_x As Long, ByVal New_y As Long, ByVal Old_X As Long, ByVal Old_Y As Long)
'Alerts if line is going over LAND
Dim dX As Long
Dim dy As Long
Dim SX As Long
Dim SY As Long
Dim err As Long
Dim E2 As Long
Dim X As Long
Dim Y As Long

Dim Pixel As Long
Dim Navigation As Boolean


Const WATER = 1
Const LAND = 0


Const OVER_WATER = False
Const OVER_LAND = True

'Bresenham's DDA
dX = Abs(New_x - Old_X)
dy = Abs(New_y - Old_Y)
If Old_X < New_x Then SX = 1 Else SX = -1
If Old_Y < New_y Then SY = 1 Else SY = -1

X = Old_X
Y = Old_Y
err = dX - dy
Do
    'check position to see if over land or water
    Pixel = MapImage(X, Y)
    
    If Pixel = LAND Then
        Navigation = OVER_LAND Or Navigation
    Else
        Navigation = OVER_WATER Or Navigation
    End If

    If X = New_x And Y = New_y Then Exit Do
    
    E2 = 2 * err
    
    If E2 > (-1 * dy) Then
        err = err - dy
        X = X + SX
    End If
    
    If E2 < dX Then
        err = err + dX
        Y = Y + SY
    End If
    
Loop
    
'If Not Navigation Then Form1.Picture1.Line (New_x, New_y)-(Old_X, Old_Y), vbGreen

LandBetweenPositions = Navigation


End Function
Public Function Colorize(i As Integer) As Long
If i < 0 Then i = 0
If i > 255 Then i = 255
Colorize = RGB(255 - i, 0, i)
End Function
Public Sub DrawScale(Box As PictureBox, s As Integer)
'draws the scale
Dim H As Long
Dim w As Long
Dim X As Long
Dim Y As Long
Dim i As Long
Dim intensity As Integer
Dim Color As Long

w = Box.ScaleWidth / (255 / s)
H = Box.ScaleHeight

Do
    intensity = intensity + s
    Color = Colorize(intensity)
    Box.Line (X, Y)-(X + w, Y + H), Color, BF
    X = X + w
Loop Until X > Box.ScaleWidth

End Sub
