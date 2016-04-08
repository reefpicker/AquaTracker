VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIBSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' ==================================================================================
' cDIBSection.cls
' Copyright © 1998 Steve McMahon (steve@vbaccelerator.com)
' Visit vbAccelerator at www.vbaccelerator.com
'
' Creates and manages a True Color (24bit) GDI DibSection.
' This is DIB in which the bitmap bits are stored in windows
' memory so can be modified.
' See the RandomiseBits and Resample methods for how to do it.
'
' For example, fading in an out a 256x256 true colour DIB by
' directly modifying the bytes runs at 38fps on my machine
' (PII 266Mhz, 32Mb RAM, 8Mb ATI Xpert@Work AGP card)
'
' Note: for best performance, when compiling an executable check
' all the boxes on the Properties-Compile tab Advanced Optimisations
' button.  This really makes a difference! (e.g. the fading example
' ran at 22fps before I did this so > 50%!).
' ==================================================================================
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)


Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type

Private Type SAFEARRAY1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds As SAFEARRAYBOUND
End Type

Private Type BITMAPINFOHEADER '40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
' Note - this is not the declare in the API viewer - modify lplpVoid to be
' Byref so we get the pointer back:
Private Declare Function CreateDIBSection Lib "gdi32" _
    (ByVal hDC As Long, _
    pBitmapInfo As BITMAPINFO, _
    ByVal un As Long, _
    lplpVoid As Long, _
    ByVal handle As Long, _
    ByVal dw As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Const BI_RGB = 0&
Private Const BI_RLE4 = 2&
Private Const BI_RLE8 = 1&
Private Const DIB_RGB_COLORS = 0 '  color table in RGBs

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

' Start of structure:
Private Const BITMAPTYPE As Integer = &H4D42
Private Type BITMAPFILEHEADER
   bfType As Integer '- type  ="BM" i.e &H4D42 - 2
   bfSize As Long ' - size in bytes of file - 6
   bfReserved1 As Integer ' - reserved, must be 0 - 8
   bfReserved2 As Integer ' - reserved, must be 0 - 10
   bfOffBits As Long ' offset from this structure to the bitmap bits - 14
End Type
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const INVALID_HANDLE_VALUE = -1
Private Const CREATE_ALWAYS = 2
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_BEGIN = 0
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Const GMEM_FIXED = &H0
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Long, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBitmap Lib "gdi32" (ByVal hDC As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO, ByVal wUsage As Long) As Long

' DrawDIB functions:
Private Declare Function DrawDibOpen Lib "MSVFW32.dll" () As Long
Private Declare Function DrawDibClose Lib "MSVFW32.dll" (ByVal hDD As Long) As Long
Private Declare Function DrawDibDraw Lib "MSVFW32.dll" (ByVal hDD As Long, ByVal hDC As Long, _
      ByVal xDst As Long, ByVal yDst As Long, ByVal dxDst As Long, ByVal dyDst As Long, _
      lpBI As Any, lpBits As Any, _
      ByVal xSrc As Long, ByVal ySrc As Long, ByVal dxSrc As Long, ByVal dySrc As Long, _
      ByVal wFlags As Long) As Long

Private m_hDIb As Long
Private m_hBmpOld As Long
Private m_hDC As Long
Private m_hDD As Long
Private m_lPtr As Long
Private m_tBI As BITMAPINFO
Dim RGB_COLOR(2) As Byte

Public Property Get UseDrawDib() As Boolean
   UseDrawDib = Not (m_hDD = 0)
End Property
Public Property Let UseDrawDib(ByVal bState As Boolean)
   If bState Then
      If m_hDD = 0 Then
         m_hDD = DrawDibOpen()
      End If
   Else
      If Not (m_hDD = 0) Then
         DrawDibClose m_hDD
      End If
   End If
End Property
Public Function CreateDIB( _
        ByVal lhDC As Long, _
        ByVal lWidth As Long, _
        ByVal lHeight As Long, _
        ByRef hDib As Long _
    ) As Boolean
    With m_tBI.bmiHeader
        .biSize = Len(m_tBI.bmiHeader)
        .biWidth = lWidth
        .biHeight = lHeight
        .biPlanes = 1
        .biBitCount = 24
        .biCompression = BI_RGB
        .biSizeImage = BytesPerScanLine * .biHeight
    End With
    hDib = CreateDIBSection( _
            lhDC, _
            m_tBI, _
            DIB_RGB_COLORS, _
            m_lPtr, _
            0, 0)
    CreateDIB = (hDib <> 0)
End Function
Public Function CreateFromPicture( _
        ByRef picThis As StdPicture _
    )
Dim lhDC As Long
Dim lhDCDesktop As Long
Dim lhbmpold As Long
Dim tBMP As BITMAP
    
    GetObjectAPI picThis.handle, Len(tBMP), tBMP
    If (Create(tBMP.bmWidth, tBMP.bmHeight)) Then
        lhDCDesktop = GetDC(GetDesktopWindow())
        If (lhDCDesktop <> 0) Then
            lhDC = CreateCompatibleDC(lhDCDesktop)
            DeleteDC lhDCDesktop
            If (lhDC <> 0) Then
                lhbmpold = SelectObject(lhDC, picThis.handle)
                LoadPictureBlt lhDC
                SelectObject lhDC, lhbmpold
                DeleteObject lhDC
            End If
        End If
    End If
End Function
Public Function CreateFromImage(ByRef picThis As StdPicture, ByVal lhDC As Long)
Dim lhbmpold As Long
Dim tBMP As BITMAP
    
    If (lhDC <> 0) Then
        
        LoadPictureBlt lhDC
      
    End If

End Function
Public Function Create( _
        ByVal lWidth As Long, _
        ByVal lHeight As Long _
    ) As Boolean
Dim bDrawDib As Boolean
   bDrawDib = UseDrawDib()
    ClearUp
    m_hDC = CreateCompatibleDC(0)
    If (m_hDC <> 0) Then
        If (CreateDIB(m_hDC, lWidth, lHeight, m_hDIb)) Then
            m_hBmpOld = SelectObject(m_hDC, m_hDIb)
            UseDrawDib = bDrawDib
            Create = True
        Else
            DeleteObject m_hDC
            m_hDC = 0
        End If
    End If
End Function
Public Property Get BytesPerScanLine() As Long
    ' Scans must align on dword boundaries:
    BytesPerScanLine = (m_tBI.bmiHeader.biWidth * 3 + 3) And &HFFFFFFFC
End Property

Public Property Get Width() As Long
    Width = m_tBI.bmiHeader.biWidth
End Property
Public Property Get Height() As Long
    Height = m_tBI.bmiHeader.biHeight
End Property

Public Sub LoadPictureBlt( _
        ByVal lhDC As Long, _
        Optional ByVal lSrcLeft As Long = 0, _
        Optional ByVal lSrcTop As Long = 0, _
        Optional ByVal lSrcWidth As Long = -1, _
        Optional ByVal lSrcHeight As Long = -1, _
        Optional ByVal eRop As RasterOpConstants = vbSrcCopy _
    )
    If lSrcWidth < 0 Then lSrcWidth = m_tBI.bmiHeader.biWidth
    If lSrcHeight < 0 Then lSrcHeight = m_tBI.bmiHeader.biHeight
    BitBlt m_hDC, 0, 0, lSrcWidth, lSrcHeight, lhDC, lSrcLeft, lSrcTop, eRop
End Sub

Public Function SavePicture(ByVal sFileName As String) As Boolean
Dim lC As Long, i As Long

   ' Save to BMP:
   SavePicture = SaveToBitmap(m_lPtr, sFileName)

End Function
Private Function SaveToBitmap(ByVal lPtrBits As Long, ByVal sFileName As String)
Dim tBH As BITMAPFILEHEADER
Dim tRGBQ As RGBQUAD
Dim hFile As Long
Dim lBytesWritten As Long
Dim lSize As Long
Dim lR As Long
Dim bErr As Boolean
Dim hMem As Long, LPTR As Long
Dim lErr As Long

   ' Prepare the BITMAPFILEHEADER
   With tBH
      .bfType = BITMAPTYPE
      .bfOffBits = 14 + Len(m_tBI)
      .bfSize = .bfOffBits + m_tBI.bmiHeader.biSizeImage
   End With
   hFile = CreateFile(sFileName, _
                 GENERIC_READ Or GENERIC_WRITE, _
                  ByVal 0&, _
                  ByVal 0&, _
                  CREATE_ALWAYS, _
                  FILE_ATTRIBUTE_NORMAL, _
                  0)
   lErr = err.LastDllError
   If (hFile = INVALID_HANDLE_VALUE) Then
      ' error
      err.Raise 17, App.EXEName & ".cDIBSection", ApiError(lErr)
   Else
      
      ' Writing the BITMAPFILEINFOHEADER is somewhat painful
      ' due to non-byte alignment of structure...
      hMem = GlobalAlloc(GPTR, 14)
      LPTR = GlobalLock(hMem)
      CopyMemory ByVal LPTR, tBH.bfType, 2
      CopyMemory ByVal LPTR + 2, tBH.bfSize, 4
      CopyMemory ByVal LPTR + 6, 0&, 4
      CopyMemory ByVal LPTR + 10, tBH.bfOffBits, 4
      lSize = 14
      lR = WriteFile(hFile, ByVal LPTR, lSize, lBytesWritten, ByVal 0&)
      GlobalUnlock hMem
      GlobalFree hMem
      
      ' Add the BITMAPINFOHEADER and colour palette:
      bErr = FileErrHandler(lR, lSize, lBytesWritten)
      If Not bErr Then
         lSize = Len(m_tBI)
         lR = WriteFile(hFile, m_tBI, lSize, lBytesWritten, ByVal 0&)
         bErr = FileErrHandler(lR, lSize, lBytesWritten)
      End If
      
      If Not bErr Then
         ' Its easy to write the bitmap data, though...
         lSize = m_tBI.bmiHeader.biSizeImage
         lR = WriteFile(hFile, ByVal lPtrBits, lSize, lBytesWritten, ByVal 0&)
         bErr = FileErrHandler(lR, lSize, lBytesWritten)
      End If
      
      
      CloseHandle hFile
      SaveToBitmap = Not (bErr)
   End If

End Function
Private Function ApiError(ByVal e As Long) As String
    Dim s As String, c As Long
    s = String(256, 0)
    c = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or _
                      FORMAT_MESSAGE_IGNORE_INSERTS, _
                      0, e, 0&, s, Len(s), ByVal 0)
    If c Then ApiError = Left$(s, c)
End Function
Private Function FileErrHandler(ByVal lR As Long, ByVal lSize As Long, ByVal lBytes As Long) As Boolean
   If (lR = 0) Or Not (lSize = lBytes) Then
      'Err.Raise
      FileErrHandler = True
   End If
End Function

Public Sub PaintPicture( _
        ByVal lhDC As Long, _
        Optional ByVal lDestLeft As Long = 0, _
        Optional ByVal lDestTop As Long = 0, _
        Optional ByVal lDestWidth As Long = -1, _
        Optional ByVal lDestHeight As Long = -1, _
        Optional ByVal lSrcLeft As Long = 0, _
        Optional ByVal lSrcTop As Long = 0, _
        Optional ByVal eRop As RasterOpConstants = vbSrcCopy _
    )
   If (lDestWidth < 0) Then lDestWidth = m_tBI.bmiHeader.biWidth
   If (lDestHeight < 0) Then lDestHeight = m_tBI.bmiHeader.biHeight
   If Not (m_hDD = 0) Then
      ' DrawDib method:
      DrawDibDraw m_hDD, lhDC, lDestLeft, lDestTop, lDestWidth, lDestHeight, m_tBI, ByVal m_lPtr, lSrcLeft, lSrcTop, lDestWidth, lDestHeight, 0
   Else
      BitBlt lhDC, lDestLeft, lDestTop, lDestWidth, lDestHeight, m_hDC, lSrcLeft, lSrcTop, eRop
   End If
End Sub

Public Property Get hDC() As Long
    hDC = m_hDC
End Property
Public Property Get hDib() As Long
    hDib = m_hDIb
End Property
Public Property Get DIBSectionBitsPtr() As Long
    DIBSectionBitsPtr = m_lPtr
End Property
Public Sub RandomiseBits( _
        Optional ByVal bGray As Boolean = False _
    )
Dim bDib() As Byte
Dim X As Long, Y As Long
Dim lC As Long
Dim tsa As SAFEARRAY2D
Dim xEND As Long
    
    ' Get the bits in the from DIB section:
    With tsa
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = m_tBI.bmiHeader.biHeight
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = BytesPerScanLine()
        .pvData = m_lPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tsa), 4

    ' random:
    Randomize Timer
    
    xEND = (Width - 1) * 3
    If (bGray) Then
        For Y = 0 To m_tBI.bmiHeader.biHeight - 1
            For X = 0 To xEND Step 3
                lC = Rnd * 255
                bDib(X, Y) = lC
                bDib(X + 1, Y) = lC
                bDib(X + 2, Y) = lC
            Next X
        Next Y
    Else
        For X = 0 To xEND Step 3
            For Y = 0 To m_tBI.bmiHeader.biHeight - 1
                bDib(X, Y) = 0
                bDib(X + 1, Y) = Rnd * 255
                bDib(X + 2, Y) = Rnd * 255
            Next Y
        Next X
    End If
    
    ' Clear the temporary array descriptor
    ' (This does not appear to be necessary, but
    ' for safety do it anyway)
    CopyMemory ByVal VarPtrArray(bDib), 0&, 4
    
End Sub

Public Sub ClearUp()
If (m_hDC <> 0) Then
     If (m_hDIb <> 0) Then
         SelectObject m_hDC, m_hBmpOld
         DeleteObject m_hDIb
     End If
     DeleteObject m_hDC
 End If
 m_hDC = 0: m_hDIb = 0: m_hBmpOld = 0: m_lPtr = 0
 If Not (m_hDD = 0) Then
   DrawDibClose m_hDD
   m_hDD = 0
End If
End Sub

Public Function Resample( _
        ByVal lNewHeight As Long, _
        ByVal lNewWidth As Long _
    ) As cDIBSection
Dim cdib As cDIBSection
    Set cdib = New cDIBSection
    If cdib.Create(lNewWidth, lNewHeight) Then
        If (lNewWidth <> m_tBI.bmiHeader.biWidth) Or (lNewHeight <> m_tBI.bmiHeader.biHeight) Then
            ' Change in size, do resample:
            ResampleDib cdib
        Else
            ' No size change so just return a copy:
            cdib.LoadPictureBlt m_hDC
        End If
        Set Resample = cdib
    End If
End Function
Public Sub InitializeMatrix()
Dim temp() As Byte
Dim tsa As SAFEARRAY2D
Dim X As Long
Dim Y As Long
Dim c As Long
Dim Top As Long
Dim iX As Integer
Dim iY As Integer


' Get the bits in the DIB section:
With tsa
    .cbElements = 1
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = m_tBI.bmiHeader.biHeight
    .Bounds(1).lLbound = 0
    .Bounds(1).cElements = BytesPerScanLine()
    .pvData = m_lPtr
End With

'copy temporary array descriptor
CopyMemory ByVal VarPtrArray(temp()), VarPtr(tsa), 4

Top = m_tBI.bmiHeader.biHeight - 1

'check which buffer to copy
For X = 0 To BytesPerScanLine() - 3 Step 3
    For Y = 0 To m_tBI.bmiHeader.biHeight - 1
        c = RGB(temp(X + 2, Y), temp(X + 1, Y), temp(X, Y))
        iX = Int(X / 3)
        iY = Top - Y
        If c = WaterColor Then
            MapImage(iX, iY) = 1
        Else
            MapImage(iX, iY) = 0
        End If
    Next Y
Next X

        
'on exit clear descriptor.  This is important so that Windows does not crash or becomes unstable

CopyMemory ByVal VarPtrArray(temp()), &H0, 4


End Sub
Public Sub TransferToFishCorridorMatrix()
Dim temp() As Byte
Dim tsa As SAFEARRAY2D
Dim X As Long
Dim Y As Long
Dim c As Long
Dim Top As Long


' Get the bits in the DIB section:
With tsa
    .cbElements = 1
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = m_tBI.bmiHeader.biHeight
    .Bounds(1).lLbound = 0
    .Bounds(1).cElements = BytesPerScanLine()
    .pvData = m_lPtr
End With

'copy temporary array descriptor
CopyMemory ByVal VarPtrArray(temp()), VarPtr(tsa), 4

Top = m_tBI.bmiHeader.biHeight - 1

'check which buffer to copy
For X = 0 To BytesPerScanLine() - 3 Step 3
    For Y = 0 To m_tBI.bmiHeader.biHeight - 1
        c = RGB(temp(X + 2, Y), temp(X + 1, Y), temp(X, Y))
        If c = vbBlue Then
            FishCorridorMap(Int(X / 3), Top - Y) = 1
        Else
            FishCorridorMap(Int(X / 3), Top - Y) = 0
        End If
    Next Y
Next X

        
'on exit clear descriptor.  This is important so that Windows does not crash or becomes unstable

CopyMemory ByVal VarPtrArray(temp()), &H0, 4


End Sub
Public Sub WriteScaledMatrix(ByRef TempImage() As Long)
Dim temp() As Byte
Dim tsa As SAFEARRAY2D
Dim X As Long
Dim Y As Long
Dim c As Byte
Dim Top As Long

'Get the bits in the DIB section:
With tsa
    .cbElements = 1
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = m_tBI.bmiHeader.biHeight
    .Bounds(1).lLbound = 0
    .Bounds(1).cElements = BytesPerScanLine()
    .pvData = m_lPtr
End With

'copy temporary array descriptor
CopyMemory ByVal VarPtrArray(temp()), VarPtr(tsa), 4

Top = m_tBI.bmiHeader.biHeight - 1

'check which buffer to copy
For X = 0 To BytesPerScanLine() - 3 Step 3
    For Y = 0 To m_tBI.bmiHeader.biHeight - 1
        c = CByte(&HFF And TempImage(Int(X / 3), Top - Y))
        temp(X, Y) = c: temp(X + 1, Y) = c: temp(X + 2, Y) = c
    Next Y
Next X

        
'on exit clear descriptor.  This is important so that Windows does not crash or becomes unstable

CopyMemory ByVal VarPtrArray(temp()), &H0, 4


End Sub
Private Sub SplitRGB(PixelColor As Long)
'this sub splits the long integer it receives into three separate long values
'it does this by masking the non-significant bits out and the shifting the values using
'the "division by 2" technique to shift bits around until they are in the lower 8 bits

Dim Green As Long
Dim Blue As Long


'split the pixelcolor long into three pixel colors (long for speed)
RGB_COLOR(0) = PixelColor And 255
Green = PixelColor And 65280
Blue = PixelColor And 16711680

'Now shift all around
'dividing a binary number by 2 is the same as shifting all bits to the right

'first channel does not require shifting.
RGB_COLOR(1) = Green / 256
RGB_COLOR(2) = Blue / 65536

'done

End Sub

Public Sub WriteMatrix(ByRef TempImage() As Long)
Dim temp() As Byte
Dim tsa As SAFEARRAY2D
Dim X As Long
Dim Y As Long
Dim c As Long
Dim Top As Long

'Get the bits in the DIB section:
With tsa
    .cbElements = 1
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = m_tBI.bmiHeader.biHeight
    .Bounds(1).lLbound = 0
    .Bounds(1).cElements = BytesPerScanLine()
    .pvData = m_lPtr
End With

'copy temporary array descriptor
CopyMemory ByVal VarPtrArray(temp()), VarPtr(tsa), 4

Top = m_tBI.bmiHeader.biHeight - 1

'check which buffer to copy
For X = 0 To BytesPerScanLine() - 3 Step 3
    For Y = 0 To m_tBI.bmiHeader.biHeight - 1
        SplitRGB (TempImage(X / 3, Y))
        temp(X, Y) = RGB_COLOR(0): temp(X + 1, Y) = RGB_COLOR(1): temp(X + 2, Y) = RGB_COLOR(2)
    Next Y
Next X

        
'on exit clear descriptor.  This is important so that Windows does not crash or becomes unstable

CopyMemory ByVal VarPtrArray(temp()), &H0, 4


End Sub
Private Function ResampleDib(ByRef cDibTo As cDIBSection) As Boolean
Dim bDibFrom() As Byte
Dim bDibTo() As Byte
Dim tSAFrom As SAFEARRAY2D
Dim tSATo As SAFEARRAY2D

    ' Get the bits in the from DIB section:
    With tSAFrom
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = m_tBI.bmiHeader.biHeight
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = BytesPerScanLine()
        .pvData = m_lPtr
    End With
    CopyMemory ByVal VarPtrArray(bDibFrom()), VarPtr(tSAFrom), 4

    ' Get the bits in the to DIB section:
    With tSATo
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cDibTo.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cDibTo.BytesPerScanLine()
        .pvData = cDibTo.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDibTo()), VarPtr(tSATo), 4

Dim xScale As Single
Dim yScale As Single

Dim X As Long, Y As Long, xEND As Long, xOut As Long

Dim fX As Single, fY As Single
Dim ifY As Long, ifX As Long
Dim dX As Single, dy As Single
Dim R As Long, R1 As Single, R2 As Single, r3 As Single, r4 As Single
Dim G As Long, g1 As Single, g2 As Single, g3 As Single, g4 As Single
Dim B As Long, b1 As Single, b2 As Single, b3 As Single, b4 As Single
Dim ir1 As Long, ig1 As Long, ib1 As Long
Dim ir2 As Long, ig2 As Long, ib2 As Long

    xScale = (Width - 1) / cDibTo.Width
    yScale = (Height - 1) / cDibTo.Height
    
    xEND = cDibTo.Width - 1
        
    For Y = 0 To cDibTo.Height - 1
        
        fY = Y * yScale
        ifY = Int(fY)
        dy = fY - ifY
        
        For X = 0 To xEND
            fX = X * xScale
            ifX = Int(fX)
            dX = fX - ifX
            
            ifX = ifX * 3
            ' Interpolate using the four nearest pixels in the source
            b1 = bDibFrom(ifX, ifY): g1 = bDibFrom(ifX + 1, ifY): R1 = bDibFrom(ifX + 2, ifY)
            b2 = bDibFrom(ifX + 3, ifY): g2 = bDibFrom(ifX + 4, ifY): R2 = bDibFrom(ifX + 5, ifY)
            b3 = bDibFrom(ifX, ifY + 1): g3 = bDibFrom(ifX + 1, ifY + 1): r3 = bDibFrom(ifX + 2, ifY + 1)
            b4 = bDibFrom(ifX + 3, ifY + 1): g4 = bDibFrom(ifX + 4, ifY + 1): r4 = bDibFrom(ifX + 5, ifY + 1)
            
            ' Interplate in x direction:
            ir1 = R1 * (1 - dy) + r3 * dy: ig1 = g1 * (1 - dy) + g3 * dy: ib1 = b1 * (1 - dy) + b3 * dy
            ir2 = R2 * (1 - dy) + r4 * dy: ig2 = g2 * (1 - dy) + g4 * dy: ib2 = b2 * (1 - dy) + b4 * dy
            ' Interpolate in y:
            R = ir1 * (1 - dX) + ir2 * dX: G = ig1 * (1 - dX) + ig2 * dX: B = ib1 * (1 - dX) + ib2 * dX
            
            ' Set output:
            If (R < 0) Then R = 0
            If (R > 255) Then R = 255
            If (G < 0) Then G = 0
            If (G > 255) Then G = 255
            If (B < 0) Then B = 0
            If (B > 255) Then
                B = 255
            End If
            xOut = X * 3
            bDibTo(xOut, Y) = B
            bDibTo(xOut + 1, Y) = G
            bDibTo(xOut + 2, Y) = R
            
        Next X
        
    Next Y

    ' Clear the temporary array descriptor
    ' (This does not appear to be necessary, but
    ' for safety do it anyway)
    CopyMemory ByVal VarPtrArray(bDibFrom), 0&, 4
    CopyMemory ByVal VarPtrArray(bDibTo), 0&, 4


End Function

Private Sub Class_Terminate()
    ClearUp
End Sub


