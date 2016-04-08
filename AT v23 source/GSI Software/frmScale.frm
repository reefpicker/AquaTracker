VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmScale 
   AutoRedraw      =   -1  'True
   Caption         =   "Color scale"
   ClientHeight    =   1104
   ClientLeft      =   192
   ClientTop       =   816
   ClientWidth     =   6432
   Icon            =   "frmScale.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1104
   ScaleWidth      =   6432
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   2400
      Top             =   -120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picScale 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   120
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   517
      TabIndex        =   0
      Top             =   360
      Width           =   6255
   End
   Begin VB.Label lblScaleValue 
      Caption         =   "0"
      Height          =   255
      Index           =   8
      Left            =   6000
      TabIndex        =   9
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblScaleValue 
      Caption         =   "0"
      Height          =   255
      Index           =   7
      Left            =   5400
      TabIndex        =   8
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblScaleValue 
      Caption         =   "0"
      Height          =   255
      Index           =   6
      Left            =   4680
      TabIndex        =   7
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblScaleValue 
      Caption         =   "0"
      Height          =   255
      Index           =   5
      Left            =   3960
      TabIndex        =   6
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblScaleValue 
      Caption         =   "0"
      Height          =   255
      Index           =   4
      Left            =   3120
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblScaleValue 
      Caption         =   "0"
      Height          =   255
      Index           =   3
      Left            =   2400
      TabIndex        =   4
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblScaleValue 
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblScaleValue 
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblScaleValue 
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.Menu mnuStyle 
      Caption         =   "Style"
      Begin VB.Menu mnuColorScheme 
         Caption         =   "Heat"
         Index           =   0
      End
      Begin VB.Menu mnuColorScheme 
         Caption         =   "Crystals"
         Index           =   1
      End
      Begin VB.Menu mnuColorScheme 
         Caption         =   "Polar"
         Index           =   2
      End
      Begin VB.Menu mnuColorScheme 
         Caption         =   "Grey Scale"
         Checked         =   -1  'True
         Index           =   3
      End
   End
   Begin VB.Menu mnuClipboard 
      Caption         =   "Edit"
      Begin VB.Menu mnuCopyScale 
         Caption         =   "Copy scale"
      End
      Begin VB.Menu mnuEnableDisableBin 
         Caption         =   "Disable bin"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuBins 
      Caption         =   "Bins"
      Begin VB.Menu mnuChangeBinNumber 
         Caption         =   "Number of bins"
      End
      Begin VB.Menu mnuChangeMaxValue 
         Caption         =   "Max bin value"
      End
      Begin VB.Menu mnuChangeMinValue 
         Caption         =   "Min bin value"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAutoSetValues 
         Caption         =   "Auto set all values"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmScale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BinNumber As Integer
Dim Selecting As Boolean
Dim BinColor As Long
Private Sub Form_Load()
ColorScale.AutoAssignColors
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Selecting Then
    Selecting = False
    ColorScale.RefreshPlot
End If

End Sub

Private Sub mnuAutoSetValues_Click()
mnuAutoSetValues.Checked = True
With ColorScale
    .AutoSet = True
    .Max = ImageProcessingEngine.SetMax
    .AutoAssignColors
    .ShowScale
    .RefreshPlot
End With
End Sub

Private Sub mnuChangeBinNumber_Click()
'change number of shades (i.e. scale steps)
Dim Value As Integer
Dim response As Variant

Value = ColorScale.NumberOfBins

response = InputBox("Enter number of bins/shades to use:", "Bins", Value)

If response = vbCancel Or response = "" Then Exit Sub
If IsNumeric(response) Then Value = CInt(response) Else Exit Sub

'only allow positive values
If Value < 0 Then Value = 1
If Value > 254 Then Value = 254

ColorScale.AutoSet = False
ColorScale.NumberOfBins = Value
ColorScale.AutoAssignColors
ColorScale.RefreshPlot
mnuAutoSetValues.Checked = False

End Sub

Private Sub mnuChangeMaxValue_Click()
'change max value for density plots
'this overrides the calculated max value
'affects all density plots

Dim Value As Long
Dim response As Variant

Value = ColorScale.Max

response = InputBox("Enter value for last bin (Max) or enter 0 to use default value:", "User defined MAX value", Value)

If response = vbCancel Or response = "" Then Exit Sub
If IsNumeric(response) Then Value = CInt(response) Else Exit Sub

'only allow positive values
If Value < 0 Then Value = 0

ColorScale.AutoSet = False
ColorScale.Max = Value
ColorScale.AutoAssignColors
ColorScale.RefreshPlot
mnuAutoSetValues.Checked = False


End Sub

Private Sub mnuChangeMinValue_Click()
'change min value
Dim Value As Long
Dim response As Variant

Value = ColorScale.Min

response = InputBox("Enter base value for first bin to use:", "Base bin value", Value)

If response = vbCancel Or response = "" Then Exit Sub
If IsNumeric(response) Then Value = CInt(response) Else Exit Sub

'only allow positive values
If Value < 0 Then Value = 0

ColorScale.AutoSet = False
ColorScale.Min = Value
ColorScale.AutoAssignColors
ColorScale.RefreshPlot
mnuAutoSetValues.Checked = False

End Sub

Private Sub mnuColorScheme_Click(Index As Integer)
Dim i As Long

'check right one! unchecked others!
For i = 0 To mnuColorScheme.UBound
    If i <> Index Then mnuColorScheme(i).Checked = False Else mnuColorScheme(i).Checked = True
Next i

'set scale
ColorScale.ColorScheme = Index
ColorScale.AutoAssignColors
ColorScale.RefreshPlot

End Sub

Private Sub mnuCopyScale_Click()
'copy to clipboard
If Selecting Then
    Selecting = False
    ColorScale.RefreshPlot
End If
Clipboard.Clear
Clipboard.SetData frmScale.picScale.Image
End Sub

Private Sub mnuEnableDisableBin_Click()
If ColorScale.IsVisible(BinNumber) Then
    ColorScale.MakeNotVisible = BinNumber
Else
    ColorScale.MakeVisible = BinNumber
End If
End Sub

Private Sub picScale_DblClick()
On Error GoTo ExitWithError

'user can change color of bin
If Selecting Then
    With CommonDialog
        .CancelError = True
        .ShowColor
        ColorScale.AssignColor(BinNumber) = .Color
    End With
End If

ExitWithError:
ColorScale.RefreshPlot
End Sub

Private Sub picScale_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'pop clipboard popup menu
If Button = vbRightButton Then
    'enable feature only available in popup version
    mnuEnableDisableBin.Visible = True
    BinNumber = ColorScale.Bin(X)
    If ColorScale.IsVisible(BinNumber) Then
        mnuEnableDisableBin.Caption = "Disable bin"
    Else
        mnuEnableDisableBin.Caption = "Enable bin"
    End If
    PopupMenu mnuClipboard
End If
mnuEnableDisableBin.Visible = False
End Sub

Private Sub picScale_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Height As Single
Dim BinX As Single
Dim response As Variant
Static PreviousBinNumber As Integer

Selecting = True
ColorScale.ShowScale
Height = picScale.ScaleHeight - 1
BinNumber = ColorScale.Bin(X)

If BinNumber = -1 Then
    response = MsgBox("Ooops. Don't drag the mouse like that!", vbOKOnly)
End If

BinColor = ColorScale.GetBinColor(BinNumber)

BinX = (BinNumber) * ColorScale.WidthOfBin

If BinX < 0 Then BinX = 0
picScale.Line (BinX, 0)-(BinX + ColorScale.WidthOfBin, Height), HighLightColor, B
picScale.Line (BinX + 1, 1)-(BinX + ColorScale.WidthOfBin - 1, Height + 1), HighLightColor Or &H808080, BF

If PreviousBinNumber <> BinNumber Then
    ColorScale.HighlightBin BinNumber
    picScale.ToolTipText = ColorScale.RangeString(BinNumber)
End If

PreviousBinNumber = BinNumber

End Sub
