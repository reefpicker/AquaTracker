VERSION 5.00
Begin VB.Form frmPickColorOfWater 
   Caption         =   "Pick color of water"
   ClientHeight    =   2250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3150
   Icon            =   "frmPickColorOfWater.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2250
   ScaleWidth      =   3150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton cmdChoose 
      Caption         =   "Choose"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Shape shpWater 
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   120
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmPickColorOfWater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdChoose_Click()
WaterColor = shpWater.FillColor
SaveSetting APPLICATION, REGISTRY_SECTION, "WATER", Str$(WaterColor)
Form1.Picture1.MousePointer = vbArrow
ChooseWaterColor = False
Form1.ScanMap
Receiver.ReDraw
Unload Me

End Sub

Private Sub Command2_Click()
Form1.Picture1.MousePointer = vbArrow
ChooseWaterColor = False
Unload Me
End Sub

Private Sub Form_Load()
shpWater.FillColor = WaterColor
Form1.Picture1.MousePointer = 14 'arrow and question


End Sub
