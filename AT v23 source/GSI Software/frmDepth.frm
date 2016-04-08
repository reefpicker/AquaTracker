VERSION 5.00
Begin VB.Form frmDepth 
   Caption         =   "Depth"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3330
   Icon            =   "frmDepth.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4065
   ScaleWidth      =   3330
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDepth 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtScaleLabel 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "1-"
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox txtScaleLabel 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "1-"
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox txtScaleLabel 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "1-"
      Top             =   480
      Width           =   495
   End
   Begin VB.PictureBox picCurrentDepth 
      Height          =   3495
      Left            =   720
      ScaleHeight     =   229
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   133
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Current Depth:"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmDepth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
DepthFormIsLoaded = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
DepthFormIsLoaded = False
End Sub
