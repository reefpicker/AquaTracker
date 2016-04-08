VERSION 5.00
Begin VB.Form frmAbioticInfo 
   Caption         =   "Abiotic info: Receiver "
   ClientHeight    =   2970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3705
   Icon            =   "frmAbioticInfo.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2970
   ScaleWidth      =   3705
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   1815
      Left            =   120
      ScaleHeight     =   1755
      ScaleWidth      =   3435
      TabIndex        =   6
      Top             =   960
      Width           =   3495
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Max:"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Min:"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Average:"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmAbioticInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub
