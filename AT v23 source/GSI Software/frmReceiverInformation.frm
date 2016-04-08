VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmReceiverInformation 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Receiver 1/255"
   ClientHeight    =   4935
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   10770
   BeginProperty Font 
      Name            =   "Arial Rounded MT Bold"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReceiverInformation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   2160
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<--Previous"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   4
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next-->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9480
      TabIndex        =   6
      Top             =   4440
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   882
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Identification"
      TabPicture(0)   =   "frmReceiverInformation.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtReceiverName"
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(2)=   "txtReceiverType"
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(4)=   "picRelativeLocation"
      Tab(0).Control(5)=   "Label20"
      Tab(0).Control(6)=   "Label7"
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Zone"
      TabPicture(1)   =   "frmReceiverInformation.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "optSource(0)"
      Tab(1).Control(1)=   "optSource(1)"
      Tab(1).Control(2)=   "txtZoneLat"
      Tab(1).Control(3)=   "txtZoneLong"
      Tab(1).Control(4)=   "picZone"
      Tab(1).Control(5)=   "cmbZoneTag"
      Tab(1).Control(6)=   "lblLong"
      Tab(1).Control(7)=   "lblLat"
      Tab(1).Control(8)=   "Label5"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Group"
      TabPicture(2)   =   "frmReceiverInformation.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Detections"
      TabPicture(3)   =   "frmReceiverInformation.frx":0496
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "lblDaysTTL"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "lblFishTTL"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "lblDetectionsTTL"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label13"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label12"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "lblQueryLabel"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Line1"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Line3"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Line4"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "Line5"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "lstFish"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "lstQueryResponse"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "lstDates"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).ControlCount=   13
      Begin VB.TextBox txtReceiverName 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -72960
         TabIndex        =   39
         Text            =   "Fixed"
         Top             =   840
         Width           =   2175
      End
      Begin VB.Frame Frame2 
         Caption         =   "Group"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   -71280
         TabIndex        =   34
         Top             =   600
         Width           =   3975
         Begin VB.TextBox txtGroupName 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1800
            TabIndex        =   36
            Text            =   "Text1"
            Top             =   360
            Width           =   1815
         End
         Begin VB.ListBox lstGroupMembers 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1815
            Left            =   240
            TabIndex        =   35
            Top             =   1200
            Width           =   3015
         End
         Begin VB.Label Label6 
            Caption         =   "Group name:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label lblOthersInGroup 
            Caption         =   "Others in group:"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   37
            Top             =   840
            Width           =   1695
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Visible/Include"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -70560
         TabIndex        =   32
         Top             =   720
         Width           =   3255
         Begin VB.CheckBox chkVisible 
            Caption         =   "Include/make visible in canvas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   33
            Top             =   360
            Width           =   2895
         End
      End
      Begin VB.TextBox txtReceiverType 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -72960
         TabIndex        =   29
         Text            =   "Fixed"
         Top             =   1440
         Width           =   2175
      End
      Begin VB.ListBox lstDates 
         Height          =   1860
         Left            =   3480
         TabIndex        =   28
         Top             =   2160
         Width           =   2895
      End
      Begin VB.OptionButton optSource 
         Caption         =   "Manual Tag"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   -73440
         TabIndex        =   25
         Top             =   1320
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton optSource 
         Caption         =   "Geographic Tag"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   -73440
         TabIndex        =   24
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox txtZoneLat 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -66120
         TabIndex        =   23
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txtZoneLong 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -66120
         TabIndex        =   22
         Top             =   2520
         Width           =   1575
      End
      Begin VB.PictureBox picZone 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -65760
         ScaleHeight     =   85
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   77
         TabIndex        =   21
         Top             =   600
         Width           =   1215
      End
      Begin VB.ListBox lstQueryResponse 
         Height          =   1860
         Left            =   6600
         TabIndex        =   19
         Top             =   2160
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Frame Frame1 
         Caption         =   "Position and Marker"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -74280
         TabIndex        =   16
         Top             =   2040
         Width           =   3255
         Begin VB.CommandButton cmdChangeMarker 
            Caption         =   "Change Marker"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   31
            Top             =   360
            Width           =   2295
         End
         Begin VB.PictureBox picMarker 
            AutoRedraw      =   -1  'True
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2520
            ScaleHeight     =   21
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   29
            TabIndex        =   30
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox txtLong 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1800
            TabIndex        =   1
            Text            =   "000.0000"
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox txtLat 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1800
            TabIndex        =   0
            Text            =   "000.0000"
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label23 
            Caption         =   "Longitude:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   18
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label Label21 
            Caption         =   "Lattitude:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   1200
            Width           =   1335
         End
      End
      Begin VB.ComboBox cmbZoneTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -73440
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   840
         Width           =   2175
      End
      Begin VB.PictureBox picRelativeLocation 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Left            =   -67200
         ScaleHeight     =   229
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   173
         TabIndex        =   7
         Top             =   720
         Width           =   2655
      End
      Begin VB.ListBox lstFish 
         Height          =   1860
         Left            =   360
         TabIndex        =   3
         Top             =   2160
         Width           =   2895
      End
      Begin VB.Line Line5 
         X1              =   6600
         X2              =   6600
         Y1              =   720
         Y2              =   1320
      End
      Begin VB.Line Line4 
         X1              =   3480
         X2              =   3480
         Y1              =   720
         Y2              =   1320
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   10680
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   10680
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label lblLong 
         Caption         =   "Longitude:"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -67440
         TabIndex        =   27
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label lblLat 
         Caption         =   "Lattitude:"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -67440
         TabIndex        =   26
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label lblQueryLabel 
         Caption         =   "Dates fish #### detected:"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6600
         TabIndex        =   20
         Top             =   1560
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.Label Label20 
         Caption         =   "Type:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74280
         TabIndex        =   15
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74280
         TabIndex        =   14
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Zone Tag:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74640
         TabIndex        =   13
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "Fish detected:"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   1560
         Width           =   2895
      End
      Begin VB.Label Label13 
         Caption         =   "Dates Active:"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   11
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label lblDetectionsTTL 
         Caption         =   "Total Detections: 9999999"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label lblFishTTL 
         Caption         =   "Fish detected: 999999"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   9
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label lblDaysTTL 
         Caption         =   "Days active: ??"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         TabIndex        =   8
         Top             =   840
         Width           =   2655
      End
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   9720
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu mnuTracks 
      Caption         =   "Tracks"
      Visible         =   0   'False
      Begin VB.Menu mnuColorize 
         Caption         =   "Colorize all"
      End
      Begin VB.Menu mnuExclude 
         Caption         =   "Exclude all"
      End
   End
End
Attribute VB_Name = "frmReceiverInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Receiver_Relative_X As Single
Dim Receiver_Relative_Y As Single
Dim ReceiverNumber As Integer
Dim CHANGED_BY_COMPUTER As Boolean
Dim TagNumber As Long
Dim LocalScale_X As Single
Dim LocalScale_Y As Single
Dim SelectColor(3) As Long
Const ECOLOGICAL = 1
Dim ActiveList As ListBox
Dim ShowFishOnQuery As Boolean



Private Sub chkVisible_Click()
If Not CHANGED_BY_COMPUTER Then
    If chkVisible.Value = vbChecked Then
        Receiver.MakeVisible ReceiverNumber
    Else
        Receiver.MakeInvisible ReceiverNumber
    End If
    
    'draw map to scale
    If ZoomRegion.Zoomed Then
        ImageProcessingEngine.DrawZoomMapToScale picRelativeLocation, Form1.Picture1
    Else
        ImageProcessingEngine.DrawMapToScale picRelativeLocation, Form1.Picture1
    End If
    DrawReceiverInSmallCanvas ReceiverNumber
    Form1.RePaint
    If JPlotIsLoaded Then frmJPlot.DrawPlot
End If

End Sub

Private Sub cmbZoneTag_Click()

Dim TagString As String


If CHANGED_BY_COMPUTER Then Exit Sub


If cmbZoneTag.ListIndex = 0 Then
    'new tag string
    TagString = InputBox("Enter new tag string to attach to receiver information:", "New Tag")
    Receiver.NewReceiverTag = TagString
    'update list
    cmbZoneTag.AddItem TagString
    'update receiver
    Receiver.AssignTag(ReceiverNumber) = cmbZoneTag.ListCount - 1
    'show it
    cmbZoneTag.ListIndex = cmbZoneTag.ListCount - 1
Else
    'update receiver
    Receiver.AssignTag(ReceiverNumber) = cmbZoneTag.ListIndex
End If

TagNumber = cmbZoneTag.ListIndex

'update geo frame
'Determine if tag is geographical/zonal and then assign accordingly
If Receiver.TagIsGeographic(TagNumber) Then
    ShowZones
    txtZoneLat.Text = Receiver.ZoneLatRange(TagNumber)
    txtZoneLong.Text = Receiver.ZoneLongRange(TagNumber)
    ImageProcessingEngine.DrawMapToScale picZone, Form1.Picture1
    Receiver.DrawZone TagNumber, picZone, Form1.Picture1
Else
    HideZones
End If

End Sub

Private Sub cmdChangeMarker_Click()
Load frmChangeMarkerForReceiver
frmChangeMarkerForReceiver.LoadReceiverMarker ReceiverNumber
frmChangeMarkerForReceiver.Show


End Sub

Private Sub cmdNext_Click()
ReceiverNumber = ReceiverNumber + 1
If ReceiverNumber > Receiver.TotalReceivers Then ReceiverNumber = 1
ShowInformation

End Sub

Private Sub cmdPrevious_Click()
ReceiverNumber = ReceiverNumber - 1
If ReceiverNumber < 1 Then ReceiverNumber = Receiver.TotalReceivers
ShowInformation

End Sub
Private Sub Form_Load()
'clear combobox for tags now
cmbZoneTag.Clear
cmbZoneTag.AddItem "Add New Tag", 0

'load receiver number and show info
ReceiverNumber = Receiver.CurrentStation_Number

Receiver.ShowTagList cmbZoneTag

'change backcolor of picmarker box
picMarker.BackColor = vbWhite
'calculate local scales in order to draw receiver
LocalScale_X = picRelativeLocation.ScaleWidth / Form1.Picture1.ScaleWidth
LocalScale_Y = picRelativeLocation.ScaleHeight / Form1.Picture1.ScaleHeight
'show information
ShowInformation

End Sub
Private Sub ShowZones()
'show zone-related buttons
optSource(0).Enabled = False
optSource(1).Enabled = True
optSource(0).Value = False
optSource(1).Value = True
lblLat.Enabled = True
lblLong.Enabled = True
txtZoneLat.Enabled = True
txtZoneLong.Enabled = True

End Sub
Private Sub HideZones()
optSource(0).Enabled = True
optSource(1).Enabled = False
optSource(0).Value = True
optSource(1).Value = False
lblLat.Enabled = False
lblLong.Enabled = False
txtZoneLat.Enabled = False
txtZoneLong.Enabled = False
txtZoneLat.Text = ""
txtZoneLong.Text = ""
picZone.Cls
End Sub
Private Sub ShowInformation()
'when loaded up, show current station receiver information
'
Dim i As Long
Dim TagText As String
Dim Notes As String
Dim GroupNumber As Long
Dim ReceiverName As String
Dim N As Long
Dim OldCount As Long
Dim FishInList(MAX_FISH) As Boolean
Dim GroupMembers(MAX_RECEIVERS_PERGROUP) As Integer
Dim R As Integer
Dim m As Long
Const RETURN_RECEIVER_NAME_ONLY = True

CHANGED_BY_COMPUTER = True

'draw map to scale
If ZoomRegion.Zoomed Then
    ImageProcessingEngine.DrawZoomMapToScale picRelativeLocation, Form1.Picture1
Else
    ImageProcessingEngine.DrawMapToScale picRelativeLocation, Form1.Picture1
End If
'Receiver name
ReceiverName = Receiver.ID(ReceiverNumber, RETURN_RECEIVER_NAME_ONLY)
txtReceiverName.Text = ReceiverName

'mobile or fixed?
txtReceiverType.Text = "Fixed Pos"



'show caption
If Receiver_Table.Invisible(ReceiverNumber) Then
    Me.Caption = "Receiver " & ReceiverName & "(" & Str$(ReceiverNumber) & "/" & Receiver.TotalReceivers & ") is set to Excluded/Not Visible"
    chkVisible.Value = vbUnchecked
    Frame3.Enabled = False
    chkVisible.Enabled = False
Else
    Me.Caption = "Receiver " & ReceiverName & "(" & Str$(ReceiverNumber) & "/" & Receiver.TotalReceivers & ")"
    chkVisible.Value = vbChecked
    chkVisible.Enabled = True
    Frame3.Enabled = True
End If

'Draw
DrawReceiverInSmallCanvas ReceiverNumber

'is it part of a group?
GroupNumber = Receiver.GroupNumber(ReceiverNumber)

If GroupNumber Then
    'list number and show all members
    lstGroupMembers.Clear
    m = Receiver.GetReceiversInGroup(GroupNumber, GroupMembers())
    i = 0
    Do
        R = GroupMembers(i)
        lstGroupMembers.AddItem Receiver.ID(R, True)
        i = i + 1
    Loop Until i >= m Or i = MAX_RECEIVERS_PERGROUP
    lblOthersInGroup.Enabled = True
    txtGroupName.Text = Receiver.ID(ReceiverNumber)
    lstGroupMembers.Enabled = True
    SSTab1.TabEnabled(2) = True
    Frame2.Enabled = True
Else
    txtGroupName.Text = ""
    lstGroupMembers.Enabled = False
    lblOthersInGroup.Enabled = False
    lstGroupMembers.Clear
    chkVisible.Enabled = True
    Frame3.Enabled = True
    Frame2.Enabled = False
    SSTab1.TabEnabled(2) = False
End If

txtLat.Text = Str$(Receiver.LA(ReceiverNumber))
txtLong.Text = Str$(Receiver.LO(ReceiverNumber))

'Shape and color
picMarker.Cls
Receiver.DrawReceiver picMarker, ReceiverNumber, LARGE_MARKER, -1, CLng(picMarker.ScaleWidth / 2), CLng(picMarker.ScaleHeight / 2)

'update tag&geo pane
TagNumber = Receiver.Tag(ReceiverNumber)
cmbZoneTag.ListIndex = Receiver.Tag(ReceiverNumber)

'Determine if tag is geographical/zonal and then assign accordingly
If Receiver.TagIsGeographic(TagNumber) Then
    ShowZones
    txtZoneLat.Text = Receiver.ZoneLatRange(TagNumber)
    txtZoneLong.Text = Receiver.ZoneLongRange(TagNumber)
    ImageProcessingEngine.DrawMapToScale picZone, Form1.Picture1
    Receiver.DrawZone TagNumber, picZone, Form1.Picture1
Else
    HideZones
End If
   

'update lists for detections
lstFish.Clear
lstDates.Clear
'show fish and dates of detections
'Fish
Receiver.TransferUniqueFishEntriesToList CInt(ReceiverNumber), FishInList()
For i = 0 To MAX_FISH
    If FishInList(i) Then lstFish.AddItem FishDatabase.Code(i)
Next i

'Dates
Receiver.TransferDatesToList CInt(ReceiverNumber), lstDates

'other detection stats
lblQueryLabel.Visible = False
lstQueryResponse.Visible = False
lblDetectionsTTL.Caption = "Total Detections: " & Str$(Receiver.Detection_TTL(ReceiverNumber))
lblFishTTL.Caption = "Fish detected: " & Str$(lstFish.ListCount)
lblDaysTTL.Caption = "Days active: " & Str$(lstDates.ListCount)
CHANGED_BY_COMPUTER = False
End Sub
Private Sub DrawReceiverInSmallCanvas(R As Integer)
Dim X As Long
Dim Y As Long
X = Receiver.X(R) * LocalScale_X
Y = Receiver.Y(R) * LocalScale_Y

Receiver.DrawReceiver picRelativeLocation, R, 1, -1, X, Y

End Sub

Private Sub lstDates_Click()
'query database for fish numbers detected on date
Dim i As Long
Dim c As Long
Dim QueryResponse(MAX_DETECTIONS_PER_RECEIVER) As String
Dim NothingElseInList As Boolean

'clear
lstQueryResponse.Clear

'get selection
For i = 0 To lstDates.ListCount - 1
    If lstDates.Selected(i) = True Then
        Receiver.QueryEntry_To_GetField ReceiverNumber, lstDates.List(i), FieldNames.Date, FieldNames.Fish, QueryResponse()
        Do
            NothingElseInList = True
            If QueryResponse(c) <> "" Then
                lstQueryResponse.AddItem QueryResponse(c)
                QueryResponse(c) = ""
                NothingElseInList = False
                ShowFishOnQuery = True
            End If
            c = c + 1
        Loop Until NothingElseInList
        
        'show info about query
        lblQueryLabel.Caption = "Fish observed on " & lstDates.List(i) & ":"
    End If
Next i

'make visible
lstQueryResponse.Visible = True
lblQueryLabel.Visible = True

End Sub
Private Sub lstFish_Click()
'query database for fish dates detected
Dim i As Long
Dim c As Long
Dim QueryResponse(MAX_DETECTIONS_PER_RECEIVER) As String
Dim NothingElseInList As Boolean


'clear
lstQueryResponse.Clear

'get selection
For i = 0 To lstFish.ListCount - 1
    If lstFish.Selected(i) = True Then
        frmFloater.cmbFishCode.ListIndex = FishDatabase.GetFishNumber(lstFish.List(i)) + 1
        Me.SetFocus
        Receiver.QueryEntry_To_GetField ReceiverNumber, lstFish.List(i), FieldNames.Fish, FieldNames.Date, QueryResponse()
        Do
            NothingElseInList = True
            If QueryResponse(c) <> "" Then
                lstQueryResponse.AddItem QueryResponse(c)
                QueryResponse(c) = ""
                NothingElseInList = False
                ShowFishOnQuery = False
            End If
            c = c + 1
        Loop Until NothingElseInList
        'show info about query
        lblQueryLabel.Caption = "Dates #" & lstFish.List(i) & " was observed:"
    End If
Next i

'make visible
lstQueryResponse.Visible = True
lblQueryLabel.Visible = True
End Sub
Private Sub lstFish_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
'special context menu
Set ActiveList = lstFish
If Button = vbRightButton Then PopupMenu mnuTracks
End Sub

Private Sub lstQueryResponse_Click()
Dim i As Integer

If ShowFishOnQuery Then
    For i = 0 To lstQueryResponse.ListCount - 1
        If lstQueryResponse.Selected(i) = True Then
            frmFloater.cmbFishCode.ListIndex = FishDatabase.GetFishNumber(lstQueryResponse.List(i)) + 1
        End If
    Next i
    Me.SetFocus
End If
End Sub

Private Sub lstQueryResponse_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
'special context menu
If ShowFishOnQuery Then
    Set ActiveList = lstQueryResponse
    If Button = vbRightButton Then PopupMenu mnuTracks
End If
End Sub

Private Sub mnuColorize_Click()
Dim i As Long
Dim f As Integer
Dim c As Long

On Error GoTo ExitWithError

'ask user to choose color
With CommonDialog
    .CancelError = True
    .ShowColor
    c = .Color
End With

'select into tracks
For i = 0 To ActiveList.ListCount - 1
    f = FishDatabase.GetFishNumber(ActiveList.List(i))
    FishDatabase.Color(f) = c
Next i

'show on canvas
frmFloater.cmbFishCode.ListIndex = 0
frmFloater.RefreshCanvas
Me.SetFocus

ExitWithError:
'NOP
End Sub

Private Sub mnuExclude_Click()
'exclude all tracks in receiver
Dim i As Long
Dim f As Integer

For i = 0 To ActiveList.ListCount - 1
    f = FishDatabase.GetFishNumber(ActiveList.List(i))
    FishDatabase.IsVisible(f) = False
Next i

'show on canvas
frmFloater.cmbFishCode.ListIndex = 0
frmFloater.RefreshCanvas
Me.SetFocus

End Sub

Private Sub optSource_Click(Index As Integer)
If optSource(ECOLOGICAL).Value = True Then
    lblLat.Enabled = True
    txtZoneLat.Enabled = True
    lblLong.Enabled = True
    txtZoneLong.Enabled = True
    picZone.Enabled = True
Else
    lblLat.Enabled = False
    txtZoneLat.Enabled = False
    lblLong.Enabled = False
    txtZoneLong.Enabled = False
    picZone.Enabled = False
End If
End Sub

Private Sub picRelativeLocation_Paint()
Dim GroupNumber As Integer

If ZoomRegion.Zoomed Then
    ImageProcessingEngine.DrawZoomMapToScale picRelativeLocation, Form1.Picture1
Else
    ImageProcessingEngine.DrawMapToScale picRelativeLocation, Form1.Picture1
End If
DrawReceiverInSmallCanvas ReceiverNumber
End Sub

Private Sub picZone_Paint()
If Receiver.TagIsGeographic(TagNumber) Then
    ImageProcessingEngine.DrawMapToScale picZone, Form1.Picture1
    Receiver.DrawZone TagNumber, picZone, Form1.Picture1
End If
End Sub
 Private Sub txtGroupName_Click()
Receiver_Selected(ReceiverNumber) = True
frmAssignReceiverToGroup.Show
End Sub


