VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{7D913FD5-17EC-11D6-AC31-88D14E8BE65B}#1.0#0"; "VISUALIZATION.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visualization Exscample"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7965
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   7965
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command17 
      Caption         =   "Extra Colors >>"
      Height          =   375
      Left            =   4320
      TabIndex        =   24
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Frame Frame7 
      Caption         =   "Peaks color"
      Height          =   615
      Left            =   5760
      TabIndex        =   38
      Top             =   3480
      Width           =   2175
      Begin VB.PictureBox picback3 
         BackColor       =   &H0000C0C0&
         Height          =   255
         Left            =   480
         ScaleHeight     =   195
         ScaleWidth      =   675
         TabIndex        =   40
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command16 
         Caption         =   "..."
         Height          =   255
         Left            =   1320
         TabIndex        =   39
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Line color"
      Height          =   1935
      Left            =   5760
      TabIndex        =   30
      Top             =   1440
      Width           =   2175
      Begin VB.OptionButton Option4 
         Caption         =   "Random"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Fire"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1320
         Width           =   1935
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Gradient"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   1080
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Normal"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   840
         Width           =   1935
      End
      Begin VB.PictureBox picback2 
         BackColor       =   &H0000FFFF&
         Height          =   255
         Left            =   480
         ScaleHeight     =   195
         ScaleWidth      =   675
         TabIndex        =   32
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command15 
         Caption         =   "..."
         Height          =   255
         Left            =   1320
         TabIndex        =   31
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Line color:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7560
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame5 
      Caption         =   "Backcolor"
      Height          =   1215
      Left            =   5760
      TabIndex        =   25
      Top             =   120
      Width           =   2175
      Begin VB.CheckBox Check1 
         Caption         =   "Gradient background"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton Command14 
         Caption         =   "..."
         Height          =   255
         Left            =   1320
         TabIndex        =   28
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox picback 
         BackColor       =   &H00000000&
         Height          =   255
         Left            =   480
         ScaleHeight     =   195
         ScaleWidth      =   675
         TabIndex        =   27
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Background color:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Extra Visualisation >>"
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Frame Frame4 
      Caption         =   "Visualisation type"
      Height          =   735
      Left            =   0
      TabIndex        =   19
      Top             =   1680
      Visible         =   0   'False
      Width           =   5655
      Begin VB.CommandButton Command12 
         Caption         =   "Curve"
         Height          =   375
         Left            =   3720
         TabIndex        =   22
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton Command11 
         Caption         =   "RoomFC"
         Height          =   375
         Left            =   1920
         TabIndex        =   21
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Sonar"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Visualisation type"
      Height          =   1095
      Left            =   0
      TabIndex        =   12
      Top             =   1680
      Width           =   5655
      Begin VB.CommandButton Command2 
         Caption         =   "Osciliscope"
         Height          =   375
         Left            =   3720
         TabIndex        =   18
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Osciliscope double"
         Height          =   375
         Left            =   1920
         TabIndex        =   17
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Osciliscope spec"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Bar filled"
         Height          =   375
         Left            =   3720
         TabIndex        =   15
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Bar"
         Height          =   375
         Left            =   1920
         TabIndex        =   14
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Spectrum"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Height          =   75
      Left            =   120
      TabIndex        =   9
      Top             =   3840
      Width           =   5535
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   3570
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      Max             =   50
      SelStart        =   5
      Value           =   5
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Xray Style"
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Yellow Style"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Exit"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   975
   End
   Begin VB.PictureBox picunder 
      BackColor       =   &H00000000&
      Height          =   1335
      Left            =   120
      ScaleHeight     =   1275
      ScaleWidth      =   5475
      TabIndex        =   1
      Top             =   120
      Width           =   5535
      Begin Visualization.Visu Visu1 
         Height          =   495
         Left            =   1680
         TabIndex        =   43
         Top             =   240
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   873
         Enabled         =   0   'False
         Style           =   2
         Draw_Typ        =   4
         Thick_Fall      =   4
         PeakColor       =   49344
         LineColor       =   65535
         BackColor       =   0
         GradientColor   =   65535
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "To detect sound you must play any music file (eg. mp3 with winamp)"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   1050
         Width           =   5415
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "To see better beat analayse use extra visualisation"
         ForeColor       =   &H0000FFFF&
         Height          =   855
         Left            =   4200
         TabIndex        =   8
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lblFps 
         BackStyle       =   0  'Transparent
         Caption         =   "0 FPS"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   75
      Left            =   0
      TabIndex        =   0
      Top             =   1560
      Width           =   5655
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   255
      Left            =   2880
      TabIndex        =   42
      Top             =   3585
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      Min             =   1
      Max             =   150
      SelStart        =   25
      Value           =   25
   End
   Begin VB.Label Label7 
      Caption         =   "Refresh Speed"
      Height          =   255
      Left            =   3120
      TabIndex        =   41
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Visualization by Dalibor Kopjar"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3960
      Width           =   5535
   End
   Begin VB.Label Label1 
      Caption         =   "Peaks Fall Of"
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   3360
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
    Visu1.BackgroundType = GradientBackground
Else
    Visu1.BackgroundType = ColorBackground
End If
End Sub

Private Sub Command1_Click()
Visu1.VisualizationType = Spectrum
End Sub

Private Sub Command10_Click()
Visu1.GraficStyle = YellowStyle
End Sub

Private Sub Command11_Click()
Visu1.VisualizationType = RoomFC
End Sub

Private Sub Command12_Click()
Visu1.VisualizationType = Curve
End Sub

Private Sub Command13_Click()
If Command13.Caption = "Extra Visualisation >>" Then
    Command13.Caption = "Normal Visualisation >>"
    Frame4.Visible = True
    Frame3.Visible = False
    
    Visu1.Top = 10
    Visu1.Height = 1250
    Label4.Visible = False
    Visu1.Left = 1500
    Visu1.Width = 2100
    
Else
    Command13.Caption = "Extra Visualisation >>"
    Frame4.Visible = False
    Frame3.Visible = True
    Visu1.Top = 250
    Visu1.Height = 500
    Label4.Visible = True
    Visu1.Left = 1900
    Visu1.Width = 1500
End If
End Sub

Private Sub Command14_Click()
cd1.ShowColor
picback.BackColor = cd1.Color
Visu1.BackColor = cd1.Color
If Check1.Value = 0 Then
Visu1.BackgroundType = ColorBackground
End If
End Sub

Private Sub Command15_Click()
cd1.ShowColor
picback2.BackColor = cd1.Color
Visu1.LineColor = cd1.Color
Visu1.GradientColor = cd1.Color
End Sub

Private Sub Command16_Click()
cd1.ShowColor
picback3.BackColor = cd1.Color
Visu1.PeakColor = cd1.Color
End Sub

Private Sub Command17_Click()
If Command17.Caption = "Extra Colors >>" Then
    Command17.Caption = "Close Extra <<"
    Me.Width = 8085
Else
    Command17.Caption = "Extra Colors >>"
    Me.Width = 5775
End If
End Sub

Private Sub Command2_Click()
Visu1.VisualizationType = Osciliscope
End Sub

Private Sub Command3_Click()
Visu1.VisualizationType = DoubleOsciliscope
End Sub

Private Sub Command4_Click()
Visu1.VisualizationType = SpecOsciliscope
End Sub

Private Sub Command5_Click()
Visu1.VisualizationType = Bar
End Sub

Private Sub Command6_Click()
Visu1.VisualizationType = BarFilled
End Sub

Private Sub Command7_Click()
Visu1.VisualizationType = Sonar
End Sub

Private Sub Command8_Click()
Unload Me
End Sub

Private Sub Command9_Click()
Visu1.GraficStyle = XrayStyle
End Sub

Private Sub Form_Load()
Frame4.Visible = False
    Frame3.Visible = True
    Visu1.Top = 250
    Visu1.Height = 500
    Label4.Visible = True
    Visu1.Left = 1900
    Visu1.Width = 1500
    
Visu1.GraficStyle = YellowStyle
Visu1.Enabled = True
Me.Width = 5775
End Sub

Private Sub Option1_Click()
Visu1.ColorDrawType = Normal
End Sub

Private Sub Option2_Click()
Visu1.ColorDrawType = Gradient
End Sub

Private Sub Option3_Click()
Visu1.ColorDrawType = RealFire
End Sub

Private Sub Option4_Click()
Visu1.ColorDrawType = RandomColor
End Sub

Private Sub Slider1_Scroll()
Visu1.PeaksFallOf = Slider1.Value
End Sub

Private Sub Slider2_Scroll()
Visu1.RefreshSpeed = Slider2.Value
End Sub

Private Sub Visu1_FramesPerSecond(ByVal optFps As Long)
lblFps.Caption = optFps & " FPS"
End Sub
