VERSION 5.00
Begin VB.UserControl Visu 
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1500
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   PropertyPages   =   "Visu1.ctx":0000
   ScaleHeight     =   450
   ScaleWidth      =   1500
   ToolboxBitmap   =   "Visu1.ctx":0026
   Begin VB.Timer Timer3 
      Interval        =   25
      Left            =   360
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   1440
      Top             =   600
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   100
      ScaleMode       =   0  'User
      ScaleWidth      =   200
      TabIndex        =   1
      Top             =   0
      Width           =   1845
      Begin VB.Timer Timer1 
         Interval        =   20
         Left            =   0
         Top             =   0
      End
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1440
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "Visu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'****************************************************************
'*                       Visualization                          *
'*                   Coded by Dalibor Kopjar                    *
'*    You are free to use the source code in your private,      *
'*  non-commercial, projects without permission. If you want    *
'* to use this code in commercial projects EXPLICIT permission  *
'*                from the author is required.                  *
'*                                                              *
'*                                                              *
'*               Copyright © Dalibor Kopjar 2002                *
'****************************************************************

Option Explicit

Dim hmixer As Long                      ' mixer handle
Dim inputVolCtrl As MIXERCONTROL        ' waveout volume control
Dim outputVolCtrl As MIXERCONTROL       ' microphone volume control
Dim rc As Long                          ' return code
Dim OK As Boolean                       ' boolean return code
Dim mxcd As MIXERCONTROLDETAILS         ' control info
Dim vol As MIXERCONTROLDETAILS_SIGNED   ' control's signed value
Dim volume As Long                      ' volume value
Dim volHmem As Long                     ' Volume Buffer
Private VU As VULights                  ' Volume Unit Values
Private FreqNum As Frequency
Private Declare Sub ReleaseCapture Lib "user32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Public Enum d_visu_type
    Spectrum = 1
    Bar = 2
    BarFilled = 3
    Osciliscope = 4
    DoubleOsciliscope = 5
    SpecOsciliscope = 6
    Sonar = 7
    RoomFC = 8
    Curve = 9
End Enum

Public Enum d_type_refresh
    FrequencyRefresh = 1
    WaveVolumeRefresh = 2
End Enum

Public Enum d_type_drawcolor
    Normal = 1
    Spec4Colors = 2
    RandomColor = 3
    Gradient = 4
    RealFire = 5
End Enum

Public Enum d_back_type
    ColorBackground = 1
    GradientBackground = 2
    Picture = 3
End Enum

Public Enum d_Style
    Custom = 1
    YellowStyle = 2
    XrayStyle = 3
End Enum

Dim i, a
Dim Dx As Long
Dim Cx As Long
Dim Dy As Long
Dim AddRate As Single

Dim Color As Long
Dim CurFps As Long
Dim UDD As Long

Dim MyFreqs(1 To 50) As Double
Dim MyThick(1 To 50) As s_MyThick
Dim MyThick2(1 To 50) As s_MyThick
Dim MyThick3(1 To 50) As s_MyThick
Dim MyColor(0 To 100) As Long
Dim MyFire(0 To 100) As Long

Const NumLevels As Long = 101
Const GradientHor = 200
Const GradientVertical = 100

Public Enum d_gra_type
    VerticalGradient = 1
    HorizonatalGradient = 2
End Enum

Dim Cl As Long
Dim AddCl As Long
Dim BackGr As Long
Dim ret As Long
Dim SnC As Long
Dim RoomCH As Integer

Public Event Click()
Public Event DoubleClick()
Public Event FramesPerSecond(ByVal optFps As Long)

Const m_def_Enabled = 1
Const m_def_RefreshSpeed = 25
Const m_def_Typ = 1
Const m_def_TypDraw = 1
Const m_def_Thick = 5
Const m_def_TypRefresh = 1
Const m_def_PeaksON As Boolean = True
Const m_def_Style = 1
Const m_def_Background = 1
Const m_def_PeakColor = vbWhite
Const m_def_LineColor = vbRed
Const m_def_BackColor = &H800000
Const m_def_Fire1 = vbGreen
Const m_def_Fire2 = vbYellow
Const m_def_Fire3 = &H80FF&
Const m_def_Fire4 = vbRed
Const m_def_GradientC = vbRed
Const m_def_BackVH = 1

Dim m_BackPic As StdPicture
Dim m_BackVH As Long
Dim m_Background As Long
Dim m_Style As Long
Dim m_PeaksON As Boolean
Dim m_TypDraw As Long
Dim m_Typ As Long
Dim m_Enabled As Boolean
Dim m_RefreshSpeed As Long
Dim m_Thick As Long
Dim m_TypRefresh As Long
Dim m_PeakColor As Long
Dim m_LineColor As Long
Dim m_BackColor As Long
Dim m_Fire1 As Long
Dim m_Fire2 As Long
Dim m_Fire3 As Long
Dim m_Fire4 As Long
Dim m_GradientC As Long

Private Function RefreshOne()
For i = 1 To UBound(MyFreqs)
    MyFreqs(i) = VU.VolLev
Next i

End Function
Private Function RefreshFreqs()
    Freq1
    Freq2
    Freq3
    Freq4
    Freq5
    Freq6
    Freq7
    Freq8
    Freq9
    Freq10
    Freq11
    Freq12
    Freq13
    Freq14
    Freq15
    Freq16
    Freq17
    Freq18
    Freq19
    Freq20
    Freq21
    Freq22
    Freq23
    Freq24
    Freq25
    Freq26
    Freq27
    Freq28
    Freq29
    Freq31
    Freq32
    Freq33
    Freq34
    Freq35
    Freq36
    Freq37
    Freq38
    Freq39
    Freq40
    Freq41
    Freq42
    Freq43
    Freq44
    Freq45
    Freq46
    Freq47
    Freq48
    Freq49
    Freq50
End Function
Private Function SetAllToNull()
For i = 1 To UBound(MyFreqs)
MyFreqs(i) = 0
Next i
End Function
Private Sub VolVal(VolIs As Long, VolFreq As Double)
For FreqNum = 0 To UBound(VU.Freq)
Next FreqNum
VolIs = volume * 327.67
VolFreq = VU.Freq(FreqNum)
VU.FreqVal = VolIs * VolFreq
End Sub

Private Function GetColor() As Long
Randomize:
   GetColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
End Function
Private Function RefreshThicks()
If m_Typ = 1 Or m_Typ = 2 Or m_Typ = 3 Then
    Cx = AddRate
    For i = 1 To UBound(MyThick)
        picMain.PSet (Cx, picMain.ScaleHeight - MyThick(i).CurY), m_PeakColor
        picMain.PSet (Cx, picMain.ScaleHeight - MyThick2(i).CurY), m_PeakColor
        picMain.PSet (Cx, picMain.ScaleHeight - MyThick3(i).CurY), m_PeakColor
       
        Cx = Cx + (AddRate * 2)
    Next i
End If
End Function

Public Property Let BackColor(ByVal New_Color As OLE_COLOR)
    m_BackColor = New_Color
    PropertyChanged "BackColor"
RefreshBack
DrawSpec
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property
Public Property Let PeakColor(ByVal New_Color As OLE_COLOR)
    m_PeakColor = New_Color
    PropertyChanged "PeakColor"
DrawSpec
End Property
Public Property Get GradientColor() As OLE_COLOR
    GradientColor = m_GradientC
End Property
Public Property Let GradientColor(ByVal New_Color As OLE_COLOR)
    m_GradientC = New_Color
    PropertyChanged "GradientColor"

MakeGradientColors m_GradientC

End Property
Public Property Get PeakColor() As OLE_COLOR
    PeakColor = m_PeakColor
End Property
Public Property Let LineColor(ByVal New_Color As OLE_COLOR)
    m_LineColor = New_Color
    PropertyChanged "LineColor"
DrawSpec
End Property
Public Property Get CustomColor1() As OLE_COLOR
    CustomColor1 = m_Fire1
End Property
Public Property Let CustomColor1(ByVal New_Color As OLE_COLOR)
    m_Fire1 = New_Color
    PropertyChanged "Fire1"
DrawSpec
End Property
Public Property Get CustomColor2() As OLE_COLOR
    CustomColor2 = m_Fire2
End Property
Public Property Let CustomColor2(ByVal New_Color As OLE_COLOR)
    m_Fire2 = New_Color
    PropertyChanged "Fire2"
DrawSpec
End Property
Public Property Get CustomColor3() As OLE_COLOR
    CustomColor3 = m_Fire3
End Property
Public Property Let CustomColor3(ByVal New_Color As OLE_COLOR)
    m_Fire3 = New_Color
    PropertyChanged "Fire3"
DrawSpec
End Property
Public Property Get CustomColor4() As OLE_COLOR
    CustomColor4 = m_Fire4
End Property
Public Property Let CustomColor4(ByVal New_Color As OLE_COLOR)
    m_Fire4 = New_Color
    PropertyChanged "Fire4"
DrawSpec
End Property
Public Property Get LineColor() As OLE_COLOR
    LineColor = m_LineColor
End Property
Public Property Get RefreshType() As d_type_refresh
    RefreshType = m_TypRefresh
End Property
Public Property Let RefreshType(ByVal New_Type As d_type_refresh)
    m_TypRefresh = New_Type
    PropertyChanged "Type_Refresh"
DrawSpec
End Property
Public Property Get RefreshSpeed() As Long
Attribute RefreshSpeed.VB_ProcData.VB_Invoke_Property = "General"
    RefreshSpeed = m_RefreshSpeed
End Property
Public Property Let RefreshSpeed(ByVal New_Speed As Long)
    If New_Speed < 1 Then Exit Property
    m_RefreshSpeed = New_Speed
    Timer3.Interval = m_RefreshSpeed
    PropertyChanged "RefreshSpeed"
End Property
Public Property Set BackgroundPicture(ByVal New_Pic As StdPicture)
    Set m_BackPic = New_Pic
    PropertyChanged "Picture"
RefreshBack
End Property
Public Property Get BackgroundPicture() As StdPicture
    Set BackgroundPicture = m_BackPic
End Property
Public Property Get VisualizationType() As d_visu_type
    VisualizationType = m_Typ
End Property
Public Property Let VisualizationType(ByVal New_Type As d_visu_type)
    If m_Typ = 7 Or m_Typ = 8 Or m_Typ = 9 Then
        If m_TypRefresh = 2 And New_Type <> 7 And New_Type <> 8 And New_Type <> 9 Then
        m_TypRefresh = 1
        End If
    End If
    
    m_Typ = New_Type
    
    If m_Typ = 7 Or m_Typ = 8 Or m_Typ = 9 Then
        m_TypRefresh = 2
        PropertyChanged "Type_Refresh"
    End If
    
    PropertyChanged "Visu_Typ"
End Property
Public Property Get GraficStyle() As d_Style
    GraficStyle = m_Style
End Property
Public Property Let GraficStyle(ByVal New_Style As d_Style)
    m_Style = New_Style
    If m_Style = 2 Then
        m_BackColor = vbBlack
        m_PeakColor = &HC0C0&
        m_GradientC = vbYellow
        m_Typ = 1
        m_TypDraw = 4
        m_PeaksON = True
        m_Thick = 4
        m_TypRefresh = 1
    
    PropertyChanged "Type_Refresh"
    PropertyChanged "PeakColor"
    PropertyChanged "BackColor"
    PropertyChanged "Draw_Typ"
    PropertyChanged "Visu_Typ"
    PropertyChanged "Thick_Fall"
    PropertyChanged "PeaksON"
    PropertyChanged "Style"
    PropertyChanged "GradientColor"
    
    MakeGradientColors m_GradientC
    RefreshBack
    
    ElseIf m_Style = 3 Then
        m_BackColor = vbWhite
        m_PeakColor = &H800000
        m_GradientC = vbBlue
        m_Typ = 1
        m_TypDraw = 4
        m_PeaksON = True
        m_Thick = 4
        m_TypRefresh = 1
        
    PropertyChanged "GradientColor"
    PropertyChanged "Type_Refresh"
    PropertyChanged "Fire1"
    PropertyChanged "Fire2"
    PropertyChanged "Fire3"
    PropertyChanged "Fire4"
    PropertyChanged "PeakColor"
    PropertyChanged "BackColor"
    PropertyChanged "Draw_Typ"
    PropertyChanged "Visu_Typ"
    PropertyChanged "Thick_Fall"
    PropertyChanged "PeaksON"
    PropertyChanged "Style"
    
    MakeGradientColors m_GradientC
    RefreshBack
    
    End If
End Property
Public Property Get ColorDrawType() As d_type_drawcolor
    ColorDrawType = m_TypDraw
End Property
Public Property Let ColorDrawType(ByVal New_Type As d_type_drawcolor)
    m_TypDraw = New_Type
    PropertyChanged "Draw_Typ"
End Property
Public Property Get BackgroundType() As d_back_type
    BackgroundType = m_Background
End Property
Public Property Let BackgroundType(ByVal New_Type As d_back_type)
    m_Background = New_Type
    RefreshBack
    PropertyChanged "Background"
End Property
Public Property Get BackgroundGradientType() As d_gra_type
    BackgroundGradientType = m_BackVH
End Property
Public Property Let BackgroundGradientType(ByVal New_Type As d_gra_type)
    m_BackVH = New_Type
    RefreshBack
    PropertyChanged "BackVH"
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    Timer1.Enabled = m_Enabled
    Timer3.Enabled = m_Enabled
    PropertyChanged "Enabled"
End Property
Public Property Get PeaksON() As Boolean
Attribute PeaksON.VB_ProcData.VB_Invoke_Property = "General"
    PeaksON = m_PeaksON
End Property
Public Property Let PeaksON(ByVal New_Peaks As Boolean)
    m_PeaksON = New_Peaks
    PropertyChanged "PeaksON"
End Property
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = "General"
    Enabled = m_Enabled
End Property
Public Property Let PeaksFallOf(ByVal New_Peak As Long)
    If New_Peak < 0 Then Exit Property
    m_Thick = New_Peak
    PropertyChanged "Thick_Fall"
End Property

Public Property Get PeaksFallOf() As Long
Attribute PeaksFallOf.VB_ProcData.VB_Invoke_Property = "General"
    PeaksFallOf = m_Thick
End Property

Private Sub picMain_Click()
RaiseEvent Click

End Sub

Private Sub picMain_DblClick()
RaiseEvent DoubleClick
End Sub

Private Sub Timer1_Timer()
If m_Enabled = False Then Exit Sub
    VU.VolLev = volume / 327.67
    
    If (volume < 0) Then volume = -volume
    ' Get the current output level
    If (1 = 1) Then
    mxcd.dwControlID = outputVolCtrl.dwControlID
    mxcd.item = outputVolCtrl.cMultipleItems
    rc = mixerGetControlDetails(hmixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
    CopyStructFromPtr volume, mxcd.paDetails, Len(volume)
    If (volume < 0) Then volume = -volume
    End If
    
End Sub
Private Function DrawRoom()
picMain.Cls
RoomCH = 0
    For i = 1 To MyFreqs(1) Step 2
    
        If m_TypDraw = 1 Then
        Color = m_LineColor
        ElseIf m_TypDraw = 2 Then
            If i <= 25 Then
            Color = m_Fire1
            ElseIf i <= 50 Then
            Color = m_Fire2
            ElseIf i <= 75 Then
            Color = m_Fire3
            Else
            Color = m_Fire4
            End If
        ElseIf m_TypDraw = 3 Then
        Color = GetColor
        ElseIf m_TypDraw = 4 Then
            If i > 100 Then
            Color = MyColor(100)
            Else
            Color = MyColor(i)
            End If
        ElseIf m_TypDraw = 5 Then
            If i > 100 Then
            Color = MyFire(100)
            Else
            Color = MyFire(i)
            End If
        End If
        
        If RoomCH = 0 Then
        picMain.Line (100 - i, 50 + (i / 2))-(200 - (100 - i), 50 + (i / 2)), Color
        picMain.Line (100 - i, 50 - (i / 2))-(200 - (100 - i), 50 - (i / 2)), Color
        picMain.Line (100 - i, 50 - (i / 2))-(100 - i, 50 + (i / 2)), Color
        picMain.Line (100 + i, 50 - (i / 2))-(100 + i, 50 + (i / 2)), Color
        RoomCH = 1
        Else
        RoomCH = 0
        End If
    Next i
    

End Function
Private Function DrawCurve()
picMain.Cls
    For i = 1 To MyFreqs(1) Step 2
    
        If m_TypDraw = 1 Then
        Color = m_LineColor
        ElseIf m_TypDraw = 2 Then
            If i <= 25 Then
            Color = m_Fire1
            ElseIf i <= 50 Then
            Color = m_Fire2
            ElseIf i <= 75 Then
            Color = m_Fire3
            Else
            Color = m_Fire4
            End If
        ElseIf m_TypDraw = 3 Then
        Color = GetColor
        ElseIf m_TypDraw = 4 Then
            If i > 100 Then
            Color = MyColor(100)
            Else
            Color = MyColor(i)
            End If
        ElseIf m_TypDraw = 5 Then
            If i > 100 Then
            Color = MyFire(100)
            Else
            Color = MyFire(i)
            End If
        End If
        
        
            picMain.Line (i * 2, 50 + (i / 2))-(200 - (i * 2), 50 - (i / 2)), Color
           
    Next i
    

End Function
Private Function DrawSonar()
picMain.Cls
SnC = 0
        If m_TypDraw = 1 Then
        Color = m_LineColor
        ElseIf m_TypDraw = 2 Then
            If MyFreqs(1) <= 25 Then
            Color = m_Fire1
            ElseIf MyFreqs(1) <= 50 Then
            Color = m_Fire2
            ElseIf MyFreqs(1) <= 75 Then
            Color = m_Fire3
            Else
            Color = m_Fire4
            End If
        ElseIf m_TypDraw = 3 Then
        Color = GetColor
        ElseIf m_TypDraw = 4 Then
            If MyFreqs(1) > 100 Then
            Color = MyColor(100)
            Else
            Color = MyColor(MyFreqs(1))
            End If
        ElseIf m_TypDraw = 5 Then
            If MyFreqs(1) > 100 Then
            Color = MyFire(100)
            Else
            Color = MyFire(MyFreqs(1))
            End If
        End If
        
        picMain.Circle (picMain.ScaleWidth / 2, picMain.ScaleHeight / 2), MyFreqs(1), Color
        
        If MyFreqs(1) > 1 Then
            For i = 1 To MyFreqs(1) Step 1
                If SnC = 4 Then
                    If m_TypDraw = 1 Then
                        Color = m_LineColor
                    ElseIf m_TypDraw = 2 Then
                            If i <= 25 Then
                            Color = m_Fire1
                            ElseIf i <= 50 Then
                            Color = m_Fire2
                            ElseIf i <= 75 Then
                            Color = m_Fire3
                            Else
                            Color = m_Fire4
                            End If
                    ElseIf m_TypDraw = 3 Then
                        Color = GetColor
                    ElseIf m_TypDraw = 4 Then
                        If i > 100 Then
                        Color = MyColor(100)
                        Else
                        Color = MyColor(i)
                        End If
                    ElseIf m_TypDraw = 5 Then
                        If i > 100 Then
                        Color = MyFire(100)
                        Else
                        Color = MyFire(i)
                        End If
                    End If
                        picMain.Circle (picMain.ScaleWidth / 2, picMain.ScaleHeight / 2), i, Color
                    SnC = 0
                Else
                    SnC = SnC + 1
                End If
            Next i
        End If
        

End Function
Private Function DrawSpec()
Dx = 100
Dy = 0

    picMain.Cls
    
    For i = 1 To UBound(MyFreqs)
    
        If m_TypDraw = 1 Then
        Color = m_LineColor
        ElseIf m_TypDraw = 2 Then
            If MyFreqs(i) <= 25 Then
            Color = m_Fire1
            ElseIf MyFreqs(i) <= 50 Then
            Color = m_Fire2
            ElseIf MyFreqs(i) <= 75 Then
            Color = m_Fire3
            Else
            Color = m_Fire4
            End If
        ElseIf m_TypDraw = 3 Then
        Color = GetColor
        ElseIf m_TypDraw = 4 Then
            If MyFreqs(i) > 100 Then
            Color = MyColor(100)
            Else
            Color = MyColor(MyFreqs(i))
            End If
        ElseIf m_TypDraw = 5 Then
            If MyFreqs(i) > 100 Then
            Color = MyFire(100)
            Else
            Color = MyFire(MyFreqs(i))
            End If
        End If
        
        picMain.Line (Dy, picMain.ScaleHeight)-(Dy + AddRate, picMain.ScaleHeight - MyFreqs(i)), Color
        picMain.Line (Dy + AddRate, picMain.ScaleHeight - MyFreqs(i))-(Dy + (AddRate * 2), picMain.ScaleHeight), Color
        
        If MyFreqs(i) > MyThick(i).CurY Then
            MyThick(i).CurY = MyFreqs(i)
            MyThick(i).FallOf = 0
            MyThick(i).StandTime = m_Thick
            MyThick2(i).CurY = MyFreqs(i)
            MyThick2(i).FallOf = 0
            MyThick2(i).StandTime = m_Thick + 1
            MyThick3(i).CurY = MyFreqs(i)
            MyThick3(i).FallOf = 0
            MyThick3(i).StandTime = m_Thick + 2
        End If
        
        Dx = picMain.ScaleHeight
        Dy = Dy + (AddRate * 2)
    Next i



End Function
Private Function DrawBar1()
Dx = picMain.ScaleHeight
Dy = AddRate
picMain.Cls

For i = 1 To UBound(MyFreqs)
        If m_TypDraw = 1 Then
        Color = m_LineColor
        ElseIf m_TypDraw = 2 Then
            If MyFreqs(i) <= 25 Then
            Color = m_Fire1
            ElseIf MyFreqs(i) <= 50 Then
            Color = m_Fire2
            ElseIf MyFreqs(i) <= 75 Then
            Color = m_Fire3
            Else
            Color = m_Fire4
            End If
        ElseIf m_TypDraw = 3 Then
        Color = GetColor
        ElseIf m_TypDraw = 4 Then
            If MyFreqs(i) > 100 Then
            Color = MyColor(100)
            Else
            Color = MyColor(MyFreqs(i))
            End If
        ElseIf m_TypDraw = 5 Then
            If MyFreqs(i) > 100 Then
            Color = MyFire(100)
            Else
            Color = MyFire(MyFreqs(i))
            End If
        End If

        picMain.Line (Dy, picMain.ScaleHeight)-(Dy, picMain.ScaleHeight - MyFreqs(i)), Color
        
    
    If MyFreqs(i) > MyThick(i).CurY Then
            MyThick(i).CurY = MyFreqs(i)
            MyThick(i).FallOf = 0
            MyThick(i).StandTime = m_Thick
            MyThick2(i).CurY = MyFreqs(i)
            MyThick2(i).FallOf = 0
            MyThick2(i).StandTime = m_Thick + 1
            MyThick3(i).CurY = MyFreqs(i)
            MyThick3(i).FallOf = 0
            MyThick3(i).StandTime = m_Thick + 2
        End If
    Dy = Dy + (AddRate * 2)
Next i

End Function
Private Function MakeGradientBack(ByVal MainColor As Long, ByVal Horz As Long)
On Error Resume Next
Dim MaxLn As Long
If Horz = 0 Then
    MaxLn = 100
Else
    MaxLn = 200
End If

    Dim R1() As Integer
    Dim G1() As Integer
    Dim B1() As Integer
    ReDim R1(MaxLn) As Integer
    ReDim G1(MaxLn) As Integer
    ReDim B1(MaxLn) As Integer
    
    Dim REnd(1 To 5) As Integer
    Dim GEnd(1 To 5) As Integer
    Dim BEnd(1 To 5) As Integer
    
    Dim i As Integer, End1 As Integer, End2 As Integer, End3 As Integer
    Dim ColorString As String
    Dim Counter As Integer
    
    If MaxLn = 200 Then
    End1 = 0
    End2 = 124
    End3 = 125
    Else
    End1 = 0
    End2 = 60
    End3 = 61
    End If
    'get RGB values
    For i = 1 To 3
        If i = 1 Then
        ColorString = Hex(vbWhite)
        ElseIf i = 2 Then
        ColorString = Hex(MainColor)
        Else
        ColorString = Hex(vbBlack)
        End If
        If Len(ColorString) = 2 Then
            BEnd(i) = 0
            GEnd(i) = 0
            REnd(i) = CLng("&H" & ColorString)
        ElseIf Len(ColorString) = 4 Then
            BEnd(i) = 0
            GEnd(i) = CLng("&H" & Left$(ColorString, 2))
            REnd(i) = CLng("&H" & Right$(ColorString, 2))
        ElseIf Len(ColorString$) = 6 Then
            BEnd(i) = CLng("&H" & Left$(ColorString, 2))
            GEnd(i) = CLng("&H" & Mid$(ColorString, 3, 2))
            REnd(i) = CLng("&H" & Right$(ColorString, 2))
        End If
    Next i
    
    'Auto calculate mixed "in-between" colors
    For i = 4 To 5 Step 1
        If REnd(i - 2) > REnd(i - 3) Then
            REnd(i) = REnd(i - 2)
        Else
            REnd(i) = REnd(i - 3)
        End If

        If GEnd(i - 2) > GEnd(i - 3) Then
            GEnd(i) = GEnd(i - 2)
        Else
            GEnd(i) = GEnd(i - 3)
        End If

        If BEnd(i - 2) > BEnd(i - 3) Then
            BEnd(i) = BEnd(i - 2)
        Else
            BEnd(i) = BEnd(i - 3)
        End If
    Next i
    
    'set color levels
    For i = 1 To End1
        R1(i) = (i - 1) * (REnd(4) - REnd(1)) / (End1 + 1) + REnd(1)
        G1(i) = (i - 1) * (GEnd(4) - GEnd(1)) / (End1 + 1) + GEnd(1)
        B1(i) = (i - 1) * (BEnd(4) - BEnd(1)) / (End1 + 1) + BEnd(1)
    Next
    Counter = 0

    For i = End1 + 1 To End2
        Counter = Counter + 1
        R1(i) = Counter * (REnd(2) - REnd(4)) / (End2 - End1 + 1) + REnd(4)
        G1(i) = Counter * (GEnd(2) - GEnd(4)) / (End2 - End1 + 1) + GEnd(4)
        B1(i) = Counter * (BEnd(2) - BEnd(4)) / (End2 - End1 + 1) + BEnd(4)
    Next
    Counter = 0

    For i = End2 + 1 To End3
        Counter = Counter + 1
        R1(i) = Counter * (REnd(5) - REnd(2)) / (End3 - End2 + 1) + REnd(2)
        G1(i) = Counter * (GEnd(5) - GEnd(2)) / (End3 - End2 + 1) + GEnd(2)
        B1(i) = Counter * (BEnd(5) - BEnd(2)) / (End3 - End2 + 1) + BEnd(2)
    Next
    Counter = 0

    For i = End3 + 1 To MaxLn
        Counter = Counter + 1
        R1(i) = Counter * (REnd(3) - REnd(5)) / (MaxLn - End3 + 1) + REnd(5)
        G1(i) = Counter * (GEnd(3) - GEnd(5)) / (MaxLn - End3 + 1) + GEnd(5)
        B1(i) = Counter * (BEnd(3) - BEnd(5)) / (MaxLn - End3 + 1) + BEnd(5)
    Next i
    
    If Horz = 0 Then
        For i = 1 To MaxLn
            picBack.Line (0, i - 1)-(picBack.ScaleWidth, i), RGB(R1(i), G1(i), B1(i)), BF
        Next i
    Else
        For i = 1 To MaxLn
            picBack.Line (i - 1, 0)-(i, picBack.ScaleHeight), RGB(R1(i), G1(i), B1(i)), BF
        Next i
    End If
    
    Set picMain.Picture = picBack.Image
    DoEvents
End Function
Private Function DrawOli1()

Dy = 0
Dx = picMain.ScaleHeight
picMain.Cls
    For i = 1 To UBound(MyFreqs)
        If m_TypDraw = 1 Then
        Color = m_LineColor
        ElseIf m_TypDraw = 2 Then
            If MyFreqs(i) <= 25 Then
            Color = m_Fire1
            ElseIf MyFreqs(i) <= 50 Then
            Color = m_Fire2
            ElseIf MyFreqs(i) <= 75 Then
            Color = m_Fire3
            Else
            Color = m_Fire4
            End If
        ElseIf m_TypDraw = 3 Then
        Color = GetColor
        ElseIf m_TypDraw = 4 Then
            If MyFreqs(i) > 100 Then
            Color = MyColor(100)
            Else
            Color = MyColor(MyFreqs(i))
            End If
        ElseIf m_TypDraw = 5 Then
            If MyFreqs(i) > 100 Then
            Color = MyFire(100)
            Else
            Color = MyFire(MyFreqs(i))
            End If
        End If
        
        If i = 1 Then
            picMain.Line (0, picMain.ScaleHeight)-(AddRate, picMain.ScaleHeight - MyFreqs(1)), Color
            Dx = picMain.ScaleHeight - MyFreqs(1)
            Dy = AddRate
        Else
            picMain.Line (Dy, Dx)-(Dy + (AddRate * 2), picMain.ScaleHeight - MyFreqs(i)), Color
            Dx = picMain.ScaleHeight - MyFreqs(i)
            Dy = Dy + (AddRate * 2)
        End If
        
        
    Next i
        picMain.Line (Dy, Dx)-(Dy + AddRate, picMain.ScaleHeight), Color

End Function
Private Function DrawOli2()
Dy = 0
Dx = picMain.ScaleHeight
picMain.Cls
    For i = 1 To UBound(MyFreqs)
        If m_TypDraw = 1 Then
        Color = m_LineColor
        ElseIf m_TypDraw = 2 Then
            If MyFreqs(i) <= 25 Then
            Color = m_Fire1
            ElseIf MyFreqs(i) <= 50 Then
            Color = m_Fire2
            ElseIf MyFreqs(i) <= 75 Then
            Color = m_Fire3
            Else
            Color = m_Fire4
            End If
        ElseIf m_TypDraw = 3 Then
        Color = GetColor
        ElseIf m_TypDraw = 4 Then
            If MyFreqs(i) > 100 Then
            Color = MyColor(100)
            Else
            Color = MyColor(MyFreqs(i))
            End If
        ElseIf m_TypDraw = 5 Then
            If MyFreqs(i) > 100 Then
            Color = MyFire(100)
            Else
            Color = MyFire(MyFreqs(i))
            End If
        End If
        
        If i = 1 Then
            picMain.Line (0, picMain.ScaleHeight / 2)-(AddRate, (picMain.ScaleHeight / 2) - MyFreqs(1) / 2), Color
            Dx = picMain.ScaleHeight / 2 - (MyFreqs(1) / 2)
            Dy = AddRate
        Else
            picMain.Line (Dy, Dx)-(Dy + (AddRate * 2), picMain.ScaleHeight / 2 - (MyFreqs(i) / 2)), Color
            Dx = picMain.ScaleHeight / 2 - (MyFreqs(i) / 2)
            Dy = Dy + (AddRate * 2)
        End If
        
        
    Next i
    picMain.Line (Dy, Dx)-(Dy + AddRate, picMain.ScaleHeight / 2), Color


Dy = 0
Dx = 50

    For i = 1 To UBound(MyFreqs)
        If m_TypDraw = 1 Then
        Color = m_LineColor
        ElseIf m_TypDraw = 2 Then
            If MyFreqs(i) <= 25 Then
            Color = m_Fire1
            ElseIf MyFreqs(i) <= 50 Then
            Color = m_Fire2
            ElseIf MyFreqs(i) <= 75 Then
            Color = m_Fire3
            Else
            Color = m_Fire4
            End If
        ElseIf m_TypDraw = 3 Then
        Color = GetColor
        ElseIf m_TypDraw = 4 Then
            If MyFreqs(i) > 100 Then
            Color = MyColor(100)
            Else
            Color = MyColor(MyFreqs(i))
            End If
        ElseIf m_TypDraw = 5 Then
            If MyFreqs(i) > 100 Then
            Color = MyFire(100)
            Else
            Color = MyFire(MyFreqs(i))
            End If
        End If
        
        If i = 1 Then
            picMain.Line (0, picMain.ScaleHeight / 2)-(AddRate, picMain.ScaleHeight / 2 + (MyFreqs(1) / 2)), Color
            Dx = picMain.ScaleHeight / 2 + (MyFreqs(1) / 2)
            Dy = AddRate
        Else
            picMain.Line (Dy, Dx)-(Dy + (AddRate * 2), picMain.ScaleHeight / 2 + (MyFreqs(i) / 2)), Color
            Dx = picMain.ScaleHeight / 2 + (MyFreqs(i) / 2)
            Dy = Dy + (AddRate * 2)
        End If
        
        
    Next i
    picMain.Line (Dy, Dx)-(Dy + AddRate, picMain.ScaleHeight / 2), Color
    
End Function

Private Function DrawOli3()
UDD = 0
Cl = 0
Dy = 0
Dx = (picMain.ScaleHeight / 2)
picMain.Cls
    For i = 1 To UBound(MyFreqs)
        If m_TypDraw = 1 Then
        Color = m_LineColor
        ElseIf m_TypDraw = 2 Then
            If MyFreqs(i) <= 25 Then
            Color = m_Fire1
            ElseIf MyFreqs(i) <= 50 Then
            Color = m_Fire2
            ElseIf MyFreqs(i) <= 75 Then
            Color = m_Fire3
            Else
            Color = m_Fire4
            End If
        ElseIf m_TypDraw = 3 Then
        Color = GetColor
        ElseIf m_TypDraw = 4 Then
            If MyFreqs(i) > 100 Then
            Color = MyColor(100)
            Else
            Color = MyColor(MyFreqs(i))
            End If
        ElseIf m_TypDraw = 5 Then
            If MyFreqs(i) > 100 Then
            Color = MyFire(100)
            Else
            Color = MyFire(MyFreqs(i))
            End If
        End If
        
        
       If UDD = 0 Then
            UDD = 1
            
            If i <> 1 Then
                picMain.Line (Dy, Dx)-(Dy + (AddRate * 2), (picMain.ScaleHeight / 2) - (MyFreqs(i) / 2)), Color
                Dy = Dy + (AddRate * 2)
                Dx = (picMain.ScaleHeight / 2) - (MyFreqs(i) / 2)
            ElseIf i = UBound(MyFreqs) Then
                picMain.Line (Dy, Dx)-(Dy + AddRate, (picMain.ScaleHeight / 2)), Color
            Else
                picMain.Line (Dy, Dx)-(Dy + AddRate, (picMain.ScaleHeight / 2) - (MyFreqs(i) / 2)), Color
                Dy = Dy + (AddRate * 2)
                Dx = (picMain.ScaleHeight / 2) - (MyFreqs(i) / 2)
            End If
       Else
            If i = UBound(MyFreqs) Then
                picMain.Line (Dy, Dx)-(Dy + (AddRate * 2), (picMain.ScaleHeight / 2)), Color
            Else
                UDD = 0
                picMain.Line (Dy, Dx)-(Dy + (AddRate * 2), (picMain.ScaleHeight / 2) + (MyFreqs(i) / 2)), Color
                Dy = Dy + (AddRate * 2)
                Dx = (picMain.ScaleHeight / 2) + (MyFreqs(i) / 2)
            End If
        End If
    Next i
        

End Function
Private Function DrawBar2()
Cl = 0
Dx = picMain.ScaleHeight
Dy = AddRate
picMain.Cls

For i = 1 To UBound(MyFreqs)
        If m_TypDraw = 1 Then
        Color = m_LineColor
        ElseIf m_TypDraw = 2 Then
            If MyFreqs(i) <= 25 Then
            Color = m_Fire1
            ElseIf MyFreqs(i) <= 50 Then
            Color = m_Fire2
            ElseIf MyFreqs(i) <= 75 Then
            Color = m_Fire3
            Else
            Color = m_Fire4
            End If
        ElseIf m_TypDraw = 3 Then
        Color = GetColor
        ElseIf m_TypDraw = 4 Then
            If MyFreqs(i) > 100 Then
            Color = MyColor(100)
            Else
            Color = MyColor(MyFreqs(i))
            End If
        ElseIf m_TypDraw = 5 Then
            If MyFreqs(i) > 100 Then
            Color = MyFire(100)
            Else
            Color = MyFire(MyFreqs(i))
            End If
        End If
        
    
        picMain.Line (Dy, picMain.ScaleHeight)-(Dy, picMain.ScaleHeight - MyFreqs(i)), Color
        picMain.Line (Dy - 1, picMain.ScaleHeight)-(Dy - 1, picMain.ScaleHeight - (MyFreqs(i) - (MyFreqs(i) / 2))), Color
        picMain.Line (Dy + 1, picMain.ScaleHeight)-(Dy + 1, picMain.ScaleHeight - (MyFreqs(i) - (MyFreqs(i) / 2))), Color
        picMain.Line (Dy + 2, picMain.ScaleHeight)-(Dy + 2, picMain.ScaleHeight - (MyFreqs(i) - (MyFreqs(i) / 3))), Color
        
    If MyFreqs(i) > MyThick(i).CurY Then
            MyThick(i).CurY = MyFreqs(i)
            MyThick(i).FallOf = 0
            MyThick(i).StandTime = m_Thick
            MyThick2(i).CurY = MyFreqs(i)
            MyThick2(i).FallOf = 0
            MyThick2(i).StandTime = m_Thick + 1
            MyThick3(i).CurY = MyFreqs(i)
            MyThick3(i).FallOf = 0
            MyThick3(i).StandTime = m_Thick + 2
        End If
    Dy = Dy + (AddRate * 2)
Next i

End Function

Private Sub Timer2_Timer()
If m_Enabled = True Then
RaiseEvent FramesPerSecond(CurFps)
End If
CurFps = 0
End Sub

Private Sub Timer3_Timer()
If m_Enabled = False Then Exit Sub

If m_TypRefresh = 1 Then
    RefreshFreqs
    Else
    RefreshOne
    End If
    
    Select Case m_Typ
        Case 1
        DrawSpec
        Case 2
        DrawBar1
        Case 3
        DrawBar2
        Case 4
        DrawOli1
        Case 5
        DrawOli2
        Case 6
        DrawOli3
        Case 7
        DrawSonar
        Case 8
        DrawRoom
        Case 9
        DrawCurve
    End Select
    
If m_Typ = 1 Or m_Typ = 2 Or m_Typ = 3 Then
    If m_PeaksON = True Then
    
    For i = 1 To UBound(MyThick)
        If MyThick(i).CurY > 0 Then
            If MyThick(i).StandTime = 0 Then
                MyThick(i).CurY = MyThick(i).CurY - MyThick(i).FallOf
                If MyThick(i).CurY < 0 Then MyThick(i).CurY = 0
            MyThick(i).FallOf = MyThick(i).FallOf + 1
            Else
            MyThick(i).StandTime = MyThick(i).StandTime - 1
            End If
        End If
    Next i
    For i = 1 To UBound(MyThick2)
        If MyThick2(i).CurY > 0 Then
            If MyThick2(i).StandTime = 0 Then
                MyThick2(i).CurY = MyThick2(i).CurY - MyThick2(i).FallOf
                If MyThick2(i).CurY < 0 Then MyThick2(i).CurY = 0
            MyThick2(i).FallOf = MyThick2(i).FallOf + 1
            Else
            MyThick2(i).StandTime = MyThick2(i).StandTime - 1
            End If
        End If
    Next i
    For i = 1 To UBound(MyThick3)
        If MyThick3(i).CurY > 0 Then
            If MyThick3(i).StandTime = 0 Then
                MyThick3(i).CurY = MyThick3(i).CurY - MyThick3(i).FallOf
                If MyThick3(i).CurY < 0 Then MyThick3(i).CurY = 0
            MyThick3(i).FallOf = MyThick3(i).FallOf + 1
            Else
            MyThick3(i).StandTime = MyThick3(i).StandTime - 1
            End If
        End If
    Next i
    
    
    RefreshThicks
    End If
End If

CurFps = CurFps + 1
DoEvents
End Sub

Private Sub UserControl_Hide()
Timer1.Enabled = False
Timer3.Enabled = False
End Sub

Private Sub UserControl_Initialize()

    rc = mixerOpen(hmixer, DEVICEID, 0, 0, 0)
    If ((MMSYSERR_NOERROR <> rc)) Then
        MsgBox "Couldn't open the mixer."
        Exit Sub
    End If
    ' Get the output volume meter
    OK = GetControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, MIXERCONTROL_CONTROLTYPE_PEAKMETER, outputVolCtrl)
    If (OK = False) Then
       MsgBox "Couldn't get waveout meter"
    End If
    ' Initialize mixercontrol structure
    mxcd.cbStruct = Len(mxcd)
    volHmem = GlobalAlloc(&H0, Len(volume))  ' Allocate a buffer for the volume value
    mxcd.paDetails = GlobalLock(volHmem)
    mxcd.cbDetails = Len(volume)
    mxcd.cChannels = 1
   
   MakeFireColors
   RefreshBack
End Sub

Private Sub MakeFireColors()
    On Error Resume Next
    
    Dim R1(1 To 101) As Integer
    Dim G1(1 To 101) As Integer
    Dim B1(1 To 101) As Integer
    
    Dim REnd(1 To 5) As Integer
    Dim GEnd(1 To 5) As Integer
    Dim BEnd(1 To 5) As Integer
    
    Dim i As Integer, End1 As Integer, End2 As Integer, End3 As Integer
    Dim ColorString As String
    Dim Counter As Integer
    
    End1 = 40
    End2 = 50
    End3 = 60
    
    'get RGB values
    For i = 1 To 3
        Select Case i
            Case 1
                ColorString = Hex(vbGreen)
            Case 2
                ColorString = Hex(vbYellow)
            Case 3
                ColorString = Hex(vbRed)
        End Select
        If Len(ColorString) = 2 Then
            BEnd(i) = 0
            GEnd(i) = 0
            REnd(i) = CLng("&H" & ColorString)
        ElseIf Len(ColorString) = 4 Then
            BEnd(i) = 0
            GEnd(i) = CLng("&H" & Left$(ColorString, 2))
            REnd(i) = CLng("&H" & Right$(ColorString, 2))
        ElseIf Len(ColorString$) = 6 Then
            BEnd(i) = CLng("&H" & Left$(ColorString, 2))
            GEnd(i) = CLng("&H" & Mid$(ColorString, 3, 2))
            REnd(i) = CLng("&H" & Right$(ColorString, 2))
        End If
    Next i
    
    'Auto calculate mixed "in-between" colors
    For i = 4 To 5 Step 1
        If REnd(i - 2) > REnd(i - 3) Then
            REnd(i) = REnd(i - 2)
        Else
            REnd(i) = REnd(i - 3)
        End If

        If GEnd(i - 2) > GEnd(i - 3) Then
            GEnd(i) = GEnd(i - 2)
        Else
            GEnd(i) = GEnd(i - 3)
        End If

        If BEnd(i - 2) > BEnd(i - 3) Then
            BEnd(i) = BEnd(i - 2)
        Else
            BEnd(i) = BEnd(i - 3)
        End If
    Next i
    
    'set color levels
    For i = 1 To End1
        R1(i) = (i - 1) * (REnd(4) - REnd(1)) / (End1 + 1) + REnd(1)
        G1(i) = (i - 1) * (GEnd(4) - GEnd(1)) / (End1 + 1) + GEnd(1)
        B1(i) = (i - 1) * (BEnd(4) - BEnd(1)) / (End1 + 1) + BEnd(1)
    Next
    Counter = 0

    For i = End1 + 1 To End2
        Counter = Counter + 1
        R1(i) = Counter * (REnd(2) - REnd(4)) / (End2 - End1 + 1) + REnd(4)
        G1(i) = Counter * (GEnd(2) - GEnd(4)) / (End2 - End1 + 1) + GEnd(4)
        B1(i) = Counter * (BEnd(2) - BEnd(4)) / (End2 - End1 + 1) + BEnd(4)
    Next
    Counter = 0

    For i = End2 + 1 To End3
        Counter = Counter + 1
        R1(i) = Counter * (REnd(5) - REnd(2)) / (End3 - End2 + 1) + REnd(2)
        G1(i) = Counter * (GEnd(5) - GEnd(2)) / (End3 - End2 + 1) + GEnd(2)
        B1(i) = Counter * (BEnd(5) - BEnd(2)) / (End3 - End2 + 1) + BEnd(2)
    Next
    Counter = 0

    For i = End3 + 1 To 101
        Counter = Counter + 1
        R1(i) = Counter * (REnd(3) - REnd(5)) / (101 - End3 + 1) + REnd(5)
        G1(i) = Counter * (GEnd(3) - GEnd(5)) / (101 - End3 + 1) + GEnd(5)
        B1(i) = Counter * (BEnd(3) - BEnd(5)) / (101 - End3 + 1) + BEnd(5)
    Next i
    
    For i = 1 To 101
       MyFire(i - 1) = RGB(R1(i), G1(i), B1(i))
    Next i
    
    DoEvents
End Sub

Private Sub UserControl_InitProperties()

m_Enabled = m_def_Enabled
m_RefreshSpeed = m_def_RefreshSpeed
m_Typ = m_def_Typ
m_TypDraw = m_def_TypDraw
m_Thick = m_def_Thick
m_TypRefresh = m_def_TypRefresh
m_PeaksON = m_def_PeaksON
m_Style = m_def_Style

m_PeakColor = m_def_PeakColor
m_LineColor = m_def_LineColor
m_BackColor = m_def_BackColor
m_Fire1 = m_def_Fire1
m_Fire2 = m_def_Fire2
m_Fire3 = m_def_Fire3
m_Fire4 = m_def_Fire4
m_GradientC = m_def_GradientC
m_Background = m_def_Background
m_BackVH = m_def_BackVH
Set m_BackPic = LoadPicture("")

MakeGradientColors m_GradientC


   ' Open the mixer specified by DEVICEID

   
   AddCl = 255 / UBound(MyFreqs)


Timer1.Enabled = True
Timer3.Enabled = True
End Sub

Private Function RefreshBack()
Select Case m_Background
    Case 1
        picMain.Picture = LoadPicture("")
        picMain.BackColor = m_BackColor
    Case 2
        If m_BackVH = 1 Then
            MakeGradientBack m_BackColor, 0
        Else
            MakeGradientBack m_BackColor, 1
        End If
    Case 3
        Set picMain.Picture = m_BackPic
End Select
End Function

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
m_RefreshSpeed = PropBag.ReadProperty("RefreshSpeed", m_def_RefreshSpeed)
Timer3.Interval = m_RefreshSpeed
m_Typ = PropBag.ReadProperty("Visu_Typ", m_def_Typ)
m_Style = PropBag.ReadProperty("Style", m_def_Style)
m_TypDraw = PropBag.ReadProperty("Draw_Typ", m_def_TypDraw)
m_Thick = PropBag.ReadProperty("Thick_Fall", m_def_Thick)
m_PeaksON = PropBag.ReadProperty("PeaksON", m_def_PeaksON)
m_TypRefresh = PropBag.ReadProperty("Type_Refresh", m_def_TypRefresh)

m_PeakColor = PropBag.ReadProperty("PeakColor", m_def_PeakColor)
m_LineColor = PropBag.ReadProperty("LineColor", m_def_LineColor)
m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
m_Fire1 = PropBag.ReadProperty("Fire1", m_def_Fire1)
m_Fire2 = PropBag.ReadProperty("Fire2", m_def_Fire2)
m_Fire3 = PropBag.ReadProperty("Fire3", m_def_Fire3)
m_Fire4 = PropBag.ReadProperty("Fire4", m_def_Fire4)
m_BackVH = PropBag.ReadProperty("BackVH", m_def_BackVH)
m_GradientC = PropBag.ReadProperty("GradientColor", m_def_GradientC)
m_Background = PropBag.ReadProperty("Background", m_def_Background)
Set m_BackPic = PropBag.ReadProperty("Picture", Nothing)

MakeGradientColors m_GradientC
RefreshBack
End Sub

Private Sub UserControl_Resize()
If UserControl.Width < 1000 Then UserControl.Width = 1000
If UserControl.Height < 250 Then UserControl.Height = 250
picMain.Left = 0
picMain.Top = 0
picMain.Width = UserControl.ScaleWidth
picMain.Height = UserControl.ScaleHeight


picMain.ScaleHeight = 100
picMain.ScaleWidth = 200

picBack.Width = picMain.Width
picBack.Height = picMain.Height
picBack.ScaleHeight = picMain.ScaleHeight
picBack.ScaleWidth = picMain.ScaleWidth

AddRate = picMain.ScaleWidth / UBound(MyFreqs)
AddRate = AddRate / 2

DrawSpec
RefreshBack
End Sub


Private Function MakeGradientColors(ByVal MainColor As Long)
On Error Resume Next
    Dim R1(1 To NumLevels) As Integer
    Dim G1(1 To NumLevels) As Integer
    Dim B1(1 To NumLevels) As Integer
    
    Dim REnd(1 To 5) As Integer
    Dim GEnd(1 To 5) As Integer
    Dim BEnd(1 To 5) As Integer
    
    Dim i As Integer, End1 As Integer, End2 As Integer, End3 As Integer
    Dim ColorString As String
    Dim Counter As Integer
    
    End1 = 1
    End2 = 101
    End3 = 101
    
    'get RGB values
    For i = 1 To 3
        If i = 1 Then
        ColorString = Hex(MainColor)
        Else
        ColorString = Hex(vbBlack)
        End If
        If Len(ColorString) = 2 Then
            BEnd(i) = 0
            GEnd(i) = 0
            REnd(i) = CLng("&H" & ColorString)
        ElseIf Len(ColorString) = 4 Then
            BEnd(i) = 0
            GEnd(i) = CLng("&H" & Left$(ColorString, 2))
            REnd(i) = CLng("&H" & Right$(ColorString, 2))
        ElseIf Len(ColorString$) = 6 Then
            BEnd(i) = CLng("&H" & Left$(ColorString, 2))
            GEnd(i) = CLng("&H" & Mid$(ColorString, 3, 2))
            REnd(i) = CLng("&H" & Right$(ColorString, 2))
        End If
    Next i
    
    'Auto calculate mixed "in-between" colors
    For i = 4 To 5 Step 1
        If REnd(i - 2) > REnd(i - 3) Then
            REnd(i) = REnd(i - 2)
        Else
            REnd(i) = REnd(i - 3)
        End If

        If GEnd(i - 2) > GEnd(i - 3) Then
            GEnd(i) = GEnd(i - 2)
        Else
            GEnd(i) = GEnd(i - 3)
        End If

        If BEnd(i - 2) > BEnd(i - 3) Then
            BEnd(i) = BEnd(i - 2)
        Else
            BEnd(i) = BEnd(i - 3)
        End If
    Next i
    
    'set color levels
    For i = 1 To End1
        R1(i) = (i - 1) * (REnd(4) - REnd(1)) / (End1 + 1) + REnd(1)
        G1(i) = (i - 1) * (GEnd(4) - GEnd(1)) / (End1 + 1) + GEnd(1)
        B1(i) = (i - 1) * (BEnd(4) - BEnd(1)) / (End1 + 1) + BEnd(1)
    Next
    Counter = 0

    For i = End1 + 1 To End2
        Counter = Counter + 1
        R1(i) = Counter * (REnd(2) - REnd(4)) / (End2 - End1 + 1) + REnd(4)
        G1(i) = Counter * (GEnd(2) - GEnd(4)) / (End2 - End1 + 1) + GEnd(4)
        B1(i) = Counter * (BEnd(2) - BEnd(4)) / (End2 - End1 + 1) + BEnd(4)
    Next
    Counter = 0

    For i = End2 + 1 To End3
        Counter = Counter + 1
        R1(i) = Counter * (REnd(5) - REnd(2)) / (End3 - End2 + 1) + REnd(2)
        G1(i) = Counter * (GEnd(5) - GEnd(2)) / (End3 - End2 + 1) + GEnd(2)
        B1(i) = Counter * (BEnd(5) - BEnd(2)) / (End3 - End2 + 1) + BEnd(2)
    Next
    Counter = 0

    For i = End3 + 1 To NumLevels
        Counter = Counter + 1
        R1(i) = Counter * (REnd(3) - REnd(5)) / (NumLevels - End3 + 1) + REnd(5)
        G1(i) = Counter * (GEnd(3) - GEnd(5)) / (NumLevels - End3 + 1) + GEnd(5)
        B1(i) = Counter * (BEnd(3) - BEnd(5)) / (NumLevels - End3 + 1) + BEnd(5)
    Next i
        
    For i = 1 To 101
        MyColor(i - 1) = RGB(R1(i), G1(i), B1(i))
    Next i
    
DoEvents
End Function
Private Sub UserControl_Show()
Timer1.Enabled = True
Timer3.Enabled = True
End Sub

Private Sub UserControl_Terminate()
   If (fRecording = True) Then
       StopInput
   End If
   GlobalFree volHmem
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
Call PropBag.WriteProperty("RefreshSpeed", m_RefreshSpeed, m_def_RefreshSpeed)
Call PropBag.WriteProperty("Visu_Typ", m_Typ, m_def_Typ)
Call PropBag.WriteProperty("Style", m_Style, m_def_Style)
Call PropBag.WriteProperty("Draw_Typ", m_TypDraw, m_def_TypDraw)
Call PropBag.WriteProperty("Thick_Fall", m_Thick, m_def_Thick)
Call PropBag.WriteProperty("Type_Refresh", m_TypRefresh, m_def_TypRefresh)
Call PropBag.WriteProperty("PeaksON", m_PeaksON, m_def_PeaksON)
Call PropBag.WriteProperty("BackVH", m_BackVH, m_def_BackVH)
Call PropBag.WriteProperty("PeakColor", m_PeakColor, m_def_PeakColor)
Call PropBag.WriteProperty("LineColor", m_LineColor, m_def_LineColor)
Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
Call PropBag.WriteProperty("Background", m_Background, m_def_Background)
Call PropBag.WriteProperty("Fire1", m_Fire1, m_def_Fire1)
Call PropBag.WriteProperty("Fire2", m_Fire2, m_def_Fire2)
Call PropBag.WriteProperty("Fire3", m_Fire3, m_def_Fire3)
Call PropBag.WriteProperty("Fire4", m_Fire4, m_def_Fire4)
Call PropBag.WriteProperty("GradientColor", m_GradientC, m_def_GradientC)
Call PropBag.WriteProperty("Picture", m_BackPic, Nothing)
End Sub
Private Function Freq1()
FreqNum = a1Freq5Hz
For VU.InOutLev = CDbl(VU.VolLev) To FreqNum
Next VU.InOutLev

MyFreqs(1) = (VU.InOutLev / 5)
MyFreqs(1) = CDbl(Left(CStr(MyFreqs(1) - 1), 3)) * 10
End Function

Private Function Freq2()
FreqNum = a2Freq10Hz
For VU.InOutLev = CDbl(VU.VolLev) To FreqNum
Next VU.InOutLev
MyFreqs(2) = (VU.InOutLev / 10)
MyFreqs(2) = CDbl(Left(CStr(MyFreqs(2) - 1), 3)) * 10
End Function

Private Function Freq3()
FreqNum = a3Freq15Hz
For VU.InOutLev = CDbl(VU.VolLev) To FreqNum
Next VU.InOutLev
MyFreqs(3) = (VU.InOutLev / 15)
MyFreqs(3) = CDbl(Left(CStr(MyFreqs(3) - 1), 3)) * 10
End Function


Private Function Freq4()
FreqNum = a4Freq22Hz
For VU.InOutLev = CDbl(VU.VolLev) To FreqNum
Next VU.InOutLev
MyFreqs(4) = (VU.InOutLev / 22)
MyFreqs(4) = CDbl(Left(CStr(MyFreqs(4) - 1), 3)) * 10
End Function
Private Function Freq5()
FreqNum = a5Freq31Hz
For VU.InOutLev = CDbl(VU.VolLev) To FreqNum
Next VU.InOutLev
MyFreqs(5) = (VU.InOutLev / 31)
MyFreqs(5) = CDbl(Left(CStr(MyFreqs(5) - 1), 3)) * 10
End Function

Private Function Freq6()
FreqNum = a6Freq40Hz
For VU.InOutLev = CDbl(VU.VolLev) To FreqNum
Next VU.InOutLev
MyFreqs(6) = (VU.InOutLev / 40)
MyFreqs(6) = CDbl(Left(CStr(MyFreqs(6) - 1), 3)) * 10
End Function

Private Function Freq7()
FreqNum = a7Freq50Hz
For VU.InOutLev = CDbl(VU.VolLev) To FreqNum
Next VU.InOutLev
MyFreqs(7) = (VU.InOutLev / 50)
MyFreqs(7) = CDbl(Left(CStr(MyFreqs(7) - 1), 3)) * 10
End Function

Private Function Freq8()
FreqNum = a8Freq60Hz
For VU.InOutLev = CDbl(VU.VolLev) To FreqNum
Next VU.InOutLev
MyFreqs(8) = (VU.InOutLev / 60)
MyFreqs(8) = CDbl(Left(CStr(MyFreqs(8) - 1), 3)) * 10
End Function
Private Function Freq9()
FreqNum = a9Freq62Hz
For VU.InOutLev = CDbl(VU.VolLev) To FreqNum
Next VU.InOutLev
MyFreqs(9) = (VU.InOutLev / 62)
MyFreqs(9) = CDbl(Left(CStr(MyFreqs(9) - 1), 3)) * 10
End Function

Private Function Freq10()
FreqNum = b1Freq70Hz
For VU.InOutLev = CDbl(VU.VolLev) To FreqNum
Next VU.InOutLev
MyFreqs(10) = (VU.InOutLev / 70)
MyFreqs(10) = CDbl(Left(CStr(MyFreqs(10) - 1), 3)) * 10
End Function
Private Function Freq11()
FreqNum = b2Freq80Hz
For VU.InOutLev = CDbl(VU.VolLev) To FreqNum
Next VU.InOutLev
MyFreqs(11) = (VU.InOutLev / 80)
MyFreqs(11) = CDbl(Left(CStr(MyFreqs(11) - 1), 3)) * 10
End Function
Private Function Freq12()
FreqNum = b3Freq85Hz
For VU.InOutLev = CDbl(VU.VolLev) To FreqNum
Next VU.InOutLev
MyFreqs(12) = (VU.InOutLev / 85)
MyFreqs(12) = CDbl(Left(CStr(MyFreqs(12) - 1), 3)) * 10
End Function
Private Function Freq13()
FreqNum = b4Freq90Hz
For VU.InOutLev = CDbl(VU.VolLev) To FreqNum
Next VU.InOutLev
MyFreqs(13) = (VU.InOutLev / 90)
MyFreqs(13) = CDbl(Left(CStr(MyFreqs(13) - 1), 3)) * 10
End Function
Private Function Freq14()
FreqNum = b5Freq95Hz
For VU.InOutLev = CDbl(VU.VolLev) To FreqNum
Next VU.InOutLev
MyFreqs(14) = (VU.InOutLev / 95)
MyFreqs(14) = CDbl(Left(CStr(MyFreqs(14) - 1), 3)) * 10
End Function
Private Function Freq15()
FreqNum = b6Freq98Hz
For VU.InOutLev = CDbl(VU.VolLev) To FreqNum
Next VU.InOutLev
MyFreqs(15) = (VU.InOutLev / 98)
MyFreqs(15) = CDbl(Left(CStr(MyFreqs(15) - 1), 3)) * 10
End Function
Private Function Freq16()
FreqNum = b7Freq100Hz
For VU.InOutLev = CDbl(VU.VolLev) To FreqNum
Next VU.InOutLev
MyFreqs(16) = (VU.InOutLev / 100)
MyFreqs(16) = CDbl(Left(CStr(MyFreqs(16) - 1), 3)) * 10
End Function
Private Function Freq17()
FreqNum = b8Freq105Hz
For VU.InOutLev = CDbl(VU.VolLev) To FreqNum
Next VU.InOutLev
MyFreqs(17) = (VU.InOutLev / 105)
MyFreqs(17) = CDbl(Left(CStr(MyFreqs(17) - 1), 3)) * 10
End Function
Private Function Freq18()
FreqNum = b9Freq115Hz
For VU.InOutLev = CDbl(VU.VolLev) To FreqNum
Next VU.InOutLev
MyFreqs(18) = (VU.InOutLev / 115)
MyFreqs(18) = CDbl(Left(CStr(MyFreqs(18) - 1), 3)) * 10
End Function
Private Function Freq19()
FreqNum = c1Freq125Hz
For VU.InOutLev = CDbl(VU.VolLev) To FreqNum
Next VU.InOutLev
MyFreqs(19) = (VU.InOutLev / 125)
MyFreqs(19) = CDbl(Left(CStr(MyFreqs(19) - 1), 3)) * 10
End Function
Private Function Freq20()
FreqNum = c2Freq150Hz
For VU.InOutLev = CDbl(VU.VolLev) To FreqNum
Next VU.InOutLev
MyFreqs(20) = (VU.InOutLev / 150)
MyFreqs(20) = CDbl(Left(CStr(MyFreqs(20) - 1), 3)) * 10
End Function
Private Function Freq21()
FreqNum = c3Freq170Hz
For VU.InOutLev = CDbl(VU.VolLev) To FreqNum
Next VU.InOutLev
MyFreqs(21) = (VU.InOutLev / 170)
MyFreqs(21) = CDbl(Left(CStr(MyFreqs(21) - 1), 3)) * 10
End Function
Private Function Freq22()
FreqNum = c4Freq200Hz
For VU.InOutLev = CDbl(VU.VolLev) To FreqNum
Next VU.InOutLev
MyFreqs(22) = (VU.InOutLev / 200)
MyFreqs(22) = CDbl(Left(CStr(MyFreqs(22) - 1), 3)) * 10
End Function
Private Function Freq23()
FreqNum = c5Freq225Hz
For VU.InOutLev = CDbl(VU.VolLev) To FreqNum
Next VU.InOutLev
MyFreqs(23) = (VU.InOutLev / 225)
MyFreqs(23) = CDbl(Left(CStr(MyFreqs(23) - 1), 3)) * 10
End Function
Private Function Freq24()
FreqNum = c6Freq250Hz
For VU.InOutLev = CDbl(VU.VolLev) To FreqNum
Next VU.InOutLev
MyFreqs(24) = (VU.InOutLev / 250)
MyFreqs(24) = CDbl(Left(CStr(MyFreqs(24) - 1), 3)) * 10
End Function
Private Function Freq25()
FreqNum = c7Freq310Hz
For VU.InOutLev = CDbl(VU.VolLev) To FreqNum
Next VU.InOutLev
MyFreqs(25) = (VU.InOutLev / 310)
MyFreqs(25) = CDbl(Left(CStr(MyFreqs(25) - 1), 3)) * 10
End Function
Private Function Freq26()
FreqNum = c8Freq350Hz
For VU.InOutLev = CDbl(VU.VolLev) To FreqNum
Next VU.InOutLev
MyFreqs(26) = (VU.InOutLev / 350)
MyFreqs(26) = CDbl(Left(CStr(MyFreqs(26) - 1), 3)) * 10
End Function
Private Function Freq27()
FreqNum = c9Freq400Hz
For VU.InOutLev = CDbl(VU.VolLev) To FreqNum
Next VU.InOutLev
MyFreqs(27) = (VU.InOutLev / 400)
MyFreqs(27) = CDbl(Left(CStr(MyFreqs(27) - 1), 3)) * 10
End Function
Private Function Freq28()
FreqNum = d1Freq450Hz
For VU.InOutLev = CDbl(VU.VolLev) To FreqNum
Next VU.InOutLev
MyFreqs(28) = (VU.InOutLev / 450)
MyFreqs(28) = CDbl(Left(CStr(MyFreqs(28) - 1), 3)) * 10
End Function
Private Function Freq29()
FreqNum = d2Freq500Hz
For VU.InOutLev = CDbl(VU.VolLev) To FreqNum
Next VU.InOutLev
MyFreqs(29) = (VU.InOutLev / 500)
MyFreqs(29) = CDbl(Left(CStr(MyFreqs(29) - 1), 3)) * 10
End Function
Private Function Freq30()
FreqNum = d3Freq600Hz
For VU.InOutLev = CDbl(VU.VolLev) To FreqNum
Next VU.InOutLev
MyFreqs(30) = (VU.InOutLev / 600)
MyFreqs(30) = CDbl(Left(CStr(MyFreqs(30) - 1), 3)) * 10
End Function
Private Function Freq31()
FreqNum = d4Freq1kHz
For VU.InOutLev = CDbl(VU.VolLev * 0.01) To FreqNum
Next VU.InOutLev
MyFreqs(31) = (VU.InOutLev / 1)
MyFreqs(31) = CDbl(Left(CStr(MyFreqs(31) - 1), 3)) * 10
End Function
Private Function Freq32()
FreqNum = d5Freq2kHz
For VU.InOutLev = CDbl(VU.VolLev * 0.02) To FreqNum
Next VU.InOutLev
MyFreqs(32) = (VU.InOutLev / 2)
MyFreqs(32) = CDbl(Left(CStr(MyFreqs(32) - 1), 3)) * 10
End Function
Private Function Freq33()
FreqNum = d6Freq3kHz
For VU.InOutLev = CDbl(VU.VolLev * 0.03) To FreqNum
Next VU.InOutLev
MyFreqs(33) = (VU.InOutLev / 3)
MyFreqs(33) = CDbl(Left(CStr(MyFreqs(33) - 1), 3)) * 10
End Function
Private Function Freq34()
FreqNum = d7Freq4kHz
For VU.InOutLev = CDbl(VU.VolLev * 0.04) To FreqNum
Next VU.InOutLev
MyFreqs(34) = (VU.InOutLev / 4)
MyFreqs(34) = CDbl(Left(CStr(MyFreqs(34) - 1), 3)) * 10
End Function
Private Function Freq35()
FreqNum = d8Freq5kHz
For VU.InOutLev = CDbl(VU.VolLev * 0.05) To FreqNum
Next VU.InOutLev
MyFreqs(35) = (VU.InOutLev / 5)
MyFreqs(35) = CDbl(Left(CStr(MyFreqs(35) - 1), 3)) * 10
End Function
Private Function Freq36()
FreqNum = d9Freq6kHz
For VU.InOutLev = CDbl(VU.VolLev * 0.06) To FreqNum
Next VU.InOutLev
MyFreqs(36) = (VU.InOutLev / 6)
MyFreqs(36) = CDbl(Left(CStr(MyFreqs(36) - 1), 3)) * 10
End Function
Private Function Freq37()
FreqNum = e1Freq7kHz
For VU.InOutLev = CDbl(VU.VolLev * 0.07) To FreqNum
Next VU.InOutLev
MyFreqs(37) = (VU.InOutLev / 7)
MyFreqs(37) = CDbl(Left(CStr(MyFreqs(37) - 1), 3)) * 10
End Function
Private Function Freq38()
FreqNum = e2Freq8kHz
For VU.InOutLev = CDbl(VU.VolLev * 0.08) To FreqNum
Next VU.InOutLev
MyFreqs(38) = (VU.InOutLev / 8)
MyFreqs(38) = CDbl(Left(CStr(MyFreqs(38) - 1), 3)) * 10
End Function
Private Function Freq39()
FreqNum = e3Freq9kHz
For VU.InOutLev = CDbl(VU.VolLev * 0.09) To FreqNum
Next VU.InOutLev
MyFreqs(39) = (VU.InOutLev / 9)
MyFreqs(39) = CDbl(Left(CStr(MyFreqs(39) - 1), 3)) * 10
End Function
Private Function Freq40()
FreqNum = e4Freq10kHz
For VU.InOutLev = CDbl(VU.VolLev * 0.1) To FreqNum
Next VU.InOutLev
MyFreqs(40) = (VU.InOutLev / 10)
MyFreqs(40) = CDbl(Left(CStr(MyFreqs(40) - 1), 3)) * 10
End Function
Private Function Freq41()
FreqNum = e5Freq11kHz
For VU.InOutLev = CDbl(VU.VolLev * 0.11) To FreqNum
Next VU.InOutLev
MyFreqs(41) = (VU.InOutLev / 11)
MyFreqs(41) = CDbl(Left(CStr(MyFreqs(41) - 1), 3)) * 10
End Function
Private Function Freq42()
FreqNum = e6Freq12kHz
For VU.InOutLev = CDbl(VU.VolLev * 0.12) To FreqNum
Next VU.InOutLev
MyFreqs(42) = (VU.InOutLev / 12)
MyFreqs(42) = CDbl(Left(CStr(MyFreqs(42) - 1), 3)) * 10
End Function
Private Function Freq43()
FreqNum = e7Freq13kHz
For VU.InOutLev = CDbl(VU.VolLev * 0.13) To FreqNum
Next VU.InOutLev
MyFreqs(43) = (VU.InOutLev / 13)
MyFreqs(43) = CDbl(Left(CStr(MyFreqs(43) - 1), 3)) * 10
End Function
Private Function Freq44()
FreqNum = e8Freq14kHz
For VU.InOutLev = CDbl(VU.VolLev * 0.14) To FreqNum
Next VU.InOutLev
MyFreqs(44) = (VU.InOutLev / 14)
MyFreqs(44) = CDbl(Left(CStr(MyFreqs(44) - 1), 3)) * 10
End Function
Private Function Freq45()
FreqNum = e9Freq15kHz
For VU.InOutLev = CDbl(VU.VolLev * 0.15) To FreqNum
Next VU.InOutLev
MyFreqs(45) = (VU.InOutLev / 15)
MyFreqs(45) = CDbl(Left(CStr(MyFreqs(45) - 1), 3)) * 10
End Function
Private Function Freq46()
FreqNum = f1Freq16kHz
For VU.InOutLev = CDbl(VU.VolLev * 0.16) To FreqNum
Next VU.InOutLev
MyFreqs(46) = (VU.InOutLev / 16)
MyFreqs(46) = CDbl(Left(CStr(MyFreqs(46) - 1), 3)) * 10
End Function
Private Function Freq47()
FreqNum = f2Freq17kHz
For VU.InOutLev = CDbl(VU.VolLev * 0.17) To FreqNum
Next VU.InOutLev
MyFreqs(47) = (VU.InOutLev / 17)
MyFreqs(47) = CDbl(Left(CStr(MyFreqs(47) - 1), 3)) * 10
End Function
Private Function Freq48()
FreqNum = f3Freq18kHz
For VU.InOutLev = CDbl(VU.VolLev * 0.18) To FreqNum
Next VU.InOutLev
MyFreqs(48) = (VU.InOutLev / 18)
MyFreqs(48) = CDbl(Left(CStr(MyFreqs(48) - 1), 3)) * 10
End Function
Private Function Freq49()
FreqNum = f4Freq19kHz
For VU.InOutLev = CDbl(VU.VolLev * 0.19) To FreqNum
Next VU.InOutLev
MyFreqs(49) = (VU.InOutLev / 19)
MyFreqs(49) = CDbl(Left(CStr(MyFreqs(49) - 1), 3)) * 10
End Function
Private Function Freq50()
FreqNum = f5Freq20kHz
For VU.InOutLev = CDbl(VU.VolLev * 0.2) To FreqNum
Next VU.InOutLev
MyFreqs(50) = (VU.InOutLev / 20)
MyFreqs(50) = CDbl(Left(CStr(MyFreqs(50) - 1), 3)) * 10
End Function
