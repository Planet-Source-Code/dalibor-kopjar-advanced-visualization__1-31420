Attribute VB_Name = "modFunctions"
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

Public Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Public Type s_MyThick
    CurY As Long
    FallOf As Long
    StandTime As Long
End Type
