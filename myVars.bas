Attribute VB_Name = "myVars"
Option Explicit

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Const SRCCOPY = &HCC0020
Public Const SRCAND = &H8800C6
Public Const SRCINVERT = &H660046
Public Const SRCPAINT = &HEE0086
Public Const SRCERASE = &H4400328
Public Const WHITENESS = &HFF0062
Public Const BLACKNESS = &H42

Type TankType
  X As Integer
  Y As Integer
End Type
  
Type ShotType
  Active As Boolean
  X As Integer
  Y As Integer
End Type
  
Global Const fTop = 8
Global Const fLeft = 8
Global Const fHeight = 293
Global Const fWidth = 473

Global Game As String
Global IntroCount As Integer
Global IntroAction As Integer
Global IntroY As Integer
Global DownKeys(255) As Boolean

Global TopRandNum As Integer
Global OptCmds As Boolean
Global OptBG As Boolean
Global Tank1 As TankType
Global Tank2 As TankType
Global Shot1() As ShotType
Global Shot2() As ShotType
