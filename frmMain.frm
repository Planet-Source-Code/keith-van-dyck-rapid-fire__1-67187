VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Rapid Fire"
   ClientHeight    =   6720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7365
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   448
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   491
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraOver 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   4395
      Left            =   0
      TabIndex        =   29
      Top             =   120
      Visible         =   0   'False
      Width           =   7095
      Begin VB.Label lblOverExit 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Exit Game"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   4440
         TabIndex        =   34
         Top             =   3280
         Width           =   2400
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Exit Game"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4440
         TabIndex        =   33
         Top             =   0
         Width           =   2415
      End
      Begin VB.Label lblOverNew 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "New Game"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   4440
         TabIndex        =   32
         Top             =   2600
         Width           =   2400
      End
      Begin VB.Shape shpOverNew 
         BackColor       =   &H8000000C&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000A&
         BorderWidth     =   2
         Height          =   615
         Left            =   4440
         Shape           =   4  'Rounded Rectangle
         Top             =   2520
         Width           =   2415
      End
      Begin VB.Label lblDisWinner 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Player 1"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   2325
         TabIndex        =   31
         Top             =   600
         Width           =   2070
      End
      Begin VB.Label lblDisCongrats 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Game Over!  The Winner Was..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   285
         Left            =   120
         TabIndex        =   30
         Top             =   120
         Width           =   3480
      End
      Begin VB.Shape shpOverExit 
         BackColor       =   &H8000000C&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000A&
         BorderWidth     =   2
         Height          =   615
         Left            =   4440
         Shape           =   4  'Rounded Rectangle
         Top             =   3240
         Width           =   2415
      End
   End
   Begin VB.Frame fraLoading 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   4560
      TabIndex        =   27
      Top             =   6120
      Visible         =   0   'False
      Width           =   615
      Begin VB.Timer tmrLoading 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   600
         Top             =   3480
      End
      Begin VB.Label lblDisLoading 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loading..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   555
         Left            =   2355
         TabIndex        =   28
         Top             =   1560
         Width           =   2145
      End
      Begin VB.Image imgSideTankLoading 
         Height          =   3930
         Left            =   600
         Picture         =   "frmMain.frx":030A
         Top             =   480
         Width           =   6000
      End
   End
   Begin VB.Frame fraPlay 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   3840
      TabIndex        =   23
      Top             =   6120
      Visible         =   0   'False
      Width           =   615
      Begin VB.PictureBox picPlay 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000007&
         Height          =   3615
         Left            =   120
         MousePointer    =   2  'Cross
         ScaleHeight     =   237
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   453
         TabIndex        =   24
         Top             =   360
         Width           =   6855
      End
      Begin VB.Timer tmrPlay 
         Enabled         =   0   'False
         Interval        =   80
         Left            =   0
         Top             =   4200
      End
      Begin VB.Label lblDisP2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Player 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   3240
         TabIndex        =   26
         Top             =   4080
         Width           =   1650
      End
      Begin VB.Label lblDisP1 
         BackStyle       =   0  'Transparent
         Caption         =   "Player 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   2160
         TabIndex        =   25
         Top             =   120
         Width           =   2010
      End
      Begin VB.Shape shpP1TOP 
         BackColor       =   &H8000000D&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   60
         Top             =   60
         Width           =   1995
      End
      Begin VB.Shape shpP2TOP 
         BackColor       =   &H8000000D&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   5040
         Top             =   4080
         Width           =   1995
      End
      Begin VB.Shape shpP2BG 
         BackColor       =   &H8000000A&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   5040
         Top             =   4080
         Width           =   1995
      End
      Begin VB.Shape shpP1BG 
         BackColor       =   &H8000000A&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   60
         Top             =   60
         Width           =   1995
      End
   End
   Begin VB.Frame fraP2 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   5280
      TabIndex        =   14
      Top             =   6120
      Visible         =   0   'False
      Width           =   615
      Begin VB.TextBox txtP2 
         BackColor       =   &H80000008&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   375
         Left            =   240
         MaxLength       =   10
         TabIndex        =   15
         Text            =   "Player 2"
         Top             =   3480
         Width           =   3735
      End
      Begin VB.Label lblDisP2Name 
         BackStyle       =   0  'Transparent
         Caption         =   "Players Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label lblDisP2Comp 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Computer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   3600
         TabIndex        =   21
         Top             =   1020
         Width           =   2415
      End
      Begin VB.Label lblDisCRP2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright© 1999-2002 Bulldog Creations, All rights reserved."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   210
         Left            =   2520
         TabIndex        =   20
         Top             =   4200
         Width           =   4410
      End
      Begin VB.Label lblP2Close 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   6960
         TabIndex        =   19
         Top             =   0
         Width           =   120
      End
      Begin VB.Label lblDisP2Main 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Player 2:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   480
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   1725
      End
      Begin VB.Shape shpP2Comp 
         BackColor       =   &H8000000D&
         BorderColor     =   &H80000009&
         Height          =   615
         Left            =   3600
         Shape           =   4  'Rounded Rectangle
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label lblDisP2Human 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Human"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   360
         TabIndex        =   17
         Top             =   1020
         Width           =   2415
      End
      Begin VB.Label lblP2Play 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Next"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         TabIndex        =   16
         Top             =   3550
         Width           =   2535
      End
      Begin VB.Shape shpP2Human 
         BackColor       =   &H8000000D&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         Height          =   615
         Left            =   360
         Shape           =   4  'Rounded Rectangle
         Top             =   960
         Width           =   2415
      End
      Begin VB.Shape shpPlay3 
         BorderColor     =   &H80000009&
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   4320
         Shape           =   4  'Rounded Rectangle
         Top             =   3480
         Width           =   2535
      End
      Begin VB.Image imgSideTankP2 
         Height          =   3930
         Left            =   600
         Picture         =   "frmMain.frx":1D9D
         Top             =   480
         Width           =   6000
      End
   End
   Begin VB.Frame fraP1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      Top             =   6120
      Visible         =   0   'False
      Width           =   615
      Begin VB.TextBox txtP1 
         BackColor       =   &H80000008&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   375
         Left            =   240
         MaxLength       =   10
         TabIndex        =   6
         Text            =   "Player 1"
         Top             =   3480
         Width           =   3735
      End
      Begin VB.Label lblDisP1Name 
         BackStyle       =   0  'Transparent
         Caption         =   "Players Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label lblDisP1Comp 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Computer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   3600
         TabIndex        =   8
         Top             =   1020
         Width           =   2415
      End
      Begin VB.Label lblDisCRP1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright© 1999-2002 Bulldog Creations, All rights reserved."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   210
         Left            =   2520
         TabIndex        =   12
         Top             =   4200
         Width           =   4410
      End
      Begin VB.Label lblP1Close 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   6960
         TabIndex        =   11
         Top             =   0
         Width           =   120
      End
      Begin VB.Label lblDisP1Main 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Player 1:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   480
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   1725
      End
      Begin VB.Shape shpP1Comp 
         BackColor       =   &H8000000D&
         BorderColor     =   &H80000009&
         Height          =   615
         Left            =   3600
         Shape           =   4  'Rounded Rectangle
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label lblDisP1Human 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Human"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   360
         TabIndex        =   9
         Top             =   1020
         Width           =   2415
      End
      Begin VB.Label lblP1Play 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Next"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         TabIndex        =   7
         Top             =   3550
         Width           =   2535
      End
      Begin VB.Shape shpPlay2 
         BorderColor     =   &H80000009&
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   4320
         Shape           =   4  'Rounded Rectangle
         Top             =   3480
         Width           =   2535
      End
      Begin VB.Shape shpP1Human 
         BackColor       =   &H8000000D&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         Height          =   615
         Left            =   360
         Shape           =   4  'Rounded Rectangle
         Top             =   960
         Width           =   2415
      End
      Begin VB.Image imgSideTankP1 
         Height          =   3930
         Left            =   600
         Picture         =   "frmMain.frx":3830
         Top             =   480
         Width           =   6000
      End
   End
   Begin VB.Frame fraIntro 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   6720
      TabIndex        =   0
      Top             =   6120
      Visible         =   0   'False
      Width           =   615
      Begin VB.PictureBox picBlueScroll 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   -1560
         Picture         =   "frmMain.frx":52C3
         ScaleHeight     =   735
         ScaleWidth      =   6495
         TabIndex        =   1
         Top             =   360
         Width           =   6495
      End
      Begin VB.Timer tmrIntro 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   240
         Top             =   4200
      End
      Begin VB.Label lblIntroX 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   6960
         TabIndex        =   4
         Top             =   0
         Width           =   120
      End
      Begin VB.Label lblIntroPlay 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Next"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         TabIndex        =   3
         Top             =   3550
         Width           =   2535
      End
      Begin VB.Image imgIntroBomb 
         Height          =   180
         Left            =   1200
         Picture         =   "frmMain.frx":5A28
         Top             =   5280
         Width           =   75
      End
      Begin VB.Shape shpPlay1 
         BorderColor     =   &H80000009&
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   4320
         Shape           =   4  'Rounded Rectangle
         Top             =   3480
         Width           =   2535
      End
      Begin VB.Label lblDisCopyright 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright© 1999-2002 Bulldog Creations, All rights reserved."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   210
         Left            =   2520
         TabIndex        =   2
         Top             =   4200
         Width           =   4410
      End
      Begin VB.Image imgIntroLogo 
         Height          =   975
         Left            =   240
         Picture         =   "frmMain.frx":5A85
         Top             =   240
         Width           =   4875
      End
      Begin VB.Image imgIntroWood 
         Height          =   1050
         Left            =   960
         Picture         =   "frmMain.frx":7324
         Top             =   5160
         Width           =   600
      End
      Begin VB.Image imgSideTank 
         Height          =   3930
         Left            =   600
         Picture         =   "frmMain.frx":7B38
         Top             =   480
         Width           =   6000
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  frmMain.BackColor = &H80000007
  Call InitBorders
  Call LoadFrame("Intro")
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  frmAddIn.Show
End Sub

Private Sub lblDisP1Comp_Click()
  If (shpP1Comp.BorderWidth = 1) Then
    shpP1Comp.BorderWidth = 3
    shpP1Comp.BackStyle = 1
    shpP1Human.BorderWidth = 1
    shpP1Human.BackStyle = 0
  End If
End Sub

Private Sub lblDisP1Human_Click()
  If (shpP1Human.BorderWidth = 1) Then
    shpP1Human.BorderWidth = 3
    shpP1Human.BackStyle = 1
    shpP1Comp.BorderWidth = 1
    shpP1Comp.BackStyle = 0
  End If
End Sub

Private Sub lblDisP2Comp_Click()
  If (shpP2Comp.BorderWidth = 1) Then
    shpP2Comp.BorderWidth = 3
    shpP2Comp.BackStyle = 1
    shpP2Human.BorderWidth = 1
    shpP2Human.BackStyle = 0
  End If
End Sub

Private Sub lblDisP2Human_Click()
  If (shpP2Human.BorderWidth = 1) Then
    shpP2Human.BorderWidth = 3
    shpP2Human.BackStyle = 1
    shpP2Comp.BorderWidth = 1
    shpP2Comp.BackStyle = 0
  End If
End Sub

Private Sub lblIntroPlay_Click()
  tmrIntro.Enabled = False
  Call PlaySound("SND\Shot.wav")
  Call LoadFrame("Player1")
End Sub

Private Sub lblIntroX_Click()
  Unload Me
End Sub

Private Sub lblOverExit_Click()
  Unload Me
End Sub

Private Sub lblOverNew_Click()
  Call LoadFrame("Intro")
End Sub

Private Sub lblP1Close_Click()
  Unload Me
End Sub

Private Sub lblP1Play_Click()
  Call LoadFrame("Player2")
  Call PlaySound("SND\Shot.wav")
  
  If (shpP1Comp.BorderWidth = 3) Then
    Game = "C"
  Else
    Game = "H"
  End If
End Sub

Private Sub lblP2Close_Click()
  Unload Me
End Sub

Private Sub lblP2Play_Click()
  Call LoadFrame("Loading")
  Call PlaySound("SND\Shot.wav")
  
  If (shpP2Comp.BorderWidth = 3) Then
    Game = Game & "C"
  Else
    Game = Game & "H"
  End If
End Sub

Private Sub picPlay_KeyDown(KeyCode As Integer, Shift As Integer)
  DownKeys(KeyCode) = True
End Sub

Private Sub picPlay_KeyUp(KeyCode As Integer, Shift As Integer)
  DownKeys(KeyCode) = False
End Sub

Private Sub tmrIntro_Timer()

  IntroCount = IntroCount + 1
  
  If (IntroAction = 0) Then
    picBlueScroll.Left = picBlueScroll.Left + 120
    If (picBlueScroll.Left >= frmMain.Width) Then
      IntroAction = 1
      IntroCount = 0
    End If
  ElseIf (IntroAction = 1) Then
    imgIntroWood.Top = imgIntroWood.Top - 60
    If (IntroCount = 30) Then
      IntroAction = 2
      IntroCount = 0
      imgIntroBomb.Top = imgIntroWood.Top
    End If
  ElseIf (IntroAction = 2) Then
    IntroY = IntroY + 1
    If (IntroY = 1) Then
      imgIntroBomb.Visible = True
      Call PlaySound("SND\Shot.wav")
    End If
    If (IntroY = 23) Then Call PlaySound("SND\Hit.wav")
    If (IntroY = 43) Then
      IntroY = -500
      imgIntroBomb.Visible = False
      imgIntroBomb.Top = imgIntroWood.Top
    End If
    If ((IntroY > 0) And (IntroY < 40)) Then
      imgIntroBomb.Top = imgIntroBomb.Top - 180
    End If
  End If
  DoEvents

End Sub

Private Sub tmrLoading_Timer()
  tmrLoading.Enabled = False
  Call LoadFrame("Play")
End Sub

Private Sub tmrPlay_Timer()
  Call CheckOptions
  If (Left(Game, 1) = "C") Then CompMove1
  If (Right(Game, 1) = "C") Then CompMove2
  Call DrawPic
  Call MoveShots
  DoEvents
End Sub
