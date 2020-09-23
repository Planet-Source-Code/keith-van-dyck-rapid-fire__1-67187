VERSION 5.00
Begin VB.Form frmHold 
   Caption         =   "Form2"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7335
   LinkTopic       =   "Form2"
   ScaleHeight     =   391
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   489
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picShotUpMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   2040
      Picture         =   "frmHold.frx":0000
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   3
      TabIndex        =   19
      Top             =   480
      Width           =   105
   End
   Begin VB.PictureBox picShotUp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   1920
      Picture         =   "frmHold.frx":0041
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   3
      TabIndex        =   18
      Top             =   480
      Width           =   105
   End
   Begin VB.PictureBox picShotDownMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   2040
      Picture         =   "frmHold.frx":0095
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   3
      TabIndex        =   17
      Top             =   240
      Width           =   105
   End
   Begin VB.PictureBox picShotDown 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   1920
      Picture         =   "frmHold.frx":00D6
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   3
      TabIndex        =   16
      Top             =   240
      Width           =   105
   End
   Begin VB.PictureBox picBG 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   4455
      Left            =   7200
      Picture         =   "frmHold.frx":012A
      ScaleHeight     =   293
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   473
      TabIndex        =   15
      Top             =   0
      Width           =   7155
   End
   Begin VB.PictureBox PicWoodMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   585
      Left            =   5160
      Picture         =   "frmHold.frx":1349A
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   14
      Top             =   120
      Width           =   360
   End
   Begin VB.PictureBox PicWood 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   585
      Left            =   4800
      Picture         =   "frmHold.frx":13547
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   13
      Top             =   120
      Width           =   360
   End
   Begin VB.PictureBox picBlueMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   585
      Left            =   4440
      Picture         =   "frmHold.frx":13A02
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   12
      Top             =   120
      Width           =   360
   End
   Begin VB.PictureBox picBlue 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   585
      Left            =   4080
      Picture         =   "frmHold.frx":13AAD
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   11
      Top             =   120
      Width           =   360
   End
   Begin VB.PictureBox picBombMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   240
      Left            =   1560
      Picture         =   "frmHold.frx":13F98
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   10
      Top             =   480
      Width           =   135
   End
   Begin VB.PictureBox picBomb 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   240
      Left            =   1440
      Picture         =   "frmHold.frx":13FE0
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   9
      Top             =   480
      Width           =   135
   End
   Begin VB.PictureBox picBorderCor 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   480
      Picture         =   "frmHold.frx":1403D
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   8
      Top             =   480
      Width           =   255
   End
   Begin VB.PictureBox picBorderBR 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   840
      Picture         =   "frmHold.frx":14083
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   7
      Top             =   720
      Width           =   255
   End
   Begin VB.PictureBox picBorderTR 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   840
      Picture         =   "frmHold.frx":140F8
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   6
      Top             =   105
      Width           =   255
   End
   Begin VB.PictureBox picBorderBL 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   240
      Picture         =   "frmHold.frx":1416C
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   5
      Top             =   720
      Width           =   255
   End
   Begin VB.PictureBox picBorderTL 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   240
      Picture         =   "frmHold.frx":141E0
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   4
      Top             =   105
      Width           =   255
   End
   Begin VB.PictureBox picBorderTop 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   210
      Left            =   600
      Picture         =   "frmHold.frx":1424D
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   3
      TabIndex        =   3
      Top             =   120
      Width           =   105
   End
   Begin VB.PictureBox picBorderBottom 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   210
      Left            =   600
      Picture         =   "frmHold.frx":1429E
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   3
      TabIndex        =   2
      Top             =   720
      Width           =   105
   End
   Begin VB.PictureBox picBorderRight 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   105
      Left            =   840
      Picture         =   "frmHold.frx":142EF
      ScaleHeight     =   3
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   1
      Top             =   480
      Width           =   210
   End
   Begin VB.PictureBox picBorderLeft 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   105
      Left            =   240
      Picture         =   "frmHold.frx":14343
      ScaleHeight     =   3
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   0
      Top             =   480
      Width           =   210
   End
End
Attribute VB_Name = "frmHold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
