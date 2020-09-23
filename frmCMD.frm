VERSION 5.00
Begin VB.Form frmCMD 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Commands"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2175
   ControlBox      =   0   'False
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   2175
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstShow 
      Appearance      =   0  'Flat
      Height          =   4515
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmCMD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDone_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  OptCmds = True
  lstShow.AddItem ("Esc: Exit Game")
  lstShow.AddItem ("F1: Toggle Background")
  lstShow.AddItem ("F2: Toggle Command List")
  lstShow.AddItem ("")
  lstShow.AddItem ("Player 1:")
  lstShow.AddItem ("  Move Up: Up Arrow")
  lstShow.AddItem ("  Move Down: Down Arrow")
  lstShow.AddItem ("  Move Left: Left Arrow")
  lstShow.AddItem ("  Move Right: Right Arrow")
  lstShow.AddItem ("  Shoot: 0")
  lstShow.AddItem ("")
  lstShow.AddItem ("Player 2:")
  lstShow.AddItem ("  Move Up: W")
  lstShow.AddItem ("  Move Down: S")
  lstShow.AddItem ("  Move Left: A")
  lstShow.AddItem ("  Move Right:  D ")
  lstShow.AddItem ("  Shoot: 1")
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  OptCmds = False
End Sub
