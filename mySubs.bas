Attribute VB_Name = "mySubs"
Option Explicit


Public Sub InitBorders()
  Dim I As Integer

  frmMain.Height = 4665
  frmMain.Width = 7365
  
  frmMain.Cls
  
  For I = 0 To frmMain.ScaleHeight / 3 - 1
    BitBlt frmMain.hdc, 0, I * frmHold.picBorderLeft.ScaleHeight, frmHold.picBorderLeft.ScaleWidth, frmHold.picBorderLeft.ScaleHeight, frmHold.picBorderLeft.hdc, 0, 0, SRCAND
    BitBlt frmMain.hdc, 0, I * frmHold.picBorderLeft.ScaleHeight, frmHold.picBorderLeft.ScaleWidth, frmHold.picBorderLeft.ScaleHeight, frmHold.picBorderLeft.hdc, 0, 0, SRCINVERT
    BitBlt frmMain.hdc, frmMain.ScaleWidth - frmHold.picBorderRight.ScaleWidth, I * frmHold.picBorderRight.ScaleHeight, frmHold.picBorderRight.ScaleWidth, frmHold.picBorderRight.ScaleHeight, frmHold.picBorderRight.hdc, 0, 0, SRCAND
    BitBlt frmMain.hdc, frmMain.ScaleWidth - frmHold.picBorderRight.ScaleWidth, I * frmHold.picBorderRight.ScaleHeight, frmHold.picBorderRight.ScaleWidth, frmHold.picBorderRight.ScaleHeight, frmHold.picBorderRight.hdc, 0, 0, SRCINVERT
  Next I
  
  For I = 0 To frmMain.ScaleWidth / 3 - 1
    BitBlt frmMain.hdc, I * frmHold.picBorderTop.ScaleWidth, 0, frmHold.picBorderTop.ScaleWidth, frmHold.picBorderTop.ScaleHeight, frmHold.picBorderTop.hdc, 0, 0, SRCAND
    BitBlt frmMain.hdc, I * frmHold.picBorderTop.ScaleWidth, 0, frmHold.picBorderTop.ScaleWidth, frmHold.picBorderTop.ScaleHeight, frmHold.picBorderTop.hdc, 0, 0, SRCINVERT
    BitBlt frmMain.hdc, I * frmHold.picBorderBottom.ScaleWidth, frmMain.ScaleHeight - frmHold.picBorderBottom.ScaleHeight, frmHold.picBorderBottom.ScaleWidth, frmHold.picBorderBottom.ScaleHeight, frmHold.picBorderBottom.hdc, 0, 0, SRCAND
    BitBlt frmMain.hdc, I * frmHold.picBorderBottom.ScaleWidth, frmMain.ScaleHeight - frmHold.picBorderBottom.ScaleHeight, frmHold.picBorderBottom.ScaleWidth, frmHold.picBorderBottom.ScaleHeight, frmHold.picBorderBottom.hdc, 0, 0, SRCINVERT
  Next I
  
  BitBlt frmMain.hdc, 0, 0, frmHold.picBorderCor.ScaleWidth, frmHold.picBorderCor.ScaleHeight, frmHold.picBorderCor.hdc, 0, 0, SRCAND
  BitBlt frmMain.hdc, 0, 0, frmHold.picBorderTL.ScaleWidth, frmHold.picBorderTL.ScaleHeight, frmHold.picBorderTL.hdc, 0, 0, SRCINVERT

  BitBlt frmMain.hdc, frmMain.ScaleWidth - frmHold.picBorderCor.ScaleWidth, 0, frmHold.picBorderCor.ScaleWidth, frmHold.picBorderCor.ScaleHeight, frmHold.picBorderCor.hdc, 0, 0, SRCAND
  BitBlt frmMain.hdc, frmMain.ScaleWidth - frmHold.picBorderTR.ScaleWidth, 0, frmHold.picBorderTR.ScaleWidth, frmHold.picBorderTR.ScaleHeight, frmHold.picBorderTR.hdc, 0, 0, SRCINVERT

  BitBlt frmMain.hdc, frmMain.ScaleWidth - frmHold.picBorderCor.ScaleWidth, frmMain.ScaleHeight - frmHold.picBorderCor.ScaleHeight, frmHold.picBorderCor.ScaleWidth, frmHold.picBorderCor.ScaleHeight, frmHold.picBorderCor.hdc, 0, 0, SRCAND
  BitBlt frmMain.hdc, frmMain.ScaleWidth - frmHold.picBorderBR.ScaleWidth, frmMain.ScaleHeight - frmHold.picBorderBR.ScaleHeight, frmHold.picBorderBR.ScaleWidth, frmHold.picBorderBR.ScaleHeight, frmHold.picBorderBR.hdc, 0, 0, SRCINVERT

  BitBlt frmMain.hdc, 0, frmMain.ScaleHeight - frmHold.picBorderCor.ScaleHeight, frmHold.picBorderCor.ScaleWidth, frmHold.picBorderCor.ScaleHeight, frmHold.picBorderCor.hdc, 0, 0, SRCAND
  BitBlt frmMain.hdc, 0, frmMain.ScaleHeight - frmHold.picBorderBL.ScaleHeight, frmHold.picBorderBL.ScaleWidth, frmHold.picBorderBL.ScaleHeight, frmHold.picBorderBL.hdc, 0, 0, SRCINVERT

End Sub

Function PlaySound(inFile As String) As Boolean
  On Error GoTo sendError
  
  Call sndPlaySound(inFile, 1)

  PlaySound = True
  Exit Function
sendError:
  PlaySound = False
  MsgBox Err.Description
  Exit Function
End Function

Function LoadFrame(inFrame As String) As Boolean
  On Error GoTo sendError
  
  Dim I As Integer
  
  frmMain.fraIntro.Visible = False
  frmMain.fraP1.Visible = False
  frmMain.fraP2.Visible = False
  frmMain.fraLoading.Visible = False
  frmMain.fraPlay.Visible = False
  frmMain.fraOver.Visible = False
  
  If (inFrame = "Intro") Then
    IntroCount = 0
    IntroAction = 0
    IntroY = 0
    frmMain.picBlueScroll.Left = -1560
    frmMain.fraIntro.Top = fTop
    frmMain.fraIntro.Left = fLeft
    frmMain.fraIntro.Height = fHeight
    frmMain.fraIntro.Width = fWidth
    frmMain.fraIntro.Visible = True
    frmMain.tmrIntro.Enabled = True
  ElseIf (inFrame = "Player1") Then
    frmMain.fraP1.Top = fTop
    frmMain.fraP1.Left = fLeft
    frmMain.fraP1.Height = fHeight
    frmMain.fraP1.Width = fWidth
    frmMain.fraP1.Visible = True
  ElseIf (inFrame = "Player2") Then
    frmMain.fraP2.Top = fTop
    frmMain.fraP2.Left = fLeft
    frmMain.fraP2.Height = fHeight
    frmMain.fraP2.Width = fWidth
    frmMain.fraP2.Visible = True
  ElseIf (inFrame = "Loading") Then
    frmMain.fraLoading.Top = fTop
    frmMain.fraLoading.Left = fLeft
    frmMain.fraLoading.Height = fHeight
    frmMain.fraLoading.Width = fWidth
    frmMain.fraLoading.Visible = True
    frmMain.tmrLoading.Enabled = True
  ElseIf (inFrame = "Play") Then
    If (Game = "CC") Then
      TopRandNum = 15
    Else
      TopRandNum = 5
    End If
    ReDim Shot1(5) As ShotType
    ReDim Shot2(5) As ShotType
    Tank1.X = (frmMain.picPlay.ScaleWidth / 2) - (frmHold.picBlue.ScaleWidth / 2) + 1
    Tank1.Y = 5
    Tank2.X = (frmMain.picPlay.ScaleWidth / 2) - (frmHold.picBlue.ScaleWidth / 2)
    Tank2.Y = frmMain.picPlay.ScaleHeight - (frmHold.PicWoodMask.ScaleHeight + 5)
    frmMain.fraPlay.Top = fTop
    frmMain.fraPlay.Left = fLeft
    frmMain.fraPlay.Height = fHeight
    frmMain.fraPlay.Width = fWidth
    frmMain.fraPlay.Visible = True
    frmMain.shpP1TOP.Width = 2000
    frmMain.shpP2TOP.Width = 2000
    frmMain.tmrPlay.Enabled = True
    OptCmds = True
    OptBG = True
    For I = 0 To 255
      DownKeys(I) = False
    Next I
    frmMain.lblDisP1.Caption = frmMain.txtP1.Text
    frmMain.lblDisP2.Caption = frmMain.txtP2.Text
    frmCMD.Show
    frmMain.picPlay.SetFocus
  ElseIf (inFrame = "Over") Then
    frmMain.fraOver.Top = fTop
    frmMain.fraOver.Left = fLeft
    frmMain.fraOver.Height = fHeight
    frmMain.fraOver.Width = fWidth
    frmMain.fraOver.Visible = True
    If (frmMain.shpP1TOP.Width > frmMain.shpP2TOP.Width) Then
      frmMain.lblDisWinner.Caption = frmMain.lblDisP1.Caption
    ElseIf (frmMain.shpP1TOP.Width < frmMain.shpP2TOP.Width) Then
      frmMain.lblDisWinner.Caption = frmMain.lblDisP2.Caption
    ElseIf (frmMain.shpP1TOP.Width = frmMain.shpP2TOP.Width) Then
      frmMain.lblDisCongrats.Caption = "It Was A Tie!"
      frmMain.lblDisWinner.Caption = "Fair Game"
    End If
  End If
  
  LoadFrame = True
  Exit Function
sendError:
  MsgBox Err.Description
  LoadFrame = False
  Exit Function
End Function

Function CheckOptions()
  On Error GoTo sendError
  
  If (DownKeys(27)) Then
    frmMain.tmrPlay.Enabled = False
    Call LoadFrame("Over")
  End If

  If (DownKeys(112)) Then
    If (OptBG = True) Then
      OptBG = False
    Else
      OptBG = True
    End If
  End If
  
  If (DownKeys(113)) Then
    If (OptCmds = True) Then
      Unload frmCMD
    Else
      frmCMD.Show
      frmMain.picPlay.SetFocus
    End If
  End If
  
  If (Left(Game, 1) = "H") Then
    If ((DownKeys(37)) And (Tank1.X > 10)) Then Tank1.X = Tank1.X - 10
    If ((DownKeys(39)) And (Tank1.X < frmMain.picPlay.ScaleWidth - (10 + frmHold.picBlueMask.ScaleWidth))) Then Tank1.X = Tank1.X + 10
    If ((DownKeys(38)) And (Tank1.Y > 10)) Then Tank1.Y = Tank1.Y - 10
    If ((DownKeys(40)) And (Tank1.Y < 60)) Then Tank1.Y = Tank1.Y + 10
    If (DownKeys(48)) Then ShootPlayer1
  End If
  
  If (Right(Game, 1) = "H") Then
    If ((DownKeys(65)) And (Tank2.X > 10)) Then Tank2.X = Tank2.X - 10
    If ((DownKeys(68)) And (Tank2.X < frmMain.picPlay.ScaleWidth - (10 + frmHold.PicWoodMask.ScaleWidth))) Then Tank2.X = Tank2.X + 10
    If ((DownKeys(87)) And (Tank2.Y > frmMain.picPlay.ScaleHeight - (60 + frmHold.PicWoodMask.ScaleHeight))) Then Tank2.Y = Tank2.Y - 10
    If ((DownKeys(83)) And (Tank2.Y < frmMain.picPlay.ScaleHeight - (10 + frmHold.PicWoodMask.ScaleHeight))) Then Tank2.Y = Tank2.Y + 10
    If (DownKeys(49)) Then ShootPlayer2
  End If
  
  Exit Function
sendError:
  MsgBox Err.Description
  Exit Function
End Function

Function CompMove1()
  On Error GoTo sendError
  
  Randomize
  
  If ((Tank2.X > Tank1.X) And (Tank1.X + frmHold.picBlueMask.ScaleWidth + 10 < frmMain.picPlay.ScaleWidth)) Then Tank1.X = Tank1.X + Int(Rnd * TopRandNum + 5)
  If ((Tank2.X < Tank1.X) And (Tank1.X - 10 > 0)) Then Tank1.X = Tank1.X - Int(Rnd * TopRandNum + 5)
  
  If ((Tank2.X > Tank1.X - frmHold.PicWoodMask.ScaleWidth) And (Tank2.X < Tank1.X + frmHold.PicWoodMask.ScaleWidth)) Then ShootPlayer1
  
  Exit Function
sendError:
  MsgBox Err.Description
  Exit Function
End Function

Function CompMove2()
  On Error GoTo sendError
  
  Randomize

  If ((Tank1.X > Tank2.X) And (Tank2.X + frmHold.PicWoodMask.ScaleWidth + 10 < frmMain.picPlay.ScaleWidth)) Then Tank2.X = Tank2.X + Int(Rnd * TopRandNum + 5)
  If ((Tank1.X < Tank2.X) And (Tank2.X - 10 > 0)) Then Tank2.X = Tank2.X - Int(Rnd * TopRandNum + 5)
  
  If ((Tank1.X > Tank2.X - frmHold.picBlueMask.ScaleWidth) And (Tank1.X < Tank2.X + frmHold.picBlueMask.ScaleWidth)) Then ShootPlayer2
  
  Exit Function
sendError:
  MsgBox Err.Description
  Exit Function
End Function

Function ShootPlayer1()
  Dim I As Integer
  Dim NewShot As Integer
  Dim Found As Boolean
  Found = False
  For I = 0 To UBound(Shot1)
    If (Shot1(I).Active = False) Then
      Found = True
      NewShot = I
      Exit For
    End If
  Next I
  If (Not (Found)) Then
    NewShot = UBound(Shot1) + 1
    ReDim Preserve Shot1(UBound(Shot1) + 3) As ShotType
  End If
  Shot1(NewShot).Active = True
  Shot1(NewShot).X = Tank1.X + (frmHold.picBlueMask.ScaleWidth / 2) - (frmHold.picShotDownMask.ScaleWidth / 2)
  Shot1(NewShot).Y = Tank1.Y + frmHold.picBlueMask.ScaleHeight
  Call PlaySound("SND\Shot.wav")
  Exit Function
sendError:
  MsgBox Err.Description
  Exit Function
End Function

Function ShootPlayer2()
  On Error GoTo sendError
      
  Dim I As Integer
  Dim NewShot As Integer
  Dim Found As Boolean
  Found = False
  For I = 0 To UBound(Shot2)
    If (Shot2(I).Active = False) Then
      Found = True
      NewShot = I
      Exit For
    End If
  Next I
  If (Not (Found)) Then
    NewShot = UBound(Shot2) + 1
    ReDim Preserve Shot2(UBound(Shot2) + 3) As ShotType
  End If
  Shot2(NewShot).Active = True
  Shot2(NewShot).X = Tank2.X + (frmHold.PicWoodMask.ScaleWidth / 2) - (frmHold.picShotUpMask.ScaleWidth / 2)
  Shot2(NewShot).Y = Tank2.Y
  Call PlaySound("SND\Shot.wav")
  Exit Function
sendError:
  MsgBox Err.Description
  Exit Function
End Function

Function DrawPic()
  On Error GoTo sendError
  
  Dim I As Integer
  
  frmMain.picPlay.Cls

  If (OptBG) Then
    BitBlt frmMain.picPlay.hdc, 0, 0, frmHold.picBG.ScaleWidth, frmHold.picBG.ScaleHeight, frmHold.picBG.hdc, 0, 0, SRCAND
    BitBlt frmMain.picPlay.hdc, 0, 0, frmHold.picBG.ScaleWidth, frmHold.picBG.ScaleHeight, frmHold.picBG.hdc, 0, 0, SRCINVERT
  End If
  
  Call BitBlt(frmMain.picPlay.hdc, Tank1.X, Tank1.Y, frmHold.picBlueMask.ScaleWidth, frmHold.picBlueMask.ScaleHeight, frmHold.picBlueMask.hdc, 0, 0, SRCAND)
  Call BitBlt(frmMain.picPlay.hdc, Tank1.X, Tank1.Y, frmHold.picBlue.ScaleWidth, frmHold.picBlue.ScaleHeight, frmHold.picBlue.hdc, 0, 0, SRCINVERT)
  
  Call BitBlt(frmMain.picPlay.hdc, Tank2.X, Tank2.Y, frmHold.PicWoodMask.ScaleWidth, frmHold.PicWoodMask.ScaleHeight, frmHold.PicWoodMask.hdc, 0, 0, SRCAND)
  Call BitBlt(frmMain.picPlay.hdc, Tank2.X, Tank2.Y, frmHold.PicWood.ScaleWidth, frmHold.PicWood.ScaleHeight, frmHold.PicWood.hdc, 0, 0, SRCINVERT)
    
  For I = 0 To UBound(Shot1)
    If (Shot1(I).Active) Then
      Call BitBlt(frmMain.picPlay.hdc, Shot1(I).X, Shot1(I).Y, frmHold.picShotDownMask.ScaleWidth, frmHold.picShotDownMask.ScaleHeight, frmHold.picShotDownMask.hdc, 0, 0, SRCAND)
      Call BitBlt(frmMain.picPlay.hdc, Shot1(I).X, Shot1(I).Y, frmHold.picShotDown.ScaleWidth, frmHold.picShotDown.ScaleHeight, frmHold.picShotDown.hdc, 0, 0, SRCINVERT)
    End If
  Next I
    
  For I = 0 To UBound(Shot2)
    If (Shot2(I).Active) Then
      Call BitBlt(frmMain.picPlay.hdc, Shot2(I).X, Shot2(I).Y, frmHold.picShotUpMask.ScaleWidth, frmHold.picShotUpMask.ScaleHeight, frmHold.picShotUpMask.hdc, 0, 0, SRCAND)
      Call BitBlt(frmMain.picPlay.hdc, Shot2(I).X, Shot2(I).Y, frmHold.picShotUp.ScaleWidth, frmHold.picShotUp.ScaleHeight, frmHold.picShotUp.hdc, 0, 0, SRCINVERT)
    End If
  Next I
    
    
    'BitBlt frmMain.hdc, 0, I * frmHold.picBorderLeft.ScaleHeight, frmHold.picBorderLeft.ScaleWidth, frmHold.picBorderLeft.ScaleHeight, frmHold.picBorderLeft.hdc, 0, 0, SRCAND
    'BitBlt frmMain.hdc, 0, I * frmHold.picBorderLeft.ScaleHeight, frmHold.picBorderLeft.ScaleWidth, frmHold.picBorderLeft.ScaleHeight, frmHold.picBorderLeft.hdc, 0, 0, SRCINVERT
 
  Exit Function
sendError:
  MsgBox Err.Description
  Exit Function
End Function

Function MoveShots()
  On Error GoTo sendError
  
  Dim I As Integer
  
  For I = 0 To UBound(Shot1)
    If (Shot1(I).Active) Then Shot1(I).Y = Shot1(I).Y + 24
    If (Shot1(I).Y > frmMain.picPlay.ScaleHeight) Then Shot1(I).Active = False
    If ((Shot1(I).Active) And (Shot1(I).X >= Tank2.X) And (Shot1(I).X < Tank2.X + frmHold.PicWoodMask.ScaleWidth) And (Shot1(I).Y + frmHold.picShotDownMask.ScaleHeight > Tank2.Y + 1) And (Shot1(I).Y + frmHold.picShotDownMask.ScaleHeight < Tank2.Y + frmHold.PicWoodMask.ScaleHeight)) Then
      Shot1(I).Active = False
      frmMain.shpP2TOP.Width = frmMain.shpP2TOP.Width - 50
      If (frmMain.shpP2TOP.Width <= 15) Then
        frmMain.tmrPlay.Enabled = False
        Call LoadFrame("Over")
      End If
    End If
  Next I
  
  For I = 0 To UBound(Shot2)
    If (Shot2(I).Active) Then Shot2(I).Y = Shot2(I).Y - 24
    If (Shot2(I).Y < 0) Then Shot2(I).Active = False
    If ((Shot2(I).Active) And (Shot2(I).X >= Tank1.X) And (Shot2(I).X < Tank1.X + frmHold.picBlueMask.ScaleWidth) And (Shot2(I).Y < Tank1.Y + frmHold.picBlueMask.ScaleHeight - 1) And (Shot2(I).Y + frmHold.picShotUpMask.ScaleHeight > Tank1.Y)) Then
      Shot2(I).Active = False
      frmMain.shpP1TOP.Width = frmMain.shpP1TOP.Width - 50
      If (frmMain.shpP1TOP.Width <= 15) Then
        frmMain.tmrPlay.Enabled = False
        Call LoadFrame("Over")
      End If
    End If
  Next I
  
  Exit Function
sendError:
  MsgBox Err.Description
  Exit Function
  

End Function



