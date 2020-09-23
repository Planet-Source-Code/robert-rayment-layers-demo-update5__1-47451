Attribute VB_Name = "DrawMod"
' DrawMod.bas
Option Explicit

Public Sub DRAW_START(PIC As PictureBox, SHP As Shape, X As Single, Y As Single)

      Select Case TheDrawStyle
      Case 0, 1, 2 ' Blend weak, Blend med, Blend strong
         SHP.Visible = False
         With SHP
            .Width = BESize
            .Height = BESize
            .Left = X - BESize \ 2
            .Top = Y - BESize \ 2
            .Visible = True
         End With
         
         Select Case TheDrawStyle
         Case 0: SStep = 4 ' Weak blend
         Case 1: SStep = 2 ' Med blend
         Case 2: SStep = 1 ' Strong blend
         End Select
         
         maxw = PIC.Width
         maxh = PIC.Height
         ReDim memBytes(1 To 4, 1 To maxw, 1 To maxh)
         GETBYTES PIC.Image
         
         aDrawMove = True
      Case 3, 4  ' Erase rect, Erase circ
         SHP.Visible = False
         With SHP
            .Width = BESize
            .Height = BESize
            .Left = X - BESize \ 2
            .Top = Y - BESize \ 2
            .Visible = True
         End With
         aDrawMove = True
      
      Case 5:  StartFreeDraw PIC, X, Y
               aDrawStart = True
      Case 6:  StartLine PIC, X, Y
               aDrawStart = True
      Case 7, 8, 9, 10: StartRectangles PIC, X, Y
               aDrawStart = True
      Case 11, 12, 13: StartCirllipses PIC, X, Y
               aDrawStart = True
      Case 14: StartArch PIC, X, Y
               aDrawStart = False
               aDrawMove = False
      Case 15: StartFRibbon PIC, X, Y     ' ///
               aDrawStart = True
      Case 16: StartBRibbon PIC, X, Y     ' \\\
               aDrawStart = True
      Case 17: StartPolyLine PIC, X, Y
               aDrawStart = True
      Case 18: StartSpline PIC, X, Y
               aDrawStart = True
      Case 19: StartSpray PIC, X, Y
               aDrawStart = True
      Case 20: StartStar PIC, X, Y
               aDrawStart = True
      Case 21, 22: StartPlusT PIC, X, Y
               aDrawStart = True
      Case 23: StartParallelogram PIC, X, Y
               aDrawStart = True
      Case 24: StartFrustrum PIC, X, Y
               aDrawStart = True
      Case 25 To 28: StartArrow PIC, X, Y
               aDrawStart = True
      
      End Select
End Sub

Public Sub DRAW_DRAW(PIC As PictureBox, X As Single, Y As Single)

      Select Case TheDrawStyle
      Case 5:  DrawFreeDraw PIC, X, Y
      Case 6:  DrawLine PIC, X, Y
      Case 7, 8, 9, 10: DrawRectangles PIC, X, Y
      Case 11, 12, 13: DrawCirllipses PIC, X, Y
      Case 14: DrawArch PIC, X, Y
      Case 15: DrawFRibbon PIC, X, Y     ' ///
      Case 16: DrawBRibbon PIC, X, Y     ' \\\
      Case 17: DrawPolyLine PIC, X, Y
      Case 18: DrawSpline PIC, X, Y
      Case 19: DrawSpray PIC, X, Y
      Case 20: DrawStar PIC, X, Y
      Case 21, 22: DrawPlusT PIC, X, Y
      Case 23: DrawParallelogram PIC, X, Y
      Case 24: DrawFrustrum PIC, X, Y
      Case 25 To 28: DrawArrow PIC, X, Y
      End Select

End Sub

Public Sub DRAW_MOVE(PIC As PictureBox, SHP As Shape, X As Single, Y As Single)
Dim ii As Long, jj As Long, k As Long


      Select Case TheDrawStyle
      Case 0, 1, 2 ' Blend weak, Blend med, Blend strong
         With SHP
            .Left = X - BESize \ 2
            .Top = Y - BESize \ 2
         End With
      
         'DO_BLENDERS
         For j = Y - BESize \ 2 To Y + BESize \ 2 Step SStep
            If j > 1 And j < PIC.Height - 2 Then
               hh = (j - Y)
               'when j=Y-BESize \ 2 hh= -BESize \ 2, j=Y+BESize \ 2 hh=BESize \ 2
               For i = X - BESize \ 2 To X + BESize \ 2 Step SStep
                  If i > 1 And i < PIC.Width - 2 Then
                     ww = (i - X)
                     'when i=X-BESize \ 2 ww= -BESize \ 2, i=X+BESize \ 2 ww=BESize \ 2
                     ' Confine to circle radius BESize \ 2
                     If ww * ww + hh * hh < BESize * BESize / 4 Then
                        ii = i + 1
                        jj = j + 1
                        If RGB(memBytes(3, ii, jj), memBytes(2, ii, jj), memBytes(1, ii, jj)) _
                           <> TransColor Then
                         For k = 1 To 3
                             memBytes(k, ii, jj) = (1& * memBytes(k, ii, jj - 1) + _
                             memBytes(k, ii, jj + 1) + memBytes(k, ii - 1, jj) + _
                             memBytes(k, ii + 1, jj) + memBytes(k, ii, jj)) \ 5
                         Next k
                        End If
                     End If
                  End If
               Next i
            End If
         Next j
      
         SetStretchBltMode PIC.hdc, HALFTONE
         
         response = StretchDIBits(PIC.hdc, _
                  X - BESize \ 2, Y - BESize \ 2, BESize, BESize, _
                  X - BESize \ 2, PIC.Height - (Y + BESize \ 2), BESize, BESize, _
                  memBytes(1, 1, 1), bmbac, 0, vbSrcCopy)
         
         PIC.Refresh
      
      Case 3, 4  ' Erase rect, Erase circ
         With SHP
            .Left = X - BESize \ 2
            .Top = Y - BESize \ 2
         End With
         PIC.DrawMode = 13
         EraseColor = PIC.ForeColor
         svDrawWidth = TheDrawWidth
         Select Case TheDrawStyle
         Case 3
            PIC.Line (X - BESize \ 2, Y - BESize \ 2)-(X + BESize \ 2, Y + BESize \ 2), EraseColor, BF
         Case 4
            zradius = BESize \ 2
            PIC.FillColor = EraseColor
            PIC.FillStyle = 0   ' Solid
            PIC.Circle (X, Y), zradius, EraseColor
            PIC.FillStyle = 1   ' Transparent
            PIC.Refresh
         End Select
         PIC.DrawWidth = svDrawWidth
      
      Case 5:  MoveFreeDraw PIC, X, Y
      Case 6:  MoveLine PIC, X, Y
      Case 7, 8, 9, 10: MoveRectangles PIC, X, Y
      Case 11, 12, 13: MoveCirllipses PIC, X, Y
      Case 14: MoveArch PIC, X, Y
      Case 15: MoveFRibbon PIC, X, Y     ' ///
      Case 16: MoveBRibbon PIC, X, Y    ' \\\
      Case 17: MovePolyLine PIC, X, Y
      Case 18: MoveSpline PIC, X, Y
      Case 19: MoveSpray PIC, X, Y
      Case 20: MoveStar PIC, X, Y
      Case 21, 22: MovePlusT PIC, X, Y
      Case 23: MoveParallelogram PIC, X, Y
      Case 24: MoveFrustrum PIC, X, Y
      Case 25 To 28: MoveArrow PIC, X, Y
      End Select
   
End Sub

Public Sub DRAW_FINAL(PIC As PictureBox, PIC2 As PictureBox, X As Single, Y As Single)

If aTrace Then PIC2.DrawWidth = TheDrawWidth

      Select Case TheDrawStyle
      Case 5:  FinalFreeDraw PIC, PIC2, X, Y
      Case 6:  FinalLine PIC, PIC2, X, Y
      Case 7, 8, 9, 10: FinalRectangles PIC, PIC2, X, Y
      Case 11, 12, 13: FinalCirllipses PIC, PIC2, X, Y
      Case 14: FinalArch PIC, PIC2, X, Y
      Case 15: FinalFRibbon PIC, PIC2, X, Y    ' ///
      Case 16: FinalBRibbon PIC, PIC2, X, Y    ' \\\
      Case 17: FinalPolyLine PIC, PIC2, X, Y
      Case 18: FinalSpline PIC, PIC2, X, Y
      Case 19: FinalSpray PIC, PIC2, X, Y
      Case 20: FinalStar PIC, PIC2, X, Y
      Case 21, 22: FinalPlusT PIC, PIC2, X, Y
      Case 23: FinalParallelogram PIC, PIC2, X, Y
      Case 24: FinalFrustrum PIC, PIC2, X, Y
      Case 25 To 28: FinalArrow PIC, PIC2, X, Y
      End Select
      
      aDrawStart = False
      aDrawMove = False


'frmTools.Label2 = "LCC=" & Str$(LClickCount) & "RCC=" & Str$(RClickCount) & " NPts=" & Str$(NumOfDrawPoints)
LClickCount = 0: RClickCount = 0: NumOfDrawPoints = 0
End Sub



'#### FREEDRAW ###################################################
Public Sub StartFreeDraw(PIC As PictureBox, X As Single, Y As Single)
   NumOfDrawPoints = 1
   ReDim ixPts(1 To 1), iyPts(1 To 1)
   ixPts(1) = X
   iyPts(1) = Y
   PIC.PSet (X, Y), vbWhite
   PIC.Refresh
End Sub

Public Sub DrawFreeDraw(PIC As PictureBox, X As Single, Y As Single)
   NumOfDrawPoints = NumOfDrawPoints + 1
   ReDim Preserve ixPts(1 To NumOfDrawPoints)
   ReDim Preserve iyPts(1 To NumOfDrawPoints)
   ixTL = X
   iyTL = Y
   ixPts(NumOfDrawPoints) = X
   iyPts(NumOfDrawPoints) = Y
   PIC.Line -(X, Y), vbWhite
   PIC.Refresh
End Sub

Public Sub MoveFreeDraw(PIC As PictureBox, X As Single, Y As Single)
   'Clear old points
   PIC.PSet (ixPts(1), iyPts(1)), vbWhite
   For i = 2 To NumOfDrawPoints
      PIC.Line -(ixPts(i), iyPts(i)), vbWhite
   Next i
   ' Move points to new position
   For i = 1 To NumOfDrawPoints
      ixPts(i) = ixPts(i) + (X - ixTL)
      iyPts(i) = iyPts(i) + (Y - iyTL)
   Next i
   ixTL = X
   iyTL = Y
   PIC.PSet (ixPts(1), iyPts(1)), vbWhite
   For i = 2 To NumOfDrawPoints
      PIC.Line -(ixPts(i), iyPts(i)), vbWhite
   Next i
   PIC.Refresh
End Sub

Public Sub FinalFreeDraw(PIC As PictureBox, PIC2 As PictureBox, X As Single, Y As Single)
   'Clear old points
   PIC.PSet (ixPts(1), iyPts(1)), vbWhite
   For i = 2 To NumOfDrawPoints
      PIC.Line -(ixPts(i), iyPts(i)), vbWhite
   Next i
   PIC.DrawMode = 13
   ' Draw Final points
   PIC.PSet (ixPts(1), iyPts(1)), DrawColor
   For i = 2 To NumOfDrawPoints
      PIC.Line -(ixPts(i), iyPts(i)), DrawColor
   Next i
'--------------------------------------------
            If aTrace Then
               'Draw final points on picFullTemp as well
               PIC2.PSet (ixPts(1) - L, iyPts(1) - T), DrawColor
               For i = 1 To NumOfDrawPoints
                  PIC2.Line -(ixPts(i) - L, iyPts(i) - T), DrawColor
               Next i
               PIC2.Refresh
            End If
'--------------------------------------------

End Sub
'#### END FREEDRAW ###################################################


'#### LINE ###################################################

Public Sub StartLine(PIC As PictureBox, X As Single, Y As Single)
   ixTL = X: iyTL = Y
   ixBR = X: iyBR = Y
   If Spacing = 0 Then
      PIC.Line (ixTL, iyTL)-(ixBR, iyBR), vbWhite
   Else
      DrawDoubleLine PIC, vbWhite
   End If
End Sub

Public Sub DrawLine(PIC As PictureBox, X As Single, Y As Single)
   If Spacing = 0 Then
      ' Clear old line
      PIC.Line (ixTL, iyTL)-(ixBR, iyBR), vbWhite
      ixBR = X: iyBR = Y
      'Draw new line
      PIC.Line (ixTL, iyTL)-(ixBR, iyBR), vbWhite
   Else
      ' Clear old double lines
      DrawDoubleLine PIC, vbWhite
      ixBR = X: iyBR = Y
      'Draw new line
      DrawDoubleLine PIC, vbWhite
   End If
   ixwidth = ixBR - ixTL: iyheight = iyBR - iyTL
End Sub

Public Sub MoveLine(PIC As PictureBox, X As Single, Y As Single)
   If Spacing = 0 Then
      'Clear old line
      PIC.Line (ixTL, iyTL)-(ixBR, iyBR), vbWhite
      ixBR = X: iyBR = Y
      ixTL = ixBR - ixwidth: iyTL = iyBR - iyheight
      'Move line to new position
      PIC.Line (ixTL, iyTL)-(ixBR, iyBR), vbWhite
   Else
      ' Clear old double lines
      DrawDoubleLine PIC, vbWhite
      ixBR = X: iyBR = Y
      ixTL = ixBR - ixwidth: iyTL = iyBR - iyheight
      'Move line to new position
      DrawDoubleLine PIC, vbWhite
   End If
End Sub

Public Sub FinalLine(PIC As PictureBox, PIC2 As PictureBox, X As Single, Y As Single)
   If Spacing = 0 Then
      'Clear old line
      PIC.Line (ixTL, iyTL)-(ixBR, iyBR), vbWhite
      PIC.DrawMode = 13
      'ReDraw final line
      PIC.Line (ixTL, iyTL)-(ixBR, iyBR), DrawColor
'--------------------------------------------
            If aTrace Then
               'Draw final line on picFullTemp as well
               PIC2.Line (ixTL - L, iyTL - T)-(ixBR - L, iyBR - T), DrawColor
               PIC2.Refresh
            End If
'--------------------------------------------
   Else
      ' Clear old double lines
      DrawDoubleLine PIC, vbWhite
      PIC.DrawMode = 13
      'ReDraw final double lines
      DrawDoubleLine PIC, DrawColor
'--------------------------------------------
            If aTrace Then
               DrawDoubleLine PIC2, DrawColor, L, T
               PIC2.Refresh
            End If
'--------------------------------------------
   End If

End Sub

Public Sub DrawDoubleLine(PIC As PictureBox, LCul As Long, Optional LL As Long = 0, Optional TT As Long = 0)

' DRAWS A SINGLE PAIR OF DOUBLE LINES

Dim xdd As Single, ydd As Single
Dim xa As Single, ya As Single
Dim xb As Single, yb As Single
   
   'Find inner parallel line pts
   Findxeye Spacing, ixTL, iyTL, ixBR, iyBR, xdd, ydd, xa, ya, xb, yb

   PIC.Line (ixTL - LL, iyTL - TT)-(ixBR - LL, iyBR - TT), LCul
   PIC.Line (xa - LL, ya - TT)-(xb - LL, yb - TT), LCul

End Sub
'#### END LINE ###################################################


'#### RECTANGLES #################################################
Public Sub StartRectangles(PIC As PictureBox, X As Single, Y As Single)
   ixTL = X: iyTL = Y
   ixBR = X: iyBR = Y
   PIC.Line (ixTL, iyTL)-(ixBR, iyBR), vbWhite, B
   If TheDrawStyle = 7 And Spacing > 0 Then
      PIC.Line (ixTL + Spacing, iyTL + Spacing)-(ixBR - Spacing, iyBR - Spacing), vbWhite, B
   End If
End Sub

Public Sub DrawRectangles(PIC As PictureBox, X As Single, Y As Single)
   Select Case TheDrawStyle
   Case 7, 9, 10
      ' Clear old rect outline
      PIC.Line (ixTL, iyTL)-(ixBR, iyBR), vbWhite, B
      If TheDrawStyle = 7 And Spacing > 0 Then
         PIC.Line (ixTL + Spacing, iyTL + Spacing)-(ixBR - Spacing, iyBR - Spacing), vbWhite, B
      End If
      ixBR = X: iyBR = Y
      'Draw new rect outline
      PIC.Line (ixTL, iyTL)-(ixBR, iyBR), vbWhite, B
      If TheDrawStyle = 7 And Spacing > 0 Then
         PIC.Line (ixTL + Spacing, iyTL + Spacing)-(ixBR - Spacing, iyBR - Spacing), vbWhite, B
      End If
      ixwidth = ixBR - ixTL: iyheight = iyBR - iyTL
   Case 8
      ' Clear old rect filled
      PIC.Line (ixTL, iyTL)-(ixBR, iyBR), vbWhite, BF
      ixBR = X: iyBR = Y
      'Draw new rect filled
      PIC.Line (ixTL, iyTL)-(ixBR, iyBR), vbWhite, BF
      ixwidth = ixBR - ixTL: iyheight = iyBR - iyTL
   End Select
End Sub

Public Sub MoveRectangles(PIC As PictureBox, X As Single, Y As Single)
   Select Case TheDrawStyle
   Case 7, 9, 10
      ' Clear old rect outline
      PIC.Line (ixTL, iyTL)-(ixBR, iyBR), vbWhite, B
      If TheDrawStyle = 7 And Spacing > 0 Then
         PIC.Line (ixTL + Spacing, iyTL + Spacing)-(ixBR - Spacing, iyBR - Spacing), vbWhite, B
      End If
      ixBR = X: iyBR = Y
      ixTL = ixBR - ixwidth: iyTL = iyBR - iyheight
      'Draw new rect outline
      PIC.Line (ixTL, iyTL)-(ixBR, iyBR), vbWhite, B
      If TheDrawStyle = 7 And Spacing > 0 Then
         PIC.Line (ixTL + Spacing, iyTL + Spacing)-(ixBR - Spacing, iyBR - Spacing), vbWhite, B
      End If
      ixwidth = ixBR - ixTL: iyheight = iyBR - iyTL
   Case 8
      ' Clear old rect filled
      PIC.Line (ixTL, iyTL)-(ixBR, iyBR), vbWhite, BF
      ixBR = X: iyBR = Y
      ixTL = ixBR - ixwidth: iyTL = iyBR - iyheight
      'Draw new rect filled
      PIC.Line (ixTL, iyTL)-(ixBR, iyBR), vbWhite, BF
      ixwidth = ixBR - ixTL: iyheight = iyBR - iyTL
   End Select
End Sub

Public Sub FinalRectangles(PIC As PictureBox, PIC2 As PictureBox, X As Single, Y As Single)
   Select Case TheDrawStyle
   Case 7
      ' Clear old rect outline
      PIC.Line (ixTL, iyTL)-(ixBR, iyBR), vbWhite, B
      If TheDrawStyle = 7 And Spacing > 0 Then
         PIC.Line (ixTL + Spacing, iyTL + Spacing)-(ixBR - Spacing, iyBR - Spacing), vbWhite, B
      End If
      PIC.DrawMode = 13
      ' Draw final rect outline
      PIC.Line (ixTL, iyTL)-(ixBR, iyBR), DrawColor, B
      If TheDrawStyle = 7 And Spacing > 0 Then
         PIC.Line (ixTL + Spacing, iyTL + Spacing)-(ixBR - Spacing, iyBR - Spacing), DrawColor, B
      End If
'--------------------------------------------
            If aTrace Then
               'Draw final rectangle on picFullTemp as well
               PIC2.Line (ixTL - L, iyTL - T)-(ixBR - L, iyBR - T), DrawColor, B
               PIC2.Refresh
            End If
'--------------------------------------------
   Case 8
      ' Clear old rect filled
      PIC.Line (ixTL, iyTL)-(ixBR, iyBR), vbWhite, BF
      PIC.DrawMode = 13
      'Draw final rect filled
      PIC.Line (ixTL, iyTL)-(ixBR, iyBR), DrawColor, BF
'--------------------------------------------
            If aTrace Then
               'Draw final rectangle on picFullTemp as well
               PIC2.Line (ixTL - L, iyTL - T)-(ixBR - L, iyBR - T), DrawColor, BF
               PIC2.Refresh
            End If
'--------------------------------------------
   Case 9   ' Horz Shaded rect
      ' Clear old rect outline
      PIC.Line (ixTL, iyTL)-(ixBR, iyBR), vbWhite, B
      PIC.DrawMode = 13
      ' Draw Final Horz shaded rect
      DrawHorzShadedRectangle PIC
'--------------------------------------------
            If aTrace Then
               DrawHorzShadedRectangle PIC2, L, T
               PIC2.Refresh
            End If
'--------------------------------------------
   Case 10  ' Vert Shaded rect
      ' Clear old rect outline
      PIC.Line (ixTL, iyTL)-(ixBR, iyBR), vbWhite, B
      PIC.DrawMode = 13
      ' Draw  Final Vert shaded rect
      DrawVertShadedRectangle PIC
'--------------------------------------------
            If aTrace Then
               DrawVertShadedRectangle PIC2, L, T
               PIC2.Refresh
            End If
'--------------------------------------------
   End Select
End Sub

Private Sub DrawHorzShadedRectangle(PIC As PictureBox, Optional LL As Long = 0, Optional TT As Long = 0)
Dim zstepy As Single
Dim xx As Single
Dim yy As Single
Dim zz As Single
   
   ' Shaded rectangle  LL & TT are X & Y offsets for Tracing
   bR = (DrawColor And &HFF&)
   bG = (DrawColor And &HFF00&) / &H100&
   bB = (DrawColor And &HFF0000) / &H10000
   zstepy = (iyBR - iyTL) / 128
   If Abs(zstepy) > 1 Then zstepy = 1 * Sgn(zstepy)
   xx = iyTL: yy = iyBR
   If Abs(zstepy) > 0 Then
      For zz = xx To yy Step zstepy
         PIC.Line (ixTL - LL, zz - TT)-(ixBR - LL, zz - TT), RGB(bR, bG, bB)
         If bR > 0 Then bR = bR - 1
         If bG > 0 Then bG = bG - 1
         If bB > 0 Then bB = bB - 1
      Next zz
   Else
      PIC.Line (ixTL - LL, iyTL - TT)-(ixBR - LL, iyBR - TT), DrawColor, B
   End If

End Sub

Private Sub DrawVertShadedRectangle(PIC As PictureBox, Optional LL As Long = 0, Optional TT As Long = 0)
Dim zstepx As Single
Dim xx As Single
Dim yy As Single
Dim zz As Single
   
   ' Shaded rectangle  LL & TT are X & Y offsets for Tracing
   bR = (DrawColor And &HFF&)
   bG = (DrawColor And &HFF00&) / &H100&
   bB = (DrawColor And &HFF0000) / &H10000
   zstepx = (ixBR - ixTL) / 128
   If Abs(zstepx) > 1 Then zstepx = 1 * Sgn(zstepx)
   xx = ixTL: yy = ixBR
   If Abs(zstepx) > 0 Then
      For zz = xx To yy Step zstepx
         PIC.Line (zz - LL, iyTL - TT)-(zz - LL, iyBR - TT), RGB(bR, bG, bB)
         If bR > 0 Then bR = bR - 1
         If bG > 0 Then bG = bG - 1
         If bB > 0 Then bB = bB - 1
      Next zz
   Else
      PIC.Line (ixTL - LL, iyTL - TT)-(ixBR - LL, iyBR - TT), DrawColor, B
   End If

End Sub
'#### END RECTANGLES #################################################


'#### CIRLLIPSES ################################################

Public Sub StartCirllipses(PIC As PictureBox, X As Single, Y As Single)
   ixTL = X: iyTL = Y
   ixBR = X: iyBR = Y
   Eval_Cirllipse_Params
   PIC.Circle (ixc, iyc), zradius, vbWhite, , , zaspect
   If TheDrawStyle = 11 And Spacing > 0 And Spacing < zradius Then
      PIC.Circle (ixc, iyc), zradius - Spacing, vbWhite, , , zaspect
   End If
End Sub

Public Sub DrawCirllipses(PIC As PictureBox, X As Single, Y As Single)
   Eval_Cirllipse_Params
   PIC.Circle (ixc, iyc), zradius, vbWhite, , , zaspect
   If TheDrawStyle = 11 And Spacing > 0 And Spacing < zradius Then
      PIC.Circle (ixc, iyc), zradius - Spacing, vbWhite, , , zaspect
   End If
   ixBR = X: iyBR = Y
   ' Draw new cirllipses & shaded outline
   Eval_Cirllipse_Params
   PIC.Circle (ixc, iyc), zradius, vbWhite, , , zaspect
   If TheDrawStyle = 11 And Spacing > 0 And Spacing < zradius Then
      PIC.Circle (ixc, iyc), zradius - Spacing, vbWhite, , , zaspect
   End If
   ixwidth = ixBR - ixTL: iyheight = iyBR - iyTL
End Sub

Public Sub MoveCirllipses(PIC As PictureBox, X As Single, Y As Single)
   ' Clear old cirllipses & shaded outline
   Eval_Cirllipse_Params
   PIC.Circle (ixc, iyc), zradius, vbWhite, , , zaspect
   If TheDrawStyle = 11 And Spacing > 0 And Spacing < zradius Then
      PIC.Circle (ixc, iyc), zradius - Spacing, vbWhite, , , zaspect
   End If
   ixBR = X: iyBR = Y
   ixTL = ixBR - ixwidth: iyTL = iyBR - iyheight
   ' Move cirllipses to new position & shaded outline
   Eval_Cirllipse_Params
   PIC.Circle (ixc, iyc), zradius, vbWhite, , , zaspect
   If TheDrawStyle = 11 And Spacing > 0 And Spacing < zradius Then
      PIC.Circle (ixc, iyc), zradius - Spacing, vbWhite, , , zaspect
   End If
   ixwidth = ixBR - ixTL: iyheight = iyBR - iyTL
End Sub

Public Sub FinalCirllipses(PIC As PictureBox, PIC2 As PictureBox, X As Single, Y As Single)
   Eval_Cirllipse_Params
   ' Clear old cirllipse
   PIC.Circle (ixc, iyc), zradius, vbWhite, , , zaspect
   If TheDrawStyle = 11 And Spacing > 0 And Spacing < zradius Then
      PIC.Circle (ixc, iyc), zradius - Spacing, vbWhite, , , zaspect
   End If
   PIC.DrawMode = 13
   
   'Case 11, 12, 13  ' Outline, Filled, Shaded Cirllipses
   Select Case TheDrawStyle
   Case 11, 13 ' Outline & Shaded Cirllipses
      If TheDrawStyle = 11 Then
         ' Draw final cirllipse
         PIC.Circle (ixc, iyc), zradius, DrawColor, , , zaspect
         If TheDrawStyle = 11 And Spacing > 0 And Spacing < zradius Then
            PIC.Circle (ixc, iyc), zradius - Spacing, DrawColor, , , zaspect
         End If
'--------------------------------------------
               If aTrace Then
                  'Draw cirllipse on picFullTemp as well
                  PIC2.Circle (ixc - L, iyc - T), zradius, DrawColor, , , zaspect
                  PIC2.Refresh
               End If
'--------------------------------------------
      
      Else  ' Shaded cirllipse
         DrawShadedCirllipse PIC
'--------------------------------------------
               If aTrace Then
                  'Draw cirllipse on picFullTemp as well
                  DrawShadedCirllipse PIC2, L, T
                  PIC2.Refresh
               End If
'--------------------------------------------
      End If
      PIC.Refresh
   Case 12     ' Filled Cirllipse
      ' Draw final filled cirllipse
      PIC.FillColor = DrawColor
      PIC.FillStyle = 0   ' Solid
      PIC.Circle (ixc, iyc), zradius, DrawColor, , , zaspect
      PIC.FillStyle = 1   ' Transparent
      PIC.Refresh
'--------------------------------------------
               If aTrace Then
                  ' Draw final filled cirllipse on picFullTemp as well
                  PIC2.FillColor = DrawColor
                  PIC2.FillStyle = 0   ' Solid
                  PIC2.Circle (ixc - L, iyc - T), zradius, DrawColor, , , zaspect
                  PIC2.FillStyle = 1   ' Transparent
                  PIC2.Refresh
               End If
   End Select
End Sub

Public Sub Eval_Cirllipse_Params()
Dim zradx As Single
Dim zrady As Single
'Public ixc, iyc, zradius, zaspect

   ixc = (ixTL + ixBR) \ 2
   iyc = (iyTL + iyBR) \ 2
   zradx = Abs(ixBR - ixTL) / 2 'Abs(X - xs)
   zrady = Abs(iyBR - iyTL) / 2 'Abs(Y - ys)
   If zradx = 0 Then
      zradius = zrady
      zaspect = 10
   ElseIf zradx >= zrady Then
      zradius = zradx
      zaspect = zrady / zradx
   Else  'zradx<zrady
      zradius = zrady
      zaspect = zrady / zradx
   End If

End Sub

Private Sub DrawShadedCirllipse(PIC As PictureBox, Optional LL As Long = 0, Optional TT As Long = 0)
Dim zstepy As Single
Dim yy As Single
   
   svDrawWidth = TheDrawWidth
   PIC.DrawWidth = 2
   
   ' Shaded circle LL & TT are X & Y offsets offsets for Tracing
   bR = (DrawColor And &HFF&)
   bG = (DrawColor And &HFF00&) / &H100&
   bB = (DrawColor And &HFF0000) / &H10000
   zstepy = zradius / 128
   If zstepy > 0 Then
      For yy = 1 To zradius Step zstepy
         PIC.Circle (ixc - LL, iyc - TT), yy, RGB(bR, bG, bB), , , zaspect
         If bR > 0 Then bR = bR - 1
         If bG > 0 Then bG = bG - 1
         If bB > 0 Then bB = bB - 1
      Next yy
   Else
      PIC.Circle (ixc - LL, iyc - TT), zradius, DrawColor, , , zaspect
   End If

   TheDrawWidth = svDrawWidth
End Sub
'#### END CIRLLIPSES ################################################

'#### ARCH ###################################################

Public Sub StartArch(PIC As PictureBox, X As Single, Y As Single)
   ixTL = X: iyTL = Y
   ixBR = X: iyBR = Y
   DrawWholeArch PIC, vbWhite
End Sub

Public Sub DrawArch(PIC As PictureBox, X As Single, Y As Single)
   ' Clear old
   DrawWholeArch PIC, vbWhite
   ixBR = X: iyBR = Y
   'Draw new
   DrawWholeArch PIC, vbWhite
   ixwidth = ixBR - ixTL: iyheight = iyBR - iyTL
End Sub

Public Sub MoveArch(PIC As PictureBox, X As Single, Y As Single)
   'Clear old
   DrawWholeArch PIC, vbWhite
   ixBR = X
   iyBR = Y
   ixTL = ixBR - ixwidth: iyTL = iyBR - iyheight
   'Move to new position
   DrawWholeArch PIC, vbWhite
End Sub

Public Sub FinalArch(PIC As PictureBox, PIC2 As PictureBox, X As Single, Y As Single)
   'Clear old
   DrawWholeArch PIC, vbWhite
   PIC.DrawMode = 13
   'ReDraw final
   DrawWholeArch PIC, DrawColor
'--------------------------------------------
               ' To picFullTemp as well
               If aTrace Then
                     DrawWholeArch PIC2, DrawColor, L, T
                     PIC2.Refresh
               End If
'--------------------------------------------
               
End Sub

Private Sub DrawWholeArch(PIC As PictureBox, ArchColor As Long, Optional LL As Long = 0, Optional TT As Long = 0)
Dim xc As Single
   
   PIC.Line (ixTL - LL, iyTL - TT)-(ixTL - LL, iyBR - TT), ArchColor 'LV
   If Spacing > 0 Then
      PIC.Line (ixTL + Spacing - LL, iyTL - TT)-(ixTL + Spacing - LL, iyBR - Spacing - TT), ArchColor 'LV
   End If
   PIC.Line (ixTL - LL, iyBR - TT)-(ixBR - LL, iyBR - TT), ArchColor  'BH
   If Spacing > 0 Then
      PIC.Line (ixTL + Spacing - LL, iyBR - Spacing - TT)-(ixBR - Spacing - LL, iyBR - Spacing - TT), ArchColor 'BH
   End If
   PIC.Line (ixBR - LL, iyBR - TT)-(ixBR - LL, iyTL - TT), ArchColor  'RV
   If Spacing > 0 Then
      PIC.Line (ixBR - Spacing - LL, iyBR - Spacing - TT)-(ixBR - Spacing - LL, iyTL - TT), ArchColor 'RV
   End If
   xc = (ixTL + ixBR) / 2
   zradius = Abs(ixBR - ixTL) / 2
   PIC.Circle (xc - LL, iyTL - TT), zradius, ArchColor, 0, pi#
   If Spacing > 0 And Spacing < zradius Then
      PIC.Circle (xc - LL, iyTL - TT), zradius - Spacing, ArchColor, 0, pi#
   End If
End Sub
'#### END ARCH ###################################################


'### F-RIBBON /// ###################################################

Public Sub StartFRibbon(PIC As PictureBox, X As Single, Y As Single)
   ixTL = X: iyTL = Y
   ixBR = X: iyBR = Y
   ReDim ixPts(1 To 1), iyPts(1 To 1)
   NumOfDrawPoints = 1
   ixPts(1) = ixTL: iyPts(1) = iyTL
   j = SRSize \ 2
   PIC.Line (ixPts(1) - j, iyPts(1) + j) _
    -(ixPts(1) + j, iyPts(1) - j), vbWhite
End Sub

Public Sub DrawFRibbon(PIC As PictureBox, X As Single, Y As Single)
   NumOfDrawPoints = NumOfDrawPoints + 1
   ReDim Preserve ixPts(1 To NumOfDrawPoints)
   ReDim Preserve iyPts(1 To NumOfDrawPoints)
   ixTL = X
   iyTL = Y
   ixPts(NumOfDrawPoints) = ixTL
   iyPts(NumOfDrawPoints) = iyTL
   j = SRSize \ 2
   PIC.Line (ixTL - j, iyTL + j) _
     -(ixTL + j, iyTL - j), vbWhite
End Sub

Public Sub MoveFRibbon(PIC As PictureBox, X As Single, Y As Single)
   'Clear old  / F-ribbon
   j = SRSize \ 2
   For i = 1 To NumOfDrawPoints
      PIC.Line (ixPts(i) - j, iyPts(i) + j) _
        -(ixPts(i) + j, iyPts(i) - j), vbWhite
   Next i
   ' New ribbon points
   For i = 1 To NumOfDrawPoints
      ixPts(i) = ixPts(i) + (X - ixTL)
      iyPts(i) = iyPts(i) + (Y - iyTL)
   Next i
   ixTL = X
   iyTL = Y
   ' Move ribbon to new position
   For i = 1 To NumOfDrawPoints
      PIC.Line (ixPts(i) - j, iyPts(i) + j) _
        -(ixPts(i) + j, iyPts(i) - j), vbWhite
   Next i
   PIC.Refresh

End Sub

Public Sub FinalFRibbon(PIC As PictureBox, PIC2 As PictureBox, X As Single, Y As Single)
' For final ribbons
Dim zdx As Single
Dim zdy As Single
Dim zlines As Single
Dim zstepx As Single
Dim zstepy As Single
Dim xx As Single
Dim yy As Single
Dim zz As Single
   
   'Clear old  / F-ribbon
   j = SRSize \ 2
   For i = 1 To NumOfDrawPoints
      PIC.Line (ixPts(i) - j, iyPts(i) + j) _
      -(ixPts(i) + j, iyPts(i) - j), vbWhite
   Next i
   
   PIC.DrawMode = 13
   
   'ReDraw solid final  / F-ribbon
   For i = 1 To NumOfDrawPoints - 1
      zdx = ixPts(i + 1) - ixPts(i)
      zdy = iyPts(i + 1) - iyPts(i)
      zlines = Abs(zdx)
      If zlines < Abs(zdy) Then zlines = Abs(zdy)
      zstepx = zdx / (zlines + 1)
      zstepy = zdy / (zlines + 1)
      xx = ixPts(i)
      yy = iyPts(i)
      For zz = 0 To zlines + 1
         PIC.Line (xx - j, yy + j) _
         -(xx + j, yy - j), DrawColor
         
'--------------------------------------------
               ' To picFullTemp as well
               If aTrace Then
                  PIC2.Line (xx - j - L, yy + j - T) _
                  -(xx + j - L, yy - j - T), DrawColor
                  PIC2.Refresh
               End If
'--------------------------------------------
         
         xx = xx + zstepx
         yy = yy + zstepy
      Next zz
   Next i
   
   PIC.Refresh

End Sub
'### END F-RIBBON /// ###################################################

'### B-RIBBON \\\ ###################################################

Public Sub StartBRibbon(PIC As PictureBox, X As Single, Y As Single)
   ixTL = X: iyTL = Y
   ixBR = X: iyBR = Y
   ReDim ixPts(1 To 1), iyPts(1 To 1)
   NumOfDrawPoints = 1
   ixPts(1) = ixTL: iyPts(1) = iyTL
   j = SRSize \ 2
   PIC.Line (ixPts(1) - j, iyPts(1) - j) _
    -(ixPts(1) + j, iyPts(1) + j), vbWhite
End Sub

Public Sub DrawBRibbon(PIC As PictureBox, X As Single, Y As Single)
   NumOfDrawPoints = NumOfDrawPoints + 1
   ReDim Preserve ixPts(1 To NumOfDrawPoints)
   ReDim Preserve iyPts(1 To NumOfDrawPoints)
   ixTL = X
   iyTL = Y
   ixPts(NumOfDrawPoints) = ixTL
   iyPts(NumOfDrawPoints) = iyTL
   j = SRSize \ 2
   PIC.Line (ixTL - j, iyTL - j) _
      -(ixTL + j, iyTL + j), vbWhite
End Sub

Public Sub MoveBRibbon(PIC As PictureBox, X As Single, Y As Single)
   j = SRSize \ 2
   'Clear old  \ B-ribbon
   For i = 1 To NumOfDrawPoints
      PIC.Line (ixPts(i) - j, iyPts(i) - j) _
        -(ixPts(i) + j, iyPts(i) + j), vbWhite
   Next i
   ' Get new points
   For i = 1 To NumOfDrawPoints
      ixPts(i) = ixPts(i) + (X - ixTL)
      iyPts(i) = iyPts(i) + (Y - iyTL)
   Next i
   ixTL = X
   iyTL = Y
   For i = 1 To NumOfDrawPoints
      PIC.Line (ixPts(i) - j, iyPts(i) - j) _
        -(ixPts(i) + j, iyPts(i) + j), vbWhite
   Next i
   PIC.Refresh

End Sub

Public Sub FinalBRibbon(PIC As PictureBox, PIC2 As PictureBox, X As Single, Y As Single)
' For final ribbons
Dim zdx As Single
Dim zdy As Single
Dim zlines As Single
Dim zstepx As Single
Dim zstepy As Single
Dim xx As Single
Dim yy As Single
Dim zz As Single


   j = SRSize \ 2
   'Clear old  \ B-ribbon
   For i = 1 To NumOfDrawPoints
      PIC.Line (ixPts(i) - j, iyPts(i) - j) _
      -(ixPts(i) + j, iyPts(i) + j), vbWhite
   Next i
   PIC.DrawMode = 13
   
   'ReDraw final  \ B-ribbon
   For i = 1 To NumOfDrawPoints - 1
      zdx = ixPts(i + 1) - ixPts(i)
      zdy = iyPts(i + 1) - iyPts(i)
      zlines = Abs(zdx)
      If zlines < Abs(zdy) Then zlines = Abs(zdy)
      zstepx = zdx / (zlines + 1)
      zstepy = zdy / (zlines + 1)
      xx = ixPts(i)
      yy = iyPts(i)
      For zz = 0 To zlines + 1
         PIC.Line (xx - j, yy - j) _
         -(xx + j, yy + j), DrawColor
         
'--------------------------------------------
               ' To picFullTemp as well
               If aTrace Then
                  PIC2.Line (xx - j - L, yy - j - T) _
                  -(xx + j - L, yy + j - T), DrawColor
                  PIC2.Refresh
               End If
'--------------------------------------------
         
         xx = xx + zstepx
         yy = yy + zstepy
      Next zz
   Next i
   
   PIC.Refresh

End Sub
'### END F-RIBBON /// ###################################################


'#### POLYLINE ######################################################

Public Sub StartPolyLine(PIC As PictureBox, X As Single, Y As Single)
   ixTL = X: iyTL = Y
   ixBR = ixTL: iyBR = iyTL
   ReDim ixPts(1 To 1), iyPts(1 To 1)
   NumOfDrawPoints = 1
   ixPts(1) = ixTL: iyPts(1) = iyTL
   PIC.Line (ixTL, iyTL)-(ixBR, iyBR), vbWhite
   
End Sub

Public Sub DrawPolyLine(PIC As PictureBox, X As Single, Y As Single)
   'Clear old polyline
   PIC.Line (ixTL, iyTL)-(ixBR, iyBR), vbWhite
   
   ixBR = X: iyBR = Y
   'Draw new polylineline
   PIC.Line (ixTL, iyTL)-(ixBR, iyBR), vbWhite
End Sub

Public Sub StartNextPolyLine(PIC As PictureBox, X As Single, Y As Single)
   
'   If X <> ixPts(NumOfDrawPoints) Or Y <> iyPts(NumOfDrawPoints) Then
   NumOfDrawPoints = NumOfDrawPoints + 1
   ReDim Preserve ixPts(1 To NumOfDrawPoints)
   ReDim Preserve iyPts(1 To NumOfDrawPoints)
   ixPts(NumOfDrawPoints) = X
   iyPts(NumOfDrawPoints) = Y
'   End If

   ixTL = X: iyTL = Y
   ixBR = X: iyBR = Y
   PIC.Line (ixTL, iyTL)-(ixBR, iyBR), vbWhite
End Sub

Public Sub UpdatePolyPoints(X As Single, Y As Single)
   NumOfDrawPoints = NumOfDrawPoints + 1
   ReDim Preserve ixPts(1 To NumOfDrawPoints)
   ReDim Preserve iyPts(1 To NumOfDrawPoints)
   ixPts(NumOfDrawPoints) = X
   iyPts(NumOfDrawPoints) = Y
   ixTL = X
   iyTL = Y
End Sub


Public Sub MovePolyLine(PIC As PictureBox, X As Single, Y As Single)
   'Clear old polylines
   For i = 1 To NumOfDrawPoints - 1
      PIC.Line (ixPts(i), iyPts(i)) _
      -(ixPts(i + 1), iyPts(i + 1)), vbWhite
   Next i
   
   ' Move PolyLines to new position
   For i = 1 To NumOfDrawPoints
      ixPts(i) = ixPts(i) + (X - ixTL)
      iyPts(i) = iyPts(i) + (Y - iyTL)
   Next i
   ixTL = X
   iyTL = Y
   ' ReDraw Polylines
   For i = 1 To NumOfDrawPoints - 1
      PIC.Line (ixPts(i), iyPts(i)) _
      -(ixPts(i + 1), iyPts(i + 1)), vbWhite
   Next i
   
End Sub


Public Sub FinalPolyLine(PIC As PictureBox, PIC2 As PictureBox, X As Single, Y As Single)
   'Clear old polylines
   For i = 1 To NumOfDrawPoints - 1
      PIC.Line (ixPts(i), iyPts(i)) _
      -(ixPts(i + 1), iyPts(i + 1)), vbWhite
   Next i
   
   PIC.DrawMode = 13
   ' ReDraw final Polyline
   If Spacing = 0 Then
      For i = 1 To NumOfDrawPoints - 1
         PIC.Line (ixPts(i), iyPts(i)) _
         -(ixPts(i + 1), iyPts(i + 1)), DrawColor
'--------------------------------------------
               ' To picFullTemp as well
               If aTrace Then
                  PIC2.Line (ixPts(i) - L, iyPts(i) - T) _
                  -(ixPts(i + 1) - L, iyPts(i + 1) - T), DrawColor
                  PIC2.Refresh
               End If
'--------------------------------------------
      Next i
   Else
      DrawDoubleLines PIC, DrawColor
'--------------------------------------------
               ' To picFullTemp as well
               If aTrace Then
                  DrawDoubleLines PIC2, DrawColor, L, T
                  PIC2.Refresh
               End If
'--------------------------------------------
   End If
   
   PIC.Refresh

End Sub
'#### END POLYLINE ######################################################

'#### SPLINE ######################################################

Public Sub StartSpline(PIC As PictureBox, X As Single, Y As Single)
   ixTL = X: iyTL = Y
   ixBR = ixTL: iyBR = iyTL
   ReDim ixPts(1 To 1), iyPts(1 To 1)
   NumOfDrawPoints = 1
   ixPts(1) = ixTL: iyPts(1) = iyTL
   PIC.Line (ixTL, iyTL)-(ixBR, iyBR), vbWhite
End Sub

Public Sub DrawSpline(PIC As PictureBox, X As Single, Y As Single)
   'Clear old polyline
   PIC.Line (ixTL, iyTL)-(ixBR, iyBR), vbWhite
   ixBR = X: iyBR = Y
   'Draw new polylineline
   PIC.Line (ixTL, iyTL)-(ixBR, iyBR), vbWhite
End Sub

Public Sub MoveSpline(PIC As PictureBox, X As Single, Y As Single)
   'Clear old polylines
   For i = 1 To NumOfDrawPoints - 1
      PIC.Line (ixPts(i), iyPts(i)) _
      -(ixPts(i + 1), iyPts(i + 1)), vbWhite
   Next i
   ' Move PolyLines to new position
   For i = 1 To NumOfDrawPoints
      ixPts(i) = ixPts(i) + (X - ixTL)
      iyPts(i) = iyPts(i) + (Y - iyTL)
   Next i
   ixTL = X
   iyTL = Y
   ' ReDraw Polylines
   For i = 1 To NumOfDrawPoints - 1
      PIC.Line (ixPts(i), iyPts(i)) _
      -(ixPts(i + 1), iyPts(i + 1)), vbWhite
   Next i
   PIC.Refresh

End Sub

Public Sub FinalSpline(PIC As PictureBox, PIC2 As PictureBox, X As Single, Y As Single)
   'Clear old polylines
   For i = 1 To NumOfDrawPoints - 1
      PIC.Line (ixPts(i), iyPts(i)) _
      -(ixPts(i + 1), iyPts(i + 1)), vbWhite
   Next i
   
   PIC.DrawMode = 13
   MakeSplinePoints
      
   ' ReDraw final Spline
   If Spacing = 0 Then
      For i = 1 To NumOfDrawPoints - 1
         PIC.Line (ixPts(i), iyPts(i)) _
         -(ixPts(i + 1), iyPts(i + 1)), DrawColor
'--------------------------------------------
               ' To picFullTemp as well
               If aTrace Then
                  PIC2.Line (ixPts(i) - L, iyPts(i) - T) _
                  -(ixPts(i + 1) - L, iyPts(i + 1) - T), DrawColor
                  PIC2.Refresh
               End If
'--------------------------------------------
      Next i
   Else
      DrawDoubleLines PIC, DrawColor
'--------------------------------------------
               ' To picFullTemp as well
               If aTrace Then
                  DrawDoubleLines PIC2, DrawColor, L, T
                  PIC2.Refresh
               End If
'--------------------------------------------
   End If
   
   PIC.Refresh

End Sub

Public Sub MakeSplinePoints()

' Public ixPts() As Integer  ' Ribbon coords
' Public iyPts() As Integer
' Public NumOfDrawPoints As Long
' Public i

' In: ixPts(NumOfDrawPoints), iyPts(NumOfDrawPoints)
'Out: New ixPts(),iyPts(),NumOfDrawPoints

Dim xaa() As Single, yaa() As Single
Dim xbb() As Single, ybb() As Single
Dim xfrac As Single
Dim xdx As Single, ydy As Single
Dim S As Integer, SUP As Integer
Dim oldpts As Long, newpts As Long


If NumOfDrawPoints <= 2 Then Exit Sub ' Single point or line
'-----------------------------------------------------
' NumOfDrawPoints > 2

   ReDim xbb(1 To NumOfDrawPoints)
   ReDim ybb(1 To NumOfDrawPoints)
   
   For i = 1 To NumOfDrawPoints
      xbb(i) = ixPts(i)
      ybb(i) = iyPts(i)
   Next i
   
   ' Develop Bezier-like points
   xfrac = 0.25
   SUP = 3
   oldpts = NumOfDrawPoints
   
   For S = 1 To SUP
      ReDim xaa(oldpts), yaa(oldpts), zaa(oldpts)
      For i = 1 To oldpts
         xaa(i) = xbb(i): yaa(i) = ybb(i)
      Next i
      newpts = 2 * oldpts - 2
      ReDim xbb(newpts), ybb(newpts)
      xbb(1) = xaa(1): ybb(1) = yaa(1)
      For i = 2 To oldpts - 1
         xdx = xaa(i) - xaa(i - 1)
         xbb(2 * i - 2) = xaa(i) - xfrac * xdx
         ydy = yaa(i) - yaa(i - 1)
         ybb(2 * i - 2) = yaa(i) - xfrac * ydy
         xdx = xaa(i + 1) - xaa(i)
         xbb(2 * i - 1) = xaa(i) + xfrac * xdx
         ydy = yaa(i + 1) - yaa(i)
         ybb(2 * i - 1) = yaa(i) + xfrac * ydy
      Next i
      xbb(newpts) = xaa(oldpts): ybb(newpts) = yaa(oldpts)
      oldpts = newpts
   Next S
   '-----------------------------------------------------

   Erase xaa(), yaa()
   
   NumOfDrawPoints = newpts
   ReDim ixPts(1 To NumOfDrawPoints)
   ReDim iyPts(1 To NumOfDrawPoints)
   For i = 1 To NumOfDrawPoints
      ixPts(i) = xbb(i)
      iyPts(i) = ybb(i)
   Next i
   
   Erase xbb(), ybb()

End Sub
'#### END SPLINE ######################################################

'#### SPRAY ###################################################
Public Sub StartSpray(PIC As PictureBox, X As Single, Y As Single)
   NumOfDrawPoints = 1
   ReDim ixPts(1 To 1), iyPts(1 To 1)
   ixPts(1) = X
   iyPts(1) = Y
   PIC.PSet (X, Y), vbWhite
   PIC.Refresh
End Sub

Public Sub DrawSpray(PIC As PictureBox, X As Single, Y As Single)
   ixTL = X
   iyTL = Y
   For j = Y - SRSize To Y + SRSize
   For i = X - SRSize To X + SRSize
      If Rnd < 0.025 Then
         PIC.PSet (i, j), vbWhite
          
         NumOfDrawPoints = NumOfDrawPoints + 1
         ReDim Preserve ixPts(1 To NumOfDrawPoints)
         ReDim Preserve iyPts(1 To NumOfDrawPoints)
         ixPts(NumOfDrawPoints) = i
         iyPts(NumOfDrawPoints) = j
      End If
   Next i
   Next j
End Sub

Public Sub MoveSpray(PIC As PictureBox, X As Single, Y As Single)
   'Clear old points
   PIC.PSet (ixPts(1), iyPts(1)), vbWhite
   For i = 2 To NumOfDrawPoints
      PIC.PSet (ixPts(i), iyPts(i)), vbWhite
   Next i
   ' Move points to new position
   For i = 1 To NumOfDrawPoints
      ixPts(i) = ixPts(i) + (X - ixTL)
      iyPts(i) = iyPts(i) + (Y - iyTL)
   Next i
   ixTL = X
   iyTL = Y
   PIC.PSet (ixPts(1), iyPts(1)), vbWhite
   For i = 2 To NumOfDrawPoints
      PIC.PSet (ixPts(i), iyPts(i)), vbWhite
   Next i
   PIC.Refresh
End Sub

Public Sub FinalSpray(PIC As PictureBox, PIC2 As PictureBox, X As Single, Y As Single)
   'Clear old points
   PIC.PSet (ixPts(1), iyPts(1)), vbWhite
   For i = 2 To NumOfDrawPoints
      PIC.PSet (ixPts(i), iyPts(i)), vbWhite
   Next i
   PIC.DrawMode = 13
   
   ' Draw Final points
   PIC.PSet (ixPts(1), iyPts(1)), DrawColor
   For i = 2 To NumOfDrawPoints
      PIC.PSet (ixPts(i), iyPts(i)), DrawColor
'--------------------------------------------
               ' To picFullTemp as well
               If aTrace Then
                  PIC2.PSet (ixPts(i) - L, iyPts(i) - T), DrawColor
                  PIC2.Refresh
               End If
'--------------------------------------------
   Next i
End Sub
'#### END SPRAY ###################################################

'#### STAR ###################################################
Public Sub StartStar(PIC As PictureBox, X As Single, Y As Single)
   ixTL = X: iyTL = Y
   ixBR = X: iyBR = Y
   DrawWholeStar PIC, vbWhite
End Sub

Public Sub DrawStar(PIC As PictureBox, X As Single, Y As Single)
   ' Clear old line
   DrawWholeStar PIC, vbWhite
   'ixBR = X:
   iyBR = Y
   'Draw new line
   DrawWholeStar PIC, vbWhite
   ixwidth = ixBR - ixTL: iyheight = iyBR - iyTL
End Sub

Public Sub MoveStar(PIC As PictureBox, X As Single, Y As Single)
   'Clear old line
   DrawWholeStar PIC, vbWhite
   ixBR = X:
   iyBR = Y
   ixTL = ixBR - ixwidth: iyTL = iyBR - iyheight
   'Move line to new position
   DrawWholeStar PIC, vbWhite
End Sub

Public Sub FinalStar(PIC As PictureBox, PIC2 As PictureBox, X As Single, Y As Single)
   'Clear old line
   DrawWholeStar PIC, vbWhite
   PIC.DrawMode = 13
   'ReDraw final line
   DrawWholeStar PIC, DrawColor
'--------------------------------------------
            If aTrace Then
               DrawWholeStar PIC2, DrawColor, L, T
               PIC2.Refresh
            End If
'--------------------------------------------
End Sub

Private Sub DrawWholeStar(PIC As PictureBox, StarColor As Long, Optional LL As Long = 0, Optional TT As Long = 0)
Dim IY As Long
Dim ztheta As Single
Dim zz As Single
   
   ' Final star LL & TT are X & Y offsets
   ztheta = 36 * pi# / 180
   ReDim ixPts(0 To 9)
   ReDim iyPts(0 To 9)
   zz = Sqr((ixBR - ixTL) ^ 2 + (iyBR - iyTL) ^ 2)
   IY = iyTL + 2 * zz
   ixPts(0) = ixTL
   iyPts(0) = IY - 2 * zz
   ixPts(1) = ixTL + zz * Sin(ztheta)
   iyPts(1) = IY - zz * Cos(ztheta)
   ixPts(9) = ixTL - zz * Sin(ztheta)
   iyPts(9) = IY - zz * Cos(ztheta)

   ixPts(2) = ixTL + 2 * zz * Sin(2 * ztheta)
   iyPts(2) = IY - 2 * zz * Cos(2 * ztheta)
   ixPts(8) = ixTL - 2 * zz * Sin(2 * ztheta)
   iyPts(8) = IY - 2 * zz * Cos(2 * ztheta)

   ixPts(3) = ixTL + zz * Sin(3 * ztheta)
   iyPts(3) = IY - zz * Cos(3 * ztheta)
   ixPts(7) = ixTL - zz * Sin(3 * ztheta)
   iyPts(7) = IY - zz * Cos(3 * ztheta)

   ixPts(4) = ixTL + 2 * zz * Sin(4 * ztheta)
   iyPts(4) = IY - 2 * zz * Cos(4 * ztheta)
   ixPts(6) = ixTL - 2 * zz * Sin(4 * ztheta)
   iyPts(6) = IY - 2 * zz * Cos(4 * ztheta)

   ixPts(5) = ixTL
   iyPts(5) = IY + zz

   PIC.PSet (ixPts(0) - LL, iyPts(0) - TT), StarColor
   For i = 1 To 9
      PIC.Line -(ixPts(i) - LL, iyPts(i) - TT), StarColor
   Next i
   PIC.Line -(ixPts(0) - LL, iyPts(0) - TT), StarColor
End Sub
'#### END STAR ###################################################


'#### PLUS & T-PIECE #############################################

Public Sub StartPlusT(PIC As PictureBox, X As Single, Y As Single)
   ixTL = X: iyTL = Y
   ixBR = X: iyBR = Y
   DrawTPiece PIC, vbWhite, ixTL, iyTL, ixBR, iyBR
End Sub

Public Sub DrawPlusT(PIC As PictureBox, X As Single, Y As Single)
   DrawTPiece PIC, vbWhite, ixTL, iyTL, ixBR, iyBR
   ixBR = X: iyBR = Y
   DrawTPiece PIC, vbWhite, ixTL, iyTL, ixBR, iyBR
   ixwidth = ixBR - ixTL: iyheight = iyBR - iyTL
End Sub

Public Sub MovePlusT(PIC As PictureBox, X As Single, Y As Single)
   ' Clear old
   DrawTPiece PIC, vbWhite, ixTL, iyTL, ixBR, iyBR
   ixBR = X: iyBR = Y
   ixTL = ixBR - ixwidth: iyTL = iyBR - iyheight
   ' New position
   DrawTPiece PIC, vbWhite, ixTL, iyTL, ixBR, iyBR
End Sub

Public Sub FinalPlusT(PIC As PictureBox, PIC2 As PictureBox, X As Single, Y As Single)
   'Clear old
   DrawTPiece PIC, vbWhite, ixTL, iyTL, ixBR, iyBR
   PIC.DrawMode = 13
   'ReDraw final
   DrawTPiece PIC, DrawColor, ixTL, iyTL, ixBR, iyBR
'--------------------------------------------
            If aTrace Then
               DrawTPiece PIC2, DrawColor, ixTL, iyTL, ixBR, iyBR, L, T
               PIC2.Refresh
            End If
'--------------------------------------------
End Sub

Public Sub DrawTPiece(PIC As PictureBox, PTColor As Long, _
   ByVal zxs As Single, ByVal zys As Single, ByVal zxe As Single, ByVal zye As Single, _
   Optional LL As Long = 0, Optional TT As Long = 0)

Dim SideLen As Long
Dim xdd As Single, ydd As Single
Dim xa As Single, ya As Single
Dim xb As Single, yb As Single
Dim xs1 As Single, ys1 As Single
Dim xx As Single, yx As Single
Dim x1 As Single, y1 As Single
Dim x2  As Single, y2 As Single
Dim xe1 As Single, ye1 As Single
Dim xe2 As Single, ye2 As Single
Dim yy As Single
Dim xs2 As Single, ys2 As Single
Dim xa1 As Single, ya1 As Single
Dim xb1 As Single, yb1 As Single
Dim xa2 As Single, ya2 As Single
Dim xb2 As Single, yb2 As Single

SideLen = 10  'Fixed Length of side piece(s)

If TheDrawStyle = 21 Then   ' + piece
   'Clear/draw TPiece to left  PLUS PIECE
   Findxeye Spacing, zxs, zys, zxe, zye, xdd, ydd, xa, ya, xb, yb
   x1 = (zxs + zxe) / 2
   y1 = (zys + zye) / 2
   x2 = (xa + xb) / 2
   y2 = (ya + yb) / 2
   Findxeye Spacing / 2, x1, y1, x2, y2, xdd, ydd, xs1, ys1, xx, yx
   Findxeye Spacing / 2, x2, y2, x1, y1, xdd, ydd, xx, yy, xe1, ye1
   Findxeye SideLen, xe1, ye1, xs1, ys1, xdd, ydd, xe2, ye2, xs2, ys2
   PIC.Line (zxs - LL, zys - TT)-(xs1 - LL, ys1 - TT), PTColor
   PIC.Line (xs1 - LL, ys1 - TT)-(xs2 - LL, ys2 - TT), PTColor
   PIC.Line (zxe - LL, zye - TT)-(xe1 - LL, ye1 - TT), PTColor
   PIC.Line (xe1 - LL, ye1 - TT)-(xe2 - LL, ye2 - TT), PTColor

ElseIf TheDrawStyle = 22 Then  ' T-piece
   'Clear/draw main line
   PIC.Line (zxs - LL, zys - TT)-(zxe - LL, zye - TT), PTColor
End If
        
'Clear/draw TPiece to right ALWAYS
Findxeye Spacing, zxs, zys, zxe, zye, xdd, ydd, xa, ya, xb, yb
x1 = (zxs + zxe) / 2
y1 = (zys + zye) / 2
x2 = (xa + xb) / 2
y2 = (ya + yb) / 2
Findxeye Spacing / 2, x1, y1, x2, y2, xdd, ydd, xx, yy, xa1, ya1
Findxeye Spacing / 2, x2, y2, x1, y1, xdd, ydd, xb1, yb1, xx, yy
Findxeye SideLen, xa1, ya1, xb1, yb1, xdd, ydd, xa2, ya2, xb2, yb2
PIC.Line (xa - LL, ya - TT)-(xa1 - LL, ya1 - TT), PTColor
PIC.Line (xa1 - LL, ya1 - TT)-(xa2 - LL, ya2 - TT), PTColor
PIC.Line (xb1 - LL, yb1 - TT)-(xb - LL, yb - TT), PTColor
PIC.Line (xb1 - LL, yb1 - TT)-(xb2 - LL, yb2 - TT), PTColor

End Sub

Public Sub Findxeye(ByVal zd As Single, ByVal x1 As Single, ByVal y1 As Single, _
   ByVal x2 As Single, ByVal y2 As Single, _
   ByRef xdd As Single, ByRef ydd As Single, ByRef xa As Single, ByRef ya As Single, _
   ByRef xb As Single, ByRef yb As Single)

Dim xd12 As Single, yd12 As Single
Dim zang As Single

If zd = 0 Then zd = 0.4

'In:  Coords of line xy1->xy2, paraspacing zd
'Out: Coords of parallel line to right xya->xyb
'     Increment xdd=x1-x0, ydd=y0-y1, needed for Findxiyi when used
'Find angle to horizontal (downwards)
'and hence increments onto parallel line

xd12 = x2 - x1: yd12 = y2 - y1
If xd12 = 0 Then
   ydd = 0: xdd = Sgn(yd12) * zd
ElseIf yd12 = 0 Then
   xdd = 0: ydd = Sgn(xd12) * zd
Else
   zang = Atn(yd12 / xd12)
   xdd = Sgn(xd12) * zd * Sin(zang)
   ydd = Sgn(xd12) * zd * Cos(zang)
End If
xa = x1 - xdd: ya = y1 + ydd
xb = x2 - xdd: yb = y2 + ydd
End Sub

Public Sub DrawDoubleLines(PIC As PictureBox, LCul As Long, Optional LL As Long = 0, Optional TT As Long = 0)


''FINDS AND DRAWS MAIN & PARALLEL LINES Spacing AWAY FROM MAIN LINE

' Public NumOfDrawPoints, ixPts(1 to NumOfDrawPoints),iyPts(1 to NumOfDrawPoints)
' Public Spacing
' In LCul as vbWhite or DrawColor
Dim x1 As Single, y1 As Single
Dim x2 As Single, y2 As Single
Dim x3 As Single, y3 As Single
Dim xi As Single, yi As Single
Dim xdd As Single, ydd As Single
Dim xa As Single, ya As Single
'Dim xa1 As Single, ya1 As Single
Dim xb As Single, yb As Single
Dim zp() As POINTAPI, zpp() As POINTAPI
Dim xprv As Single, yprv As Single
Dim xx As Single, yx As Single
Dim hpen As Long, hpenold As Long
'Dim response As Long
Dim NumPts As Long

NumPts = NumOfDrawPoints

'Set 1st outer parallel line pts
x1 = ixPts(1): y1 = iyPts(1)
If NumPts = 1 Then
   NumPts = 2
   x2 = ixBR: y2 = iyBR
Else
   x2 = ixPts(2): y2 = iyPts(2)
End If

'Find 1st inner parallel line pts
Findxeye Spacing, x1, y1, x2, y2, xdd, ydd, xa, ya, xb, yb
'xa1 = xa: ya1 = ya    'Save start of paraline to close end

If NumPts = 2 Then 'Single segment
   PIC.Line (x1 - LL, y1 - TT)-(x2 - LL, y2 - TT), LCul
   PIC.Line (xa - LL, ya - TT)-(xb - LL, yb - TT), LCul
Else

   hpen = CreatePen(PIC.DrawStyle, PIC.DrawWidth, LCul)
   hpenold = SelectObject(PIC.hdc, hpen)
   ReDim zp(1 To NumPts) As POINTAPI
   ReDim zpp(1 To NumPts) As POINTAPI
   
   xprv = xa: yprv = ya
   For j = 1 To NumPts - 2
      x1 = ixPts(j): y1 = iyPts(j)
      x2 = ixPts(j + 1): y2 = iyPts(j + 1)
      x3 = ixPts(j + 2): y3 = iyPts(j + 2)
      Findxeye Spacing, x1, y1, x2, y2, xdd, ydd, xa, ya, xx, yx
      Findxeye Spacing, x2, y2, x3, y3, xdd, ydd, xx, yx, xb, yb
      Findxiyi x1, y1, x2, y2, x3, y3, xdd, ydd, xa, ya, xb, yb, xi, yi

      If x1 < -30000 Then x1 = -30000
      If x1 > 30000 Then x1 = 30000
      If y1 < -30000 Then y1 = -30000
      If y1 > 30000 Then y1 = 30000

      zp(j).X = x1 - LL: zp(j).Y = y1 - TT

      If x2 < -30000 Then x2 = -30000
      If x2 > 30000 Then x2 = 30000
      If y2 < -30000 Then y2 = -30000
      If y2 > 30000 Then y2 = 30000

      zp(j + 1).X = x2 - LL: zp(j + 1).Y = y2 - TT

      zpp(j).X = xprv - LL: zpp(j).Y = yprv - TT

      If xi < -30000 Then xi = -30000
      If xi > 30000 Then xi = 30000
      If yi < -30000 Then yi = -30000
      If yi > 30000 Then yi = 30000

      zpp(j + 1).X = xi - LL: zpp(j + 1).Y = yi - TT

      xprv = xi: yprv = yi                         'x,yprv -> x,yi
   Next j
   
   response = Polyline(PIC.hdc, zp(1), NumPts - 1)
   response = Polyline(PIC.hdc, zpp(1), NumPts - 1)
   response = SelectObject(PIC.hdc, hpenold)
   response = DeleteObject(hpen)

   'Last parallel line pt
   PIC.Line (x2 - LL, y2 - TT)-(x3 - LL, y3 - TT), LCul
   PIC.Line (xi - LL, yi - TT)-(xb - LL, yb - TT), LCul
   Erase zp, zpp
End If
End Sub


Public Sub Findxiyi(ByVal x1 As Single, ByVal y1 As Single, _
ByVal x2 As Single, ByVal y2 As Single, _
ByVal x3 As Single, ByVal y3 As Single, _
ByVal xdd As Single, ByVal ydd As Single, _
ByVal xa As Single, ByVal ya As Single, _
ByVal xb As Single, ByVal yb As Single, _
ByRef xi As Single, ByRef yi As Single)

'In:  Coords of 2 intersecting lines xya-(xyi)-xyb zd from xy1-xy2-xy3
'      From Findxeye:-
'      xa,ya start coords of the 1st parallel line zd from xy1-xy2
'      xb,yb end coords of the 2nd parallel line zd away from xy2-xy3
'      xdd,ydd incr to x2,y2 to give xi,yi IF the lines have the same slope
'Out: xi,yi intersection of the 2 parallel lines  xya-xyi-xyb

Dim xd12 As Single, yd12 As Single
Dim xd23 As Single, yd23 As Single
Dim xd13 As Single, yd13 As Single
Dim zm12 As Single ', zc1 As Single
Dim zm23 As Single ', zc2 As Single
'Dim zm13 As Single , zc3 As Single


'Find slopes
xd12 = x2 - x1: yd12 = y2 - y1
xd23 = x3 - x2: yd23 = y3 - y2
xd12 = x2 - x1: yd12 = y2 - y1
xd13 = x3 - x1: yd13 = y3 - y1

'Find slopes
If xd12 <> 0 Then
   zm12 = yd12 / xd12
'   zc1 = y1 - zm12 * x1
Else  'Vertical
   zm12 = Sgn(yd12) * 10000
'   zc1 = Sgn(zm12) * 10000
End If

If xd23 <> 0 Then
   zm23 = yd23 / xd23
'   zc2 = y2 - zm23 * x2
Else   'Vertical
   zm23 = Sgn(yd23) * 10000
'   zc2 = Sgn(zm23) * 10000
End If

' For joining ends
'If xd13 <> 0 Then
'   zm13 = yd13 / xd13
'   zc3 = y3 - zm13 * x3
'Else   'Vertical
'   zm13 = Sgn(yd13) * 10000
'   zc3 = Sgn(zm13) * 10000
'End If

'Find intersection
If zm12 <> zm23 Then
      
      If Abs(zm12) > 9000 Then
         xi = xa
         yi = zm23 * xi - xb * zm23 + yb
      ElseIf Abs(zm23) > 9000 Then
         xi = xb
         yi = zm12 * xi - xa * zm12 + ya
      Else
         xi = xa * zm12 - ya - xb * zm23 + yb
         xi = xi / (zm12 - zm23)
         yi = zm23 * xi - xb * zm23 + yb
      End If

Else
   xi = x2 - xdd
   yi = y2 + ydd
End If

'If xi,yi lies outside triangle x1y1,x2y2,x3y3 then xiyi=(x1y1+x3y3)/2
'Restrict magnitude (NB could deal with shallow angles better)

'sin12 = Sgn(zd12 * xi + yi + zc1)
'sin23 = Sgn(zd23 * xi + yi + zc2)
'sin13 = Sgn(zd13 * xi + yi + zc3)
'If sin12 = sin23 And sin12 = sin13 Then
'Else
'   xi = (x1 + x3) / 2: yi = (y1 + y3) / 2
'End If

If xi > 32500 Then xi = 32500
If xi < -32500 Then xi = -32500
If yi > 32500 Then yi = 32500
If yi > 32500 Then yi = 32500
End Sub


'#### PARALLELOGRAM #########################################################

Public Sub StartParallelogram(PIC As PictureBox, X As Single, Y As Single)
   ixTL = X: iyTL = Y
   ixBR = ixTL: iyBR = iyTL
   ReDim ixPts(1 To 1), iyPts(1 To 1)
   NumOfDrawPoints = 1
   ixPts(1) = ixTL: iyPts(1) = iyTL
   PIC.Line (ixTL, iyTL)-(ixBR, iyBR), vbWhite
End Sub

Public Sub DrawParallelogram(PIC As PictureBox, X As Single, Y As Single)

   'Clear old polyline
   PIC.Line (ixTL, iyTL)-(ixBR, iyBR), vbWhite
   
   ixBR = X: iyBR = Y
   'Draw new polylineline
   PIC.Line (ixTL, iyTL)-(ixBR, iyBR), vbWhite
End Sub

'Public Sub StartNextPolyLine(PIC As PictureBox,   X As Single, Y As Single)
'Public Sub UpdatePolyPoints(X As Single, Y As Single)

Public Sub CompleteParallelogram(PIC As PictureBox, X As Single, Y As Single)
   ' ixPts(1 to 3), iyPts(1 to 3)
   ' Clear point 1-3
   For i = 1 To 2
      PIC.Line (ixPts(i), iyPts(i))-(ixPts(i + 1), iyPts(i + 1)), vbWhite
   Next i
   
   NumOfDrawPoints = 4
   ReDim Preserve ixPts(1 To 4)
   ReDim Preserve iyPts(1 To 4)
   ixPts(4) = ixPts(1) - (ixPts(2) - ixPts(3))
   iyPts(4) = iyPts(1) + (iyPts(3) - iyPts(2))
   
   ' Draw parallelogram
   For i = 1 To 3
      PIC.Line (ixPts(i), iyPts(i))-(ixPts(i + 1), iyPts(i + 1)), vbWhite
   Next i
   PIC.Line (ixPts(4), iyPts(4))-(ixPts(1), iyPts(1)), vbWhite
   

End Sub

Public Sub MoveParallelogram(PIC As PictureBox, X As Single, Y As Single)

   ' Clear parallelogram
   For i = 1 To 3
      PIC.Line (ixPts(i), iyPts(i))-(ixPts(i + 1), iyPts(i + 1)), vbWhite
   Next i
   PIC.Line (ixPts(4), iyPts(4))-(ixPts(1), iyPts(1)), vbWhite

   ' Move points to new position
   For i = 1 To NumOfDrawPoints
      ixPts(i) = ixPts(i) + (X - ixTL)
      iyPts(i) = iyPts(i) + (Y - iyTL)
   Next i
   ixTL = X
   iyTL = Y

   ' New parallelogram
   For i = 1 To 3
      PIC.Line (ixPts(i), iyPts(i))-(ixPts(i + 1), iyPts(i + 1)), vbWhite
   Next i
   PIC.Line (ixPts(4), iyPts(4))-(ixPts(1), iyPts(1)), vbWhite

End Sub

Public Sub FinalParallelogram(PIC As PictureBox, PIC2 As PictureBox, X As Single, Y As Single)
   ' Clear parallelogram
   For i = 1 To 3
      PIC.Line (ixPts(i), iyPts(i))-(ixPts(i + 1), iyPts(i + 1)), vbWhite
   Next i
   PIC.Line (ixPts(4), iyPts(4))-(ixPts(1), iyPts(1)), vbWhite

   PIC.DrawMode = 13
   
   ' Final parallelogram
   For i = 1 To 3
      PIC.Line (ixPts(i), iyPts(i))-(ixPts(i + 1), iyPts(i + 1)), DrawColor
   Next i
   PIC.Line (ixPts(4), iyPts(4))-(ixPts(1), iyPts(1)), DrawColor

'--------------------------------------------
            ' To picFullTemp as well
            If aTrace Then
               For i = 1 To 3
                  PIC2.Line (ixPts(i) - L, iyPts(i) - T)-(ixPts(i + 1) - L, iyPts(i + 1) - T), DrawColor
               Next i
               PIC2.Line (ixPts(4) - L, iyPts(4) - T)-(ixPts(1) - L, iyPts(1) - T), DrawColor
               PIC2.Refresh
            End If
'--------------------------------------------
End Sub
'#### END PARALLELOGRAM #########################################################


'#### FRUSTRUM #########################################################

Public Sub StartFrustrum(PIC As PictureBox, X As Single, Y As Single)
   ixTL = X: iyTL = Y
   ixBR = ixTL: iyBR = iyTL
   ReDim ixPts(1 To 1), iyPts(1 To 1)
   NumOfDrawPoints = 1
   ixPts(1) = ixTL: iyPts(1) = iyTL
   PIC.Line (ixTL, iyTL)-(ixBR, iyBR), vbWhite
End Sub

Public Sub DrawFrustrum(PIC As PictureBox, X As Single, Y As Single)

   'Clear old polyline
   PIC.Line (ixTL, iyTL)-(ixBR, iyBR), vbWhite
   
   ixBR = X: iyBR = Y
   'Draw new polylineline
   PIC.Line (ixTL, iyTL)-(ixBR, iyBR), vbWhite
End Sub

'Public Sub StartNextPolyLine(PIC As PictureBox,   X As Single, Y As Single)
'Public Sub UpdatePolyPoints(X As Single, Y As Single)

Public Sub CompleteFrustrum(PIC As PictureBox, X As Single, Y As Single)

Dim x1 As Single, x2 As Single, x3 As Single, x4 As Single
Dim y1 As Single, y2 As Single, y3 As Single, y4 As Single
Dim za As Single
Dim xr2 As Single
Dim xr3 As Single, yr3 As Single
Dim xr4 As Single, yr4 As Single
   
   ' ixPts(1 to 3), iyPts(1 to 3)
   ' Clear point 1-3
   For i = 1 To 2
      PIC.Line (ixPts(i), iyPts(i))-(ixPts(i + 1), iyPts(i + 1)), vbWhite
   Next i
   
   NumOfDrawPoints = 4
   ReDim Preserve ixPts(1 To 4)
   ReDim Preserve iyPts(1 To 4)
   
   x1 = ixPts(1): y1 = iyPts(1)
   x2 = ixPts(2): y2 = iyPts(2)
   x3 = ixPts(3): y3 = iyPts(3)
   'Find angle of 1st line to horizontal
   za = zAtn2(y2 - y1, x2 - x1)
   'Translate to make x1,y1 the origin
   x2 = x2 - x1
   y2 = y2 - y1
   x3 = x3 - x1
   y3 = y3 - y1
   'Rotate za about 1st pt
   xr2 = x2 * Cos(za) + y2 * Sin(za)
   xr3 = x3 * Cos(za) + y3 * Sin(za)
   yr3 = -x3 * Sin(za) + y3 * Cos(za)
   'Find 4th point
   xr4 = (xr2 - xr3)
   yr4 = yr3
   'Rotate back
   x2 = xr2 * Cos(-za)
   x3 = xr3 * Cos(-za) + yr3 * Sin(-za)
   y3 = -xr3 * Sin(-za) + yr3 * Cos(-za)
   x4 = xr4 * Cos(-za) + yr4 * Sin(-za)
   y4 = -xr4 * Sin(-za) + yr4 * Cos(-za)
   'Translate back
   ixPts(2) = x2 + x1: iyPts(2) = y2 + y1
   ixPts(3) = x3 + x1: iyPts(3) = y3 + y1
   ixPts(4) = x4 + x1
   iyPts(4) = y4 + y1
   
   ' Draw Frustrum
   For i = 1 To 3
      PIC.Line (ixPts(i), iyPts(i))-(ixPts(i + 1), iyPts(i + 1)), vbWhite
   Next i
   PIC.Line (ixPts(4), iyPts(4))-(ixPts(1), iyPts(1)), vbWhite
   
End Sub

Public Sub MoveFrustrum(PIC As PictureBox, X As Single, Y As Single)

   ' Clear Frustrun
   For i = 1 To 3
      PIC.Line (ixPts(i), iyPts(i))-(ixPts(i + 1), iyPts(i + 1)), vbWhite
   Next i
   PIC.Line (ixPts(4), iyPts(4))-(ixPts(1), iyPts(1)), vbWhite

   ' Move points to new position
   For i = 1 To NumOfDrawPoints
      ixPts(i) = ixPts(i) + (X - ixTL)
      iyPts(i) = iyPts(i) + (Y - iyTL)
   Next i
   ixTL = X
   iyTL = Y

   ' New Frustrum
   For i = 1 To 3
      PIC.Line (ixPts(i), iyPts(i))-(ixPts(i + 1), iyPts(i + 1)), vbWhite
   Next i
   PIC.Line (ixPts(4), iyPts(4))-(ixPts(1), iyPts(1)), vbWhite

End Sub

Public Sub FinalFrustrum(PIC As PictureBox, PIC2 As PictureBox, X As Single, Y As Single)
   ' Clear Frustrum
   For i = 1 To 3
      PIC.Line (ixPts(i), iyPts(i))-(ixPts(i + 1), iyPts(i + 1)), vbWhite
   Next i
   PIC.Line (ixPts(4), iyPts(4))-(ixPts(1), iyPts(1)), vbWhite

   PIC.DrawMode = 13
   
   ' Final Frustrum
   For i = 1 To 3
      PIC.Line (ixPts(i), iyPts(i))-(ixPts(i + 1), iyPts(i + 1)), DrawColor
   Next i
   PIC.Line (ixPts(4), iyPts(4))-(ixPts(1), iyPts(1)), DrawColor

'--------------------------------------------
            ' To picFullTemp as well
            If aTrace Then
               For i = 1 To 3
                  PIC2.Line (ixPts(i) - L, iyPts(i) - T)-(ixPts(i + 1) - L, iyPts(i + 1) - T), DrawColor
               Next i
               PIC2.Line (ixPts(4) - L, iyPts(4) - T)-(ixPts(1) - L, iyPts(1) - T), DrawColor
               PIC2.Refresh
            End If
'--------------------------------------------
End Sub
'#### END FRUSTRUM #########################################################


'#### ARROWS ###################################################

Public Sub StartArrow(PIC As PictureBox, X As Single, Y As Single)
   ixTL = X: iyTL = Y
   ixBR = X: iyBR = Y
   Arrow PIC, vbWhite
End Sub

Public Sub DrawArrow(PIC As PictureBox, X As Single, Y As Single)
   ' Clear old arrow
   Arrow PIC, vbWhite
   ixBR = X: iyBR = Y
   'Draw new arrow
   Arrow PIC, vbWhite
   ixwidth = ixBR - ixTL: iyheight = iyBR - iyTL
End Sub

Public Sub MoveArrow(PIC As PictureBox, X As Single, Y As Single)
   ' Clear old arrow
   Arrow PIC, vbWhite
   ixBR = X: iyBR = Y
   ixTL = ixBR - ixwidth: iyTL = iyBR - iyheight
   'Move arrow to new position
   Arrow PIC, vbWhite
End Sub

Public Sub FinalArrow(PIC As PictureBox, PIC2 As PictureBox, X As Single, Y As Single)
   ' Clear old arrow
   Arrow PIC, vbWhite
   PIC.DrawMode = 13
   ' Redrwa final arrow
   Arrow PIC, DrawColor
'--------------------------------------------
            ' To picFullTemp as well
            If aTrace Then
               Arrow PIC2, DrawColor, L, T
               PIC2.Refresh
            End If
'--------------------------------------------
End Sub

Public Sub Arrow(PIC As PictureBox, LCul As Long, Optional LL As Long = 0, Optional TT As Long = 0)

Dim xs As Single, ys As Single
Dim xe As Single, ye As Single
Dim zarrang As Single, zarrlen As Single
Dim zang1 As Single, zang2 As Single, zang3 As Single
Dim xd1 As Single, xd2 As Single
Dim yd1 As Single, yd2 As Single
Dim x3 As Single, y3 As Single
Dim x4 As Single, y4 As Single
Dim x5 As Single, y5 As Single

   zarrang = 0.333
   zarrlen = TheDrawWidth * 10
   xs = ixTL: ys = iyTL
   xe = ixBR: ye = iyBR
   'Find angle of 1st line to horizontal
   zang1 = zAtn2(ye - ys, xe - xs)
   
   
   zang2 = zang1 - zarrang: zang3 = pi# / 2 - zang1 - zarrang
   xd1 = zarrlen * Cos(zang2): yd1 = zarrlen * Sin(zang2)
   xd2 = zarrlen * Sin(zang3): yd2 = zarrlen * Cos(zang3)

   Select Case zang1
   Case 0 To pi# / 2, 3 * pi# / 2 To 2 * pi#: zang1 = -zang1
   Case pi# / 2 To 3 * pi# / 2: zang1 = pi# - zang1
   End Select

   x3 = xe - xd1
   y3 = ye - yd1
   x4 = xe - xd2
   y4 = ye - yd2
   x5 = (x3 + x4) / 2
   y5 = (y3 + y4) / 2

   Select Case TheDrawStyle
   Case 25 To 27  'Main line xs,ys-xe,ye but TriArrow
      '--> 0+1+2   Single
      PIC.Line (xs - LL, ys - TT)-(xe - LL, ye - TT), LCul
      PIC.Line (xe - LL, ye - TT)-(x3 - LL, y3 - TT), LCul
      PIC.Line (xe - LL, ye - TT)-(x4 - LL, y4 - TT), LCul
   End Select

   Select Case TheDrawStyle
   Case 26
      '>--> Feathered
      PIC.Line (xs - LL, ys - TT)-(xs - xd1 - LL, ys - yd1 - TT), LCul
      PIC.Line (xs - LL, ys - TT)-(xs - xd2 - LL, ys - yd2 - TT), LCul
   Case 27
      '<--> Double
      PIC.Line (xs - LL, ys - TT)-(xs + xd1 - LL, ys + yd1 - TT), LCul
      PIC.Line (xs - LL, ys - TT)-(xs + xd2 - LL, ys + yd2 - TT), LCul
   Case 28
      '--|> Triangle
      PIC.Line (xs - LL, ys - TT)-(x5 - LL, y5 - TT), LCul  'Main line xs,ys-x5,y5
      PIC.Line (xe - LL, ye - TT)-(x3 - LL, y3 - TT), LCul
      PIC.Line (xe - LL, ye - TT)-(x4 - LL, y4 - TT), LCul
      PIC.Line (x3 - LL, y3 - TT)-(x4 - LL, y4 - TT), LCul
   End Select

End Sub
'#### END ARROWS ###################################################


'### FILL #####################################################

Public Sub Fill(PIC As PictureBox, X As Single, Y As Single, Optional LL As Long = 0, Optional TT As Long = 0)
   ' Fill with FillColor = DrawColor at X,Y
   PIC.DrawStyle = vbSolid
   PIC.DrawMode = 13
   PIC.DrawWidth = 1
   PIC.FillColor = DrawColor
   PIC.FillStyle = vbFSSolid
   
   ' FLOODFILLSURFACE = 1
   ' Fills with FillColor so long as point surrounded by
   ' color = PIC.Point(X, Y)
   
   If PIC.Point(X - LL, Y - TT) <> TransColor Then
      ExtFloodFill PIC.hdc, X - LL, Y - TT, PIC.Point(X - LL, Y - TT), FLOODFILLSURFACE
   ElseIf Not aMerged And Left$(StoreFileSpec$(PicNum), 4) = "TLay" Then
      ExtFloodFill PIC.hdc, X - LL, Y - TT, PIC.Point(X - LL, Y - TT), FLOODFILLSURFACE
   End If
   
   PIC.FillStyle = vbFSTransparent  'Default (Transparent)
   PIC.DrawWidth = TheDrawWidth
   PIC.ForeColor = DrawColor
   PIC.Refresh
End Sub
'### END FILL #####################################################

Public Function zAtn2(ByVal Y As Single, ByVal X As Single) As Single
'0° to right, 0 to -pi#(-180°) anticlockwise, 0 to +pi#(+180°) clockwise
If X = 0 Then
    If Abs(Y) > Abs(X) Then   'Must be an overflow
        If Y > 0 Then zAtn2 = pi# / 2 Else zAtn2 = -pi# / 2
    Else
        zAtn2 = 0   'Must be an underflow
    End If
Else
    zAtn2 = Atn(Y / X)
    If (X < 0) Then
        If (Y < 0) Then zAtn2 = zAtn2 - pi# Else zAtn2 = zAtn2 + pi#
    End If
End If
End Function

