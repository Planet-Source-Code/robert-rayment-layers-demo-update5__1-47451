Attribute VB_Name = "EffectsMod"
' EffectsMod.bas

Option Explicit
         
' From Form1, picLevel_MouseMove
' DISPLAY_EFFECTS picFull(PicNum), picTemp ' Individual in Merged Mode
' MERGE
' Else
' DISPLAY_EFFECTS picDisplay, picTemp ' Non-merged Individual or Whole Merged

Public Sub DISPLAY_EFFECTS(PIC As PictureBox, PIC2 As PictureBox)
' Public FilterParam

Dim Seed As Long

' Shape factors
Dim zFac As Single
Dim NewW As Long
Dim ww As Long
Dim hh As Long
Dim wadj As Long
Dim hadj As Long


' In:  EffectsType = Index & picDisplay visible
      
      Wsrc = PIC.Width: Hsrc = PIC.Height
      svmaxw = maxw: svmaxh = maxh
      
      
      Select Case EffectsType
      
      Case 0 To 5, 7, 8  ' Inverter, Sharp-Soft & Dark-Bright
                        ' & bR,bG,bB +/- & Relief, Metallic
         maxw = Wsrc: maxh = Hsrc
         
         ReDim memBack(1 To Wsrc, 1 To Hsrc)
         memBack(1, 1) = TransColor
         memBack(Wsrc, Hsrc) = TransColor
         'Background Pic To memBack (PIC2 = ORG picDisplay)
         GETDIBS PIC2.Image, 0
         
         ReDim memPic(1 To Wsrc, 1 To Hsrc)
         i = bmbac.bmiH.biSizeImage  ' = ImageSize
         CopyMemory memPic(1, 1), memBack(1, 1), i
         ptrmemPic = VarPtr(memPic(1, 1))

'         ' Same
'         ReDim memBytes(1 To 4, 1 To Wsrc, 1 To Hsrc)
'         i = bmbac.bmiH.biSizeImage  ' = ImageSize
'         CopyMemory memBytes(1, 1, 1), memBack(1, 1), i
'         ptrmemPic = VarPtr(memBytes(1, 1, 1))
         
         ptrmemBack = VarPtr(memBack(1, 1))
         FillASMEffects
         response = CallWindowProc(ptrMC3, ptrStruc3, 2&, 3&, EffectsType)
         'Action in ASM is from memPic() to memBack()
      
         maxw = svmaxw: maxh = svmaxh
      
      Case 6   ' Diffuse
         
         maxw = Wsrc: maxh = Hsrc
         
         ReDim memBack(1 To Wsrc, 1 To Hsrc)
         memBack(1, 1) = TransColor
         memBack(Wsrc, Hsrc) = TransColor
         'Background Pic To memBack (PIC2 = ORG picDisplay)
         GETDIBS PIC2.Image, 0

         'VBDiffuser   ' tests: comment out next 4 lines to try
                        ' see PublicMod.bas
                        
         ptrmemBack = VarPtr(memBack(1, 1))
         FillASMEffects
         Seed = Rnd * 255
         response = CallWindowProc(ptrMC3, ptrStruc3, ptrMC3, Seed, EffectsType)
         'Action in ASM is from memBack() to memBack()
         
         maxw = svmaxw: maxh = svmaxh
      
      Case 9  ' Flute Up \/
            
         PIC.Picture = LoadPicture
         
         If FilterParam = 128 Then
            BitBlt PIC.hdc, 0, 0, Wsrc, Hsrc, PIC2.hdc, 0, 0, vbSrcCopy
            PIC.Refresh
            Exit Sub
         End If

         ' Define Flute UP
         zFac = Wsrc / FilterParam   ' 8
         For i = 0 To Hsrc - 1
            j = Int(3 * zFac * i / Hsrc)
            NewW = Wsrc - 2 * j
            SetStretchBltMode PIC.hdc, HALFTONE
            StretchBlt PIC.hdc, j, i, NewW, 1, _
               PIC2.hdc, 0, i, Wsrc, 1, vbSrcCopy
         Next i

         PIC.Refresh
      
      Case 10  ' Flute down  /\

         PIC.Picture = LoadPicture
         
         If FilterParam = 128 Then
            BitBlt PIC.hdc, 0, 0, Wsrc, Hsrc, PIC2.hdc, 0, 0, vbSrcCopy
            PIC.Refresh
            Exit Sub
         End If
         
         ' Define Flute DOWN
         zFac = Wsrc / FilterParam  ' 8
         For i = 0 To Hsrc - 1
            j = Int(3 * zFac * (1 - i / Hsrc))
            NewW = Wsrc - 2 * j
            SetStretchBltMode PIC.hdc, HALFTONE
            StretchBlt PIC.hdc, j, i, NewW, 1, _
               PIC2.hdc, 0, i, Wsrc, 1, vbSrcCopy
         Next i
         
         PIC.Refresh
      
      Case 11  ' Ripple
      
         If FilterParam > 0 Then
            
            PIC.Picture = LoadPicture
         
            zFac = Wsrc / 20 'FilterParam
            For i = 0 To Hsrc - 1
         
               j = Int(zFac * (1 + Sin(pi# * (1 + FilterParam * i / Hsrc))))
               NewW = Wsrc - 2 * j
               SetStretchBltMode PIC.hdc, HALFTONE
               StretchBlt PIC.hdc, j, i, NewW, 1, _
                  PIC2.hdc, 0, i, Wsrc, 1, vbSrcCopy
            Next i
         Else
            SetStretchBltMode PIC.hdc, HALFTONE
            StretchBlt PIC.hdc, 0, 0, Wsrc, Hsrc, _
               PIC2.hdc, 0, 0, Wsrc, Hsrc, vbSrcCopy
         End If
         
         PIC.Refresh
         
      Case 12  ' Rounded rectangle
         
         PIC.Picture = LoadPicture
         
         zFac = FilterParam * Hsrc \ 256  ' Corner radius
         For j = 0 To Hsrc - 1
            If j < zFac Then
               i = zFac - Sqr(zFac * zFac - (zFac - j) * (zFac - j))
            ElseIf j > Hsrc - zFac Then
               i = zFac - Sqr(zFac * zFac - (zFac - Hsrc + j) * (zFac - Hsrc + j))
            Else
               i = 0
            End If
            NewW = Wsrc - 2 * i
            
            SetStretchBltMode PIC.hdc, HALFTONE
            StretchBlt PIC.hdc, i, j, NewW, 1, _
               PIC2.hdc, 0, j, Wsrc, 1, vbSrcCopy
         Next j
         
         PIC.Refresh
      
      Case 13  ' Tile
         
         PIC.Picture = LoadPicture
         ww = Wsrc / FilterParam: hh = Hsrc / FilterParam
         For j = 0 To Hsrc - 1 Step hh + 1
            hadj = hh
            If j + hadj > Hsrc - 1 Then hadj = Hsrc - j - 1
            For i = 0 To Wsrc - 1 Step ww + 1
               wadj = ww
               If i + wadj > Wsrc - 1 Then wadj = Wsrc - i - 1
               SetStretchBltMode PIC.hdc, HALFTONE
               StretchBlt PIC.hdc, i, j, wadj, hadj, _
                  PIC2.hdc, 0, 0, Wsrc, Hsrc, vbSrcCopy
            Next i
         Next j
         
         PIC.Refresh


      Case 14  ' Horizontal Shading

         maxw = Wsrc: maxh = Hsrc
         
         ReDim memBack(1 To Wsrc, 1 To Hsrc)
         memBack(1, 1) = TransColor
         memBack(Wsrc, Hsrc) = TransColor
         'Background Pic To memBack (PIC2 = ORG picDisplay)
         GETDIBS PIC2.Image, 0
         
         ptrmemBack = VarPtr(memBack(1, 1))
         FillASMEffects
         response = CallWindowProc(ptrMC3, ptrStruc3, 2&, 3&, 9&)
         'Action in ASM is from memBack() to memBack()
         
'         ReDim memBytes(1 To 4, 1 To Wsrc, 1 To Hsrc)
'         i = bmbac.bmiH.biSizeImage  ' = ImageSize
'         CopyMemory memBytes(1, 1, 1), memBack(1, 1), i
'         VB_H_SHADE
      
         maxw = svmaxw: maxh = svmaxh
      
      Case 15  ' Vertical Shading

         maxw = Wsrc: maxh = Hsrc
         
         ReDim memBack(1 To Wsrc, 1 To Hsrc)
         memBack(1, 1) = TransColor
         memBack(Wsrc, Hsrc) = TransColor
         'Background Pic To memBack (PIC2 = ORG picDisplay)
         GETDIBS PIC2.Image, 0
         
         ptrmemBack = VarPtr(memBack(1, 1))
         FillASMEffects
         response = CallWindowProc(ptrMC3, ptrStruc3, 2&, 3&, 10&)
         'Action in ASM is from memBack() to memBack()
         
'         ReDim memBytes(1 To 4, 1 To Wsrc, 1 To Hsrc)
'         i = bmbac.bmiH.biSizeImage  ' = ImageSize
'         CopyMemory memBytes(1, 1, 1), memBack(1, 1), i
'         VB_V_SHADE
         
         maxw = svmaxw: maxh = svmaxh

'      Case 16
'      Case 17
'      Case 18
'      Case 19
      
      
      
      ' Singles
      Case 20   ' Elliptic O
            
         PIC.Picture = LoadPicture
         
         ' Define ellipse
         For i = 0 To Hsrc - 1
            zFac = (((i - (Hsrc \ 2)) / (Hsrc \ 2)) ^ 2)
            If zFac > 1 Then zFac = 1
            j = Int((Wsrc \ 2) * (1 - Sqr(1 - zFac)))
            NewW = Wsrc - 2 * j
            SetStretchBltMode PIC.hdc, HALFTONE
            StretchBlt PIC.hdc, j, i, NewW, 1, _
               PIC2.hdc, 0, i, Wsrc, 1, vbSrcCopy
         Next i
         
         PIC.Refresh
      
      Case 21  ' Flip horizontal
      
         PIC.Picture = LoadPicture

         SetStretchBltMode PIC.hdc, HALFTONE
         StretchBlt PIC.hdc, 0, 0, Wsrc, Hsrc, _
            PIC2.hdc, Wsrc - 1, 0, -Wsrc, Hsrc, vbSrcCopy
         
         PIC.Refresh
      
      Case 22  ' Mirror right hand half
      
         PIC.Picture = LoadPicture
         
         ' Mirror right hand half   < ] to [ ]
         ' Right hand half unchanged
         ' Right hand half mirrored to Left hand half
         SetStretchBltMode PIC.hdc, HALFTONE
         StretchBlt PIC.hdc, 0, 0, Wsrc \ 2, Hsrc, _
            PIC2.hdc, Wsrc - 1, 0, -Wsrc \ 2, Hsrc, vbSrcCopy
         SetStretchBltMode PIC.hdc, HALFTONE
         StretchBlt PIC.hdc, Wsrc \ 2, 0, Wsrc \ 2, Hsrc, _
            PIC2.hdc, Wsrc \ 2, 0, Wsrc \ 2, Hsrc, vbSrcCopy
      
         PIC.Refresh
      
      Case 23  ' Mirror left hand half  < ] to < >
      
         PIC.Picture = LoadPicture
         
         ' Mirror left hand half
         ' Left hand half unchanged
         ' Left hand half mirrored to Right hand half
         SetStretchBltMode PIC.hdc, HALFTONE
         StretchBlt PIC.hdc, 0, 0, Wsrc \ 2, Hsrc, _
            PIC2.hdc, 0, 0, Wsrc \ 2, Hsrc, vbSrcCopy
         SetStretchBltMode PIC.hdc, HALFTONE
         StretchBlt PIC.hdc, Wsrc \ 2, 0, Wsrc \ 2, Hsrc, _
            PIC2.hdc, Wsrc \ 2, 0, -Wsrc \ 2, Hsrc, vbSrcCopy
         
         PIC.Refresh
           
            
      End Select
      
      
      
      If EffectsType < 9 Or EffectsType = 14 Or EffectsType = 15 Then
         ' Effects using ASM to memBack()
         
         With PIC
            '.Picture = LoadPicture
            .Width = Wsrc
            .Height = Hsrc
            .Refresh
         End With
            
         SetStretchBltMode PIC.hdc, HALFTONE
         
         ' Blit memBack to picDisplay
         StretchDIBits PIC.hdc, _
            0&, 0&, Wsrc, Hsrc, _
            0&, 0&, Wsrc, Hsrc, _
            memBack(1, 1), bmbac, 0, vbSrcCopy
         
         PIC.Refresh
         
         Erase memBack()
         Erase memPic()
      
      End If
         

 End Sub

