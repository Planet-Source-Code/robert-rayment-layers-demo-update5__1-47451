Attribute VB_Name = "PublicMod"
' PublicMod.bas

Option Explicit

' Sleep API Delay for resizing & rotating
Public TLIM As Long

' Loading pics
Public PathSpec$  ' App path
Public APath$     ' Folder containing multi-selected pictures
Public FileSpec$()  ' Last opened picture
Public NumPicsSelected As Long

' Pic info
Public StoreFileSpec$()    ' Stored picture filespecs
Public PicWidth() As Long
Public PicHeight() As Long

Public NumOfStoredPics As Long
Public PicNum As Long      ' Current stored picture displayed, optSelect
Public MaxNumOfPics As Long   ' Settable at Form_Load
Public W As Long  ' A pic width
Public H As Long  ' A pic height
Public T As Long  ' A pic Top
Public L As Long  ' A pic left
Public maxw As Long  ' Max pic width   'set by first background
Public maxh As Long  ' Max pic height  'picture
Public svmaxw As Long  ' Save Max pic width   'set by first background
Public svmaxh As Long  ' Save Max pic height  'picture
Public SmallPicW As Long   ' Thumb pic box sizes
Public SmallPicH As Long
Public ww As Long  ' Alternative pic width
Public hh As Long  ' Alternative pic height

Public WHMin As Long
Public WHMax As Long
Public WHChange As Long

' SHOW ALL allows individual pic selection
Public aShowON As Boolean
Public aShowALLON As Boolean

' Clipping
Public aClipON As Boolean

' Clipping & Drawing & Lasso
Public ixTL As Long, iyTL As Long   ' For clipping pic 0, background
Public ixBR As Long, iyBR As Long   ' * lasso
Public ixwidth As Long, iyheight As Long
Public prevX As Long, prevY As Long


' Lasso
Public aLasso As Boolean
Public NumOfSLines As Long
Public ixSL0 As Long ' Start coords of S lines
Public iySL0 As Long
Public SCul As Long  ' S line color

' Merge
Public aMerged As Boolean ' Flag when merging displayed
Public aResizeMerge As Boolean
Public memBack() As Long
Public memPic() As Long
Public PicAlpha() As Long  ' 0 - 256
Public iza As Long      ' 0   -> 256

' Resizing screen
Public STX As Long   ' = Screen.TwipsPerPixelX
Public STY As Long   ' = Screen.TwipsPerPixelY
Public ExtraBorder As Long
Public ExtraHeight As Long ' For different Win sys
Public ORG_ScreenWidth As Long

' Colors
Public TransColor As Long ' Common transparent color
Public TextColor As Long
Public DrawColor As Long ' Drawing color for picDisplay
Public aPicAColor As Boolean
Public bR As Byte
Public bG As Byte
Public bB As Byte

' Drawing
Public aDRAW As Boolean
Public TheDrawStyle As Long
Public TheDrawWidth As Long
Public SStep As Long       ' Blend stength
Public ixc As Long         ' Cirllipse parameter
Public iyc As Long         ' Cirllipse parameter
Public zaspect As Single   ' Cirllipse parameter
Public zradius As Single   ' Cirllipse parameter
Public ixPts() As Integer  ' Ribbon coords
Public iyPts() As Integer
Public NumOfDrawPoints As Long
Public svTheDrawStyle As Long
Public svDrawWidth As Long
Public aTrace As Long            ' Copies some drawing to picFullTemp
Public ADDED As Boolean    ' Indicates a shape has been added in TRace mode

Public BESize As Long
Public SRSize As Long
Public Spacing As Long
Public aDrawStart As Boolean
Public aDrawMove As Boolean
Public EraseColor As Long
Public LClickCount As Long
Public RClickCount As Long
Public aHairs As Boolean

Public aShowInstructions As Boolean

' Effects
Public EffectsType As Long
Public memBytes() As Byte
Public FilterParam As Long
Public aEffects As Boolean
Public aIndividual As Boolean

' Transparent layers TLayers
Public TLayerWidth As Long    ' Default sizes
Public TLayerHeight As Long

' Resizing pic
Public aResize As Boolean

' Rotate pic
Public StepAngle As Long

' Magnifier
Public aMagON As Boolean

' Text
Public TheText$      ' Picked up from frmText.frm

' Help
Public aHelp As Boolean  ' Help form flag

Public frmToolsTop As Long
Public frmToolsLeft As Long

' Gen
Public i As Long  ' Gen loop counter
Public j As Long  ' Gen loop counter
Public a$         ' Gen string
Public oldMODE As Long  ' Gen blitting mode
Public response As Long ' Gen response to MsgBox & APIs
Public aDone As Boolean
Public aKeying As Boolean

Public Const pi# = 3.14159625

Public Sub FillbacStruc(ByVal mwidth As Long, ByVal mheight As Long)
Dim ScanLineBytes As Long
  
  With bmbac.bmiH
   .biSize = 40
   .biwidth = mwidth
   .biheight = -mheight
   .biPlanes = 1
   .biBitCount = 32     ' BGRA
   .biCompression = 0
   
   ' Ensure expansion to 4B boundary
   ScanLineBytes = (((mwidth * .biBitCount) + 31) \ 32) * 4

   .biSizeImage = ScanLineBytes * Abs(.biheight)
   .biXPelsPerMeter = 0
   .biYPelsPerMeter = 0
   .biClrUsed = 0
   .biClrImportant = 0
 End With

End Sub

Public Sub FillPicStruc(ByVal mwidth As Long, ByVal mheight As Long)
Dim ScanLineBytes As Long
  
  With bmpic.bmiH
   .biSize = 40
   .biwidth = mwidth
   .biheight = -mheight
   .biPlanes = 1
   .biBitCount = 32     ' BGRA
   .biCompression = 0
   
   ' Ensure expansion to 4B boundary
   ScanLineBytes = (((mwidth * .biBitCount) + 31) \ 32) * 4

   .biSizeImage = ScanLineBytes * Abs(.biheight)
 End With

End Sub

Public Sub GETDIBS(ByVal PICIM As Long, PIC As Long)
' PICIM = picture.Image
' PIC = 0 for background else other pics
Dim NewDC As Long
Dim OldH As Long

NewDC = CreateCompatibleDC(0&)
OldH = SelectObject(NewDC, PICIM)


If PIC = 0 Then   ' Background
   ' Public maxw,maxh, memBack
   FillbacStruc maxw, maxh ' fill BITMAPINFOHEADER
   If GetDIBits(NewDC, PICIM, 0, maxh, memBack(1, 1), bmbac, 1&) = 0 Then
      MsgBox "Background DIB Error in GETDIBS"
      End
   End If
Else  ' Other pics
   ' Public W,H, memPic
   FillPicStruc W, H ' fill BITMAPINFOHEADER
   If GetDIBits(NewDC, PICIM, 0, H, memPic(1, 1), bmpic, 1&) = 0 Then
     MsgBox "Other pics DIB Error in GETDIBS" & Str$(PIC)
     End
   End If
End If

' Clear mem
SelectObject NewDC, OldH
DeleteDC NewDC

End Sub


Public Sub GETBYTES(ByVal PICIM As Long)
Dim NewDC As Long
Dim OldH As Long

NewDC = CreateCompatibleDC(0&)
OldH = SelectObject(NewDC, PICIM)
   
   
   FillbacStruc maxw, maxh ' fill BITMAPINFOHEADER
   If GetDIBits(NewDC, PICIM, 0, maxh, memBytes(1, 1, 1), bmbac, 1&) = 0 Then
      MsgBox "DIB Error in GETBYTES", vbCritical, " Layers - Drawing"
      End
   End If

' Clear mem
SelectObject NewDC, OldH
DeleteDC NewDC

End Sub

Public Sub FixScrollbars(picC As PictureBox, picP As PictureBox, HS As HScrollBar, VS As VScrollBar)
   ' picC = Container = picFrame
   ' picP = Picture   = picDisplay
      HS.Max = picP.Width - picC.Width + 4   ' +4 to allow for border
      VS.Max = picP.Height - picC.Height + 4 ' +4 to allow for border
      HS.LargeChange = picC.Width \ 10
      HS.SmallChange = 1
      VS.LargeChange = picC.Height \ 10
      VS.SmallChange = 1
      HS.Top = picC.Top + picC.Height + 1
      HS.Left = picC.Left
      HS.Width = picC.Width
      If picP.Width < picC.Width Then
         HS.Visible = False
         'HS.Enabled = False
      Else
         HS.Visible = True
         'HS.Enabled = True
      End If
      VS.Top = picC.Top
      VS.Left = picC.Left - VS.Width - 1
      VS.Height = picC.Height
      If picP.Height < picC.Height Then
         VS.Visible = False
         'VS.Enabled = False
      Else
         VS.Visible = True
         'VS.Enabled = True
      End If
End Sub

Public Sub FixExtension(FSpec$)
Dim p As Long
Dim ext$

If Len(FSpec$) = 0 Then Exit Sub

   p = InStr(1, FSpec$, ".")
   
   If p = 0 Then
      FSpec$ = FSpec$ & ".bmp"
   Else
      ext$ = LCase$(Mid$(FSpec$, p))
      If ext$ <> ".bmp" Then FSpec$ = Mid$(FSpec$, 1, p - 1) & ".bmp"
   End If

End Sub


Public Function ShortFileSpec(FSpec$, Maxlen As Long) As String
Dim section1 As Long
Dim section2 As Long
Dim section3 As Long

ShortFileSpec = FSpec$
If Len(FSpec$) <= Maxlen Then Exit Function

section1 = Maxlen \ 4
section2 = Maxlen \ 4
section3 = Maxlen \ 2

ShortFileSpec = Left$(FSpec$, section1)
ShortFileSpec = ShortFileSpec & String$(section2, ".")
ShortFileSpec = ShortFileSpec & Right$(FSpec$, section3)

End Function


Public Sub GetExtras(BStyle As Byte)

Dim Border As Long
Dim CapHeight As Long
Dim MenuHeight As Long

' Public ExtraBorder, ExtraHeight

' BStyle 1 to 5 (not 0)
' BStyle = Form1.BorderStyle

Border = GetSystemMetrics(SM_CXDLGFRAME)
If BStyle = 2 Or BStyle = 5 Then Border = Border + 1 ' Sizable
If BStyle > 3 Then
   CapHeight = GetSystemMetrics(SM_CYSMCAPTION) ' Small cap - ToolWindow
Else
   CapHeight = GetSystemMetrics(SM_CYCAPTION)   ' Standard cap
End If
ExtraBorder = 2 * Border
ExtraHeight = CapHeight + ExtraBorder

MenuHeight = GetSystemMetrics(SM_CYMENU)
ExtraHeight = CapHeight + MenuHeight + ExtraBorder

' Win98  ExtraBorder=6 or 8, ExtraHeight= 41 - 46
' WinXP  ExtraBorder=6 or 8, ExtraHeight= 44 - 54

End Sub

Public Sub Extract_Store_FileSpecs(ByVal FileSpecString$) ', ByVal StartPicNum As Long)
'Public FileSpec$(), NumPicsSelected

' FileSpecString$ = multi-selected filespec string
' eg "C:\Program Files\Common\File1.bmp"  ' One file
' eg "C:\Program Files\Common|File1.bmp|File2.jpg"  ' Two files, | = Null char

Dim pNull1 As Long  ' Instr pointers
Dim pNull2 As Long
Dim FName$

   NumPicsSelected = -1 'StartPicNum
   
   pNull2 = 1
   pNull1 = InStr(pNull2, FileSpecString$, vbNullChar)
   If pNull1 = 0 Then
      NumPicsSelected = NumPicsSelected + 1
      ReDim FileSpec$(NumPicsSelected)
      FileSpec$(NumPicsSelected) = FileSpecString$
   Else  ' pNull1<>0
      
      APath$ = Left$(FileSpecString$, pNull1 - 1) & "\"
      pNull1 = pNull1 + 1
      
      Do
         pNull2 = InStr(pNull1, FileSpecString$, vbNullChar)
         If pNull2 = 0 Then
            NumPicsSelected = NumPicsSelected + 1
            ReDim Preserve FileSpec$(NumPicsSelected)
            FName$ = Mid$(FileSpecString$, pNull1, Len(FileSpecString$) - pNull1 + 1)
            FileSpec$(NumPicsSelected) = FName$
            Exit Do
         Else
            NumPicsSelected = NumPicsSelected + 1
            ReDim Preserve FileSpec$(NumPicsSelected)
            FName$ = Mid$(FileSpecString$, pNull1, pNull2 - pNull1)
            FileSpec$(NumPicsSelected) = FName$
         End If
         pNull1 = pNull2 + 1
      Loop
   
   End If

NumPicsSelected = NumPicsSelected + 1

End Sub


'' VB VB VB  ALTERNATiVE CODE ############################
'
'Public Sub VBDiffuser()
'' From DISPLAY_EFFECTS
'
''1.   ' Diffuser
''     incY = FilterParam * (Rnd - 0.5)
''     incX = FilterParam * (Rnd - 0.5)
'
''2.   ' Vert fluted glass
''     incY=0
''     incX = FilterParam * Cos(ixsource)
'
''3.   ' Horz & Vert Fluting, sort of pixelated
''     incY = FilterParam * Sin(ixsource)
''     incX = FilterParam * Sin(ixsource)
'
'Dim incX As Single
'Dim incY As Single
'
'Dim ixsource As Long
'Dim iysource As Long
'Dim ixdest As Long   ' TopLeft corner of rectangle
'Dim iydest As Long
'
'For iysource = 1 To maxh Step 1
'
'   For ixsource = 1 To maxw Step 1
'
'      If memBack(ixsource, iysource) <> TransColor Then
'         incY = FilterParam * Sin(ixsource)  '(Rnd - 0.5)
'         If (iysource + incY) < 1 Then incY = 0
'         If (iysource + incY) > maxh Then incY = 0
'
'         incX = FilterParam * Sin(iysource) '(Rnd - 0.5)
'         If (ixsource + incX) < 1 Then incX = 0
'         If (ixsource + incX) > maxw Then incX = 0
'         memBack(ixsource, iysource) = memBack(ixsource + incX, iysource + incY)
'
'      End If
'
'   Next ixsource
'Next iysource
'
'End Sub

'===============================================================================
'Public Sub VBRelief()
'' From DISPLAY_EFFECTS
'' ReDim memBack(1 To Wsrc, 1 To Hsrc)
'' Background Pic To memBack (picTemp = ORG picDisplay)
'' GETDIBS picTemp.Image, 0
'' ReDim memBytes(1 To 4, 1 To maxw, 1 To maxh)
'' GETBYTES picTemp.Image
'
'Dim ixsource As Long
'Dim iysource As Long
'Dim ixdest As Long   ' TopLeft corner of rectangle
'Dim iydest As Long
'
'Dim LBlue As Long
'Dim LGreen As Long
'Dim LRed As Long
'
'For iysource = 2 To maxh - 1
'   For ixsource = 2 To maxw - 1
'
'      If memBack(ixsource, iysource) <> TransColor Then
'
'         LBlue = 2 * memBytes(1, ixsource - 1, iysource - 1) + _
'         memBytes(1, ixsource, iysource - 1) + memBytes(1, ixsource - 1, iysource) _
'         - 2 * memBytes(1, ixsource + 1, iysource + 1) - memBytes(1, ixsource, iysource + 1) _
'         - memBytes(1, ixsource + 1, iysource)
'
'         LGreen = 2 * memBytes(2, ixsource - 1, iysource - 1) + _
'         memBytes(2, ixsource, iysource - 1) + memBytes(2, ixsource - 1, iysource) _
'         - 2 * memBytes(2, ixsource + 1, iysource + 1) - memBytes(2, ixsource, iysource + 1) _
'         - memBytes(2, ixsource + 1, iysource)
'
'         LRed = 2 * memBytes(3, ixsource - 1, iysource - 1) + _
'         memBytes(3, ixsource, iysource - 1) + memBytes(3, ixsource - 1, iysource) _
'         - 2 * memBytes(3, ixsource + 1, iysource + 1) - memBytes(3, ixsource, iysource + 1) _
'         - memBytes(3, ixsource + 1, iysource)
'
'         LBlue = (memBytes(1, ixsource, iysource) + LBlue) \ 2 + FilterParam
'         LGreen = (memBytes(2, ixsource, iysource) + LGreen) \ 2 + FilterParam '50
'         LRed = (memBytes(3, ixsource, iysource) + LRed) \ 2 + FilterParam '50
'
'         If LBlue < 0 Then LBlue = 0
'         If LBlue > 255 Then LBlue = 255
'         If LGreen < 0 Then LGreen = 0
'         If LGreen > 255 Then LGreen = 255
'         If LRed < 0 Then LRed = 0
'         If LRed > 255 Then LRed = 255
'
'         memBack(ixsource, iysource) = RGB(LBlue, LGreen, LRed)
'
'      End If
'
'   Next ixsource
'Next iysource
'
'End Sub

'===============================================================================
'===============================================================================
' VB VB VB  MERGE CODE
'Dim ix As Long
'Dim iy As Long
'Dim picCul As Long
'Dim mbacCul As Long
'
'Dim bpicred As Long
'Dim bpicgreen As Long
'Dim bpicblue As Long
'Dim bacred As Long
'Dim bacgreen As Long
'Dim bacblue As Long
'
'Dim ixbac As Long
'Dim iybac As Long
''''''''''''''''''''''''''
'
'   ReDim memBack(1 To maxw, 1 To maxh)
'   memBack(1, 1) = TransColor
'   memBack(maxw, maxh) = TransColor
'   ' NB Needs to be (1 To maxw) ??. (maxw) gives a streaky output ??!!
'
'   'Background Pic To memBack
'   GETDIBS picFull(0).Image, 0
'
'   For i = 1 To NumOfStoredPics - 1
      
      ' Following pictures
      ' Coords
'      W = picFull(i).Width
'      H = picFull(i).Height
'      T = picFull(i).Top
'      L = picFull(i).Left
'
'      ReDim memPic(1 To W, 1 To H)
'      'Get pic To memPic
'      GETDIBS picFull(i).Image, i
'
'      iza = PicAlpha(i)    ' 0   -> 256
'
'         '....... MERGE ASM based on this VB code ..
'         ' Add all colors to memBack if not TransColor

'         For iy = H To 1 Step -1   ' picFull(i).Height
'
'            iybac = iy + T    ' picFull(i).Top
'
'            If iybac >= 1 Then
'            If iybac <= maxh Then
'
'               For ix = W To 1 Step -1   ' picFull(i).Width
'
'                  ixbac = ix + L    ' picFull(i).Left
'
'                  If ixbac >= 1 Then
'                  If ixbac <= maxw Then
'
'                     picCul = memPic(ix, iy)
'
'                     If picCul <> TransColor Then
'
'                        bpicblue = (picCul And &HFF&)
'                        bpicgreen = (picCul And &HFF00&) / &H100&
'                        bpicred = (picCul And &HFF0000) / &H10000
'
'                        mbacCul = memBack(ixbac, iybac)
'
'                        bacblue = (mbacCul And &HFF&)
'                        bacgreen = (mbacCul And &HFF00&) / &H100&
'                        bacred = (mbacCul And &HFF0000) / &H10000
'
'                        ' Cross fade, ia = 0 to 256
'                        bacblue = ia * (bpicblue - bacblue) \ 256 + bacblue
'                        bacgreen = ia * (bpicgreen - bacgreen) \ 256 + bacgreen
'                        bacred = ia * (bpicred - bacred) \ 256 + bacred
'
'                        ' Take account of all merged layers underneath
'                        memBack(ixbac, iybac) = RGB(bacblue, bacgreen, bacred)
'
'                     End If   ' If picCul <> TransColor Then
'
'                  End If   ' If ixbac <= maxw Then
'                  End If   ' If ixbac >= 1 Then
'
'               Next ix
'
'            End If   ' If iybac <= maxh Then
'            End If   ' If iybac >= 1 Then
'
'         Next iy
'         '............................................................
'  Next i
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Public Sub VB_H_SHADE()
'
''' From DISPLAY_EFFECTS
''         maxw = Wsrc: maxh = Hsrc
''
''         ReDim memBack(1 To Wsrc, 1 To Hsrc)
''         memBack(1, 1) = TransColor
''         memBack(Wsrc, Hsrc) = TransColor
''         'Background Pic To memBack (PIC2 = ORG picDisplay)
''         GETDIBS PIC2.Image, 0
''
''         ReDim memBytes(1 To 4, 1 To Wsrc, 1 To Hsrc)
''         i = bmbac.bmiH.biSieImage  ' = ImageSie
''         CopyMemory memBytes(1, 1, 1), memBack(1, 1), i
''
''         Action on memBytes() -> memBack()
'' Public FilterParam
'
'Dim ixs As Long
'Dim iys As Long
'
'Dim iB As Long
'Dim iG As Long
'Dim iR As Long
'
'Dim im As Long
'Dim ic As Long
'
'im = Abs(4 * FilterParam)  ' FilterParam = -64 -> 0 -> +64
'If im > 255 Then im = 255
'
'For iys = maxh To 1 Step -1
'   For ixs = maxw To 1 Step -1
'
'      ic = CLng(Sgn(FilterParam) * (512 * ixs / maxw - 256))  ' Single Edge shading
'      'ic = Abs(512 * ixs / maxw - 256)    ' Centre shading
'      'ic = Abs(256 - 512 * ixs / maxw)   ' Centre shading
'      If ic > 256 Then ic = 256
'      If ic < 0 Then ic = 0
'
'      iR = memBytes(1, ixs, iys)
'      iR = im * (ic - iR) \ 256 + iR
'      If iR < 0 Then iR = 0
'      If iR > 255 Then iR = 255
'      iG = memBytes(2, ixs, iys)
'      iG = im * (ic - iG) \ 256 + iG
'      If iG < 0 Then iG = 0
'      If iG > 255 Then iG = 255
'      iB = memBytes(3, ixs, iys)
'      iB = im * (ic - iB) \ 256 + iB
'      If iB < 0 Then iB = 0
'      If iB > 255 Then iB = 255
'
'      If RGB(iR, iG, iB) = TransColor Then
'         If iR > 0 Then iR = iR - 1 Else iR = iR + 1
'      End If
'
'      memBack(ixs, iys) = RGB(iR, iG, iB)
'
'   Next ixs
'Next iys
'
'End Sub
'
'Public Sub VB_V_SHADE()
'
''' From DISPLAY_EFFECTS
''         maxw = Wsrc: maxh = Hsrc
''
''         ReDim memBack(1 To Wsrc, 1 To Hsrc)
''         memBack(1, 1) = TransColor
''         memBack(Wsrc, Hsrc) = TransColor
''         'Background Pic To memBack (PIC2 = ORG picDisplay)
''         GETDIBS PIC2.Image, 0
''
''         ReDim memBytes(1 To 4, 1 To Wsrc, 1 To Hsrc)
''         i = bmbac.bmiH.biSieImage  ' = ImageSie
''         CopyMemory memBytes(1, 1, 1), memBack(1, 1), i
''
''         Action on memBytes() -> memBack()
'' Public FilterParam
'
'Dim ixs As Long
'Dim iys As Long
'
'Dim iB As Long
'Dim iG As Long
'Dim iR As Long
'
'Dim im As Long
'Dim ic As Long
'
'im = Abs(4 * FilterParam)  ' FilterParam = -64 -> 0 -> +64
'If im > 255 Then im = 255
'
'For ixs = maxw To 1 Step -1
'   For iys = maxh To 1 Step -1
'
'      ic = CLng(Sgn(FilterParam) * (512 * iys / maxh - 256))
'      If ic > 256 Then ic = 256
'      If ic < 0 Then ic = 0
'
'      iR = memBytes(1, ixs, iys)
'      iR = im * (ic - iR) \ 256 + iR
'      If iR < 0 Then iR = 0
'      If iR > 255 Then iR = 255
'      iG = memBytes(2, ixs, iys)
'      iG = im * (ic - iG) \ 256 + iG
'      If iG < 0 Then iG = 0
'      If iG > 255 Then iG = 255
'      iB = memBytes(3, ixs, iys)
'      iB = im * (ic - iB) \ 256 + iB
'      If iB < 0 Then iB = 0
'      If iB > 255 Then iB = 255
'
'      If RGB(iR, iG, iB) = TransColor Then
'         If iR > 0 Then iR = iR - 1 Else iR = iR + 1
'      End If
'
'      memBack(ixs, iys) = RGB(iR, iG, iB)
'
'   Next iys
'Next ixs
'
'End Sub
'
