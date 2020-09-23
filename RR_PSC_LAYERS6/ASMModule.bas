Attribute VB_Name = "ASMModule"
' ASMModule.bas

Option Explicit
Option Base 1


'-----------------------------------------------------------------------------
' For calling machine code
Public Declare Function CallWindowProc Lib "USER32" Alias "CallWindowProcA" _
(ByVal lpMCode As Long, _
ByVal Long1 As Long, ByVal Long2 As Long, _
ByVal Long3 As Long, ByVal Long4 As Long) As Long

' Use:
' response = CallWindowProc(ptrMC, ptrStruc, 2&, 3&, 4&)
'                                  8         12  16  20
'-----------------------------------------------------------------------------

Public ptrmemBack As Long
Public ptrmemPic As Long

Public Wsrc As Long
Public Hsrc As Long
Public Wdes As Long
Public Hdes As Long
Public ztheta As Single
Public ptop As Long
Public pleft As Long
Public pright As Long
Public pbottom As Long

'-------------------------------
Public Type ASMTypeXFade
   W As Long
   H As Long
   T As Long
   L As Long
   iza As Long
   ptrmemPic As Long
   
   ptrmemBack As Long
   maxw As Long
   maxh As Long
   TransColor As Long
End Type

Public ASMalpha As ASMTypeXFade

'-------------------------------
Public Type ASMRot   ' Input
   Wsrc As Long      ' memPic width picFull(PicNum)
   Hsrc As Long
   Wdes As Long      ' memBack width
   Hdes As Long
   ptrmemPic As Long
   ptrmemBack As Long
   ztheta As Single  ' rotation angle (rad)
   TransColor As Long
                     ' Output:
   ptop As Long      ' min rect after rotation
   pleft As Long     ' to be extracted from memBack
   pright As Long    ' to new picFull(PicNum)
   pbottom As Long
End Type

Public ASMRotation As ASMRot

'-------------------------------
Public Type ASMSomeEffects
   Wsrc As Long         ' Source & dest same width & height
   Hsrc As Long
   ptrmemPic As Long    ' SOURCE: memPic width picDisplay
   ptrmemBack As Long   ' DEST: memBack width picDisplay
                        ' If Merge -> picDisplay
                        ' Else -> picDisplay & -> picFull(PicNum)
   FilterParam As Long
   TransColor As Long   ' No Effects on TransColor
End Type

Public ASMEffects As ASMSomeEffects
'-------------------------------


Public bMergeMC() As Byte
Public ptrMC As Long     ' = VarPtr(bMergeMC(1))
Public ptrStruc As Long  ' = VarPtr(ASMalpha.W)

Public bRotateMC() As Byte
Public ptrMC2 As Long    ' = VarPtr(bRotateMC(1))
Public ptrStruc2 As Long ' = VarPtr(ASMRotation.Wsrc)

Public bFEffectsMC() As Byte
Public ptrMC3 As Long    ' = VarPtr(bFEffectsMC(1))
Public ptrStruc3 As Long ' = VarPtr(ASMEffects.Wsrc)


Public Sub FillFixedASMalpha()
With ASMalpha
   .ptrmemBack = ptrmemBack
   .maxw = maxw
   .maxh = maxh
   .TransColor = TransColor
End With
End Sub

Public Sub FillVaryingASMalpha()
With ASMalpha
   .W = W
   .H = H
   .T = T
   .L = L
   .iza = iza
   .ptrmemPic = ptrmemPic
End With
End Sub

Public Sub FillASMRotation()
With ASMRotation
   .Wsrc = Wsrc
   .Hsrc = Hsrc
   .Wdes = Wdes
   .Hdes = Hdes
   .ptrmemPic = ptrmemPic
   .ptrmemBack = ptrmemBack
   .ztheta = ztheta
   .TransColor = TransColor

   .ptop = ptop
   .pleft = pleft
   .pright = pright
   .pbottom = pbottom
End With
End Sub

Public Sub FillASMEffects()
With ASMEffects
   .Wsrc = Wsrc
   .Hsrc = Hsrc
   .ptrmemPic = ptrmemPic     ' SOURCE: memPic
   .ptrmemBack = ptrmemBack   ' DEST:   memBack
   .FilterParam = FilterParam
   .TransColor = TransColor
End With
'response = CallWindowProc(ptrMC3, ptrStruc3, 2&, 3&, EffectsType)
'EffectsType=1 Sharp-Soft
'EffectsType=2 Dark-Bright

End Sub

Public Sub Loadmcode(InFile$, MCCode() As Byte)

''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''Load Machine code from bin file
'Dim BB As Byte
'Loadmcode PathSpec$ & "Merge.bin", bMergeMC()
'ptrMC = VarPtr(bMergeMC(1))
'ptrStruc = VarPtr(ASMalpha.W)
''BB = bMergeMC(1)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Dim MCSize As Long
'Load machine code into MCCode() byte array
On Error GoTo InFileErr
If Len(Dir$(InFile$)) = 0 Then
   MsgBox InFile$ & " missing", vbCritical, "Layer ASM"
   DoEvents
   Unload Form1
   End
End If
Open InFile$ For Binary As #1
MCSize = LOF(1)
If MCSize = 0 Then
InFileErr:
   MsgBox InFile$ & " zero size", vbCritical, "Layer ASM"
   Close
   Kill InFile$
   DoEvents
   Unload Form1
   End
End If

ReDim MCCode(1 To MCSize)
Get #1, , MCCode
Close #1
On Error GoTo 0
End Sub

