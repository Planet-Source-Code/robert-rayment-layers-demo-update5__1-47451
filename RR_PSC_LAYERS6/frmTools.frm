VERSION 5.00
Begin VB.Form frmTools 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00008000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Drawing Tools"
   ClientHeight    =   6555
   ClientLeft      =   150
   ClientTop       =   105
   ClientWidth     =   2310
   ControlBox      =   0   'False
   DrawWidth       =   2
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   437
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   154
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRGB 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   1785
      MaxLength       =   3
      TabIndex        =   53
      Text            =   "B"
      Top             =   5820
      Width           =   450
   End
   Begin VB.TextBox txtRGB 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   1785
      MaxLength       =   3
      TabIndex        =   52
      Text            =   "G"
      Top             =   5535
      Width           =   450
   End
   Begin VB.TextBox txtRGB 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   1785
      MaxLength       =   3
      TabIndex        =   51
      Text            =   "R"
      Top             =   5265
      Width           =   450
   End
   Begin VB.CheckBox chkReset 
      BackColor       =   &H00008000&
      Caption         =   "Reset"
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   6165
      Width           =   555
   End
   Begin VB.PictureBox PicSliders 
      AutoRedraw      =   -1  'True
      Height          =   210
      Index           =   3
      Left            =   60
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   48
      Top             =   5880
      Width           =   1575
   End
   Begin VB.PictureBox PicSliders 
      AutoRedraw      =   -1  'True
      Height          =   210
      Index           =   2
      Left            =   60
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   47
      Top             =   5640
      Width           =   1575
   End
   Begin VB.PictureBox PicSliders 
      AutoRedraw      =   -1  'True
      Height          =   210
      Index           =   1
      Left            =   60
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   46
      Top             =   5415
      Width           =   1575
   End
   Begin VB.PictureBox PicSliders 
      AutoRedraw      =   -1  'True
      Height          =   210
      Index           =   0
      Left            =   60
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   45
      Top             =   5175
      Width           =   1575
   End
   Begin VB.PictureBox picCBTemp 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   2505
      Picture         =   "frmTools.frx":0000
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   144
      TabIndex        =   44
      Top             =   7605
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.PictureBox picCB 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   45
      Picture         =   "frmTools.frx":067B
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   144
      TabIndex        =   43
      Top             =   4200
      Width           =   2160
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H00C0FFC0&
      Height          =   420
      Index           =   29
      Left            =   1860
      Picture         =   "frmTools.frx":72BD
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   " Fill "
      Top             =   2280
      Width           =   420
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H00C0FFC0&
      Height          =   420
      Index           =   28
      Left            =   1395
      Picture         =   "frmTools.frx":793F
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   " Triangle arrow "
      Top             =   2295
      Width           =   420
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H00C0FFC0&
      Height          =   420
      Index           =   27
      Left            =   930
      Picture         =   "frmTools.frx":7FC1
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   " Double arrow "
      Top             =   2295
      Width           =   420
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H00C0FFC0&
      Height          =   420
      Index           =   26
      Left            =   495
      Picture         =   "frmTools.frx":8643
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   " Feathered arrow "
      Top             =   2295
      Width           =   420
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H00C0FFC0&
      Height          =   420
      Index           =   25
      Left            =   75
      Picture         =   "frmTools.frx":8CC5
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   " Single arrow "
      Top             =   2295
      Width           =   420
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H00C0FFC0&
      Height          =   420
      Index           =   24
      Left            =   1815
      Picture         =   "frmTools.frx":9347
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   " Frustrum "
      Top             =   1875
      Width           =   420
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H00C0FFC0&
      Height          =   420
      Index           =   23
      Left            =   1395
      Picture         =   "frmTools.frx":99C9
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   " Parallelogram "
      Top             =   1890
      Width           =   420
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H00C0FFC0&
      Height          =   420
      Index           =   22
      Left            =   945
      Picture         =   "frmTools.frx":A04B
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   " T-piece "
      Top             =   1890
      Width           =   420
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H00C0FFC0&
      Height          =   420
      Index           =   21
      Left            =   510
      Picture         =   "frmTools.frx":A6CD
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   " Cross "
      Top             =   1890
      Width           =   420
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H00C0FFC0&
      Height          =   420
      Index           =   19
      Left            =   1860
      Picture         =   "frmTools.frx":AD4F
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   " Spray "
      Top             =   1440
      Width           =   420
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      Height          =   1500
      Left            =   30
      TabIndex        =   20
      Top             =   2625
      Width           =   2205
      Begin VB.VScrollBar VScroll1 
         Height          =   240
         Index           =   2
         LargeChange     =   2
         Left            =   525
         Max             =   0
         Min             =   16
         SmallChange     =   2
         TabIndex        =   32
         Top             =   825
         Width           =   195
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   240
         Index           =   0
         Left            =   525
         Max             =   4
         Min             =   16
         TabIndex        =   23
         Top             =   285
         Value           =   4
         Width           =   195
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   240
         Index           =   1
         Left            =   525
         Max             =   1
         Min             =   16
         TabIndex        =   22
         Top             =   555
         Value           =   1
         Width           =   195
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   240
         Index           =   3
         LargeChange     =   2
         Left            =   525
         Max             =   1
         Min             =   16
         SmallChange     =   2
         TabIndex        =   21
         Top             =   1095
         Value           =   1
         Width           =   195
      End
      Begin VB.Label LabSettings 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   75
         TabIndex        =   37
         Top             =   825
         Width           =   375
      End
      Begin VB.Label LabSettings 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   75
         TabIndex        =   36
         Top             =   1095
         Width           =   375
      End
      Begin VB.Label LabSettings 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   75
         TabIndex        =   35
         Top             =   555
         Width           =   375
      End
      Begin VB.Label LabSettings 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   75
         TabIndex        =   34
         Top             =   300
         Width           =   375
      End
      Begin VB.Label Label1 
         BackColor       =   &H00008000&
         Caption         =   " X Double spacing"
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Index           =   2
         Left            =   705
         TabIndex        =   33
         Top             =   825
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00008000&
         Caption         =   "Blend && Erase size"
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Index           =   0
         Left            =   720
         TabIndex        =   26
         Top             =   285
         Width           =   1380
      End
      Begin VB.Label Label1 
         BackColor       =   &H00008000&
         Caption         =   "DrawWidth"
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Index           =   1
         Left            =   720
         TabIndex        =   25
         Top             =   555
         Width           =   1410
      End
      Begin VB.Label Label1 
         BackColor       =   &H00008000&
         Caption         =   "Z Size"
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Index           =   3
         Left            =   750
         TabIndex        =   24
         Top             =   1080
         Width           =   1275
      End
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H00C0FFC0&
      Height          =   420
      Index           =   20
      Left            =   120
      Picture         =   "frmTools.frx":B3D1
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   " Star "
      Top             =   1890
      Width           =   420
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H00C0FFC0&
      Height          =   420
      Index           =   18
      Left            =   1440
      Picture         =   "frmTools.frx":BA53
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   " Spline "
      Top             =   1455
      Width           =   420
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H00C0FFC0&
      Height          =   420
      Index           =   17
      Left            =   1005
      Picture         =   "frmTools.frx":C0D5
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   " Polyline "
      Top             =   1440
      Width           =   420
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H00C0FFC0&
      Height          =   420
      Index           =   16
      Left            =   540
      Picture         =   "frmTools.frx":C757
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   " \ B-ribbon "
      Top             =   1440
      Width           =   420
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H00C0FFC0&
      Height          =   420
      Index           =   15
      Left            =   120
      Picture         =   "frmTools.frx":CDD9
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   " / F-ribbon "
      Top             =   1485
      Width           =   435
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H00C0FFC0&
      Height          =   420
      Index           =   14
      Left            =   1860
      Picture         =   "frmTools.frx":D45B
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   " Arch "
      Top             =   1050
      Width           =   420
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H00C0FFC0&
      Height          =   420
      Index           =   13
      Left            =   1425
      Picture         =   "frmTools.frx":DADD
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   " Shaded cirllipse "
      Top             =   1065
      Width           =   420
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H00C0FFC0&
      Height          =   420
      Index           =   12
      Left            =   1005
      Picture         =   "frmTools.frx":E15F
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   " Filled cirllipse "
      Top             =   1050
      Width           =   420
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H00C0FFC0&
      Height          =   420
      Index           =   11
      Left            =   570
      Picture         =   "frmTools.frx":E7E1
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   " Cirllipse "
      Top             =   1050
      Width           =   420
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H00C0FFC0&
      Height          =   420
      Index           =   10
      Left            =   120
      Picture         =   "frmTools.frx":EE63
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   " Vert shaded rectangle "
      Top             =   1050
      Width           =   435
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H00C0FFC0&
      Height          =   420
      Index           =   9
      Left            =   1860
      Picture         =   "frmTools.frx":F4E5
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   " Horz shaded rectangle "
      Top             =   585
      Width           =   420
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H00C0FFC0&
      Height          =   420
      Index           =   8
      Left            =   1440
      Picture         =   "frmTools.frx":FB67
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   " Filled rectangle "
      Top             =   585
      Width           =   420
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H00C0FFC0&
      Height          =   420
      Index           =   7
      Left            =   1005
      Picture         =   "frmTools.frx":101E9
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   " Rectangle "
      Top             =   585
      Width           =   420
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H00C0FFC0&
      Height          =   420
      Index           =   6
      Left            =   570
      Picture         =   "frmTools.frx":1086B
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   " Line "
      Top             =   585
      Width           =   420
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H00C0FFC0&
      Height          =   420
      Index           =   5
      Left            =   120
      Picture         =   "frmTools.frx":10EED
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   " Freedraw "
      Top             =   585
      Width           =   435
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H00C0FFC0&
      Height          =   420
      Index           =   4
      Left            =   1860
      Picture         =   "frmTools.frx":1156F
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   " Erase circle "
      Top             =   120
      Width           =   420
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H00C0FFC0&
      Height          =   420
      Index           =   3
      Left            =   1425
      Picture         =   "frmTools.frx":11BF1
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   " Erase rectangle "
      Top             =   120
      Width           =   420
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H00C0FFC0&
      Height          =   420
      Index           =   2
      Left            =   1005
      Picture         =   "frmTools.frx":12273
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   " Blend strong "
      Top             =   120
      Width           =   420
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H00C0FFC0&
      Height          =   420
      Index           =   1
      Left            =   570
      Picture         =   "frmTools.frx":128F5
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   " Blend medium "
      Top             =   120
      Width           =   420
   End
   Begin VB.OptionButton optTools 
      BackColor       =   &H00C0FFC0&
      Height          =   420
      Index           =   0
      Left            =   135
      Picture         =   "frmTools.frx":12F77
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   " Blend weak "
      Top             =   120
      Width           =   435
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No Trace"
      Height          =   435
      Left            =   1755
      TabIndex        =   54
      Top             =   6075
      Width           =   480
   End
   Begin VB.Label LabShowColor 
      BorderStyle     =   1  'Fixed Single
      Height          =   270
      Left            =   75
      TabIndex        =   49
      Top             =   6150
      Width           =   930
   End
End
Attribute VB_Name = "frmTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmTools  frmTools.frm

Option Explicit

Dim picW As Long
Dim picH As Long
Dim zstep As Single
Dim Cul As Long
Dim zCul As Single
Dim EffectsT As Long

Private Sub Form_Load()
   
   ' Keep on top
   i = SetWindowPos(frmTools.hwnd, HWND_TOPMOST, _
   frmToolsLeft, frmToolsTop, 4800 / STX, 3600 / STY, wFlags)

   aDrawStart = False
   aDrawMove = False
   LClickCount = 0
   RClickCount = 0
   
   TheDrawStyle = 5
   TheDrawWidth = 1
   BESize = 12
   Spacing = 0
   SRSize = 12
   Form1.picDisplay.ForeColor = Form1.picDisplay.BackColor
   EraseColor = Form1.picDisplay.ForeColor
   
   VScroll1(0) = BESize
   VScroll1(1) = TheDrawWidth
   VScroll1(2) = Spacing
   VScroll1(3) = SRSize
   
   LabSettings(0) = BESize
   LabSettings(1) = TheDrawWidth
   LabSettings(2) = Spacing   'X
   LabSettings(3) = SRSize    'Z
   
   For j = 0 To 5
      For i = 0 To 4
         With optTools(j * 5 + i)
            .Width = 28
            .Height = 28
            .Left = 3 + (i * 30)
            .Top = 3 + (j * 30)
         End With
      Next i
   Next j
   TheDrawStyle = 5
   optTools(5).Value = True
   
   Form1.Shape1(0).Visible = False
   Form1.Shape1(1).Visible = False
   Form1.Shape1(2).Visible = False
   Form1.Shape2(0).Visible = False
   Form1.Shape2(1).Visible = False
   
   ' Set up sliders
   Wsrc = picCB.Width: Hsrc = picCB.Height
   
   picW = PicSliders(0).Width
   picH = PicSliders(0).Height
   
   zstep = 2.2
   zCul = 40
   For i = 0 To picW
      Cul = CLng(zCul)
      PicSliders(0).Line (i, 0)-(i, picH), RGB(Cul, Cul, Cul)
      PicSliders(1).Line (i, 0)-(i, picH), RGB(Cul, 0, 0)
      PicSliders(2).Line (i, 0)-(i, picH), RGB(0, Cul, 0)
      PicSliders(3).Line (i, 0)-(i, picH), RGB(0, 0, Cul)
      
      If i Mod 8 = 0 Then
         PicSliders(0).Line (i, 0)-(i, PicSliders(0).Height), RGB(255 - Cul, 255 - Cul, 255 - Cul)
         PicSliders(1).Line (i, 0)-(i, PicSliders(0).Height), RGB(255 - Cul, 255 - Cul, 255 - Cul)
         PicSliders(2).Line (i, 0)-(i, PicSliders(0).Height), RGB(255 - Cul, 255 - Cul, 255 - Cul)
         PicSliders(3).Line (i, 0)-(i, PicSliders(0).Height), RGB(255 - Cul, 0, 0)
      End If

      zCul = zCul + zstep
      If zCul > 255 Then zCul = 255
   Next i


End Sub

Private Sub chkReset_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   chkReset.Value = 0

   BitBlt picCB.hdc, 0, 0, Wsrc, Hsrc, picCBTemp.hdc, 0, 0, vbSrcCopy
   picCB.Refresh
   
End Sub

Private Sub LabShowColor_Click()
   
   DrawColor = LabShowColor.BackColor
   Form1.picColor(1).BackColor = DrawColor
   Form1.picColor(1).Refresh

End Sub

Private Sub picCB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   Cul = GetPixel(picCB.hdc, X, Y)
   
   ' Show RGB
   bR = (Cul And &HFF&)
   bG = (Cul And &HFF00&) / &H100&
   bB = (Cul And &HFF0000) / &H10000
   
   aKeying = False
   txtRGB(0).Text = Trim$(Str$(bR))
   txtRGB(1).Text = Trim$(Str$(bG))
   txtRGB(2).Text = Trim$(Str$(bB))
   aKeying = True
   
   LabShowColor.BackColor = RGB(bR, bG, bB) 'Cul
   LabShowColor.Refresh

End Sub

Private Sub picCB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

   DrawColor = GetPixel(picCB.hdc, X, Y)
   Form1.picColor(1).BackColor = DrawColor
   Form1.picColor(1).Refresh

End Sub

Private Sub PicSliders_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

   If Button = vbLeftButton Then
      EffectsT = Index + 2
      Select Case EffectsT
      Case 2, 3, 4, 5  ' Dark-Bright & RGB +/-
         FilterParam = (X - 64) * 4 ' -255 -> +255
         If FilterParam < -255 Then FilterParam = -255
         If FilterParam > 255 Then FilterParam = 255

         Wsrc = picCB.Width: Hsrc = picCB.Height
         svmaxh = maxh
         svmaxw = maxw
         maxh = Hsrc
         maxw = Wsrc
         ReDim memBack(1 To Wsrc, 1 To Hsrc)
         memBack(1, 1) = TransColor
         memBack(Wsrc, Hsrc) = TransColor
         'Background Pic To memBack (picTemp = ORG picDisplay)
         GETDIBS picCBTemp.Image, 0

         ReDim memPic(1 To Wsrc, 1 To Hsrc)
         i = Wsrc * Hsrc * 4 ' = ImageSize
         CopyMemory memPic(1, 1), memBack(1, 1), i
         ptrmemPic = VarPtr(memPic(1, 1))
         ptrmemBack = VarPtr(memBack(1, 1))
         FillASMEffects
         response = CallWindowProc(ptrMC3, ptrStruc3, 2&, 3&, EffectsT)
         'Action in ASM is from memPic() to memBack()
   
         With picCB
            .Picture = LoadPicture
            .Refresh
         End With

         SetStretchBltMode picCB.hdc, HALFTONE
         
         ' Blit memBack to picCB
         StretchDIBits picCB.hdc, _
            0&, 0&, Wsrc, Hsrc, _
            0&, 0&, Wsrc, Hsrc, _
            memBack(1, 1), bmbac, 0, vbSrcCopy
         
         picCB.Refresh
   
         maxh = svmaxh
         maxw = svmaxw
         
         Erase memBack()
         Erase memPic()
   
      End Select
      
   End If

End Sub

Private Sub txtRGB_Change(Index As Integer)
   
   If Not aKeying Then Exit Sub
   
   If Not IsNumeric(txtRGB(Index).Text) Then txtRGB(Index).Text = "0"
   If Len(txtRGB(Index).Text) = 0 Then
      Cul = 0
      txtRGB(Index).Text = "0"
   Else
      Cul = Val(txtRGB(Index).Text)
      If Cul > 255 Then
         Cul = 255
         txtRGB(Index).Text = "255"
      End If
   End If
   If Cul < 0 Then
      Cul = 0
      txtRGB(Index).Text = "0"
   End If
   
   bR = Val(txtRGB(0).Text)
   bG = Val(txtRGB(1).Text)
   bB = Val(txtRGB(2).Text)
   
   LabShowColor.BackColor = RGB(bR, bG, bB)
    
End Sub



'### SELECT SHAPE STYLE #################################################
Private Sub optTools_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   TheDrawStyle = Index
   'Show_Instructions
         
   aDrawStart = False
   aDrawMove = False
   LClickCount = 0
   RClickCount = 0

   Form1.Shape1(0).Visible = False
   Form1.Shape1(1).Visible = False
   Form1.Shape1(2).Visible = False
   Form1.Shape2(0).Visible = False
   Form1.Shape2(1).Visible = False

End Sub
'### END SELECT SHAPE STYLE #################################################

Private Sub VScroll1_Change(Index As Integer)

   Select Case Index
   Case 0   ' Blend & Erase size  BESize
      BESize = VScroll1(0)
      LabSettings(0) = BESize
   Case 1   ' TheDrawWidth
      TheDrawWidth = VScroll1(1)
      LabSettings(1) = TheDrawWidth
   Case 2   ' Spacing  'X
      Spacing = VScroll1(2)
      LabSettings(2) = Spacing
   Case 3   ' SRSize  'Z
      SRSize = VScroll1(3)
      LabSettings(3) = SRSize
   End Select

End Sub



