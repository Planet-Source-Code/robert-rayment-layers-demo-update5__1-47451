VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00008000&
   Caption         =   " LAYERS"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10020
   DrawWidth       =   2
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   554
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   668
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkPlusHairs 
      BackColor       =   &H00008000&
      Caption         =   "+ Hairs"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   108
      Top             =   4350
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   3
      Left            =   9375
      Style           =   1  'Graphical
      TabIndex        =   106
      Top             =   255
      Width           =   120
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   2
      Left            =   9210
      Style           =   1  'Graphical
      TabIndex        =   105
      Top             =   285
      Width           =   120
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   1
      Left            =   9390
      Style           =   1  'Graphical
      TabIndex        =   104
      Top             =   75
      Width           =   120
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   0
      Left            =   9225
      Style           =   1  'Graphical
      TabIndex        =   103
      Top             =   75
      Width           =   120
   End
   Begin VB.CheckBox chkMagnifier 
      BackColor       =   &H00008000&
      Caption         =   "Magnifier"
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   102
      Top             =   4650
      Width           =   975
   End
   Begin VB.CheckBox chkToggleInstructions 
      BackColor       =   &H00008000&
      Caption         =   "Instructions"
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   101
      Top             =   4020
      Width           =   975
   End
   Begin VB.CheckBox chkTrace 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Trace OFF"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   99
      Top             =   5595
      Width           =   960
   End
   Begin VB.Frame fraShowPics 
      BackColor       =   &H00008000&
      Caption         =   "Select pic"
      ForeColor       =   &H00C0FFFF&
      Height          =   570
      Left            =   90
      TabIndex        =   91
      Top             =   6465
      Width           =   1005
      Begin VB.CommandButton cmdShowPic 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   390
         Picture         =   "Main.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   93
         ToolTipText     =   " Show next pic "
         Top             =   240
         Width           =   285
      End
      Begin VB.CommandButton cmdShowPic 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   90
         Picture         =   "Main.frx":1184
         Style           =   1  'Graphical
         TabIndex        =   92
         ToolTipText     =   " Show prev pic "
         Top             =   240
         Width           =   285
      End
      Begin VB.Label LabSelPicNum 
         Alignment       =   2  'Center
         BackColor       =   &H00008080&
         Caption         =   "0"
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   690
         TabIndex        =   95
         Top             =   255
         Width           =   225
      End
   End
   Begin VB.Frame fraSetTLIM 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   1260
      Left            =   1305
      TabIndex        =   79
      Top             =   585
      Width           =   915
      Begin VB.CommandButton cmdTLIM 
         BackColor       =   &H00C0FFFF&
         Caption         =   "X"
         Height          =   240
         Left            =   270
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   945
         Width           =   330
      End
      Begin VB.TextBox txtTLIM 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   81
         Text            =   "200"
         Top             =   570
         Width           =   585
      End
      Begin VB.Label Label4 
         BackColor       =   &H00008000&
         Caption         =   " Loop delay  ~ 200-2000"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   360
         Left            =   90
         TabIndex        =   80
         Top             =   150
         Width           =   795
      End
   End
   Begin VB.Frame fraDraw 
      BackColor       =   &H00008000&
      Caption         =   "Draw pic 0"
      ForeColor       =   &H00C0FFFF&
      Height          =   1770
      Left            =   30
      TabIndex        =   64
      Top             =   75
      Width           =   1170
      Begin VB.CommandButton cmdDraw 
         BackColor       =   &H00C0FFFF&
         Caption         =   "A&DD to 0"
         Height          =   300
         Index           =   4
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   97
         Top             =   1395
         Width           =   945
      End
      Begin VB.CommandButton cmdDraw 
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Cancel"
         Height          =   285
         Index           =   3
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   1095
         Width           =   795
      End
      Begin VB.CommandButton cmdDraw 
         BackColor       =   &H0080FF80&
         Caption         =   "&Accept"
         Height          =   285
         Index           =   0
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   210
         Width           =   795
      End
      Begin VB.CommandButton cmdDraw 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Redo"
         Height          =   285
         Index           =   1
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   495
         Width           =   795
      End
      Begin VB.CommandButton cmdDraw 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Undo"
         Height          =   285
         Index           =   2
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   795
         Width           =   795
      End
   End
   Begin VB.Frame fraLevelBar 
      BackColor       =   &H00008000&
      Caption         =   " Level "
      ForeColor       =   &H00C0FFFF&
      Height          =   855
      Left            =   4920
      TabIndex        =   52
      Top             =   6240
      Width           =   3255
      Begin VB.CommandButton cmdLevelBar 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cancel"
         Height          =   255
         Index           =   1
         Left            =   2205
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   525
         Width           =   660
      End
      Begin VB.CommandButton cmdLevelBar 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Accept"
         Height          =   255
         Index           =   0
         Left            =   2205
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   195
         Width           =   660
      End
      Begin VB.PictureBox picLevelBar 
         AutoRedraw      =   -1  'True
         Height          =   270
         Left            =   90
         ScaleHeight     =   14
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   129
         TabIndex        =   53
         Top             =   450
         Width           =   2000
         Begin VB.CommandButton cmdSingleEffect 
            BackColor       =   &H0000C000&
            Caption         =   "CLICK ME"
            Height          =   255
            Left            =   -15
            Style           =   1  'Graphical
            TabIndex        =   78
            Top             =   15
            Width           =   1965
         End
      End
      Begin VB.Label LabFilterParam 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Left            =   1605
         TabIndex        =   55
         Top             =   195
         Width           =   465
      End
      Begin VB.Label LabLevelBar 
         BackColor       =   &H00008000&
         Caption         =   "Dark-Bright"
         ForeColor       =   &H00C0FFFF&
         Height          =   225
         Left            =   180
         TabIndex        =   54
         Top             =   210
         Width           =   1305
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00C0FFFF&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   7  'Diagonal Cross
         Height          =   525
         Left            =   2910
         Top             =   225
         Width           =   270
      End
   End
   Begin VB.Frame fraInstructions 
      BackColor       =   &H00008000&
      Caption         =   " Instructions "
      ForeColor       =   &H00C0FFFF&
      Height          =   930
      Left            =   1920
      TabIndex        =   49
      Top             =   5280
      Width           =   5775
      Begin VB.CommandButton cmdHideInstructions 
         BackColor       =   &H00C0FFFF&
         Caption         =   "X"
         Height          =   225
         Left            =   5430
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   600
         Width           =   270
      End
      Begin VB.Label LabInstructions 
         BackColor       =   &H00C0FFFF&
         Caption         =   "LabInstructions"
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   150
         TabIndex        =   51
         Top             =   225
         Width           =   5160
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00C0FFFF&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   7  'Diagonal Cross
         Height          =   285
         Left            =   5400
         Top             =   210
         Width           =   285
      End
   End
   Begin VB.Frame fraColors 
      BackColor       =   &H00008000&
      Caption         =   "Colors"
      ForeColor       =   &H00C0FFFF&
      Height          =   1515
      Left            =   105
      TabIndex        =   39
      Top             =   1890
      Width           =   960
      Begin VB.PictureBox picColor 
         AutoRedraw      =   -1  'True
         BackColor       =   &H000000FF&
         Height          =   255
         Index           =   2
         Left            =   75
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   42
         ToolTipText     =   "Text color"
         Top             =   570
         Width           =   345
      End
      Begin VB.PictureBox picColor 
         AutoRedraw      =   -1  'True
         BackColor       =   &H000000FF&
         Height          =   255
         Index           =   1
         Left            =   75
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   41
         ToolTipText     =   "Draw color"
         Top             =   945
         Width           =   345
      End
      Begin VB.PictureBox picColor 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Index           =   0
         Left            =   75
         ScaleHeight     =   14
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   40
         ToolTipText     =   "Transparent color"
         Top             =   210
         Width           =   330
      End
      Begin VB.Label LabRGB 
         BackColor       =   &H00008000&
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   225
         Index           =   2
         Left            =   660
         TabIndex        =   48
         Top             =   1260
         Width           =   225
      End
      Begin VB.Label LabRGB 
         BackColor       =   &H00008000&
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   225
         Index           =   1
         Left            =   375
         TabIndex        =   47
         Top             =   1260
         Width           =   225
      End
      Begin VB.Label LabRGB 
         BackColor       =   &H00008000&
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   210
         Index           =   0
         Left            =   60
         TabIndex        =   46
         Top             =   1260
         Width           =   225
      End
      Begin VB.Label LabColors 
         BackColor       =   &H00008000&
         Caption         =   "Text"
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Index           =   2
         Left            =   495
         TabIndex        =   45
         Top             =   600
         Width           =   420
      End
      Begin VB.Label LabColors 
         BackColor       =   &H00008000&
         Caption         =   "Draw"
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Index           =   1
         Left            =   510
         TabIndex        =   44
         Top             =   960
         Width           =   420
      End
      Begin VB.Label LabColors 
         BackColor       =   &H00008000&
         Caption         =   "Trans"
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Index           =   0
         Left            =   480
         TabIndex        =   43
         Top             =   255
         Width           =   420
      End
   End
   Begin VB.CheckBox chkToggleThumbBar 
      BackColor       =   &H00008000&
      Caption         =   "Thumb Bar"
      ForeColor       =   &H00C0FFFF&
      Height          =   300
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   3690
      Width           =   975
   End
   Begin VB.CheckBox chkSHOWALL 
      BackColor       =   &H00C0FFFF&
      Caption         =   " SHOW ALL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   600
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   " Allows individual picture selection "
      Top             =   5850
      Width           =   840
   End
   Begin VB.CheckBox chkMERGE 
      BackColor       =   &H00C0FFFF&
      Caption         =   "MERGE ALL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   600
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   " Merge pictures using alphas "
      Top             =   4980
      Width           =   840
   End
   Begin VB.Frame fraAccRedoCancel 
      BackColor       =   &H00008000&
      Caption         =   "Lasso pic 0"
      ForeColor       =   &H00C0FFFF&
      Height          =   1320
      Left            =   4560
      TabIndex        =   32
      Top             =   165
      Width           =   1155
      Begin VB.CommandButton cmdAccRedoCancel 
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Cancel"
         Height          =   315
         Index           =   2
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   915
         Width           =   795
      End
      Begin VB.CommandButton cmdAccRedoCancel 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Redo"
         Height          =   315
         Index           =   1
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   570
         Width           =   795
      End
      Begin VB.CommandButton cmdAccRedoCancel 
         BackColor       =   &H0080FF80&
         Caption         =   "&Accept"
         Height          =   315
         Index           =   0
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   225
         Width           =   795
      End
   End
   Begin VB.Frame fraRotate 
      BackColor       =   &H00008000&
      Caption         =   "Rotate pic 0"
      ForeColor       =   &H00C0FFFF&
      Height          =   1575
      Left            =   5790
      TabIndex        =   26
      Top             =   165
      Visible         =   0   'False
      Width           =   1155
      Begin VB.CommandButton cmdAngleChange 
         BackColor       =   &H00C0FFFF&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   870
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   495
         Width           =   180
      End
      Begin VB.CommandButton cmdAngleChange 
         BackColor       =   &H00C0FFFF&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   630
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   495
         Width           =   180
      End
      Begin VB.CommandButton cmdAngleChange 
         BackColor       =   &H00C0FFFF&
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
         Height          =   195
         Index           =   1
         Left            =   390
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   495
         Width           =   180
      End
      Begin VB.CommandButton cmdAngleChange 
         BackColor       =   &H00C0FFFF&
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
         Height          =   195
         Index           =   0
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   495
         Width           =   180
      End
      Begin VB.CommandButton cmdRot 
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   1
         Left            =   840
         Picture         =   "Main.frx":14C6
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   780
         Width           =   210
      End
      Begin VB.CommandButton cmdRot 
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   0
         Left            =   135
         Picture         =   "Main.frx":1808
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   780
         Width           =   210
      End
      Begin VB.TextBox txtRotate 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   345
         MaxLength       =   4
         TabIndex        =   29
         Text            =   "-180"
         Top             =   750
         Width           =   435
      End
      Begin VB.CommandButton cmdRotAccCan 
         BackColor       =   &H0080FF80&
         Height          =   270
         Index           =   0
         Left            =   180
         Picture         =   "Main.frx":1B4A
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1170
         Width           =   375
      End
      Begin VB.CommandButton cmdRotAccCan 
         BackColor       =   &H00C0C0FF&
         Height          =   270
         Index           =   1
         Left            =   675
         Picture         =   "Main.frx":1E8C
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1170
         Width           =   345
      End
      Begin VB.Label Label1 
         BackColor       =   &H00008000&
         Caption         =   "Step  Deg"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   165
         Left            =   225
         TabIndex        =   30
         Top             =   270
         Width           =   675
      End
   End
   Begin VB.Frame fraSizer 
      BackColor       =   &H00008000&
      Caption         =   "Size pic 0"
      ForeColor       =   &H00C0FFFF&
      Height          =   1710
      Left            =   6990
      TabIndex        =   19
      Top             =   165
      Visible         =   0   'False
      Width           =   1170
      Begin VB.CommandButton cmdWHChange 
         BackColor       =   &H00C0FFFF&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   885
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   405
         Width           =   195
      End
      Begin VB.CommandButton cmdWHChange 
         BackColor       =   &H00C0FFFF&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   615
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   405
         Width           =   210
      End
      Begin VB.CommandButton cmdWHChange 
         BackColor       =   &H00C0FFFF&
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
         Height          =   195
         Index           =   1
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   405
         Width           =   210
      End
      Begin VB.CommandButton cmdWHChange 
         BackColor       =   &H00C0FFFF&
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
         Height          =   195
         Index           =   0
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   405
         Width           =   210
      End
      Begin VB.CommandButton cmdWH 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   3
         Left            =   870
         Picture         =   "Main.frx":21CE
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   975
         Width           =   240
      End
      Begin VB.CommandButton cmdWH 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   2
         Left            =   120
         Picture         =   "Main.frx":2510
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   975
         Width           =   210
      End
      Begin VB.CommandButton cmdWH 
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   1
         Left            =   855
         Picture         =   "Main.frx":2852
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   645
         Width           =   225
      End
      Begin VB.CommandButton cmdWH 
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Index           =   0
         Left            =   120
         Picture         =   "Main.frx":2B94
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   645
         Width           =   210
      End
      Begin VB.CommandButton cmdReSize 
         BackColor       =   &H00C0C0FF&
         Height          =   255
         Index           =   1
         Left            =   645
         Picture         =   "Main.frx":2ED6
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1350
         Width           =   360
      End
      Begin VB.CommandButton cmdReSize 
         BackColor       =   &H0080FF80&
         Height          =   255
         Index           =   0
         Left            =   180
         Picture         =   "Main.frx":3218
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1350
         Width           =   345
      End
      Begin VB.TextBox txtSize 
         Height          =   285
         Index           =   1
         Left            =   330
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   945
         Width           =   495
      End
      Begin VB.TextBox txtSize 
         Height          =   300
         Index           =   0
         Left            =   330
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   615
         Width           =   495
      End
      Begin VB.Label Label3 
         BackColor       =   &H00008000&
         Caption         =   "Step pix"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   165
         Left            =   270
         TabIndex        =   74
         Top             =   210
         Width           =   600
      End
   End
   Begin VB.Frame fraThumbBar 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Left            =   2055
      TabIndex        =   0
      Top             =   3435
      Width           =   6795
      Begin VB.PictureBox picChangeAlpha 
         AutoRedraw      =   -1  'True
         Height          =   300
         Left            =   2190
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   252
         TabIndex        =   100
         Top             =   330
         Width           =   3840
      End
      Begin VB.CommandButton cmdRatch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "<-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1215
         Style           =   1  'Graphical
         TabIndex        =   88
         ToolTipText     =   "Slide thumbs"
         Top             =   360
         Width           =   315
      End
      Begin VB.CommandButton cmdRatch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "->"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1545
         Style           =   1  'Graphical
         TabIndex        =   87
         ToolTipText     =   "Slide thumbs"
         Top             =   360
         Width           =   315
      End
      Begin VB.PictureBox picButtonContainer 
         BackColor       =   &H00008000&
         Height          =   1455
         Left            =   15
         ScaleHeight     =   93
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   71
         TabIndex        =   17
         Top             =   285
         Width           =   1125
         Begin VB.CommandButton cmdResizeMerged 
            BackColor       =   &H00E0E0E0&
            Caption         =   "RESIZE M"
            Height          =   225
            Left            =   30
            Style           =   1  'Graphical
            TabIndex        =   90
            ToolTipText     =   " Resize Merged "
            Top             =   1155
            Width           =   990
         End
         Begin VB.CommandButton cmdClipMerged 
            BackColor       =   &H00E0E0E0&
            Caption         =   "CLIP M"
            Height          =   225
            Left            =   30
            Style           =   1  'Graphical
            TabIndex        =   89
            ToolTipText     =   " Clip Merged "
            Top             =   930
            Width           =   990
         End
         Begin VB.CommandButton cmdTBRotate 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Rotate 0"
            Height          =   225
            Left            =   30
            Style           =   1  'Graphical
            TabIndex        =   84
            Top             =   705
            Width           =   990
         End
         Begin VB.CommandButton cmdTBResize 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Resize 0"
            Height          =   225
            Left            =   30
            Style           =   1  'Graphical
            TabIndex        =   83
            Top             =   480
            Width           =   990
         End
         Begin VB.CommandButton cmdTBLasso 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Lasso 0"
            Height          =   240
            Left            =   30
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   240
            Width           =   990
         End
         Begin VB.CommandButton cmdTBClip 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Clip 0"
            Height          =   270
            Left            =   30
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   -30
            Width           =   990
         End
      End
      Begin VB.CommandButton cmdHideThumbBar 
         BackColor       =   &H0080FFFF&
         Caption         =   "X"
         Height          =   210
         Index           =   1
         Left            =   6540
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   30
         Width           =   255
      End
      Begin VB.CommandButton cmdHideThumbBar 
         BackColor       =   &H0080FFFF&
         Caption         =   "X"
         Height          =   210
         Index           =   0
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   30
         Width           =   255
      End
      Begin VB.PictureBox picThumbContainer 
         BackColor       =   &H00C0C0C0&
         Height          =   1110
         Left            =   1125
         ScaleHeight     =   70
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   111
         TabIndex        =   1
         Top             =   630
         Width           =   1719
         Begin VB.CommandButton cmdSwapPics 
            Height          =   195
            Index           =   0
            Left            =   630
            Picture         =   "Main.frx":355A
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Swap"
            Top             =   375
            Width           =   300
         End
         Begin VB.CommandButton cmdMerge 
            BackColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   0
            Left            =   615
            Picture         =   "Main.frx":36A4
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "<- Merge"
            Top             =   705
            Width           =   300
         End
         Begin VB.PictureBox picThumb 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   0
            Left            =   15
            ScaleHeight     =   48
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   48
            TabIndex        =   24
            Top             =   300
            Width           =   720
         End
         Begin VB.OptionButton optSelect 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "pic 0"
            DownPicture     =   "Main.frx":37EE
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Index           =   0
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   30
            Width           =   525
         End
         Begin VB.Shape ShapeSmall 
            BorderColor     =   &H000000FF&
            Height          =   780
            Left            =   0
            Top             =   270
            Width           =   780
         End
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1860
         Picture         =   "Main.frx":4130
         ToolTipText     =   "-> Mouse down on shade bar to set alpha ->"
         Top             =   330
         Width           =   300
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00C0FFFF&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   7  'Diagonal Cross
         Height          =   210
         Index           =   3
         Left            =   1230
         Top             =   30
         Width           =   5190
      End
      Begin VB.Label LabMoveThumbBar 
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "       Thumb Bar"
         ForeColor       =   &H0080FFFF&
         Height          =   300
         Left            =   15
         MousePointer    =   5  'Size
         TabIndex        =   12
         Top             =   -15
         Width           =   6825
      End
      Begin VB.Label LabAlphaValue 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LabAlphaValue"
         Height          =   315
         Left            =   6045
         TabIndex        =   9
         Top             =   315
         Width           =   675
      End
   End
   Begin VB.HScrollBar HS 
      Height          =   240
      Left            =   825
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   7110
      Width           =   5370
   End
   Begin VB.VScrollBar VS 
      Height          =   5325
      Left            =   1500
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   75
      Width           =   255
   End
   Begin VB.PictureBox picFrame 
      AutoRedraw      =   -1  'True
      Height          =   6810
      Left            =   1905
      ScaleHeight     =   450
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   60
      Width           =   7260
      Begin VB.PictureBox picDisplay 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   5220
         Left            =   30
         ScaleHeight     =   348
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   452
         TabIndex        =   3
         Top             =   -15
         Width           =   6780
         Begin VB.PictureBox picFullBackUp 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   720
            Left            =   1830
            ScaleHeight     =   48
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   57
            TabIndex        =   98
            Top             =   2400
            Width           =   855
         End
         Begin VB.PictureBox picFullTempBack 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   720
            Left            =   810
            ScaleHeight     =   48
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   57
            TabIndex        =   96
            Top             =   2385
            Width           =   855
         End
         Begin VB.PictureBox picFullTemp 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   720
            Left            =   810
            ScaleHeight     =   48
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   57
            TabIndex        =   94
            Top             =   1545
            Width           =   855
         End
         Begin VB.PictureBox picTemp 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   570
            Left            =   75
            ScaleHeight     =   38
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   49
            TabIndex        =   14
            Top             =   495
            Width           =   735
         End
         Begin VB.PictureBox picFull 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Index           =   0
            Left            =   855
            ScaleHeight     =   51
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   53
            TabIndex        =   15
            Top             =   465
            Width           =   795
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00FFFFFF&
            BorderStyle     =   3  'Dot
            DrawMode        =   7  'Invert
            X1              =   12
            X2              =   39
            Y1              =   173
            Y2              =   173
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FFFFFF&
            BorderStyle     =   3  'Dot
            DrawMode        =   7  'Invert
            X1              =   26
            X2              =   26
            Y1              =   157
            Y2              =   188
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H00FFFFFF&
            BorderStyle     =   3  'Dot
            BorderWidth     =   2
            DrawMode        =   7  'Invert
            FillColor       =   &H00C0C0C0&
            Height          =   300
            Index           =   1
            Left            =   3630
            Shape           =   2  'Oval
            Top             =   1860
            Width           =   300
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   2
            DrawMode        =   7  'Invert
            FillColor       =   &H00C0C0C0&
            Height          =   300
            Index           =   0
            Left            =   3150
            Shape           =   1  'Square
            Top             =   1890
            Width           =   300
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   5
            DrawMode        =   7  'Invert
            FillColor       =   &H00FFFFFF&
            FillStyle       =   7  'Diagonal Cross
            Height          =   300
            Index           =   2
            Left            =   2685
            Shape           =   3  'Circle
            Top             =   1920
            Width           =   300
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   3
            DrawMode        =   7  'Invert
            FillColor       =   &H00FFFFFF&
            FillStyle       =   7  'Diagonal Cross
            Height          =   300
            Index           =   1
            Left            =   2265
            Shape           =   3  'Circle
            Top             =   1905
            Width           =   300
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FFFFFF&
            DrawMode        =   7  'Invert
            FillColor       =   &H00FFFFFF&
            FillStyle       =   7  'Diagonal Cross
            Height          =   300
            Index           =   0
            Left            =   1860
            Shape           =   3  'Circle
            Top             =   1890
            Width           =   300
         End
         Begin VB.Shape SR 
            BorderColor     =   &H00FFFFFF&
            BorderStyle     =   3  'Dot
            DrawMode        =   7  'Invert
            FillColor       =   &H00FFFFFF&
            FillStyle       =   6  'Cross
            Height          =   285
            Left            =   330
            Top             =   1800
            Width           =   285
         End
         Begin VB.Line S 
            Index           =   0
            X1              =   24
            X2              =   34
            Y1              =   86
            Y2              =   102
         End
      End
   End
   Begin VB.Label LabTog 
      BackColor       =   &H00008000&
      Caption         =   "Togglers"
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   255
      TabIndex        =   107
      Top             =   3450
      Width           =   600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      Index           =   1
      X1              =   84
      X2              =   84
      Y1              =   2
      Y2              =   475
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   82
      X2              =   82
      Y1              =   1
      Y2              =   474
   End
   Begin VB.Label LabGreen 
      Height          =   570
      Index           =   1
      Left            =   975
      TabIndex        =   86
      Top             =   5850
      Width           =   135
   End
   Begin VB.Label LabGreen 
      Height          =   570
      Index           =   0
      Left            =   975
      TabIndex        =   85
      Top             =   4980
      Width           =   135
   End
   Begin VB.Label LabNSP 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " NumOfStoredPics ="
      Height          =   285
      Left            =   4290
      TabIndex        =   8
      Top             =   7650
      Width           =   1980
   End
   Begin VB.Label LabXY 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LabXY"
      Height          =   285
      Left            =   6285
      TabIndex        =   7
      Top             =   7680
      Width           =   3690
   End
   Begin VB.Label LabInfo 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LabInfo"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   6
      Top             =   7650
      Width           =   4290
   End
   Begin VB.Menu mnuFile 
      Caption         =   " &FILE "
      Begin VB.Menu mnuOpen 
         Caption         =   "&Load pic 0"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuReLoad 
         Caption         =   "&ReLoad pic 0"
         Shortcut        =   ^R
      End
      Begin VB.Menu zBrk2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLoadTLayer 
         Caption         =   "Load &Transparent Layer 0"
         Shortcut        =   ^T
      End
      Begin VB.Menu zbrk3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveStoredPic 
         Caption         =   "&Save stored pic 0"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveDisplay 
         Caption         =   "Save &Display"
         Shortcut        =   ^D
      End
      Begin VB.Menu zBrk4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClipBoard 
         Caption         =   "&Copy display to clipboard"
         Index           =   0
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuClipBoard 
         Caption         =   "&Paste clipboard to pic 0"
         Index           =   1
         Shortcut        =   ^P
      End
      Begin VB.Menu zbrk5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSetLoopDelay 
         Caption         =   "S&et Loop Delay"
         Shortcut        =   ^E
      End
      Begin VB.Menu zBrk6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuColor 
      Caption         =   "&Colors"
      Begin VB.Menu mnuColorN 
         Caption         =   " &Draw color"
         Index           =   0
      End
      Begin VB.Menu mnuColorN 
         Caption         =   " &Text color"
         Index           =   1
      End
      Begin VB.Menu mnuColorN 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuColorN 
         Caption         =   " &PicA Draw color"
         Index           =   3
      End
      Begin VB.Menu mnuColorN 
         Caption         =   "-"
         Index           =   4
      End
   End
   Begin VB.Menu mnuDrawToggle 
      Caption         =   "&Draw_OFF"
   End
   Begin VB.Menu mnuText 
      Caption         =   "&Text"
   End
   Begin VB.Menu mnuEffects 
      Caption         =   "&Effects"
      Begin VB.Menu mnuPICorMERGED 
         Caption         =   "Effects on INDIVIDUAL pic 0"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuPICorMERGED 
         Caption         =   "Effects on whole MERGED picture"
         Index           =   1
      End
      Begin VB.Menu mnuPICorMERGED 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuSingleEffects 
         Caption         =   " &Single Effects"
         Begin VB.Menu mnuSEffects 
            Caption         =   "&0  Invert"
            Index           =   0
         End
         Begin VB.Menu mnuSEffects 
            Caption         =   "&1  Elliptic"
            Index           =   1
         End
         Begin VB.Menu mnuSEffects 
            Caption         =   "&2  Flip horizontal"
            Index           =   2
         End
         Begin VB.Menu mnuSEffects 
            Caption         =   "&3  Mirror right half"
            Index           =   3
         End
         Begin VB.Menu mnuSEffects 
            Caption         =   "&4  Mirror left half"
            Index           =   4
         End
      End
      Begin VB.Menu zBrk10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEffectsN 
         Caption         =   " &1  Sharp-Soft"
         Index           =   1
      End
      Begin VB.Menu mnuEffectsN 
         Caption         =   " &2  Dark-Bright"
         Index           =   2
      End
      Begin VB.Menu mnuEffectsN 
         Caption         =   " &3  Red +/-"
         Index           =   3
      End
      Begin VB.Menu mnuEffectsN 
         Caption         =   " &4  Green +/-"
         Index           =   4
      End
      Begin VB.Menu mnuEffectsN 
         Caption         =   " &5  Blue +/-"
         Index           =   5
      End
      Begin VB.Menu mnuEffectsN 
         Caption         =   " &6  Diffuse"
         Index           =   6
      End
      Begin VB.Menu mnuEffectsN 
         Caption         =   " &7  Relief"
         Index           =   7
      End
      Begin VB.Menu mnuEffectsN 
         Caption         =   " &8  Metallic"
         Index           =   8
      End
      Begin VB.Menu mnuEffectsN 
         Caption         =   " &9  Flute Up  \ /"
         Index           =   9
      End
      Begin VB.Menu mnuEffectsN 
         Caption         =   " &A Flute Down / \"
         Index           =   10
      End
      Begin VB.Menu mnuEffectsN 
         Caption         =   " &B Ripple"
         Index           =   11
      End
      Begin VB.Menu mnuEffectsN 
         Caption         =   " &C Rounded rectangle"
         Index           =   12
      End
      Begin VB.Menu mnuEffectsN 
         Caption         =   " &D Tile"
         Index           =   13
      End
      Begin VB.Menu mnuEffectsN 
         Caption         =   " &E Horizontal shading ||"
         Index           =   14
      End
      Begin VB.Menu mnuEffectsN 
         Caption         =   " &F Vertical shading =="
         Index           =   15
      End
   End
   Begin VB.Menu mnuACTIONS 
      Caption         =   "&ACTIONS"
      Begin VB.Menu mnuClip 
         Caption         =   " &Clip pic"
      End
      Begin VB.Menu mnuLasso 
         Caption         =   " &Lasso pic"
      End
      Begin VB.Menu mnuResizeApic 
         Caption         =   " &Resize pic"
      End
      Begin VB.Menu mnuRotatePic 
         Caption         =   " R&otate pic"
      End
      Begin VB.Menu zbrk12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClipMerged 
         Caption         =   " C&LIP MERGED"
      End
      Begin VB.Menu mnuResizeMerged 
         Caption         =   " R&ESIZE MERGED"
      End
      Begin VB.Menu zBrk14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRepositionLostPic 
         Caption         =   " Re-&position (lost) pic"
      End
      Begin VB.Menu zbrk16 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClearALLpics 
         Caption         =   " Clear &All pics but 0"
      End
      Begin VB.Menu mnuClearNpics 
         Caption         =   " Clear p&ic"
      End
      Begin VB.Menu zBrk20 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuTrace 
      Caption         =   "T&RACE"
      Begin VB.Menu mnuTraceN 
         Caption         =   "Trace O&N (Merged Mode)"
         Index           =   0
      End
      Begin VB.Menu mnuTraceN 
         Caption         =   "Trace O&FF"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu mnuTraceN 
         Caption         =   "&Add Trace to pic"
         Enabled         =   0   'False
         Index           =   2
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Form1  Main.frm

' LAYERS DEMO  by  Robert Rayment  Aug 2003

' Updates
' 16 Aug
' FileOps correction for cancelled file input @ mnuOpen
' 15 Aug
' Clip auto-scrolling adjusted
' 11 Aug
' 1. Re-load needed to cancel FileOps when finished else stopped
'    Drawing
' 8 Aug
' 1. Check for FileOps at picDisplay_Mouse.. to avoid inteference
'    when Double-Clicking to Open a file over the display.
' 7 Aug
' 1. Redo saving picTemp @ LClickCount=0 in picDisplay Mouse_Up (Sent to PSC)
' 2. TheDrawStyle changed from 0 to 5 (Freedraw) in Form_Load (Sent to PSC)
' 3. Addition to cmdMerge (Merging pairs) to re-select picNum correctly (Sent to PSC)

Option Explicit

Dim OS As New OSDialog

Dim aClipBoard As Boolean
Dim aClipBoardUsed As Boolean

Dim aFileOps As Boolean

Dim fraX As Long  ' For moving frames
Dim fraY As Long
Dim fraLeft As Long
Dim fraTop As Long

Dim FW As Long
Dim FH As Long

Dim picFullX As Long ' For moving pics
Dim picFullY As Long
Dim jpic As Long     ' pic index for moving merged pics

Dim LeftPicNum As Long    ' To keep track of Thumb-bar brace





Private Sub Form_Load()

   ''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' Preferences & Defaults
   
   MaxNumOfPics = 21  ' optSelect(0 - 20)
   'MaxNumOfPics = 31  ' optSelect(0 - 30)
   
   TLIM = 20   ' Sleep API for delaying resizing & rotating
               ' Set to a larger value for faster machines
   
   ' Fix common transparent color
   TransColor = RGB(223, 223, 223) ' NB in memBack & memPic
                                   ' BGRA =223,223,223,0
   TLayerWidth = 256    ' Default Transparent
   TLayerHeight = 256   ' layer size
   
   aShowInstructions = True
   
   ' Set Resize limits & default increments
   WHMin = 8
   WHMax = 1024
   WHChange = 8
   
   ' Set Default Rotation increment deg
   StepAngle = 5
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''
   
   Me.ScaleMode = vbPixels
   
   STX = Screen.TwipsPerPixelX
   STY = Screen.TwipsPerPixelY
   
   ' TEST BUTTONS see end of Form1
   For i = 0 To 3
      Command1(i).Visible = False
   Next i
   
   
   ' Merge ShowAll indicators
   LabGreen(0).BackColor = RGB(0, 128, 0)
   LabGreen(1).BackColor = RGB(0, 128, 0)
   
   aTrace = False
   cmdDraw(4).Enabled = False ' ADD trace
   mnuTrace.Enabled = False
   picFullTemp.Visible = False
   picFullTempBack.Visible = False
   picFullBackUp.Visible = False
   
   picColor(0).Cls
   picColor(0).BackColor = TransColor
   picColor(0).Refresh
   
   picThumb(0).BackColor = TransColor
   picThumb(0).Refresh
   
   picDisplay.BackColor = TransColor
   picDisplay.Refresh
   
   LineUpThumbBarElements
   
   ' Alpha picbox shade bar
   picChangeAlpha.Width = (256 + 6) * STX
   For i = 0 To 255
      picChangeAlpha.Line (i, 0)-(i, picChangeAlpha.Height), RGB(i, i, i)
      If i Mod 5 = 0 Then
         picChangeAlpha.Line (i, 0)-(i, picChangeAlpha.Height), RGB(255 - i, 255 - i, 255 - i)
      End If
   Next i
   picChangeAlpha.Refresh
   
   aClipBoard = False
   aClipBoardUsed = False
   aFileOps = False
   aMerged = False
   aResizeMerge = False
   aShowON = True
   aShowALLON = False
   aClipON = False
   aLasso = False
   aDRAW = False
   aPicAColor = False
   aEffects = False
   aHelp = False
   ADDED = False
   aMagON = False
   aHairs = False
   ' + hairs
   Line2.Visible = False
   Line3.Visible = False
   
   mnuPICorMERGED(0).Checked = True
   aIndividual = True
   mnuPICorMERGED(1).Checked = False
   mnuPICorMERGED(1).Enabled = False
   
   
   ' To detect screen res change on the fly
   ORG_ScreenWidth = GetSystemMetrics(SM_CXSCREEN)
'-----------------------------------
   
   picDisplay.Top = 0   'In picFrame
   picDisplay.Left = 0
   
   picDisplay.Print "picDisplay"
'-----------------------------------

   PathSpec$ = App.Path
   If Right$(PathSpec$, 1) <> "\" Then PathSpec$ = PathSpec$ & "\"
   
   optSelect(0).Value = True
   ' Also sets PicNum
   
   NumOfStoredPics = 0
   LabNSP = " NumOfStoredPics =" & Str$(NumOfStoredPics)
   
   ReDim PicAlpha(1 To MaxNumOfPics - 1)
   ' Init PicAlpha values @ 50%
   For i = 1 To MaxNumOfPics - 1
      PicAlpha(i) = 128
   Next i
   LabAlphaValue = Str$(100 * 128 / 256) & " %"
   
   ' Disable all but first optSelect
   ' so pics must be loaded in order
   For i = 1 To MaxNumOfPics - 1
      optSelect(i).Enabled = False
   Next i
   
   ' Disable Swap & Merge buttons
   For i = 0 To MaxNumOfPics - 2
      cmdSwapPics(i).Enabled = False
      cmdMerge(i).Enabled = False
   Next i
   
   ReDim StoreFileSpec$(0)
   ReDim PicWidth(0)
   ReDim PicHeight(0)
   
   mnuDrawToggle.Enabled = False
   mnuText.Enabled = False
   mnuEffects.Enabled = False
   mnuACTIONS.Enabled = False
   mnuRepositionLostPic.Enabled = False
   mnuClearALLpics.Enabled = False
   mnuClearNpics.Enabled = False
   chkMERGE.Enabled = False
   
   SmallPicW = picThumb(0).Width
   SmallPicH = picThumb(0).Height

   cmdTBClip.Enabled = False
   cmdTBLasso.Enabled = False
   cmdTBResize.Enabled = False
   cmdTBRotate.Enabled = False
   
   chkSHOWALL.Enabled = False
   chkMERGE.Enabled = False
   
   TheDrawStyle = 0
   TheDrawWidth = 1
   EffectsType = 5
   
'-----------------------------------
   fraSizer.Visible = False
   fraSizer.Top = 5
   fraSizer.Left = 2
   fraRotate.Visible = False
   fraRotate.Top = 5
   fraRotate.Left = 2
   fraAccRedoCancel.Visible = False
   fraAccRedoCancel.Top = 5
   fraAccRedoCancel.Left = 2
   fraInstructions.Visible = False
   fraDraw.Visible = False
   fraSetTLIM.Visible = False
   
   ' Effects
   cmdSingleEffect.Visible = False
   fraLevelBar.Visible = False
   For i = 0 To 255 Step 2
      j = RGB(i, i, i)
      If i = 128 Then j = RGB(0, 250, 250)
      picLevelBar.Line (i \ 2, 0)-(i \ 2, picLevelBar.Height), j
      If i Mod 8 = 0 Then
         picLevelBar.Line (i \ 2, 0)-(i \ 2, picLevelBar.Height), RGB(255 - i, 255 - i, 255 - i)
      End If

   Next i
   picLevelBar.Refresh
   
'-----------------------------------
   ' Check color setting
   GetObjectAPI picDisplay.Image, Len(bmp), bmp

   If bmp.bmBitsPixel <= 16 Then
      response = MsgBox("BETTER IF IN 24 or 32-BIT TRUE COLOR" & vbCr & "CONTINUE ?", vbQuestion + vbYesNo, "Layers")
      If response = vbNo Then Form_Unload False
   End If
'-----------------------------------
   
   For i = 0 To MaxNumOfPics - 1
      With picFull(i)
         .Top = 0
         .Left = 0
         .Visible = False
      End With
   Next i
   picTemp.Visible = False

   
   ' Paint color
   picColor(1).Cls
   picColor(1).BackColor = RGB(255, 0, 0)
   DrawColor = RGB(255, 0, 0)
   
   ' Text color
   picColor(2).Cls
   picColor(2).BackColor = RGB(255, 0, 0)
   TextColor = RGB(255, 0, 0)

   ' Lasso
   S(0).Visible = False
   NumOfSLines = 1
   
   ' Clip
   SR.Visible = False
   
   ' Blender & Eraser shapes
   Shape1(0).Visible = False
   Shape1(1).Visible = False
   Shape1(2).Visible = False
   Shape2(0).Visible = False
   Shape2(1).Visible = False

' CreateRoundRectRgn Lib "gdi32" _
'(X1, Y1, X2, Y2, X3, Y3) As Long

' Shape Check boxes
   response = CreateRoundRectRgn(4, 4, chkMERGE.Width - 4, chkMERGE.Height - 4, 15, 15)
   SetWindowRgn chkMERGE.hwnd, response, True
   DeleteObject response

   response = CreateRoundRectRgn(4, 4, chkSHOWALL.Width - 4, chkSHOWALL.Height - 4, 15, 15)
   SetWindowRgn chkSHOWALL.hwnd, response, True
   DeleteObject response

''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''Load Machine code from bin file
'Dim AA$, BB$      ' Test
Loadmcode PathSpec$ & "Merge.bin", bMergeMC()
ptrMC = VarPtr(bMergeMC(1))
ptrStruc = VarPtr(ASMalpha.W)
'AA$ = Hex$(bMergeMC(1))
'BB$ = Hex$(bMergeMC(2))

Loadmcode PathSpec$ & "Rotate.bin", bRotateMC()
ptrMC2 = VarPtr(bRotateMC(1))
ptrStruc2 = VarPtr(ASMRotation.Wsrc)
'AA$ = Hex$(bRotateMC(1))
'BB$ = Hex$(bRotateMC(2))

Loadmcode PathSpec$ & "FEffects.bin", bFEffectsMC()
ptrMC3 = VarPtr(bFEffectsMC(1))
ptrStruc3 = VarPtr(ASMEffects.Wsrc)
'AA$ = Hex$(bFEffectsMC(1))
'BB$ = Hex$(bFEffectsMC(2))

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Me.Caption = " LAYERS DEMO  by Robert Rayment"

   ' Starting position of frmTools
   FW = GetSystemMetrics(SM_CXSCREEN)
   frmToolsLeft = FW - 2520 / STX   ' 640
   frmToolsTop = 72

End Sub

Private Sub Form_Resize()
'Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'Public Const SM_CXSCREEN = 0 'X Size of screen
'Public Const SM_CYSCREEN = 1 'Y Size of Screen

'Dim FW As Long
'Dim FH As Long

   If WindowState = vbMinimized Then
      If aHelp Then
         Unload frmHelp
         aHelp = False
      End If
      If aDRAW Then
          mnuDrawToggle_Click ' Closes frmTools
      End If
      ' Text form is modal
      Exit Sub
   End If
   
   If Me.Width < 7700 Then Me.Width = 7700
   
   
   fraThumbBar.Visible = False
   picDisplay.Visible = False
   
   STX = Screen.TwipsPerPixelX
   STY = Screen.TwipsPerPixelY
   
   GetExtras Me.BorderStyle
   ' Public ExtraBorder, ExtraHeight
   
   FW = GetSystemMetrics(SM_CXSCREEN)
   FH = GetSystemMetrics(SM_CYSCREEN)
   
   If WindowState <> vbMaximized Then
      Me.Height = 0.885 * FH * STX + ExtraHeight
   Else
      If FW <> ORG_ScreenWidth Then
         MsgBox " WINDOW MAXIMIZED - " & vbCr & " SET NORMAL SIZE BEFORE" & vbCr & " CHANGING SCREEN RES AGAIN"
      End If
   End If
   
   ' 16, 48,100 Win98 Cap+Menu+Border
   
      picFrame.Width = Me.Width / STX - ExtraBorder - 108 - picFrame.Left
      picFrame.Height = Me.Height / STY - ExtraHeight - 54 - picFrame.Top
      LabInfo.Top = picFrame.Height + picFrame.Top + 36
      LabNSP.Top = LabInfo.Top 'Me.Height / STY - ExtraHeight - LabNSP.Height
      LabXY.Top = LabNSP.Top
      FixScrollbars picFrame, picDisplay, HS, VS
   
      ' Checker picFrame
      picFrame.BackColor = vbWhite
      For j = 0 To picFrame.Height Step 32
      For i = 0 To picFrame.Width Step 32
         picFrame.Line (i, j)-(i + 16, j + 16), &HD0E0D0, BF
         picFrame.Line (i + 16, j + 16)-(i + 32, j + 32), &HD0E0D0, BF
      Next i
      Next j
      picFrame.Refresh
      
      picDisplay.Visible = True
        
      Line1(0).y2 = LabInfo.Top
      Line1(1).y2 = LabInfo.Top
      
      Me.Show

      fraThumbBar.Top = Form1.Height / STY - fraThumbBar.Height - 60
      fraThumbBar.Visible = True
      
      frmToolsLeft = FW - 2520 / STX   ' 640
      frmToolsTop = 72
   
End Sub


Private Sub CLEAR_Incomplete_Actions()

   If aPicAColor And Not aDRAW Then
      aPicAColor = False
      fraAccRedoCancel.Visible = False
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
      
   If aLasso Then
      If NumOfSLines > 1 Then
         For i = 2 To NumOfSLines: Unload S(i - 1): Next
      End If
      S(0).Visible = False
      NumOfSLines = 1
      aLasso = False
      fraAccRedoCancel.Visible = False
      fraInstructions.Visible = False
   End If
   
   If aClipON Then
      ' At Clip Start or Clip Redo
      SR.Visible = False
      aClipON = False
      fraAccRedoCancel.Visible = False
      fraInstructions.Visible = False
   End If
   
   mnuDrawToggle.Enabled = True
   If aDRAW Then
      aDRAW = Not aDRAW
      
      Unload frmTools
      DoEvents
      
      fraInstructions.Visible = False
      mnuDrawToggle.Caption = "&Draw_OFF"
      
      mnuText.Enabled = True
      mnuEffects.Enabled = True
      mnuACTIONS.Enabled = True
      mnuColor.Enabled = True
      picDisplay.MousePointer = vbDefault
      fraDraw.Visible = False
   End If
   
   picDisplay.MousePointer = vbDefault
   
   fraSizer.Visible = False
   aResize = False
   
   fraRotate.Visible = False
   fraAccRedoCancel.Visible = False
   fraLevelBar.Visible = False
   
   aClipON = False
   aLasso = False
   
   aEffects = False
   mnuEffects.Enabled = True
   
   DoEvents

End Sub

Private Sub chkMERGE_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

   ' MERGE ALL
   chkMERGE.Value = 0
   
   mnuClipBoard(0).Enabled = True
   mnuPICorMERGED(1).Enabled = True

   ' Merge ShowAll indicator
   LabGreen(0).BackColor = RGB(128, 255, 128)
   LabGreen(1).BackColor = RGB(0, 128, 0)
   
   
   If aFileOps = True Then
      aFileOps = False
      Sleep 200   ' To avoid unwanted MERGE on loading pic
   End If
   
   CLEAR_Incomplete_Actions
   
   aMerged = True
   
   optSelect_Click (PicNum)
   
   picDisplay.SetFocus
   
   mnuTrace.Enabled = True
   
   If PicNum < NumOfStoredPics Then
      mnuTraceN(2).Caption = "&Add Trace to pic" & Str$(PicNum)
   End If

   
   MERGE

End Sub

Private Sub chkSHOWALL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

   ' SHOW ALL
   chkSHOWALL.Value = 0
   
   ' Merge ShowAll indicator
   LabGreen(0).BackColor = RGB(0, 128, 0)
   LabGreen(1).BackColor = RGB(128, 255, 128)


   If aFileOps = True Then
      aFileOps = False
      Sleep 200   ' To avoid unwanted SHOWALL on loading pic
   End If
   
   picDisplay.SetFocus
   
   mnuTraceN_Click 1
   mnuTrace.Enabled = False
   
   
   SHOWALL

   mnuClipBoard(0).Enabled = False

   mnuPICorMERGED(0).Checked = True
   aIndividual = True
   mnuPICorMERGED(1).Checked = False
   mnuPICorMERGED(1).Enabled = False

End Sub

Private Sub cmdShowPic_Click(Index As Integer)
' Show prev & next picture in Show mode

   Select Case Index
   Case 0   ' Prev pic
      If PicNum > 0 Then
         PicNum = PicNum - 1
         optSelect_Click (PicNum)
         optSelect(PicNum).Value = True
         LabSelPicNum = Str$(PicNum)
      End If
   Case 1   ' Next pic
      If PicNum <= NumOfStoredPics - 2 Then
         PicNum = PicNum + 1
         optSelect_Click (PicNum)
         optSelect(PicNum).Value = True
         LabSelPicNum = Str$(PicNum)
      End If
   End Select

End Sub


Private Sub cmdMerge_Click(Index As Integer)
' Index =0  0 <- Merge 1
'
' Index =2  2 <- Merge 3
' etc
' Index Max = MaxNumOfPics-1
'

' Merge picFull(Index+1) on to  picFull(Index)
' using picTemp
   
CLEAR_Incomplete_Actions

Screen.MousePointer = vbHourglass
DoEvents
   
   For i = 1 To NumOfStoredPics - 1
      picFull(i).Visible = False
   Next i
   
   svmaxw = maxw
   svmaxh = maxh
   
   maxw = picFull(Index).Width
   maxh = picFull(Index).Height
   
   ReDim memBack(1 To maxw, 1 To maxh)
   memBack(1, 1) = TransColor
   memBack(maxw, maxh) = TransColor
   ' NB Needs to be (1 To maxw) ??. (maxw) gives a streaky output ??!!
   
   ' picFull(Index) To memBack using maxw,maxh
   GETDIBS picFull(Index).Image, 0
   
   ptrmemBack = VarPtr(memBack(1, 1))

   FillFixedASMalpha
   
   W = picFull(Index + 1).Width
   H = picFull(Index + 1).Height
   T = picFull(Index + 1).Top - picFull(Index).Top
   L = picFull(Index + 1).Left - picFull(Index).Left
   
   ReDim memPic(1 To W, 1 To H)
   'Get picFull(Index+1) To memPic using W,H
   GETDIBS picFull(Index + 1).Image, 1
   
   iza = PicAlpha(Index + 1)   ' 0   -> 256
   
   ptrmemPic = VarPtr(memPic(1, 1))
   
   FillVaryingASMalpha

   response = CallWindowProc(ptrMC, ptrStruc, 2&, 3&, 4&)
   ' picFull(Index+1) merged with picFull(Index) is in
   ' memBack()  NB picFull(Index) size unchanged.
   
   ' Move memBack() to picFull(Index)
   
   SetStretchBltMode picFull(Index).hdc, HALFTONE
   
   StretchDIBits picFull(Index).hdc, _
      0&, 0&, maxw, maxh, _
      0&, 0&, maxw, maxh, _
      memBack(1, 1), bmbac, 0, vbSrcCopy
   
   picFull(Index).Refresh
   
   Erase memBack()
   Erase memPic()
   
   PicNum = Index
   
   If PicNum > 0 Then
      maxw = svmaxw
      maxh = svmaxh
   End If
   
   ' Prob unnec
   PicWidth(PicNum) = picFull(PicNum).Width
   PicHeight(PicNum) = picFull(PicNum).Height
   
   If aMerged Then
      
      ' Show new Thumb
      SetStretchBltMode picThumb(PicNum).hdc, HALFTONE
      
      StretchBlt picThumb(PicNum).hdc, 0, 0, _
         SmallPicW, SmallPicH, picFull(PicNum).hdc, _
         0, 0, PicWidth(PicNum), PicHeight(PicNum), vbSrcCopy
      
      picThumb(PicNum).Refresh
      
      MERGE
      
      optSelect(PicNum).Value = True
   
   Else  ' New picFull(PicNum)
   
      picFull_To_PicDisplay
   
      Display2picFull_picThumb
      optSelect_Click (PicNum)
      optSelect(PicNum).Value = True
   
  End If
  
FixScrollbars picFrame, picDisplay, HS, VS
Screen.MousePointer = vbDefault
DoEvents

End Sub


Private Sub SHOWALL()

' Show all pictures for positioning

CLEAR_Incomplete_Actions

mnuResizeMerged.Enabled = False

If NumOfStoredPics >= 2 Then

  ' PicNum = NumOfStoredPics
  ' optSelect(PicNum).Value = True
   
   aMerged = False
   
   optSelect_Click (PicNum)
   
   ' Background to picDisplay
   With picDisplay
      .Picture = LoadPicture
      .Width = picFull(0).Width
      .Height = picFull(0).Height
   End With
   
   BitBlt picDisplay.hdc, 0, 0, _
      picDisplay.Width, picDisplay.Height, _
      picFull(0).hdc, 0, 0, vbSrcCopy
   
   ' Make all picFulls visible
   ' picDisplay is their container
   For i = 1 To NumOfStoredPics - 1
      picFull(i).Visible = True
   Next i

   aShowON = True
   aShowALLON = True

End If

FixScrollbars picFrame, picDisplay, HS, VS

' Merge ShowAll indicator
LabGreen(0).BackColor = RGB(0, 128, 0)
LabGreen(1).BackColor = RGB(128, 255, 128)

End Sub


'#### Move picture brace ################################

Private Sub cmdRatch_Click(Index As Integer)
   
   If Index = 0 Then ' <
      If LeftPicNum < MaxNumOfPics - 1 Then
         picThumbContainer.Left = picThumbContainer.Left - 54 * STX
         LeftPicNum = LeftPicNum + 1
      End If
   Else  ' >
      If LeftPicNum > 0 Then
         picThumbContainer.Left = picThumbContainer.Left + 54 * STX
         LeftPicNum = LeftPicNum - 1
      End If
   End If
   
End Sub

Private Sub mnuACTIONS_Click()
   
   CLEAR_Incomplete_Actions

End Sub


'#### COLOR STUFF ##########################################

Private Sub LabColors_Click(Index As Integer)
   If mnuColor.Enabled Then PopupMenu mnuColor
End Sub

Private Sub mnuColor_Click()

If Not aDRAW Then
   Screen.MousePointer = vbDefault
   CLEAR_Incomplete_Actions
   mnuColorN(1).Enabled = True
Else
   mnuColorN(1).Enabled = False
End If

End Sub

Private Sub mnuColorN_Click(Index As Integer)
Dim CF As CFDialog
Dim TheColor As Long
Dim Mpos As POINTAPI

Select Case Index
Case 0
   If Not aDRAW Then
      LabInstructions = " Select Draw Color"
   End If
   
Case 1: LabInstructions = " Select Text Color"
Case 3:
        LabInstructions = " Pick Color from anywhere: Press SHIFT KEY to set Draw Color" _
           & vbCr & Space$(54) & "CTRL KEY or  [X]  to Cancel"
End Select
If aShowInstructions Then fraInstructions.Visible = True

   Select Case Index
   
   Case 0, 1
    
      Set CF = New CFDialog
   
      If CF.VBChooseColor(TheColor, , , , Me.hwnd) Then
         
         Select Case Index
         Case 0   ' DrawColor
            picColor(1).BackColor = TheColor
            DrawColor = TheColor
            'XorColor = DrawColor Xor 0
         Case 1   ' TextColor
            picColor(2).BackColor = TheColor
            TextColor = TheColor
         End Select
      
      End If
      
      Set CF = Nothing

   Case 3   ' PicA Draw color
      
      If aDRAW Then
         picDisplay.MousePointer = vbDefault
         svTheDrawStyle = TheDrawStyle
      End If
      
      
      Screen.MousePointer = vbCustom
      Screen.MouseIcon = LoadResPicture("PICACOLOR", vbResCursor)
      
      aPicAColor = True
      
      Do
         Call GetCursorPos(Mpos)
         response = GetDC(0&)   ' Get Device Context to whole screen
         TheColor = GetPixel(response, Mpos.X, Mpos.Y)
         picColor(1).BackColor = TheColor
         picColor(1).Refresh
         ReleaseDC 0&, response    ' Important to avoid using up resources
         ' Show RGB
         bR = (TheColor And &HFF&)
         bG = (TheColor And &HFF00&) / &H100&
         bB = (TheColor And &HFF0000) / &H10000
         LabRGB(0).Caption = Trim$(Str$(bR))
         LabRGB(1).Caption = Trim$(Str$(bG))
         LabRGB(2).Caption = Trim$(Str$(bB))
         
         If GetAsyncKeyState(VK_SHIFT) <> 0 Then   ' Shift key pressed
                                                   ' Set DrawColor
            If aDRAW Then
               aPicAColor = False
               Screen.MousePointer = vbDefault
               DrawColor = TheColor
               TheDrawStyle = svTheDrawStyle
               FIX_DRAW_CURSORS
               Show_Instructions
               frmTools.optTools(TheDrawStyle).Value = True
            Else
               DrawColor = picColor(1).BackColor
               aPicAColor = False
               Screen.MousePointer = vbDefault
            End If
         
         ElseIf GetAsyncKeyState(VK_CONTROL) <> 0 Then   ' Ctrl key pressed
                                                         ' Cancel
            If aDRAW Then
               aPicAColor = False
               Screen.MousePointer = vbDefault
               picColor(1).BackColor = DrawColor
               TheDrawStyle = svTheDrawStyle
               FIX_DRAW_CURSORS
               Show_Instructions
               frmTools.optTools(TheDrawStyle).Value = True
            Else
               picColor(1).BackColor = DrawColor
               aPicAColor = False
               Screen.MousePointer = vbDefault
            End If
         End If
         
         DoEvents
      Loop Until Not aPicAColor
      
   End Select

   If Not aDRAW Then fraInstructions.Visible = False

End Sub









Private Sub picColor_Click(Index As Integer)
   
   Select Case Index
   Case 0
      TransColor = picColor(Index).BackColor
   Case 1
      DrawColor = picColor(Index).BackColor
   Case 3
      TextColor = picColor(Index).BackColor
   End Select
   
End Sub
'#### END COLOR STUFF ##########################################


'#### Bring off-screen pic to coords 2,2 in picDisplay ################

Private Sub mnuRepositionLostPic_Click()
   
   Clear_LassoLines
   
   If PicNum < NumOfStoredPics Then
      picFull(PicNum).Top = 2
      picFull(PicNum).Left = 2
   Else
      MsgBox "Select a picture first", vbInformation, "Reposition pic"
   End If

End Sub

'#### START PIC SIZING STUFF ######################################

Private Sub cmdResizeMerged_Click()
' On Thumb bar
   mnuResizeMerged_Click

End Sub

Private Sub cmdTBResize_Click()
' On Thumb bar
   mnuResizeApic_Click

End Sub

Private Sub cmdWHChange_Click(Index As Integer)
   
   Select Case Index
   Case 0: WHChange = 1
   Case 1: WHChange = 2
   Case 2: WHChange = 5
   Case 3: WHChange = 8
   End Select
   DoEvents
   
End Sub


Private Sub mnuResizeMerged_Click()

   CLEAR_Incomplete_Actions
   
   aResizeMerge = True

   fraSizer.Caption = "Size"
   fraSizer.Visible = True
   ww = picDisplay.Width
   hh = picDisplay.Height

   txtSize(0).Text = Str$(ww)
   txtSize(1).Text = Str$(hh)
   
   If ww <= WHMax And ww >= WHMin And _
      hh <= WHMax And hh >= WHMin Then

         With picTemp
            .Picture = LoadPicture
            .Width = ww
            .Height = hh
            .Refresh
         End With

         ' Save original for Cancel
         ' & also to resize from
         BitBlt picTemp.hdc, 0, 0, ww, hh, _
         picDisplay.hdc, 0, 0, vbSrcCopy
         picTemp.Refresh
   Else
         MsgBox "Picture too big or too small, maybe clip it", vbInformation, "Layers - Sizing"
         fraSizer.Visible = False
         aResizeMerge = False
         Exit Sub
   End If

End Sub

Private Sub mnuResizeApic_Click()

   CLEAR_Incomplete_Actions
   
   If Not aMerged Then
      If PicNum >= NumOfStoredPics Then
         MsgBox " Select a picture first", vbExclamation, " Layer - Resizing"
         Exit Sub
      End If
      optSelect_Click (PicNum)
   End If
      
   If Not aResize Then
      
      If Not aResizeMerge And aMerged Then
         picFull_To_picTemp
      Else
         picDisplay_To_picTemp
      End If
   
   End If
   
   aResize = True
   
   If PicNum < NumOfStoredPics Then
      fraSizer.Caption = "Size pic" & Str$(PicNum)
      fraSizer.Visible = True
      ww = picFull(PicNum).Width
      hh = picFull(PicNum).Height
   Else
      MsgBox "Select a picture first", vbInformation, "Layers - Sizing"
   End If

   txtSize(0).Text = Str$(ww)
   txtSize(1).Text = Str$(hh)
   
   If ww < WHMin Or ww > WHMax Or _
      hh < WHMin Or hh > WHMax Then
      
      MsgBox "Picture too big or too small, maybe clip it", vbInformation, "Layers - Sizing"
      fraSizer.Visible = False
      aResize = True
      Exit Sub
   End If

End Sub

Private Sub picThumb_Click(Index As Integer)

If Index < NumOfStoredPics Then
   optSelect_Click Index
   optSelect(Index).Value = True
End If

End Sub

Private Sub txtSize_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

aKeying = False
If KeyCode = 13 Then
   KeyCode = 0
   aKeying = True
   cmdWH_MouseDown Index, 1, 0, 0, 0
End If

End Sub

Private Sub cmdWH_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

aDone = False

Do

   If Not IsNumeric(txtSize(0).Text) Then
      If Not aResizeMerge And aMerged Then
         txtSize(0).Text = Str$(picFull(PicNum).Width)
      Else
         txtSize(0).Text = Str$(picDisplay.Width)
      End If
   End If
   If Not IsNumeric(txtSize(1).Text) Then
      If Not aResizeMerge And aMerged Then
         txtSize(1).Text = Str$(picFull(PicNum).Height)
      Else
         txtSize(1).Text = Str$(picDisplay.Height)
      End If
   End If
   
   If Not aKeying Then
      Select Case Index
      Case 0  ' Smaller Width
         txtSize(0).Text = Val(txtSize(0).Text) - WHChange
      Case 1    ' Larger Width
         txtSize(0).Text = Val(txtSize(0).Text) + WHChange
      Case 2   ' Smaller  height
         txtSize(1).Text = Val(txtSize(1).Text) - WHChange
      Case 3   ' Large height
         txtSize(1).Text = Val(txtSize(1).Text) + WHChange
      End Select
   End If
   
   ww = Val(txtSize(0).Text)
   hh = Val(txtSize(1).Text)
   
   If ww < WHMin Then
      ww = WHMin
      txtSize(0).Text = Str$(WHMin)
   End If
   If hh < WHMin Then
      hh = WHMin
      txtSize(1).Text = Str$(WHMin)
   End If

   If ww > WHMax Then
      ww = WHMax
      txtSize(0).Text = Str$(WHMax)
   End If
   If hh > WHMax Then
      hh = WHMax
      txtSize(1).Text = Str$(WHMax)
   End If
   
   If Not aResizeMerge And aMerged Then
      ' Changing picture size in merge mode
      ' picTemp to picFull(PicNum)
      With picFull(PicNum)
         .Width = ww
         .Height = hh
         .Picture = LoadPicture
      End With
      
      SetStretchBltMode picFull(PicNum).hdc, HALFTONE
      
      StretchBlt picFull(PicNum).hdc, 0&, 0&, ww, hh, _
         picTemp.hdc, 0&, 0&, picTemp.Width, picTemp.Height, vbSrcCopy
      
      picFull(PicNum).Refresh

      DoEvents
      
      MERGE

   Else  ' aResizeMerge or individual pictures

      With picDisplay
         .Width = ww
         .Height = hh
         .Picture = LoadPicture
      End With

      DoEvents
      
      SetStretchBltMode picDisplay.hdc, HALFTONE
     
      StretchBlt picDisplay.hdc, 0&, 0&, ww, hh, _
         picTemp.hdc, _
         0&, 0&, picTemp.Width, picTemp.Height, vbSrcCopy
      
      picDisplay.Refresh

   End If   ' If aResizeMerge Then

   Sleep TLIM
   
   DoEvents

Loop Until aDone Or aKeying

aKeying = False

End Sub

Private Sub cmdWH_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

   aDone = True
   DoEvents

End Sub

Private Sub cmdReSize_Click(Index As Integer)
   
   'CLEAR_Incomplete_Actions
   
   Select Case Index
   Case 0   ' Accept
      
      If aResizeMerge Then
         ' Leave pic Display on screen for Save picDisplay
      
      ElseIf Not aResizeMerge And aMerged Then
         ' picFull(PicNum) changed in Merged Mode
         W = picFull(PicNum).Width
         H = picFull(PicNum).Height
         PicWidth(PicNum) = W
         PicHeight(PicNum) = H

         If PicNum = 0 Then
            maxw = W
            maxh = H
         End If
         
      Else
         ' Individual new picFull size on picDisplay
         W = picDisplay.Width
         H = picDisplay.Height
         PicWidth(PicNum) = W
         PicHeight(PicNum) = H
         
         If PicNum = 0 Then
            maxw = W
            maxh = H
         End If
          
         Display2picFull_picThumb   ' Using current PicNum, W & H
         FixScrollbars picFrame, picDisplay, HS, VS
         LabInfo = ShortFileSpec(a$, 36) & "   W =" & Str$(W) & "  H =" & Str$(H)
      
      End If
      
   Case 1   ' Cancel
      
      If aResizeMerge Then
         ' Leave pic Display on screen for Save picDisplay - Cancelled
         ' Original picDisplay in picTemp
         
         ww = picTemp.Width
         hh = picTemp.Height
         
         With picDisplay
            .Width = ww
            .Height = hh
            .Picture = LoadPicture
         End With
         BitBlt picDisplay.hdc, 0, 0, ww, hh, _
            picTemp.hdc, 0, 0, vbSrcCopy
         picDisplay.Refresh
         
      ElseIf Not aResizeMerge And aMerged Then
         ' picFull(PicNum) changed in Merged Mode - Cancelled
         ' Original picFull(PicNum) in picTemp
         W = picTemp.Width
         H = picTemp.Height
         
         With picFull(PicNum)
            .Width = W
            .Height = H
            .Picture = LoadPicture
         End With
         BitBlt picFull(PicNum).hdc, 0, 0, W, H, _
            picTemp.hdc, 0, 0, vbSrcCopy
         picFull(PicNum).Refresh
      
         PicWidth(PicNum) = W
         PicHeight(PicNum) = H
         
         MERGE
      
      Else
         ' New picFull size on picDisplay - Cancelled
         ' Original picFull(PicNum) in picTemp
         W = picTemp.Width
         H = picTemp.Height
         
         With picFull(PicNum)
            .Width = W
            .Height = H
            .Picture = LoadPicture
         End With
         BitBlt picFull(PicNum).hdc, 0, 0, W, H, _
            picTemp.hdc, 0, 0, vbSrcCopy
         picFull(PicNum).Refresh
         optSelect_Click (PicNum)
         
      End If
      
   End Select

   fraSizer.Visible = False
   aResize = False
   aResizeMerge = False
   FixScrollbars picFrame, picDisplay, HS, VS

End Sub
'#### END PIC SIZING STUFF ######################################

'#### ROTATE STUFF  ####################################

Private Sub cmdTBRotate_Click()
   
   mnuRotatePic_Click

End Sub

Private Sub cmdAngleChange_Click(Index As Integer)
' Gets new angle change

   Select Case Index
   Case 0: StepAngle = 1
   Case 1: StepAngle = 2
   Case 2: StepAngle = 5
   Case 3: StepAngle = 8
   End Select
   DoEvents
End Sub

Private Sub mnuRotatePic_Click()

   CLEAR_Incomplete_Actions
   
   If Not aMerged Then
      If PicNum >= NumOfStoredPics Then
         MsgBox " Select a picture first", vbExclamation, " Layer - Rotating"
         Exit Sub
      End If
      optSelect_Click (PicNum)
   End If
   
   If aMerged Then
      picFull_To_picTemp
      Wsrc = picTemp.Width
      Hsrc = picTemp.Height
   Else
      picDisplay_To_picTemp
      Wdes = picTemp.Width
      Hdes = picTemp.Height
   End If
   
   If PicNum < NumOfStoredPics Then
      fraRotate.Caption = "Rotate pic" & Str$(PicNum)
      fraRotate.Visible = True
   Else
      MsgBox "Select a picture first", vbInformation, "Layers - Rotating"
   End If

   txtRotate.Text = "0"

End Sub

Private Sub txtRotate_KeyUp(KeyCode As Integer, Shift As Integer)

   aKeying = False
   If KeyCode = 13 Then
      KeyCode = 0
      aKeying = True
      cmdRot_MouseDown 0, 1, 0, 0, 0
   End If

End Sub

Private Sub cmdRot_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

aDone = False

Do
   
   If Not IsNumeric(txtRotate.Text) Then txtRotate.Text = "0"
   
   If Not aKeying Then
      Select Case Index
      Case 0
         txtRotate.Text = Val(txtRotate.Text) - StepAngle
      Case 1
         txtRotate.Text = Val(txtRotate.Text) + StepAngle
      End Select
   End If
   
   ztheta = Val(txtRotate.Text)
   If ztheta > 360 Then
      ztheta = ztheta - 360
      txtRotate.Text = Str$(ztheta)
   End If
   If ztheta < -360 Then
      ztheta = ztheta = ztheta + 360
      txtRotate.Text = Str$(ztheta)
   End If
   ztheta = ztheta * pi# / 180
   
      If aMerged Then
         ' Rotating picture in merge mode
         ' rotate picTemp into picFull(PicNum)
         
         Wsrc = picTemp.Width
         Hsrc = picTemp.Height
         ReDim memPic(1 To Wsrc, 1 To Hsrc)
         W = Wsrc
         H = Hsrc
         GETDIBS picTemp.Image, 1
         ptrmemPic = VarPtr(memPic(1, 1))
         
         Wdes = Sqr(Wsrc * Wsrc + Hsrc * Hsrc) + 2
         Hdes = Wdes
         ReDim memBack(1 To Wdes, 1 To Hdes)
         memBack(1, 1) = TransColor
         memBack(Wdes, Hdes) = TransColor
         FillbacStruc Wdes, Hdes ' fill BITMAPINFOHEADER
         ptrmemBack = VarPtr(memBack(1, 1))
         
         pleft = 0
         pright = picFull(PicNum).Width
         ptop = 0
         pbottom = picFull(PicNum).Height
   
         FillASMRotation
   
         response = CallWindowProc(ptrMC2, ptrStruc2, 2&, 3&, ptrMC2)
         ' ASM rotates memPic (picTemp) into memBack (->picFull(PicNum)

         ' Reduce rotated rectangle
         pleft = ASMRotation.pleft
          If pleft > 1 Then pleft = pleft - 2
         pright = ASMRotation.pright
          If pright < Wdes - 1 Then pright = pright + 2
         Wdes = pright - pleft '+ 1

         pbottom = ASMRotation.pbottom
          If pbottom > 1 Then pbottom = pbottom - 2
         ptop = ASMRotation.ptop
          If ptop < Hdes - 1 Then ptop = ptop + 2
         Hdes = ptop - pbottom '+ 1
         
         With picFull(PicNum)
            .Picture = LoadPicture
            .Width = Wdes
            .Height = Hdes
            .Refresh
         End With
            
         SetStretchBltMode picFull(PicNum).hdc, HALFTONE
         
         ' Blit memBack to picDisplay
         StretchDIBits picFull(PicNum).hdc, _
            0&, 0&, Wdes, Hdes, _
            pleft, pbottom, Wdes, Hdes, _
            memBack(1, 1), bmbac, 0, vbSrcCopy
         picFull(PicNum).Refresh
         
         DoEvents
         
         MERGE
      
      Else
         ' Individual picture
         Wsrc = picFull(PicNum).Width
         Hsrc = picFull(PicNum).Height
         ReDim memPic(1 To Wsrc, 1 To Hsrc)
         memPic(1, 1) = TransColor
         memPic(Wsrc, Hsrc) = TransColor
         W = Wsrc
         H = Hsrc
         GETDIBS picFull(PicNum).Image, 1
         ptrmemPic = VarPtr(memPic(1, 1))
         
         Wdes = Sqr(Wsrc * Wsrc + Hsrc * Hsrc) + 2
         Hdes = Wdes
         ReDim memBack(1 To Wdes, 1 To Hdes)
         memBack(1, 1) = TransColor
         memBack(Wdes, Hdes) = TransColor
         FillbacStruc Wdes, Hdes ' fill BITMAPINFOHEADER
         ptrmemBack = VarPtr(memBack(1, 1))
         
         pleft = 0
         pright = picFull(PicNum).Width
         ptop = 0
         pbottom = picFull(PicNum).Height
   
         FillASMRotation
   
         response = CallWindowProc(ptrMC2, ptrStruc2, 2&, 3&, ptrMC2)
         ' ASM rotates memPic (picFull(PicNum))into memBack (->picDisPlay))
   
         ' Reduce rotated rectangle
         pleft = ASMRotation.pleft
          If pleft > 1 Then pleft = pleft - 2
         pright = ASMRotation.pright
          If pright < Wdes - 1 Then pright = pright + 2
         Wdes = pright - pleft '+ 1

         pbottom = ASMRotation.pbottom
          If pbottom > 1 Then pbottom = pbottom - 2
         ptop = ASMRotation.ptop
          If ptop < Hdes - 1 Then ptop = ptop + 2
         Hdes = ptop - pbottom '+ 1
      
         With picDisplay
            .Picture = LoadPicture
            .Width = Wdes
            .Height = Hdes
            .Refresh
         End With
         
         SetStretchBltMode picDisplay.hdc, HALFTONE
         
         ' Blit memBack to picDisplay
         StretchDIBits picDisplay.hdc, _
            0&, 0&, Wdes, Hdes, _
            pleft, pbottom, Wdes, Hdes, _
            memBack(1, 1), bmbac, 0, vbSrcCopy
            
         picDisplay.Refresh
         
      End If

   Sleep TLIM
   
   DoEvents

Loop Until aDone Or aKeying

aKeying = False

End Sub

Private Sub cmdRot_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

   aDone = True
   DoEvents

End Sub


Private Sub cmdRotAccCan_Click(Index As Integer)

   'CLEAR_Incomplete_Actions
   
   ' No aResizeMerge for rotate
   
   Select Case Index
   Case 0   ' Accept
      
      If aMerged Then
         ' picFull(PicNum) changed in Merged Mode
         W = picFull(PicNum).Width
         H = picFull(PicNum).Height
         PicWidth(PicNum) = W
         PicHeight(PicNum) = H
      
      
      Else
         ' New picFull(PicNum) picture on picDisplay
         Erase memBack()
         Erase memPic()
         
         W = Wdes
         H = Hdes
         ' Set new W & H
         PicWidth(PicNum) = W
         PicHeight(PicNum) = H
         
         If PicNum = 0 Then
            maxw = W 'pright - pleft
            maxh = H 'pbottom - ptop
         End If
         
         Display2picFull_picThumb   ' Using current PicNum, W & H
         
         FixScrollbars picFrame, picDisplay, HS, VS
         LabInfo = ShortFileSpec(a$, 36) & "   W =" & Str$(W) & "  H =" & Str$(H)
   
      End If
   
   Case 1   ' Cancel
         
      If aMerged Then
         ' picFull(PicNum) changed in Merged Mode - Cancelled
         ' Original picFull(PicNum) in picTemp

         W = picTemp.Width
         H = picTemp.Height
         
         With picFull(PicNum)
            .Width = W
            .Height = H
            .Picture = LoadPicture
         End With
         BitBlt picFull(PicNum).hdc, 0, 0, W, H, _
            picTemp.hdc, 0, 0, vbSrcCopy
         picFull(PicNum).Refresh
      
         PicWidth(PicNum) = W
         PicHeight(PicNum) = H
         
         MERGE
      
      
      Else
         ' New picFull size on picDisplay - Cancelled
         ' Original picFull(PicNum)unchanged
         
         ' Re-set old W & H
         W = PicWidth(PicNum)
         H = PicHeight(PicNum)
         
         With picDisplay
            .Picture = LoadPicture
            .Width = W
            .Height = H
            .Refresh
         End With
            
         BitBlt picDisplay.hdc, 0, 0, W, H, _
         picFull(PicNum).hdc, 0, 0, vbSrcCopy
         
         Display2picFull_picThumb   ' Using current PicNum, W & H
         
         FixScrollbars picFrame, picDisplay, HS, VS
         
         
         'ww = pright - pleft + 1
         'hh = pbottom - ptop + 1
         
         If PicNum = 0 Then
            maxw = W 'pright - pleft
            maxh = H 'pbottom - ptop
         End If
      
      End If
   
   End Select

   FixScrollbars picFrame, picDisplay, HS, VS
   LabInfo = ShortFileSpec(a$, 36) & "   W =" & Str$(W) & "  H =" & Str$(H)
   
   fraRotate.Visible = False

End Sub
'#### ROTATE STUFF  ####################################


'#### DRAW TOGGLER ###############################################

Private Sub mnuDrawToggle_Click()
   
   If aPicAColor Then
      aPicAColor = False
      fraAccRedoCancel.Visible = False
      Screen.MousePointer = vbDefault
   End If
   '-------------------------------------------------
   If NumOfSLines > 1 Then
      For i = 2 To NumOfSLines: Unload S(i - 1): Next
   End If
   S(0).Visible = False
   NumOfSLines = 1
   aLasso = False
   SR.Visible = False
   
   picDisplay.MousePointer = vbDefault
   
   fraSizer.Visible = False
   aResize = False
   fraRotate.Visible = False
   aClipON = False
   aLasso = False
   fraAccRedoCancel.Visible = False
   DoEvents
   
   If Not aMerged Then
      If NumOfStoredPics = 0 Then
         MsgBox "Load some pictures", vbInformation, "Layers - Drawing"
         Exit Sub
      End If
      If PicNum >= NumOfStoredPics Then
         MsgBox "Select a picture first or Merge All", vbInformation, "Layers - Drawing"
         Exit Sub
      End If
   End If
   
   '-------------------------------------------------
   aDRAW = Not aDRAW
   If aDRAW Then
      
      If aShowALLON Then
         optSelect_Click (PicNum)
      End If
      aDRAW = True
      
      If aMerged Then
         fraDraw.Caption = "Draw"
      Else
         fraDraw.Caption = "Draw pic" & Str$(PicNum)
      End If
      fraDraw.Visible = True
      fraThumbBar.Visible = False
      
      If aShowInstructions Then
         fraInstructions.Visible = True
         Show_Instructions
      End If
      FIX_DRAW_CURSORS
      
      picDisplay_To_picTemp

      'If aTrace Then
      '   mnuTraceN_Click 0
      'End If
      
      frmTools.Show vbModeless
      
      mnuDrawToggle.Caption = "&Draw_ ON"
      mnuText.Enabled = False
      mnuEffects.Enabled = False
      mnuACTIONS.Enabled = False

   Else
      
      Unload frmTools
      
      mnuDrawToggle.Caption = "&Draw_OFF"
      mnuText.Enabled = True
      mnuEffects.Enabled = True
      mnuACTIONS.Enabled = True
      mnuColor.Enabled = True
      
      picDisplay.MousePointer = vbDefault
      fraAccRedoCancel.Visible = False
      fraDraw.Visible = False
      fraInstructions.Visible = False
      fraThumbBar.Visible = True
      
      ' Hide + hairs
      Line2.Visible = False
      Line3.Visible = False
   
   End If
   
   Screen.MousePointer = vbDefault
   DoEvents

End Sub
'#### END DRAW TOGGLER ###############################################


'#### TRACING SET UP ############################################

Private Sub chkTrace_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   chkTrace.Value = 0
   If mnuTrace.Enabled Then PopupMenu mnuTrace
   
End Sub

Private Sub mnuTraceN_Click(Index As Integer)

   Select Case Index
   Case 0   ' Trace ON (Merged Mode)
      
      If aDRAW Then
         MsgBox " Set TRACE up before DRAWING!", vbInformation, " Layers - TRACE"
         Exit Sub
      End If
      
      If NumOfStoredPics > 1 And PicNum < NumOfStoredPics Then
         
         aTrace = True
         mnuTraceN(0).Checked = True
         mnuTraceN(1).Checked = False
         mnuTraceN(2).Enabled = True
         mnuTraceN(2).Caption = "&Add Trace to pic" & Str$(PicNum)
         cmdDraw(4).Enabled = True
         cmdDraw(4).Caption = "A&DD to" & Str$(PicNum)
         chkTrace.ForeColor = 255
         chkTrace.Caption = "Trace ON"
         
         ' picFull(PicNum) to PicFullTemp, also for drawing on
         picFull_To_picFullTemp
         
         ' Also picFull(PicNum) to picFullBackUp for Undo
         picFull_To_picFullBackUp

      Else
         MsgBox "Select or load more than 1 picture", vbInformation, "Layers - Trace Switch"
         Exit Sub
      End If
      
   Case 1   ' Trace OFF
      aTrace = False
      mnuTraceN(0).Checked = False
      mnuTraceN(1).Checked = True
      mnuTraceN(2).Enabled = False
      cmdDraw(4).Enabled = False
      chkTrace.ForeColor = 0
      chkTrace.Caption = "Trace OFF"
   
      ' picFullTemp ??
   
   Case 2   ' Add Trace to pic #
      
      cmdDraw_Click 4
      
   End Select

End Sub
'#### END TRACING SET UP ############################################


'#### DRAW ACCEPT,REDO,UNDO & CANCEL #################################

Private Sub cmdDraw_Click(Index As Integer)

If aDRAW Then
   Select Case Index
   Case 0   ' Accept
   
      If aPicAColor Then Exit Sub
      
      EraseDrawArrays
      
      maxw = svmaxw
      maxh = svmaxh
      
      If Not aMerged Then
         Display2picFull_picThumb   ' picDisplay to picFull(PicNum) &
      End If                        ' to picThumb(PicNum) (thumbs)
      
      picDisplay.MousePointer = vbDefault
      picDisplay.DrawMode = 13
      picDisplay.DrawWidth = 1
      fraAccRedoCancel.Visible = False
      fraInstructions.Visible = False
      mnuDrawToggle.Enabled = True
      mnuDrawToggle_Click  ' Come out of Draw mode
      mnuTraceN_Click 1
      mnuColor_Click
      ' Hide + hairs
      aHairs = False
      Line2.Visible = False
      Line3.Visible = False
      
   Case 1   ' Redo
      
      If aPicAColor Then Exit Sub
      
      Select Case TheDrawStyle
      Case 0, 1, 2   ' Blenders
         Erase memBytes()
      Case 5, 14, 15, 16, 17, 18, 19, 20 ' FREEDRAW,ARCH,RIBBONS,POLYLINES,SPLINE,SPRAY,STAR
         Erase ixPts()
         Erase iyPts()
      End Select
      
      If Not aMerged Then
         
         picTemp_To_picDisplay
      
      ElseIf aMerged And Not aTrace Then
         
         picTemp_To_picDisplay
      
      ElseIf aTrace And Not ADDED Then
      
         picTemp_To_picDisplay
         picFullTempBack_To_picFullTemp
   
      ElseIf aTrace And ADDED Then

         ' picFullTemp_To_picFullTempBack done
         ' before each shape
         
         ' Restore last picFullTempBack
         picFullTempBack_To_picFullTemp
            
         ' Also Restore last picFull(PicNum) from picFullTempBack
         picFullTempBack_To_picFull
         
         ' Restore Thumb
         SetStretchBltMode picThumb(PicNum).hdc, HALFTONE
         
         StretchBlt picThumb(PicNum).hdc, 0, 0, _
            SmallPicW, SmallPicH, picFull(PicNum).hdc, _
            0, 0, W, H, vbSrcCopy
         
         picThumb(PicNum).Refresh
         
         MERGE
         
         picDisplay_To_picTemp
         
      End If
   
   Case 2   ' Undo
      
      If aPicAColor Then Exit Sub
      
      If aMerged Then
      
         If aTrace Then
            
            ' Restore picFullTemp from picFullBackUp
            picFullBackUp_To_picFullTemp
            
            ' Restart picFullTempBack as well
            picFullBackUp_To_picFullTempBack
            
            ' Also picFull(PicNum)from picFullBackUp
            picFullBackUp_To_picFull
            
            
            ' Restore Thumb
            SetStretchBltMode picThumb(PicNum).hdc, HALFTONE
            
            StretchBlt picThumb(PicNum).hdc, 0, 0, _
               SmallPicW, SmallPicH, picFull(PicNum).hdc, _
               0, 0, W, H, vbSrcCopy
            
            picThumb(PicNum).Refresh
         
         End If
         
         MERGE ' Merged for TRACE & NON-TRACE
      
         picDisplay_To_picTemp
         
      Else  ' Individual pic Undo
         
         BitBlt picDisplay.hdc, 0, 0, picDisplay.Width, picDisplay.Height, _
            picFull(PicNum).hdc, 0, 0, vbSrcCopy
         
         picDisplay.Refresh
      
      End If
      
      
      Select Case DrawStyle
      Case 0, 1, 2   ' Blenders
         Erase memBytes()
      Case 5, 14, 15, 16, 17, 18, 19, 20 ' FREEDRAW,ARCH,RIBBONS,POLYLINES,SPLINE,SPRAY,STAR
         Erase ixPts()
         Erase iyPts()
      End Select
      
   Case 3   ' Cancel
      
      If aPicAColor Then
         aPicAColor = False
         Screen.MousePointer = vbDefault
         picColor(1).BackColor = DrawColor
         TheDrawStyle = svTheDrawStyle
         FIX_DRAW_CURSORS
         Show_Instructions
         frmTools.optTools(TheDrawStyle).Value = True
      Else
      
         If aMerged Then
            
            If aTrace Then
               
               ' Restore picFull(PicNum)from picFullBackUp
               picFullBackUp_To_picFull
               
            End If
            
            MERGE
         
         Else
            BitBlt picDisplay.hdc, 0, 0, picDisplay.Width, picDisplay.Height, _
               picFull(PicNum).hdc, 0, 0, vbSrcCopy
            
            picDisplay.Refresh
         End If
         
         EraseDrawArrays
         
         maxw = svmaxw
         maxh = svmaxh
   
         picDisplay.MousePointer = vbDefault
         picDisplay.DrawMode = 13
         picDisplay.DrawWidth = 1
         fraAccRedoCancel.Visible = False
         fraInstructions.Visible = False
         mnuDrawToggle.Enabled = True
         mnuDrawToggle_Click  ' DRAW_OFF
         mnuTraceN_Click 1
         mnuColor_Click
         ' Hide + hairs
         aHairs = False
         Line2.Visible = False
         Line3.Visible = False
      
      End If
   
   Case 4      ' ADD to pic #
         
      If aPicAColor Then Exit Sub
      
      Select Case TheDrawStyle
      Case 0 To 5, 29: Exit Sub ' Blenders, Erasers & Fill not traceable
      End Select
      
      ' Tranfers picFullTemp to picFull(PicNum)
      
      ' Assuming valid picNum & in Merge Mode
      ' picFullTemp.Visible = True  ' Test
   
         
      ADDED = True
      
      ' picFullTemp takes drwaings in trace mode
      ' ADD to picFull
      
      picFullTemp_To_picFull
      
      ' picFull to Thumb
      
      ' Show new Thumb
      SetStretchBltMode picThumb(PicNum).hdc, HALFTONE
      
      StretchBlt picThumb(PicNum).hdc, 0, 0, _
         SmallPicW, SmallPicH, picFull(PicNum).hdc, _
         0, 0, W, H, vbSrcCopy
      
      picThumb(PicNum).Refresh
      
      If aMerged Then
         MERGE
         picDisplay_To_picTemp
      End If
      
   End Select
   
End If   ' If aDRAW Then

End Sub
'#### DRAW ACCEPT,REDO,UNDO & CANCEL #################################


'#### picDisplay ACTION, CLIP,LASSO,MOVE PICS & DRAWING #####################

Private Sub picDisplay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'LabXY = aFileOps & aDRAW
   
   If aFileOps = True Then Exit Sub

   If aClipON Then CLIPPER_Start_End X, Y: Exit Sub
   
   If aLasso Then LASSO_START X, Y: Exit Sub
   
   If aResize Then Exit Sub
   
   If aMerged And Button = vbLeftButton And Not aDRAW Then
      
      MOVE_MERGE_START X, Y
      
      Exit Sub
   
   End If
   
If aShowON Then
   ' picFull() displayed then
End If

End Sub

Private Sub picDisplay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If aFileOps = True Then Exit Sub
   LabXY = Str$(X) & Str$(Y)

   If aClipON Then CLIPPER_Move Button, X, Y: Exit Sub
   
   If aDRAW Then
   
      If aHairs Then
         Line2.x1 = X
         Line2.x2 = X
         Line2.y1 = 0
         Line2.y2 = picDisplay.Height
         Line3.x1 = 0
         Line3.x2 = picDisplay.Width
         Line3.y1 = Y
         Line3.y2 = Y
      End If

      FIX_DRAW_CURSORS
      Show_Instructions
      If aDrawStart Then
         If Not aDrawMove Then
            DRAW_DRAW picDisplay, X, Y
         Else
            Select Case TheDrawStyle
            Case 0, 1, 2 ' Blend weak, Blend med, Blend strong
               DRAW_MOVE picDisplay, Shape1(TheDrawStyle), X, Y
            Case 3, 4  ' Erase rect, Erase circ
               DRAW_MOVE picDisplay, Shape2(TheDrawStyle - 3), X, Y
            Case Else
               DRAW_MOVE picDisplay, Shape1(TheDrawStyle), X, Y
            End Select
         End If   ' If Not aDrawMove Then
      End If   ' If aDrawStart Then
      Exit Sub
   End If   ' If aDRAW Then
   
   If aLasso Then LASSO_MOVE Button, X, Y: Exit Sub
   If aResize Then Exit Sub
   
   If aMerged And Button = vbLeftButton Then
      MOVE_MERGE_MOVE X, Y
      Exit Sub
   End If
   
If aShowON Then
   ' picFull() displayed then
End If

End Sub

Private Sub picDisplay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   If aFileOps = True Then Exit Sub
   If aClipON Then Exit Sub
   
   If aDRAW Then
   
      If Button = vbLeftButton Then
         
         If LClickCount = 0 Then
            picDisplay_To_picTemp   ' For non-trace Redo
         End If
         
         If aShowInstructions Then Show_Instructions
         
         LClickCount = LClickCount + 1
         If TheDrawStyle = 29 Then
            Fill picDisplay, X, Y
            aDrawStart = False
            aDrawMove = False
            LClickCount = 0
            Exit Sub
         End If
         
         If LClickCount = 1 Then
            picDisplay.DrawMode = 7
            picDisplay.DrawWidth = TheDrawWidth
'--------------------------------------------
               If aTrace Then
                  picFullTemp_To_picFullTempBack
                  L = picFull(PicNum).Left  ' Offset for ADD
                  T = picFull(PicNum).Top
                  ADDED = False
               End If
'--------------------------------------------
            Select Case TheDrawStyle
            Case 0, 1, 2 ' Blend weak, Blend med, Blend strong
                  
               Shape1(0).Visible = False
               Shape1(1).Visible = False
               Shape1(2).Visible = False
               
               DRAW_START picDisplay, Shape1(TheDrawStyle), X, Y
            Case 3, 4  ' Erase rect, Erase circ
               
               Shape2(0).Visible = False
               Shape2(1).Visible = False
               
               DRAW_START picDisplay, Shape2(TheDrawStyle - 3), X, Y
            Case Else
               DRAW_START picDisplay, Shape1(TheDrawStyle), X, Y
            End Select
            aDrawStart = True
         End If
         
         If TheDrawStyle = 17 Or TheDrawStyle = 18 Then  ' PolyLines & Splines
            If LClickCount > 1 Then
               StartNextPolyLine picDisplay, X, Y
            End If
         
         ElseIf TheDrawStyle = 23 Or TheDrawStyle = 24 Then ' Parallelogram & Frustrum
            
            If LClickCount = 2 Then
               StartNextPolyLine picDisplay, X, Y
            ElseIf LClickCount = 3 Then
               UpdatePolyPoints X, Y
               If TheDrawStyle = 23 Then
                  CompleteParallelogram picDisplay, X, Y
               Else  ' 24 Frustrum
                  CompleteFrustrum picDisplay, X, Y
               End If
               aDrawMove = True
            End If
            
         ElseIf LClickCount = 2 Then
            aDrawMove = True
         End If
       
      ElseIf Button = vbRightButton Then
            
         RClickCount = RClickCount + 1
         If TheDrawStyle = 17 Or TheDrawStyle = 18 Then  ' PolyLines & Splines
            
            If RClickCount = 1 Then
               UpdatePolyPoints X, Y
               aDrawMove = True
            Else
               DRAW_FINAL picDisplay, picFullTemp, X, Y
               aDrawStart = False
               aDrawMove = False
               LClickCount = 0
               RClickCount = 0
               picDisplay.DrawWidth = 1
               picDisplay.DrawMode = 13
            End If
         
         ElseIf TheDrawStyle = 23 Or TheDrawStyle = 24 Then ' Parallelogram & Frustrum
            
            If LClickCount = 3 Then
               DRAW_FINAL picDisplay, picFullTemp, X, Y
               aDrawStart = False
               aDrawMove = False
               LClickCount = 0
               RClickCount = 0
               picDisplay.DrawWidth = 1
               picDisplay.DrawMode = 13
            End If
         
         Else
         
            Select Case TheDrawStyle
            Case 0, 1, 2 ' Blend weak, Blend med, Blend strong
               Shape1(TheDrawStyle).Visible = False
               aDrawStart = False
               aDrawMove = False
               LClickCount = 0
               RClickCount = 0
               picDisplay.DrawWidth = 1
               picDisplay.DrawMode = 13
            Case 3, 4  ' Erase rect, Erase circ
               Shape2(TheDrawStyle - 3).Visible = False
               aDrawStart = False
               aDrawMove = False
               LClickCount = 0
               RClickCount = 0
               picDisplay.DrawWidth = 1
               picDisplay.DrawMode = 13
            Case Else
               DRAW_FINAL picDisplay, picFullTemp, X, Y
               aDrawStart = False
               aDrawMove = False
               LClickCount = 0
               RClickCount = 0
               picDisplay.DrawWidth = 1
               picDisplay.DrawMode = 13
            End Select
         End If
      End If   ' If Button = vbLeftButton Then
   End If   ' If aDRAW Then
   
   If aLasso Then LASSO_UP: Exit Sub
   If aResize Then Exit Sub
   
   If aMerged And Button = vbLeftButton Then
      If jpic > 0 Then
         Erase memBack()
         Erase memPic()
      End If
      MousePointer = vbDefault
   End If
   
End Sub
'#### END picDisplay ACTION, CLIP,LASSO,MOVE PICS & DRAWING #####################

Private Sub EraseDrawArrays()

      Select Case DrawStyle
      Case 0, 1, 2   ' Blenders
         Erase memBytes()
      Case 5, 14, 15, 16, 17, 18, 19, 20 ' FREEDRAW,ARCH,RIBBONS,POLYLINES,SPLINE,SPRAY,STAR
         Erase ixPts()
         Erase iyPts()
      End Select
      
      ' Reduce back up pics
      With picFullBackUp
         .Width = 48
         .Height = 48
         .Picture = LoadPicture
         .Refresh
      End With
      With picFullTempBack
         .Width = 48
         .Height = 48
         .Picture = LoadPicture
         .Refresh
      End With

End Sub


'### GENERAL FRAME MOVER #####################################
Private Sub fraMOVER(fra As Frame, Button As Integer, X As Single, Y As Single)

   If Button = vbLeftButton Then
      
      fraLeft = fra.Left + (X - fraX) \ STX
      If fraLeft < 0 Then fraLeft = 0
      If fraLeft + fra.Width > Me.Width \ STX Then
         fraLeft = Me.Width \ STX - fra.Width - 8
      End If
      fra.Left = fraLeft
      
      fraTop = fra.Top + (Y - fraY) \ STY
      If fraTop < 8 Then fraTop = 8
      If fraTop + fra.Height > Me.Height \ STY Then
         fraTop = Me.Height \ STY - fra.Height - 8
      End If
      fra.Top = fraTop
      
   End If

End Sub
'### END GENERAL FRAME MOVER #####################################


'####  MERGING #######################################

Private Sub MERGE()
' MERGE
'Public memBack() As Long
'Public memPic() As Long
'Public TransColor As Long
'Public PicAlpha() As Long
'Public iza As Long      ' 0   -> 256


'CLEAR_Incomplete_Actions

If NumOfStoredPics = 1 Then
   MsgBox " MUST HAVE MORE THAN ONE PIC STORED", vbInformation, "Layers - Merging"
   Exit Sub
End If

'Screen.MousePointer = vbHourglass
'DoEvents
   
   ' First picture is background
   maxw = picFull(0).Width
   maxh = picFull(0).Height
   
   
   If picDisplay.Width <> maxw Or picDisplay.Height <> maxh Then
      With picDisplay
         .Picture = LoadPicture
         .Width = maxw
         .Height = maxh
      End With
      picDisplay.Refresh
   
   End If
   
   
   ReDim memBack(1 To maxw, 1 To maxh)
   memBack(1, 1) = TransColor
   memBack(maxw, maxh) = TransColor
   ' NB Needs to be (1 To maxw) ??. (maxw) gives a streaky output ??!!
   
   'Background Pic To memBack
   GETDIBS picFull(0).Image, 0
   
   ptrmemBack = VarPtr(memBack(1, 1))

   FillFixedASMalpha
   
   For i = 1 To NumOfStoredPics - 1
      ' Following pictures
      ' Coords
      W = picFull(i).Width
      H = picFull(i).Height
      T = picFull(i).Top
      L = picFull(i).Left
      
      ReDim memPic(1 To W, 1 To H)
      'Get pic To memPic
      GETDIBS picFull(i).Image, i
      
      iza = PicAlpha(i)    ' 0   -> 256
      
      ptrmemPic = VarPtr(memPic(1, 1))
      FillVaryingASMalpha
   
      response = CallWindowProc(ptrMC, ptrStruc, 2&, 3&, 4&)

   Next i
   
   For i = 1 To NumOfStoredPics - 1
      picFull(i).Visible = False
   Next i
      
   SetStretchBltMode picDisplay.hdc, HALFTONE
   
   ' Blit memBack to picDisplay
   StretchDIBits picDisplay.hdc, _
      0&, 0&, maxw, maxh, _
      0&, 0&, maxw, maxh, _
      memBack(1, 1), bmbac, 0, vbSrcCopy
   
   picDisplay.Refresh
   
   Erase memBack()
   Erase memPic()
   
   aMerged = True
   
   aShowON = False
   aShowALLON = False

   FixScrollbars picFrame, picDisplay, HS, VS
 '  Screen.MousePointer = vbDefault
   
   DoEvents

End Sub
'####  END MERGING #######################################


'#### EFFECTS on picDisplay ie PicNum or Merged pic ##############


Private Sub cmdSingleEffect_Click()
 
 picLevelBar_MouseDown 1, 0, 0, 0

End Sub

Private Sub mnuEffects_Click()
' Main Menu Header

End Sub

Private Sub mnuPICorMERGED_Click(Index As Integer)

   If Index = 0 Then ' Effects on individual pics Merged or not
      mnuPICorMERGED(0).Checked = True
      mnuPICorMERGED(1).Checked = False
      aIndividual = True
   Else     ' Effects on whole MERGED picture
      mnuPICorMERGED(0).Checked = False
      mnuPICorMERGED(1).Checked = True
      aIndividual = False
   End If

End Sub

Private Sub mnuSingleEffects_Click()
' -> mnuSEffects
End Sub

Private Sub mnuSEffects_Click(Index As Integer)

' 0->0   ' Invert
' 1->20  ' Elliptic
' 2->21  ' Flip
' 3->22  ' Mirror right half
' 4->23  ' Mirror left half
   
   Select Case Index
   Case 0: mnuEffectsN_Click 0
   Case 1: mnuEffectsN_Click 20
   Case 2: mnuEffectsN_Click 21
   Case 3: mnuEffectsN_Click 22
   Case 4: mnuEffectsN_Click 23
'   Case 5: mnuEffectsN_Click 24
   
   End Select
   
End Sub

Private Sub mnuEffectsN_Click(Index As Integer)

CLEAR_Incomplete_Actions


   If NumOfStoredPics = 0 Then
      MsgBox "Load picture(s)first", vbInformation, "Layers - Effects"
      Exit Sub
   End If
   
   If aIndividual Then
      
      If PicNum >= NumOfStoredPics Then
         MsgBox "Select a picture first", vbInformation, "Layers - Effects"
         Exit Sub
      End If
      
      If aShowALLON Then
         optSelect_Click (PicNum)
      End If
      
      If Not aEffects Then
         If aMerged Then
            picFull_To_picTemp
            With picFull(PicNum)
               .BackColor = TransColor
               .ForeColor = TransColor
            End With
         Else
            picDisplay_To_picTemp
         End If
      End If
      

   Else  ' Whole Merged picture

      If Not aEffects Then picDisplay_To_picTemp
      
   End If
   
   aEffects = True


EffectsType = Index

mnuEffects.Enabled = False
cmdSingleEffect.Visible = False

   Select Case EffectsType
   ' Single
   Case 0: LabLevelBar = "Inverter"
      cmdSingleEffect.Visible = True
   ' Variable
   Case 1: LabLevelBar = "Sharp-Soft"
   Case 2: LabLevelBar = "Dark-Bright"
   Case 3: LabLevelBar = "bR +/-"
   Case 4: LabLevelBar = "bG +/-"
   Case 5: LabLevelBar = "bB +/-"
   Case 6: LabLevelBar = "Diffuse"
   Case 7: LabLevelBar = "Relief"
   Case 8: LabLevelBar = "Metallic"
   Case 9: LabLevelBar = "Flute Up \ /"
   Case 10: LabLevelBar = "Flute Down / \"
   Case 11: LabLevelBar = "Ripple"
   Case 12: LabLevelBar = "Rounded rect"
   Case 13: LabLevelBar = "Tile"
   Case 14: LabLevelBar = "Horz Shading"
   Case 15: LabLevelBar = "Vert Shading"
'   Case 16
'   Case 17
'   Case 18
'   Case 19
   
   ' Singles
   Case 20: LabLevelBar = "Elliptic"
      cmdSingleEffect.Visible = True   ' CLICK ME ' BAR
   Case 21: LabLevelBar = "Flip horizontal"
      cmdSingleEffect.Visible = True
   Case 22: LabLevelBar = "Mirror right half"
      cmdSingleEffect.Visible = True
   Case 23: LabLevelBar = "Mirror left half"
      cmdSingleEffect.Visible = True
   
   
   
   End Select
   
   Select Case EffectsType
   Case 0, 20 To 23: fraLevelBar.Caption = " Click mouse"
   Case Else: fraLevelBar.Caption = " Mouse Button down && slide"
   End Select
   
      fraLevelBar.Top = 8
      fraLevelBar.Visible = True

End Sub


Private Sub picLevelBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   picLevelBar_MouseMove Button, Shift, X, Y
      
End Sub

Private Sub picLevelBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If X < 0 Then X = 0
If X > 128 Then X = 128

   Select Case EffectsType
   
   Case 0  ' Inverter
      FilterParam = 255
      LabFilterParam = Str$(FilterParam)
   
   Case 1   ' Sharp-Soft
      FilterParam = (8 * X) \ 128 + 1 '1 -> 4 Sharp 5 -> 8 Soft
      If FilterParam > 8 Then FilterParam = 8
      LabFilterParam = Str$(FilterParam)
   
   Case 2, 3, 4, 5  ' Dark-Bright & RGB +/-
      FilterParam = (X - 64) * 4 ' -255 -> +255
      If FilterParam < -255 Then FilterParam = -255
      If FilterParam > 255 Then FilterParam = 255
      LabFilterParam = Str$(FilterParam)
   
   Case 6  ' Diffuse
      FilterParam = (16 * X) \ 128 + 1 '1 -> 16
      If FilterParam > 16 Then FilterParam = 16
      LabFilterParam = Str$(FilterParam)
   
   Case 7, 8  ' Relief, Metallic
      FilterParam = 40 + X \ 2   ' 40 -> 104
      LabFilterParam = Str$(FilterParam)
   
   Case 9, 10 ' Flute Up & Down
      FilterParam = X + 1
      LabFilterParam = Str$(FilterParam)
   
   Case 11  ' Ripple
      
      FilterParam = (40 * X) \ 128 '- 1
      LabFilterParam = Str$(FilterParam)
   
   Case 12  ' Rounded rectangle
      FilterParam = X + 1
      LabFilterParam = Str$(FilterParam)
   
   Case 13  ' Tile
      FilterParam = X \ 8 + 2
      LabFilterParam = Str$(FilterParam)
   
   Case 14, 15 ' H & V Shading

      FilterParam = X - 64
      LabFilterParam = Str$(FilterParam)

'   Case 15
'   Case 16
'   Case 17
'   Case 18
'   Case 19
   
   
   ' Singles
   Case 20 To 23  ' Elliptic, Flip, MirrorRH, MirrorLH
      FilterParam = 0
      LabFilterParam = Str$(FilterParam)
   
   
   End Select

   LabFilterParam.Refresh
   
   If Button = vbLeftButton Then
      If aMerged And aIndividual Then
         DISPLAY_EFFECTS picFull(PicNum), picTemp ' Individual in Merged Mode
         MERGE
      Else
         DISPLAY_EFFECTS picDisplay, picTemp ' Non-merged Individual or Whole Merged
      End If
   
   End If

End Sub

Private Sub cmdLevelBar_Click(Index As Integer)

   cmdSingleEffect.Visible = False
   Select Case Index
   Case 0   ' Accept Effects
      If aMerged Then
         ' Leave picDisplay or
         If aIndividual Then
            ' Show new Thumb
            SetStretchBltMode picThumb(PicNum).hdc, HALFTONE
            
            StretchBlt picThumb(PicNum).hdc, 0, 0, _
               SmallPicW, SmallPicH, picFull(PicNum).hdc, _
               0, 0, W, H, vbSrcCopy
            
            picThumb(PicNum).Refresh
         End If
         
      Else  ' Non-Merged
         'picDisplay 2  picFull
         Display2picFull_picThumb
      End If
   
   Case 1   ' Cancel effects
      
      If aMerged And aIndividual Then
         picTemp_To_picFull
         MERGE
      Else
         picTemp_To_picDisplay
      End If
      
   End Select
   
   fraLevelBar.Visible = False
   mnuEffects.Enabled = True
   aEffects = False
   
End Sub

Private Sub fraLevelBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   fraX = X
   fraY = Y

End Sub

Private Sub fraLevelBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   fraMOVER fraLevelBar, Button, X, Y

End Sub
'#### END EFFECTS on picDisplay ie PicNum or Merged pic ##############


'#### SWAP PICS ##########################################

Private Sub cmdSwapPics_Click(Index As Integer)

' Index =0 swap 0 & 1
'
' Index =2 swap 2 & 3
' etc
' Index Max = MaxNumOfPics-1
'

   CLEAR_Incomplete_Actions
   
   
   ' Swap picFull(Index) & picFull(Index+1)
   ' using picTemp
   
   ' picFull(Index) to picTemp
   W = PicWidth(Index)
   H = PicHeight(Index)
   With picTemp
      .Picture = LoadPicture
      .Width = W
      .Height = H
   End With
   
   BitBlt picTemp.hdc, 0, 0, _
      W, H, picFull(Index).hdc, 0, 0, vbSrcCopy
   
   picTemp.Refresh
   
   ' picFull(Index+1) to picFull(Index)
   W = PicWidth(Index + 1)
   H = PicHeight(Index + 1)
   With picFull(Index)
      .Picture = LoadPicture
      .Width = W
      .Height = H
   End With
  
   BitBlt picFull(Index).hdc, 0, 0, _
      W, H, picFull(Index + 1).hdc, 0, 0, vbSrcCopy
   
   ' picTemp to picFull(Index+1)
   W = PicWidth(Index)
   H = PicHeight(Index)
   With picFull(Index + 1)
      .Picture = LoadPicture
      .Width = W
      .Height = H
   End With
   
   BitBlt picFull(Index + 1).hdc, 0, 0, _
      W, H, picTemp.hdc, 0, 0, vbSrcCopy
   
   ' Swap FileSpecs
   a$ = StoreFileSpec$(Index)
   StoreFileSpec$(Index) = StoreFileSpec$(Index + 1)
   StoreFileSpec$(Index + 1) = a$
   ' Swap W
   i = PicWidth(Index)
   PicWidth(Index) = PicWidth(Index + 1)
   PicWidth(Index + 1) = i
   ' Swap H
   i = PicHeight(Index)
   PicHeight(Index) = PicHeight(Index + 1)
   PicHeight(Index + 1) = i
   
   If Index = 0 Then
      maxw = picFull(0).Width
      maxh = picFull(0).Height
      picDisplay.Width = maxw
      picDisplay.Height = maxh
      picDisplay.Picture = LoadPicture
   End If
   
   'Show New Thumbs
   For i = Index To Index + 1
      
      SetStretchBltMode picThumb(i).hdc, HALFTONE
      
      StretchBlt picThumb(i).hdc, 0, 0, _
         SmallPicW, SmallPicH, picFull(i).hdc, _
         0, 0, PicWidth(i), PicHeight(i), vbSrcCopy
      
      picThumb(i).Refresh
   
   Next i

   FixScrollbars picFrame, picDisplay, HS, VS
   
   ShowInfo

   ' Show picDisplay if NOT Merged
   ' Else Re-Merge after swapping
   If Not aMerged Then
      optSelect_Click Index
      optSelect(Index).Value = True
   Else
      MERGE
   End If
   
   'aShowON = False

End Sub
'#### END SWAP PICS ##########################################


'#### TRANSPARENCY ALPHA  ################################

Private Sub picChangeAlpha_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
i = CLng(X)

   If PicNum = NumOfStoredPics Then Exit Sub
   
   If PicNum > 0 Then
      If i < 0 Then i = 0
      If i > 256 Then i = 256
      LabAlphaValue = Str$(100 * (256 - i) \ 256) & " %"
      LabAlphaValue.Refresh
      
      If Button = vbLeftButton Then
      
         PicAlpha(PicNum) = i
         
         ' Re-Merge if in Merged mode
         If aMerged Then MERGE
      
      End If
   End If

End Sub

Private Sub picChangeAlpha_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
i = CLng(X)
   
   If PicNum = NumOfStoredPics Then Exit Sub
   
   If PicNum > 0 Then
      If i < 0 Then i = 0
      If i > 256 Then i = 256
      LabAlphaValue = Str$(100 * (256 - i) \ 256) & " %"
      LabAlphaValue.Refresh
      
      If Button = vbLeftButton Then
      
         PicAlpha(PicNum) = i
         
         ' Re-Merge if in Merged mode
         If aMerged Then MERGE
      
      End If
   End If

End Sub
'#### END TRANSPARENCY ALPHA  ################################


'#### Toggle & Hide Thumb-bar #####################################

Private Sub chkToggleThumbBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

   'CLEAR_Incomplete_Actions
   chkToggleThumbBar.Value = 0
   picDisplay.SetFocus
   fraThumbBar.Visible = Not fraThumbBar.Visible

End Sub

Private Sub cmdHideThumbBar_Click(Index As Integer)
' On Thumb bar
   fraThumbBar.Visible = False
End Sub

'#### Move Thumb-bar #############################################

Private Sub LabMoveThumbBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   fraX = X
   fraY = Y

End Sub

Private Sub LabMoveThumbBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   fraMOVER fraThumbBar, Button, X, Y
   
End Sub


'#### TOGGLE + HAIRS ######################################

Private Sub chkPlusHairs_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   chkPlusHairs.Value = 0
   
   aHairs = Not aHairs
   
   If aHairs Then
      If aDRAW Then
         Line2.Visible = True
         Line3.Visible = True
      End If
   Else
      Line2.Visible = False
      Line3.Visible = False
   End If
   
   picDisplay.SetFocus

End Sub


'#### DELETE PICS #############################################

Private Sub mnuClearALLpics_Click()
   
   CLEAR_Incomplete_Actions
   
   ' Clear pics 1 to last or just PicNum=1
      
   For i = 1 To NumOfStoredPics - 1
      
      With picFull(i)
         .Width = 48
         .Height = 48
      End With
      picFull(i).Picture = LoadPicture
      picFull(i).Refresh
      
      With picThumb(i)
         .Width = 48
         .Height = 48
      End With
      picThumb(i).Picture = LoadPicture
      picThumb(i).Refresh
   
   Next i
   
   ' Disable all but first optSelect
   ' so pics must be loaded in order
   For i = 2 To MaxNumOfPics - 1
      optSelect(i).Enabled = False
   Next i
   
   ' Disable Swap & Merge buttons
   For i = 0 To MaxNumOfPics - 2
      cmdSwapPics(i).Enabled = False
      cmdMerge(i).Enabled = False
   Next i
   
   ReDim Preserve StoreFileSpec$(0)
   
   PicNum = 0
   
   aMerged = False
   aShowON = True
   aShowALLON = False
   
   ' Merge ShowAll indicator
   LabGreen(0).BackColor = RGB(0, 128, 0)
   LabGreen(1).BackColor = RGB(0, 128, 0)
   
   optSelect_Click 0
   
   optSelect(0).Value = True
   
   NumOfStoredPics = 1
   LabNSP = " NumOfStoredPics =" & Str$(NumOfStoredPics)

   chkSHOWALL.Enabled = False
   chkMERGE.Enabled = False
   mnuRepositionLostPic.Enabled = False
   mnuClearALLpics.Enabled = False
   mnuClearNpics.Enabled = False

End Sub


Private Sub mnuClearNpics_Click()

   CLEAR_Incomplete_Actions
   
   If PicNum >= NumOfStoredPics Then
      MsgBox "Select a picture first", vbInformation, "Clear pic"
      Exit Sub
   End If
   
   If PicNum = 0 Then
      MsgBox "Can't clear background, only replace or swap", vbInformation, "Clear pic"
      Exit Sub
   End If
   
   If PicNum < NumOfStoredPics Then
      ' move down all picFulls() but last one
      ' overwriting picFull(picNum)
      
      For j = PicNum To NumOfStoredPics - 2
         
         W = picFull(j + 1).Width
         H = picFull(j + 1).Height
         
         With picFull(j)
            .Picture = LoadPicture
            .Width = W
            .Height = H
            .Refresh
         End With

         BitBlt picFull(j).hdc, 0, 0, W, H, _
            picFull(j + 1).hdc, 0, 0, vbSrcCopy
         picFull(j).Refresh
         
         StoreFileSpec$(j) = StoreFileSpec$(j + 1)
         PicWidth(j) = W
         PicHeight(j) = H
         PicAlpha(j) = PicAlpha(j + 1)
         
         With picThumb(j)
            .Width = SmallPicW
            .Height = SmallPicH
            .Picture = LoadPicture
            .Refresh
         End With
         
         ' Show each replace picFull(j)
         ' on Thumb Bar
         ' picFull(j) to picThumb(j)
         SetStretchBltMode picThumb(j).hdc, HALFTONE
         
         StretchBlt picThumb(j).hdc, 0, 0, _
            SmallPicW, SmallPicH, picFull(j).hdc, _
            0, 0, W, H, vbSrcCopy
         picThumb(j).Refresh
         
      Next j
      
      GoTo ClearLastPic
   
   End If
   
   If PicNum >= NumOfStoredPics Or NumOfStoredPics = 2 Then  ' Last pic
ClearLastPic:
      i = NumOfStoredPics - 1 ' Last PicNum with a picture
         
      With picFull(i)
         .Width = SmallPicW
         .Height = SmallPicH
         .Picture = LoadPicture
         .Refresh
         .Visible = False
      End With
      
      With picThumb(i)
         .Width = SmallPicW
         .Height = SmallPicH
         .Picture = LoadPicture
         .Refresh
      End With
      
      ' Disable last optSelect
      'i = MaxNumOfPics - 1
      If i < MaxNumOfPics - 1 Then
         optSelect(i + 1).Enabled = False
      End If
      
      ' Disable last Swap & Merge button
      'i = MaxNumOfPics - 2
      If i = NumOfStoredPics - 1 Then
         cmdSwapPics(i - 1).Enabled = False
         cmdMerge(i - 1).Enabled = False
      End If
      
      PicNum = i - 1
      
      NumOfStoredPics = NumOfStoredPics - 1
      ReDim Preserve StoreFileSpec$(PicNum)
      ReDim Preserve PicWidth(PicNum)
      ReDim Preserve PicHeight(PicNum)
      LabNSP = " NumOfStoredPics =" & Str$(NumOfStoredPics)
      
      If aMerged Then
         If NumOfStoredPics > 1 Then
            MERGE
         Else
            aMerged = False
            ' Merge ShowAll indicator
            LabGreen(0).BackColor = RGB(0, 128, 0)
            LabGreen(1).BackColor = RGB(0, 128, 0)
            aShowON = True
            chkSHOWALL.Enabled = False
            chkMERGE.Enabled = False
         End If
      End If
         
      If Not aMerged Then
         If Not aShowALLON Then
            optSelect_Click (PicNum)
         Else
            SHOWALL
         End If
      End If
      
      optSelect(PicNum).Value = True
      
      
   End If

End Sub
'#### END DELETE PIC PicNum #############################################


'#### Select a pic ####################################

Private Sub optSelect_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   optSelect_Click Index
End Sub

Private Sub optSelect_Click(Index As Integer)

' Indicates PicNum
' Enable/Disable Thumb bar commands
' Enable/Disable equivalent menu commands
' Move picFull(PicNum) to picDisplay
' FixScrollBars
' Highlight with shape rect
   
   CLEAR_Incomplete_Actions
   
   mnuClipBoard(0).Enabled = True
   mnuDrawToggle.Enabled = True
   
   mnuPICorMERGED(0).Checked = True
   mnuPICorMERGED(0).Caption = "Effects on INDIVIDUAL pic" & Str$(PicNum)
   aIndividual = True
   
   If Not aMerged Then
      mnuPICorMERGED(1).Checked = False
      mnuPICorMERGED(1).Enabled = False
   End If
   
   mnuSaveStoredPic.Enabled = True
   mnuSaveDisplay.Enabled = True

   
   
   PicNum = Index
   
   If PicNum < NumOfStoredPics Then
      aShowALLON = False
      LabGreen(1).BackColor = RGB(0, 128, 0)
      mnuTraceN(2).Caption = "&Add Trace to pic" & Str$(PicNum)
      If PicNum > 0 Then
         i = PicAlpha(PicNum)
         LabAlphaValue = Str$(100 * (256 - i) \ 256) & " %"
         LabAlphaValue.Refresh
      End If
      
   End If
   
   ' CLIP & RESIZE MERGE
   If Not aMerged Then
      cmdClipMerged.Enabled = False
      cmdClipMerged.BackColor = RGB(224, 224, 224)
      mnuClipMerged.Enabled = False
      
      cmdResizeMerged.Enabled = False
      cmdResizeMerged.BackColor = RGB(224, 224, 224)
      mnuResizeMerged.Enabled = False
   Else  ' Merged
      cmdClipMerged.Enabled = True
      cmdClipMerged.BackColor = RGB(255, 255, 255)
      mnuClipMerged.Enabled = True
      
      cmdResizeMerged.Enabled = True
      cmdResizeMerged.BackColor = RGB(255, 255, 255)
      mnuResizeMerged.Enabled = True
   End If
   
   ' Captions
   If Not aMerged Then
      ' TB Thumb Bar
      cmdTBClip.Caption = " Clip" & Str$(PicNum)
      mnuClip.Caption = " &Clip pic" & Str$(PicNum)
      cmdTBLasso.Caption = " Lasso" & Str$(PicNum)
      mnuLasso.Caption = " &Lasso pic" & Str$(PicNum)
      cmdTBResize.Caption = "Resize" & Str$(PicNum)
      mnuResizeApic.Caption = " &Resize pic" & Str$(PicNum)
      cmdTBRotate.Caption = "Rotate" & Str$(PicNum)
      mnuRotatePic.Caption = " R&otate pic" & Str$(PicNum)
   Else ' Merged
      cmdTBClip.Caption = " Clip"
      mnuClip.Caption = " &Clip"
      cmdTBLasso.Caption = " Lasso"
      mnuLasso.Caption = " Lasso"
      cmdTBResize.Caption = "Resize" & Str$(PicNum)
      mnuResizeApic.Caption = " &Resize pic" & Str$(PicNum)
      cmdTBRotate.Caption = "Rotate" & Str$(PicNum)
      mnuRotatePic.Caption = " R&otate pic" & Str$(PicNum)
   End If
   
   ' Enabled & BackColor
   If Not aMerged Then
      Select Case PicNum
      Case 0
        cmdTBClip.Enabled = True
        cmdTBClip.BackColor = RGB(255, 255, 255)
        mnuClip.Enabled = True
        cmdTBLasso.Enabled = False
        cmdTBLasso.BackColor = RGB(224, 224, 224)
        mnuLasso.Enabled = False
       cmdTBLasso.Caption = " Lasso"
       mnuLasso.Caption = " Lasso"
        cmdTBResize.Enabled = True
        cmdTBResize.BackColor = RGB(255, 255, 255)
        mnuResizeApic.Enabled = True
        cmdTBRotate.Enabled = False
        cmdTBRotate.BackColor = RGB(224, 224, 224)
        mnuRotatePic.Enabled = False
       cmdTBRotate.Caption = "Rotate"
       mnuRotatePic.Caption = " Rotate"
      Case Is < NumOfStoredPics  ' All True
        cmdTBClip.Enabled = True
        cmdTBClip.BackColor = RGB(255, 255, 255)
        mnuClip.Enabled = True
        cmdTBLasso.Enabled = True
        cmdTBLasso.BackColor = RGB(255, 255, 255)
        mnuLasso.Enabled = True
        cmdTBResize.Enabled = True
        cmdTBResize.BackColor = RGB(255, 255, 255)
        mnuResizeApic.Enabled = True
        cmdTBRotate.Enabled = True
        cmdTBRotate.BackColor = RGB(255, 255, 255)
        mnuRotatePic.Enabled = True
      Case Else   ' PicNum = NumOfStoredPics  All False
        cmdTBClip.Enabled = False
        cmdTBClip.BackColor = RGB(224, 224, 224)
        mnuClip.Enabled = False
        cmdTBLasso.Enabled = False
        cmdTBLasso.BackColor = RGB(224, 224, 224)
        mnuLasso.Enabled = False
        cmdTBResize.Enabled = False
        cmdTBResize.BackColor = RGB(224, 224, 224)
        mnuResizeApic.Enabled = False
        cmdTBRotate.Enabled = False
        cmdTBRotate.BackColor = RGB(224, 224, 224)
        mnuRotatePic.Enabled = False
      End Select
      
   Else   ' Merged
      Select Case PicNum  ' Only Resize pic True
      Case 0
        cmdTBClip.Enabled = False
        cmdTBClip.BackColor = RGB(224, 224, 224)
        mnuClip.Enabled = False
        cmdTBLasso.Enabled = False
        cmdTBLasso.BackColor = RGB(224, 224, 224)
        mnuLasso.Enabled = False
        cmdTBResize.Enabled = True
        cmdTBResize.BackColor = RGB(255, 255, 255)
        mnuResizeApic.Enabled = True
        cmdTBRotate.Enabled = False
        cmdTBRotate.BackColor = RGB(224, 224, 224)
        mnuRotatePic.Enabled = False
   
      Case Is < NumOfStoredPics ' Only Resize & Rotate True
        cmdTBClip.Enabled = False
        cmdTBClip.BackColor = RGB(224, 224, 224)
        mnuClip.Enabled = False
        cmdTBLasso.Enabled = False
        cmdTBLasso.BackColor = RGB(224, 224, 224)
        mnuLasso.Enabled = False
        cmdTBResize.Enabled = True
        cmdTBResize.BackColor = RGB(255, 255, 255)
        mnuResizeApic.Enabled = True
        cmdTBRotate.Enabled = True
        cmdTBRotate.BackColor = RGB(255, 255, 255)
        mnuRotatePic.Enabled = True
      Case Else   ' PicNum = NumOfStoredPics  All False
        cmdTBClip.Enabled = False
        cmdTBClip.BackColor = RGB(224, 224, 224)
        mnuClip.Enabled = False
        cmdTBLasso.Enabled = False
        cmdTBLasso.BackColor = RGB(224, 224, 224)
        mnuLasso.Enabled = False
        cmdTBResize.Enabled = False
        cmdTBResize.BackColor = RGB(224, 224, 224)
        mnuResizeApic.Enabled = False
        cmdTBRotate.Enabled = False
        cmdTBRotate.BackColor = RGB(224, 224, 224)
        mnuRotatePic.Enabled = False
      End Select
   End If   ' Not aMerged & Merged
   
   '  Reset BackColor for these two on Start up
   If NumOfStoredPics = 0 Then
      cmdTBClip.BackColor = RGB(224, 224, 224)
      cmdTBResize.BackColor = RGB(224, 224, 224)
   End If
  
   mnuRepositionLostPic.Caption = " Re-&position (lost) pic" & Str$(PicNum)
   mnuClearNpics.Caption = " Clear p&ic" & Str$(PicNum)
   
   If PicNum = 0 Or PicNum >= NumOfStoredPics Then
      mnuRepositionLostPic.Enabled = False
      mnuClearNpics.Enabled = False
   Else
      mnuRepositionLostPic.Enabled = True
      mnuClearNpics.Enabled = True
   End If
   
   If NumOfStoredPics > 1 Then
      mnuClearALLpics.Enabled = True
   Else
      mnuClearALLpics.Enabled = False
   End If
   
   If PicNum > 0 And fraThumbBar.Visible Then
      picThumbContainer.SetFocus ' Remove button rect
   End If
   
   ' Loading & Saving
   mnuOpen.Caption = "&Load pic" & Str$(PicNum)
   mnuReLoad.Caption = "&Reload pic" & Str$(PicNum)
   mnuLoadTLayer.Caption = "Load &Transparent Layer" & Str$(PicNum)
   mnuSaveStoredPic.Caption = "&Save stored pic" & Str$(PicNum)
   mnuClipBoard(1).Caption = "&Paste clipboard to pic" & Str$(PicNum)
   
   If PicNum >= NumOfStoredPics Then
      mnuReLoad.Enabled = False
      mnuSaveStoredPic.Enabled = False
   Else
      mnuReLoad.Enabled = True
      mnuSaveStoredPic.Enabled = True
   End If
   
   
   
   If aShowON Then   ' ie not merged
   
      If PicNum < NumOfStoredPics Then   ' Display stored pic
      
         picDisplay.Picture = LoadPicture ' Nec to clear old pic mem
         
         W = PicWidth(Index)
         H = PicHeight(Index)
         
         With picDisplay
            .Picture = LoadPicture
            .Width = W
            .Height = H
            .Refresh
         End With
         
         For i = 0 To NumOfStoredPics - 1
            picFull(i).Visible = False
         Next i
         
         ' Show just picFull(PicNum) in pic Display
         BitBlt picDisplay.hdc, 0, 0, _
            W, H, picFull(PicNum).hdc, 0, 0, vbSrcCopy
         
         picDisplay.Refresh
      
         FixScrollbars picFrame, picDisplay, HS, VS
         
      End If
   End If
   
   If PicNum < NumOfStoredPics Then ShowInfo
   
   LabSelPicNum = Str$(PicNum)
   
   ShapeSmall.Left = picThumb(Index).Left - 2

End Sub
'#### END Select a pic ####################################

Private Sub ShowInfo() 'PicNum

   If PicNum = NumOfStoredPics Then
      ww = PicWidth(PicNum - 1)
      hh = PicHeight(PicNum - 1)
      a$ = StoreFileSpec$(PicNum - 1)
   Else
      ww = PicWidth(PicNum)
      hh = PicHeight(PicNum)
      a$ = StoreFileSpec$(PicNum)
   End If
      
   LabInfo = ShortFileSpec(a$, 36) & "   W =" & Str$(ww) & "  H =" & Str$(hh)

End Sub


'### COPY & PASTE CLIPBOARD #########################################

Private Sub mnuClipBoard_Click(Index As Integer)

If Index = 0 Then ' Copy
   Clipboard.Clear
   Clipboard.SetData picDisplay.Image, vbCFBitmap
   aClipBoardUsed = True
   DoEvents
Else  ' Paste
   aClipBoard = True
   mnuOpen_Click
   aClipBoard = False
End If

End Sub
'### END COPY & PASTE CLIPBOARD #########################################


'#### LOAD & RE-LOAD a pic #########################################

Private Sub mnuFile_Click()
   
   If aDRAW Then
      mnuOpen.Enabled = False
      mnuReLoad.Enabled = False
      mnuLoadTLayer.Enabled = False
      mnuSaveStoredPic.Enabled = False
      mnuSaveDisplay.Enabled = False
   Else
      If NumOfStoredPics = 0 Then
         mnuReLoad.Enabled = False
         mnuRepositionLostPic.Enabled = False
         mnuSaveStoredPic.Enabled = False
         mnuSaveDisplay.Enabled = False
         mnuClipBoard(0).Enabled = False  ' Copy to Clipboard
      Else
         mnuOpen.Enabled = True
         mnuReLoad.Enabled = True
         mnuLoadTLayer.Enabled = True
         mnuSaveStoredPic.Enabled = True
         mnuSaveDisplay.Enabled = True
         mnuRepositionLostPic.Enabled = True
         mnuClipBoard(0).Enabled = True  ' Copy to Clipboard
      End If
   End If
   
End Sub

Private Sub mnuOpen_Click()
Dim Title$, Filt$, InDir$
Dim FileSpecString$
Dim k As Integer

   
CLEAR_Incomplete_Actions
   
   ' LOAD STANDARD VB PICTURES INTO picDisplay
   ' OR CLIPBOARD TO picDisplay

On Error GoTo LoadError:

   aFileOps = True
   DoEvents
   chkMERGE.Enabled = False
   chkSHOWALL.Enabled = False
   

   If Not aClipBoard Then
   
ReOpen:
   
      Title$ = "Load picture(s) to/from picture" & Str$(PicNum)
      Filt$ = "Pics bmp,jpg,gif,ico,cur,wmf,emf|*.bmp;*.jpg;*.gif;*.ico;*.cur;*.wmf;*.emf"
      InDir$ = FileSpecString$
      
      OS.ShowOpen FileSpecString$, Title$, Filt$, InDir$, "", Me.hwnd, True
      
      If Len(FileSpecString$) = 0 Then
         Close
         MsgBox "Either none or too many files selected", vbInformation, "Layers - Loading"
         aFileOps = False
         Exit Sub
      End If
   
   End If

   If aClipBoard Then FileSpecString$ = "ClipBoard" & Str$(PicNum)

   'Public FileSpec$(), NumPicsSelected
   Extract_Store_FileSpecs FileSpecString$  ', ByVal StartPicNum As Long)
   ' returns NumPicsSelected
   
   For k = NumPicsSelected To 1 Step -1
      
      If NumPicsSelected = 1 Then
          FileSpecString$ = FileSpec$(k - 1)
      Else
          FileSpecString$ = APath$ & FileSpec$(k - 1)
      End If
      
      picDisplay.Picture = LoadPicture
      picDisplay.Width = 16
      picDisplay.Height = 16
      picDisplay.Refresh
      
      If Not aClipBoard Then
            picDisplay.Picture = LoadPicture(FileSpecString$)
      Else
            picDisplay.Picture = Clipboard.GetData(vbCFBitmap)
      End If
      picDisplay.Refresh
      
      'DoEvents
      
      GetObjectAPI picDisplay.Image, Len(bmp), bmp
      
      W = bmp.bmWidth
      H = bmp.bmHeight
   
'----------------------------------------------------------------
      If PicNum = NumOfStoredPics Then  ' Bump info
         
         NumOfStoredPics = NumOfStoredPics + 1
         LabNSP = " NumOfStoredPics =" & Str$(NumOfStoredPics)
         ReDim Preserve StoreFileSpec$(PicNum)
         ReDim Preserve PicWidth(PicNum)
         ReDim Preserve PicHeight(PicNum)
   
      End If
   
      ' Update info from mnuOpen
      StoreFileSpec$(PicNum) = FileSpecString$
      PicWidth(PicNum) = W
      PicHeight(PicNum) = H
      
      If PicNum = 0 Then
         maxw = W
         maxh = H
      End If
      
      With picFull(PicNum)
         .Top = 0
         .Left = 0
      End With
   
      Display2picFull_picThumb   ' Using current PicNum

      ' Enabled next picture in order
      PicNum = PicNum + 1
      
      If PicNum < MaxNumOfPics - 1 Then
         optSelect(PicNum).Enabled = True
         optSelect(PicNum).ForeColor = vbRed
      End If
   
      ' Enable swapping when enough pics loaded
      If NumOfStoredPics > 1 And PicNum < MaxNumOfPics Then
         If PicNum > 1 Then
            cmdSwapPics(PicNum - 2).Enabled = True
            cmdMerge(PicNum - 2).Enabled = True
         End If
      End If
'----------------------------------------------------------------
nextpic:
   Next k
   
'----------------------------------------------------------------
COMPLETE_LOADING
   
If aFileOps = True Then
   aFileOps = False
   Sleep 200   ' To avoid unwanted MERGE on loading pic
End If
   

If NumOfStoredPics > 1 Then
   chkMERGE.Enabled = True
   chkSHOWALL.Enabled = True
End If

If aMerged Then MERGE
If aShowALLON Then SHOWALL

On Error GoTo 0
aFileOps = False
Exit Sub
'===================
LoadError:

MsgBox " Can't load " & FileSpecString$ & vbCr & " Possibly 32-bit icon", vbInformation, "Layers - Loading"
If k = 1 Then
   Resume ReOpen
Else
   Resume nextpic
End If

End Sub

'#### RELOAD #######################################################

Private Sub mnuReLoad_Click()
Dim FileSpecString$

CLEAR_Incomplete_Actions
   
   aFileOps = True
   DoEvents
   chkMERGE.Enabled = False
   chkSHOWALL.Enabled = False
   
   If PicNum > NumOfStoredPics - 1 Then
      MsgBox " Select a picture first", vbExclamation, " Layers Re-load"
      aFileOps = False
      Exit Sub
   End If
   
   FileSpecString$ = StoreFileSpec$(PicNum)
   
   If Left$(FileSpecString$, 6) = "TLayer" Then
      MsgBox " CANNOT RE-LOAD A TLAYER", vbInformation, "Layers Re-load"
      Exit Sub
   End If
   
   picDisplay.Picture = LoadPicture
   picDisplay.Picture = LoadPicture(FileSpecString$)
   DoEvents
   
   FixScrollbars picFrame, picDisplay, HS, VS
   
   GetObjectAPI picDisplay.Image, Len(bmp), bmp
   
   W = bmp.bmWidth
   H = bmp.bmHeight
   
   If PicNum = 0 Then
      maxw = W
      maxh = H
   End If
   
   LabInfo = ShortFileSpec(FileSpecString$, 36) & "   W =" & Str$(W) & "  H =" & Str$(H)
   
   CompleteStoring
   
   If NumOfStoredPics > 1 Then
      chkMERGE.Enabled = True
      chkSHOWALL.Enabled = True
   End If

   If aMerged Then MERGE
   If aShowALLON Then SHOWALL

   aFileOps = False

End Sub

'#### Load transparent layer #################################

Private Sub mnuLoadTLayer_Click()

Dim FileSpecString$

CLEAR_Incomplete_Actions
   
aFileOps = True
   
   W = TLayerWidth   ' 256  default
   H = TLayerHeight  ' 256
   FileSpecString$ = "TLayer" & Str$(PicNum)
   
   picFull(PicNum).Picture = LoadPicture
   picFull(PicNum).BackColor = TransColor
   ' ->
   picFull(PicNum).Width = W
   picFull(PicNum).Height = H
   
   picDisplay.Picture = LoadPicture
   picDisplay.Width = W
   picDisplay.Height = H
   picDisplay.Cls
   picDisplay.Print "TLayer" & Str$(PicNum)  ' Temporary
      
'----------------------------------------------------------------
   If PicNum = NumOfStoredPics Then  ' Bump info
      
      NumOfStoredPics = NumOfStoredPics + 1
      LabNSP = " NumOfStoredPics =" & Str$(NumOfStoredPics)
      ReDim Preserve StoreFileSpec$(PicNum)
      ReDim Preserve PicWidth(PicNum)
      ReDim Preserve PicHeight(PicNum)

   End If

   ' Update info from mnuOpen
   StoreFileSpec$(PicNum) = FileSpecString$
   PicWidth(PicNum) = W
   PicHeight(PicNum) = H

   If PicNum = 0 Then
      maxw = W
      maxh = H
   End If
   
   Display2picFull_picThumb   ' Using current PicNum
   
   ' Enabled next picture in order
   PicNum = PicNum + 1
   
   If PicNum < MaxNumOfPics - 1 Then
      optSelect(PicNum).Enabled = True
      optSelect(PicNum).ForeColor = vbRed
   End If

   ' Enable swapping when enough pics loaded
   If NumOfStoredPics > 1 And PicNum < MaxNumOfPics Then
      If PicNum > 1 Then
         cmdSwapPics(PicNum - 2).Enabled = True
         cmdMerge(PicNum - 2).Enabled = True
      End If
   End If
'----------------------------------------------------------------
COMPLETE_LOADING

aFileOps = False

If aMerged Then MERGE
If aShowALLON Then SHOWALL

End Sub
'#### END Load transparent layer #################################

'#### New Transparent Layer with TEXT ############################

Private Sub mnuText_Click()
Dim FileSpecString$

CLEAR_Incomplete_Actions
   
   If aShowALLON Then
      optSelect_Click (PicNum)   ' NO, turns off aDRAW
   End If
   
   frmText.Show vbModal

   If Len(TheText$) = 0 Then Exit Sub   ' Cancel @ frmText
   
   ' Public Info for frmText Return
'   W = Label1.Width
'   H = Label1.Height
'   PicNum = NumOfStoredPics - 1
'   TheText$ = Text1.Text
'   TextFont.fntName = CurFont.Name
'   TextFont.fntSize = CurFont.Size
'   TextFont.fntItalic = CurFont.Italic
'   TextFont.fntBold = CurFont.Bold

   picFull(PicNum).Picture = LoadPicture
   picFull(PicNum).BackColor = TransColor
   ' ->
   picFull(PicNum).Width = W
   picFull(PicNum).Height = H
   
   picFull(PicNum).FontName = TextFont.fntName
   picFull(PicNum).FontSize = TextFont.fntSize
   picFull(PicNum).FontItalic = TextFont.fntItalic
   picFull(PicNum).FontBold = TextFont.fntBold
   
   picFull(PicNum).ForeColor = TextColor
   picFull(PicNum).CurrentX = 0
   picFull(PicNum).CurrentY = 0
   picFull(PicNum).Print TheText$;
   
   ' Show TextColor
   picColor(2).BackColor = TextColor
 
'----------------------------------------------------------------
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' picFull(PicNum) to picThumb(PicNum)
   
   
   SetStretchBltMode picThumb(PicNum).hdc, HALFTONE
   
   StretchBlt picThumb(PicNum).hdc, 0, 0, _
      SmallPicW, SmallPicH, picFull(PicNum).hdc, _
      0, 0, W, H, vbSrcCopy
   picThumb(PicNum).Refresh

   SetStretchBltMode picThumb(PicNum).hdc, oldMODE

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      
   If PicNum = NumOfStoredPics Then  ' Bump info
      
      NumOfStoredPics = NumOfStoredPics + 1
      LabNSP = " NumOfStoredPics =" & Str$(NumOfStoredPics)
      ReDim Preserve StoreFileSpec$(PicNum)
      ReDim Preserve PicWidth(PicNum)
      ReDim Preserve PicHeight(PicNum)

   End If

   FileSpecString$ = "TLayer" & Str$(PicNum)
   ' Update info from mnuOpen
   StoreFileSpec$(PicNum) = FileSpecString$
   PicWidth(PicNum) = W
   PicHeight(PicNum) = H
   
   ' PicFull(PicNum) to picDisplay
   optSelect_Click (PicNum)
   
   '------------------------------------
   ' Enabled next picture in order
   PicNum = PicNum + 1
   
   If PicNum < MaxNumOfPics - 1 Then
      optSelect(PicNum).Enabled = True
      optSelect(PicNum).ForeColor = vbRed
   End If

   ' Enable swapping when enough pics loaded
   If NumOfStoredPics > 1 And PicNum < MaxNumOfPics Then
      If PicNum > 1 Then
         cmdSwapPics(PicNum - 2).Enabled = True
         cmdMerge(PicNum - 2).Enabled = True
      End If
   End If
'----------------------------------------------------------------
COMPLETE_LOADING

End Sub
'#### END Transparent with TEXT #######################################

Private Sub COMPLETE_LOADING()

   'COMPLETE_LOADING
   
   ' Enable Clip PicNum
   ' Set mnuFile stuff captions with PicNum
   ' Move picFull(PicNum) to picDisplay
   ' Highlight with shape rect
   optSelect(PicNum) = True
   
  ' Enable file menu stuff & some ACTIONS
   If NumOfStoredPics > 0 Then
      mnuDrawToggle.Enabled = True
      mnuText.Enabled = True
      mnuEffects.Enabled = True
      mnuSaveDisplay.Enabled = True
      mnuACTIONS.Enabled = True
   End If
   
   ' Enable clear & Merging pics when more than 1 stored
   If NumOfStoredPics > 1 Then
      chkSHOWALL.Enabled = True
      chkMERGE.Enabled = True
      'mnuRepositionLostPic.Enabled = True
      'mnuClearALLpics.Enabled = True
      'mnuClearNpics.Enabled = True
   End If
   

   ' FixScrollbars on displayed picture
   If PicNum = NumOfStoredPics Then
      ww = PicWidth(PicNum - 1)
      hh = PicHeight(PicNum - 1)
      a$ = StoreFileSpec$(PicNum - 1)
   Else
      ww = PicWidth(PicNum)
      hh = PicHeight(PicNum)
      a$ = StoreFileSpec$(PicNum)
   End If
      
   FixScrollbars picFrame, picDisplay, HS, VS
   LabInfo = ShortFileSpec(a$, 36) & "   W =" & Str$(ww) & "  H =" & Str$(hh)
   
   DoEvents
   
End Sub
'#### END LOAD & TLAYER & TLAYER+TEXT & RE-LOAD a pic ###


'#### Store picDisplay & Show Thumbs ####################

Private Sub Display2picFull_picThumb() 'PicNum

' Input PicNum

   ' W x H is the picDisplay size
   ' with loaded picture or altered
   ' picture
   
   ' Move picDisplay into picFull
   picFull(PicNum).Picture = LoadPicture
   
   With picFull(PicNum)
      .Width = W
      .Height = H
   End With
   
   ' Store in picFull()
   BitBlt picFull(PicNum).hdc, 0, 0, _
      W, H, picDisplay.hdc, 0, 0, vbSrcCopy
   
   picFull(PicNum).Refresh
   
   DoEvents
   
   ' Show Thumb in picThumb()
   
   With picThumb(PicNum)
      .Picture = LoadPicture
      .Width = SmallPicW
      .Height = SmallPicH
   End With
   picThumb(PicNum).Refresh
   
   picDisplay.Refresh
   
   DoEvents
   
   ' picDisplay to picThumb(picNum)
   SetStretchBltMode picThumb(PicNum).hdc, HALFTONE
   
   StretchBlt picThumb(PicNum).hdc, 0, 0, _
      SmallPicW, SmallPicH, picDisplay.hdc, _
      0, 0, W, H, vbSrcCopy
   picThumb(PicNum).Refresh

End Sub
'#### END Store picDisplay & Show Thumbs ####################

Private Sub CompleteStoring()
' Input PicNum

   Display2picFull_picThumb   ' picDisplay to picFull(PicNum) &
                              ' to picThumb(PicNum) (thumbs)

   If PicNum = NumOfStoredPics Then  ' Bump info
      NumOfStoredPics = NumOfStoredPics + 1
      LabNSP = " NumOfStoredPics =" & Str$(NumOfStoredPics)
      ReDim Preserve StoreFileSpec$(PicNum)
      ReDim Preserve PicWidth(PicNum)
      ReDim Preserve PicHeight(PicNum)
      If NumOfStoredPics = 2 Then chkSHOWALL.Enabled = True
   
   End If

   ' UPDATE INFO & BUTTONS
   
   ' Store input or new W & H
   PicWidth(PicNum) = W
   PicHeight(PicNum) = H

   ' Enabled next picture in order
   If PicNum < MaxNumOfPics - 1 Then
      optSelect(PicNum + 1).Enabled = True
      optSelect(PicNum + 1).ForeColor = vbRed
   End If

   ' Enable swapping when enough pics loaded
   If PicNum >= 1 And PicNum < MaxNumOfPics Then
      cmdSwapPics(PicNum - 1).Enabled = True
      cmdMerge(PicNum - 1).Enabled = True
   End If

   ' Enable clear & Merging pics when more than 1 stored
   If NumOfStoredPics > 1 Then
      mnuACTIONS.Enabled = True
      chkMERGE.Enabled = True
   End If

   ShowInfo

End Sub


'#### SAVING  ##################################################

Private Sub mnuSaveStoredPic_Click()
Dim Title$, Filt$, InDir$
Dim FileSpecString$
   
CLEAR_Incomplete_Actions
   
   If PicNum >= NumOfStoredPics Then
      MsgBox " MUST LOAD & STORE OR SELECT A PIC BEFORE SAVING", vbInformation, "Layers - Save pic"
      Exit Sub
   End If

   '  SAVE 24bpp BMP
   FileSpecString$ = vbNullString
   Title$ = "Save stored pic" & Str$(PicNum) & " as bmp"
   Filt$ = "Save bmp|*.bmp"
   InDir$ = FileSpecString$
   
   OS.ShowSave FileSpecString$, Title$, Filt$, InDir$, "", Me.hwnd
   
   If Len(FileSpecString$) = 0 Then
      Close
      Exit Sub
   End If
   
   FixExtension FileSpecString$
   
   SavePicture picFull(PicNum).Image, FileSpecString$

End Sub

Private Sub mnuSaveDisplay_Click()
Dim Title$, Filt$, InDir$
Dim FileSpecString$
   
CLEAR_Incomplete_Actions
   
   If aShowALLON Then
      MsgBox " CAN'T SAVE ALL PICTURES" & vbCr & " in SHOW ALL mode", vbInformation, "Layers - Save display"
      Exit Sub
   End If
   '  SAVE 24bpp BMP
   
   FileSpecString$ = vbNullString
   Title$ = "Save Display as bmp"
   Filt$ = "Save bmp|*.bmp"
   InDir$ = FileSpecString$
   
   OS.ShowSave FileSpecString$, Title$, Filt$, InDir$, "", Me.hwnd
   
   If Len(FileSpecString$) = 0 Then
      Close
      Exit Sub
   End If
   
   FixExtension FileSpecString$
   
   SavePicture picDisplay.Image, FileSpecString$

End Sub
'#### END SAVING  ##################################################


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   LabXY = Str$(X) & Str$(Y)
End Sub


'#### CLIPPING ########################################

Private Sub cmdClipMerged_Click()
' On Thumb bar
   mnuClip_Click

End Sub

Private Sub mnuClipMerged_Click()

   mnuClip_Click

End Sub

Private Sub cmdTBClip_Click()
 
   mnuClip_Click

End Sub

Private Sub mnuClip_Click()
   
   CLEAR_Incomplete_Actions
   
   If Not aMerged Then optSelect_Click (PicNum)   ' Brings up pic PicNum in picDisplay
   
   aClipON = True

   If aShowInstructions Then
      LabInstructions = " CLIP: Hold Left button down - Draw - Left button up," & vbCr & _
      "  then Accept, Redo (or just redraw) or Cancel"
      fraInstructions.Visible = True
   End If
   
   If aMerged Then
      fraAccRedoCancel.Caption = "Clip"
   Else
      fraAccRedoCancel.Caption = "Clip" & Str$(PicNum)
   End If
   fraAccRedoCancel.Visible = True
   
   ixTL = 0:  iyTL = 0
   ixwidth = W:  iyheight = H
   
   fraThumbBar.Visible = False

End Sub

Private Sub CLIPPER_Start_End(X As Single, Y As Single)

   picDisplay.MousePointer = vbCrosshair
   
   ixTL = CLng(X)
   iyTL = CLng(Y)
   ixBR = ixTL
   iyBR = iyTL
   prevX = ixBR
   prevY = iyBR
   aClipON = True
   
   With SR
      .Left = ixTL: .Top = iyTL
      .Width = ixBR - ixTL: .Height = iyBR - iyTL
      .Visible = True
      .BorderWidth = 2 'SLWidth
   End With
   
   ixwidth = ixBR - ixTL:  iyheight = iyBR - iyTL
   
End Sub

Private Sub CLIPPER_Move(Button As Integer, X As Single, Y As Single)

   picDisplay.MousePointer = vbCrosshair
   
   ixBR = CLng(X)
   iyBR = CLng(Y)
   
   If Button <> vbLeftButton Then Exit Sub
   
   If ixBR <= 0 Then ixBR = 1   ' Leave a gap
   If ixBR >= picDisplay.Width - 1 Then ixBR = picDisplay.Width - 2
   
   If picDisplay.Width > picFrame.Width Then
      If ixBR > picFrame.Width Then
         If ixBR > prevX Then
            If HS.Value < HS.Max Then HS.Value = HS.Value + 1
         End If
      End If
   End If
   
   If iyBR <= 0 Then iyBR = 1
   If iyBR >= picDisplay.Height - 1 Then iyBR = picDisplay.Height - 2

   If picDisplay.Height > picFrame.Height Then
      If iyBR > picFrame.Height Then
         If iyBR > prevY Then
            If VS.Value < VS.Max Then VS.Value = VS.Value + 1
         End If
      End If
   End If
   
   
   If ixBR < ixTL Then ixBR = ixTL
   If iyBR < iyTL Then iyBR = iyTL
   
   With SR
      .Width = ixBR - ixTL: .Height = iyBR - iyTL
      .Visible = True
      .BorderWidth = 2 'SLWidth
   End With
   
   ixwidth = ixBR - ixTL:  iyheight = iyBR - iyTL

   prevX = ixBR
   prevY = iyBR

End Sub
'#### END CLIPPING ########################################


'### BLITTERS ######################################################################

' 1.  picDisplay_To_picTemp

' 2.  picFull_To_picTemp
' 3.  picFull_To_picFullTemp
' 4.  picFull_To_picFullBackUp

' 6.  picTemp_To_picDisplay

' 7.  picFullTemp_To_picFull()   ' Takes Traced drawings
' 8.  picFullBackUp_To_picFull
' 9.  picFullTempBack_To_picFull

' 10. picFullTempBack_To_picFullTemp
' 11. picFullBackUp_To_picFullTemp

' 12. picFull_To_picFullTempBack

' 13. picFullBackUp_To_picFullTempBack

' 14. picFullTemp_To_picFullTempBack

' 15. picTemp_To_picFull

'1.
Private Sub picDisplay_To_picTemp()

   With picTemp
      .Width = picDisplay.Width
      .Height = picDisplay.Height
      .Picture = LoadPicture
      .BackColor = TransColor
      .ForeColor = TransColor
   End With
   
   BitBlt picTemp.hdc, 0, 0, picTemp.Width, picTemp.Height, _
      picDisplay.hdc, 0, 0, &H42
      
   BitBlt picTemp.hdc, 0, 0, picTemp.Width, picTemp.Height, _
      picDisplay.hdc, 0, 0, vbSrcCopy
      
   picTemp.Refresh
   
End Sub

'2.
Private Sub picFull_To_picTemp()

   With picTemp
      .Width = picFull(PicNum).Width
      .Height = picFull(PicNum).Height
      .Picture = LoadPicture
      .BackColor = TransColor
      .ForeColor = TransColor
   End With
   
   BitBlt picTemp.hdc, 0, 0, picFull(PicNum).Width, picFull(PicNum).Height, _
      picFull(PicNum).hdc, 0, 0, &H42
      
   BitBlt picTemp.hdc, 0, 0, picFull(PicNum).Width, picFull(PicNum).Height, _
      picFull(PicNum).hdc, 0, 0, vbSrcCopy
   
   picTemp.Refresh

End Sub

'3.
Private Sub picFull_To_picFullTemp()
   
   W = picFull(PicNum).Width
   H = picFull(PicNum).Height
   T = picFull(PicNum).Top
   L = picFull(PicNum).Left

   With picFullTemp
      .Picture = LoadPicture
      .Width = W
      .Height = H
      .Refresh
   End With
   
   BitBlt picFullTemp.hdc, 0, 0, W, H, _
      picFull(PicNum).hdc, 0, 0, vbSrcCopy
   
   picFullTemp.Refresh
   picFullTemp.DrawWidth = 1
   picFullTemp.DrawMode = 13

End Sub

'4.
Private Sub picFull_To_picFullBackUp()
         
   With picFullBackUp
      .Picture = LoadPicture
      .Width = W
      .Height = H
      .Refresh
   End With
   
   BitBlt picFullBackUp.hdc, 0, 0, W, H, _
      picFull(PicNum).hdc, 0, 0, vbSrcCopy
   
   picFullBackUp.Refresh
   picFullBackUp.DrawWidth = 1
   picFullBackUp.DrawMode = 13

End Sub

'5.
Private Sub picFull_To_PicDisplay()

   W = picFull(PicNum).Width
   H = picFull(PicNum).Height
   
   With picDisplay
      .Width = W
      .Height = H
      .Picture = LoadPicture
      .Refresh
   End With
   
   ' Show on picDisplay
   BitBlt picDisplay.hdc, 0, 0, W, H, _
      picFull(PicNum).hdc, 0, 0, vbSrcCopy
   
   picDisplay.Refresh

End Sub

'6.
Private Sub picTemp_To_picDisplay()

   With picDisplay
      .Width = picTemp.Width
      .Height = picTemp.Height
      .Picture = LoadPicture
   End With
   BitBlt picDisplay.hdc, 0, 0, picDisplay.Width, picDisplay.Height, _
      picTemp.hdc, 0, 0, vbSrcCopy
   
   picDisplay.Refresh

End Sub

'7.
Private Sub picFullTemp_To_picFull()

   W = picFull(PicNum).Width
   H = picFull(PicNum).Height
   T = picFull(PicNum).Top
   L = picFull(PicNum).Left

   With picFull(PicNum)
      .Picture = LoadPicture
      .Width = W
      .Height = H
      .DrawWidth = TheDrawWidth
      .Refresh
   End With
   
   BitBlt picFull(PicNum).hdc, 0, 0, W, H, _
      picFullTemp.hdc, 0, 0, vbSrcCopy
   
   picFull(PicNum).Refresh

End Sub

'8.
Private Sub picFullBackUp_To_picFull()

   W = picFull(PicNum).Width
   H = picFull(PicNum).Height
   
   picFull(PicNum).Picture = LoadPicture

   BitBlt picFull(PicNum).hdc, 0, 0, W, H, _
      picFullBackUp.hdc, 0, 0, vbSrcCopy
   
   picFullTemp.Refresh

   ' Show new Thumb
   SetStretchBltMode picThumb(PicNum).hdc, HALFTONE

   StretchBlt picThumb(PicNum).hdc, 0, 0, _
      SmallPicW, SmallPicH, picFull(PicNum).hdc, _
      0, 0, W, H, vbSrcCopy

   picThumb(PicNum).Refresh

End Sub

'9.
Private Sub picFullTempBack_To_picFull()

   W = picFull(PicNum).Width
   H = picFull(PicNum).Height
   
   With picFull(PicNum)
      .Picture = LoadPicture
      .Width = W
      .Height = H
      .Refresh
   End With
   
   BitBlt picFull(PicNum).hdc, 0, 0, W, H, _
      picFullTempBack.hdc, 0, 0, vbSrcCopy
   
   picFull(PicNum).Refresh
      
   ' Show new Thumb
   SetStretchBltMode picThumb(PicNum).hdc, HALFTONE
   
   StretchBlt picThumb(PicNum).hdc, 0, 0, _
      SmallPicW, SmallPicH, picFull(PicNum).hdc, _
      0, 0, W, H, vbSrcCopy
   
   picThumb(PicNum).Refresh
   
End Sub

'10.
Private Sub picFullTempBack_To_picFullTemp()
         
   W = picFull(PicNum).Width
   H = picFull(PicNum).Height

   ' Restore last picFullTemp from picFullTempBack
   With picFullTemp
      .Picture = LoadPicture
      .Width = W
      .Height = H
      .Refresh
   End With
   
   BitBlt picFullTemp.hdc, 0, 0, W, H, _
      picFullTempBack.hdc, 0, 0, vbSrcCopy
   
   picFullTemp.Refresh

End Sub

'11.
Private Sub picFullBackUp_To_picFullTemp()
            
   W = picFull(PicNum).Width
   H = picFull(PicNum).Height
   T = picFull(PicNum).Top
   L = picFull(PicNum).Left

   With picFullTemp
      .Picture = LoadPicture
      .Width = W
      .Height = H
      .Refresh
   End With

   BitBlt picFullTemp.hdc, 0, 0, W, H, _
      picFullBackUp.hdc, 0, 0, vbSrcCopy
   
   picFullTemp.Refresh

End Sub

' 12.
Private Sub picFull_To_picFullTempBack()

   W = picFull(PicNum).Width
   H = picFull(PicNum).Height
   T = picFull(PicNum).Top
   L = picFull(PicNum).Left

   With picFullTempBack
      .Picture = LoadPicture
      .Width = W
      .Height = H
      .Refresh
   End With
   
   BitBlt picFullTempBack.hdc, 0, 0, W, H, _
      picFull(PicNum).hdc, 0, 0, vbSrcCopy

   picFullTempBack.Refresh
   
End Sub

'13.
Private Sub picFullBackUp_To_picFullTempBack()

   W = picFullBackUp.Width
   H = picFullBackUp.Height
   T = picFullBackUp.Top
   L = picFullBackUp.Left

   With picFullTempBack
      .Picture = LoadPicture
      .Width = W
      .Height = H
      .Refresh
   End With
   
   BitBlt picFullTempBack.hdc, 0, 0, W, H, _
      picFullBackUp.hdc, 0, 0, vbSrcCopy

   picFullTempBack.Refresh

End Sub

'14.
Private Sub picFullTemp_To_picFullTempBack()

   W = picFullTemp.Width
   H = picFullTemp.Height

   With picFullTempBack
      .Picture = LoadPicture
      .Width = W
      .Height = H
      .Refresh
   End With
   
   BitBlt picFullTempBack.hdc, 0, 0, W, H, _
      picFullTemp.hdc, 0, 0, vbSrcCopy

   picFullTempBack.Refresh

End Sub

'15.
Private Sub picTemp_To_picFull()
   
   W = picFull(PicNum).Width
   H = picFull(PicNum).Height
   
   With picFull(PicNum)
      .Picture = LoadPicture
      .Width = W
      .Height = H
      .Refresh
   End With
   
   BitBlt picFull(PicNum).hdc, 0, 0, W, H, _
      picTemp.hdc, 0, 0, vbSrcCopy
   
   picFull(PicNum).Refresh
      
'   ' Show new Thumb
'   SetStretchBltMode picThumb(PicNum).hdc, HALFTONE
'
'   StretchBlt picThumb(PicNum).hdc, 0, 0, _
'      SmallPicW, SmallPicH, picFull(PicNum).hdc, _
'      0, 0, W, H, vbSrcCopy
'
'   picThumb(PicNum).Refresh

End Sub
'### END BLITTERS ######################################################################


'#### DRAWING CURSORS & INSTRUCTIONS ####################

Private Sub FIX_DRAW_CURSORS()

      Select Case TheDrawStyle
      Case 0 To 4
         picDisplay.MousePointer = vbDefault
      Case 5 To 28
         picDisplay.MousePointer = vbCustom
         picDisplay.MouseIcon = LoadResPicture("PENCIL", vbResCursor)
      Case 29  ' Fill
         picDisplay.MousePointer = vbCustom
         picDisplay.MouseIcon = LoadResPicture("FILL", vbResCursor)
      End Select

End Sub

Private Sub chkToggleInstructions_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   chkToggleInstructions.Value = 0
   
   aShowInstructions = Not aShowInstructions
   If aShowInstructions Then
      fraInstructions.Visible = True
   Else
      fraInstructions.Visible = False
   End If

End Sub

Private Sub Show_Instructions()

If aPicAColor Then
   LabInstructions = " Pick Color from anywhere: Press SHIFT KEY to set Draw Color" _
     & vbCr & Space$(54) & "CTRL KEY or  [X]  to Cancel"
Else
   
   Select Case TheDrawStyle
   Case 0
      LabInstructions = "0   BLEND WEAK" & vbCr & _
      "..LC-MOVE-RC to finish"
   Case 1
      LabInstructions = "1   BLEND MEDIUM" & vbCr & _
      "..LC-MOVE-RC to finish"
   Case 2
      LabInstructions = "2   BLEND STRING" & vbCr & _
      "..LC-MOVE-RC to finish"
   Case 3
      LabInstructions = "3   ERASE RECTANGLE" & vbCr & _
      "..LC-MOVE-RC to finish"
   Case 4
      LabInstructions = "4   ERASE CIRCLE" & vbCr & _
      "..LC-MOVE-RC to finish"
   Case 5 To 16, 19 To 22, 25 To 28
      Select Case TheDrawStyle
      Case 5: LabInstructions = "5   FREEDRAW"
      Case 6: LabInstructions = "6   LINE"
      Case 7: LabInstructions = "7   RECTANGLE"
      Case 8: LabInstructions = "8   FILLED RECTANGLE"
      Case 9: LabInstructions = "9   HORZ SHADED RECTANGLE"
      Case 10: LabInstructions = "10   VERT SHADED RECTANGLE"
      Case 11: LabInstructions = "11   CIRLLIPSE"
      Case 12: LabInstructions = "12   FILLED CIRLLIPSE"
      Case 13: LabInstructions = "13   SHADED CIRLLIPSE"
      Case 14: LabInstructions = "14   ARCH"
      Case 15: LabInstructions = "15   / F-RIBBON"
      Case 16: LabInstructions = "16   \ B-RIBBON"
      Case 19: LabInstructions = "19   SPRAY"
      Case 20: LabInstructions = "20   STAR"
      Case 21: LabInstructions = "21   PLUS SHAPE"
      Case 22: LabInstructions = "22   T-PIECE"
      Case 25: LabInstructions = "25   SINGLE ARROW"
      Case 26: LabInstructions = "26   FEATHERED ARROW"
      Case 27: LabInstructions = "27   DOUBLE ARROW"
      Case 28: LabInstructions = "28   TRIANGLE ARROW"
      End Select
         
      LabInstructions = LabInstructions & vbCr & _
         "..LC-DRAW-LC to end-MOVE to locate-" & vbCr & "..RC to fix"
   Case 17: LabInstructions = "17   POLYLINES" & vbCr & _
      "..LC-DRAW- Repeat for next segments - 1st" & vbCr & _
      "..RC to end -MOVE to locate 2nd RC to fix"
   Case 18: LabInstructions = "18   SPLINE" & vbCr & _
      "..LC-DRAW Repeat for next segments - 1st" & vbCr & _
      "..RC to end -MOVE to locate 2nd RC to fix"
   Case 23: LabInstructions = "23   PARALLELOGRAM" & vbCr & _
      "..LC-DRAW-LC-DRAW-LC to complete" & vbCr & _
      "..-MOVE to locate RC to fix"
   Case 24: LabInstructions = "24   FRUSTRUM" & vbCr & _
      "..LC-DRAW-LC-DRAW-LC to complete" & vbCr & _
      "..-MOVE to locate RC to fix"
   Case 29: LabInstructions = "29   FILL" & vbCr & _
      "..LC only"
   End Select
   
End If
End Sub


Private Sub fraInstructions_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   fraX = X
   fraY = Y

End Sub

Private Sub fraInstructions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   fraMOVER fraInstructions, Button, X, Y

End Sub

Private Sub cmdHideInstructions_Click()
    
    fraInstructions.Visible = False
    If aPicAColor Then
      If aDRAW Then
         aPicAColor = False
         Screen.MousePointer = vbDefault
         picColor(1).BackColor = DrawColor
         TheDrawStyle = svTheDrawStyle
         FIX_DRAW_CURSORS
         Show_Instructions
         If aShowInstructions Then fraInstructions.Visible = True
         frmTools.optTools(TheDrawStyle).Value = True
      Else
         picColor(1).BackColor = DrawColor
         aPicAColor = False
         Screen.MousePointer = vbDefault
      End If
    End If

End Sub
'#### END DRAWING CURSORS & INSTRUCTIONS ####################



'#### START LASSO ##########################################

Private Sub cmdTBLasso_Click()
   
   mnuLasso_Click

End Sub

Private Sub mnuLasso_Click()

   CLEAR_Incomplete_Actions

   fraAccRedoCancel.Caption = "Lasso pic" & Str$(PicNum)
   
   optSelect_Click (PicNum)   ' Brings up pic PicNum in picDisplay
   
   aLasso = True

   If aShowInstructions Then
      LabInstructions = " LASSO: Hold left-button down, Draw, Up to complete"
      fraInstructions.Visible = True
   End If
   
   picDisplay.MousePointer = vbCustom
   picDisplay.MouseIcon = LoadResPicture("LASSO", vbResCursor)
   
   fraAccRedoCancel.Visible = True
   
   fraThumbBar.Visible = False

End Sub

Private Sub LASSO_START(X As Single, Y As Single)

ixSL0 = CLng(X)
iySL0 = CLng(Y)

   Clear_LassoLines
   aLasso = True
   
   With S(0)
      .x1 = ixSL0: .y1 = iySL0
      .x2 = ixSL0: .y2 = iySL0
      .Visible = True
      .BorderWidth = 2 'SLWidth
   End With

   ' Start bounding rect coords
   ixTL = ixSL0: iyTL = iySL0
   ixBR = ixSL0: iyBR = iySL0

End Sub

Private Sub LASSO_MOVE(Button As Integer, X As Single, Y As Single)
Dim IX As Long
Dim IY As Long

   IX = CLng(X)
   IY = CLng(Y)
   
   If Button <> vbLeftButton Then Exit Sub
   
   If IX <= 0 Then IX = 1   ' Leave a gap
   If IX >= picDisplay.Width - 1 Then IX = picDisplay.Width - 2
   If IY <= 0 Then IY = 1
   If IY >= picDisplay.Height - 1 Then IY = picDisplay.Height - 2
   
   ' Show B/W/R S lines
   Select Case Rnd
   Case Is < 0.3: SCul = vbBlack
   Case Is < 0.6: SCul = vbWhite
   Case Else
      SCul = vbRed
   End Select

   NumOfSLines = NumOfSLines + 1
   Load S(NumOfSLines - 1)
   With S(NumOfSLines - 1)
      .x1 = S(NumOfSLines - 2).x2: .y1 = S(NumOfSLines - 2).y2
      .x2 = IX: .y2 = IY
      .BorderColor = SCul
      .Visible = True
   End With

   ' Find bounding rect coords of selection
   If IX < ixTL Then ixTL = IX
   If IX > ixBR Then ixBR = IX
   If IY < iyTL Then iyTL = IY
   If IY > iyBR Then iyBR = IY

End Sub

Private Sub LASSO_UP()

   ' Connect last to first point
   NumOfSLines = NumOfSLines + 1
   Load S(NumOfSLines - 1)
   With S(NumOfSLines - 1)
      .x1 = S(NumOfSLines - 2).x2
      .y1 = S(NumOfSLines - 2).y2
      .x2 = S(0).x1
      .y2 = S(0).y1
      .BorderColor = SCul
      .Visible = True
   End With

End Sub

'#### Extract Lasso Area ################################

Private Sub Extract_Lasso_Area()
Dim px As Long
Dim py As Long
Dim Cul As Long
Dim FillPtcul As Long

' Have:
' bounding rect on picDisplay of picNum
' ixTL,iyTL.........
' .........ixBR,iyBR

' S(i) lines i= 0 to NumOfSlines-1
   ww = picDisplay.Width
   hh = picDisplay.Height
'1)
   ' Size picTemp, blacken
   With picTemp
      .Picture = LoadPicture
      .Width = ww
      .Height = hh
      .BackColor = 0 ' black
      .DrawWidth = 2
      .ForeColor = RGB(255, 255, 255) ' White
   End With
'2)
   ' Draw White Lassoed shape on picTemp
   picTemp.PSet (S(0).x1, S(0).y1)
   For i = 1 To NumOfSLines - 1
      picTemp.Line -(S(i).x1, S(i).y1), RGB(255, 255, 255)
   Next i
   picTemp.Line -(S(0).x1, S(0).y1), RGB(255, 255, 255)
   
   '------------------------------------
   ' Have a white outline of select image
   ' Make mask: change outer region black to white
   picTemp.DrawStyle = vbSolid
   picTemp.DrawMode = 13
   picTemp.DrawWidth = 1
   picTemp.FillColor = vbWhite
   picTemp.FillStyle = vbFSSolid
   
   px = 0: py = 0    'FloodFill point
   Cul = picTemp.Point(px, py)
   If Cul <> 0 Then ' ie black
      For i = 1 To 10
         px = px + 1: py = py + 1
         Cul = picTemp.Point(px, py)
         If Cul = 0 Then Exit For 'Black found
      Next
      If i = 11 Then ' Black not found
         picTemp.FillStyle = vbFSTransparent  'Default (Transparent)
         MsgBox "Can't make mask.  Try aDrawing again", , "Lasso"
         Exit Sub
      End If
   End If
   FillPtcul = 0  ' Black
   'FLOODFILLSURFACE = 1
   'Fills with FillColor so long as point surrounded by FillPtcul&
   response = ExtFloodFill(picTemp.hdc, px, py, FillPtcul&, FLOODFILLSURFACE)
   picTemp.Refresh
   picTemp.FillStyle = vbFSTransparent  'Default (Transparent)
   picTemp.DrawWidth = 1
   ' Test if OK
   ' picTemp.Visible = True
   '------------------------------------
'3)
   ' Whiten picFull(PicNum)
   With picFull(PicNum)
      .Picture = LoadPicture
      .Width = ww
      .Height = hh
      .BackColor = vbWhite
   End With
   picFull(PicNum).Refresh
   
   '------------------------------------
'4)
   ' Get selected area to picFull(picNum)
   BitBlt picFull(PicNum).hdc, 0, 0, ww, hh, _
      picTemp.hdc, 0, 0, vbSrcAnd
   BitBlt picFull(PicNum).hdc, 0, 0, ww, hh, _
      picDisplay.hdc, 0, 0, vbSrcPaint
   
   '------------------------------------
'5)
   ' Fill white area with TransColor
   picFull(PicNum).DrawStyle = vbSolid
   picFull(PicNum).DrawMode = 13
   picFull(PicNum).DrawWidth = 1
   picFull(PicNum).FillColor = TransColor
   picFull(PicNum).FillStyle = vbFSSolid
   px = 0: py = 0    'FloodFill point
   Cul = picFull(PicNum).Point(px, py)
   If Cul <> vbWhite Then ' ie black
      For i = 1 To 10
         px = px + 1: py = py + 1
         Cul = picFull(PicNum).Point(px, py)
         If Cul = 0 Then Exit For 'Black found
      Next
      If i = 11 Then ' vbWhite not found
         picFull(PicNum).FillStyle = vbFSTransparent  'Default (Transparent)
         MsgBox "Can't make mask.  Try aDrawing again", , "Lasso"
         Exit Sub
      End If
   End If
   FillPtcul = vbWhite  ' Black
   'FLOODFILLSURFACE = 1
   'Fills with FillColor so long as point surrounded by FillPtcul&
   response = ExtFloodFill(picFull(PicNum).hdc, px, py, FillPtcul, FLOODFILLSURFACE)
   picFull(PicNum).Refresh
   picFull(PicNum).FillStyle = vbFSTransparent  'Default (Transparent)
   picFull(PicNum).DrawWidth = 1
   picFull(PicNum).Refresh
   ' Test if OK
   'picFull(PicNum).Visible = True
   '------------------------------------

   '------------------------------------
'6)
   ' Blit selected rect to picDisplay
   W = ixBR - ixTL
   H = iyBR - iyTL
   With picDisplay
      .Picture = LoadPicture
      .Width = W
      .Height = H
      .BackColor = 0
   End With
   picDisplay.Refresh
   
   BitBlt picDisplay.hdc, 0, 0, W, H, _
      picFull(PicNum).hdc, ixTL, iyTL, vbSrcCopy
   picDisplay.Refresh
   '------------------------------------
'7)
   ' & wrap up
   FixScrollbars picFrame, picDisplay, HS, VS
   CompleteStoring   ' Display2picFull_picThumb Also saves new W & H
   
   Me.Caption = " LAYERS"
   Screen.MousePointer = 0

End Sub
'#### END Extract Lasso Area ################################

Private Sub Clear_LassoLines()
   
   If NumOfSLines > 1 Then
      For i = 2 To NumOfSLines: Unload S(i - 1): Next
   End If
   S(0).Visible = False
   NumOfSLines = 1
   aLasso = False

End Sub
'#### END START LASSO ##########################################


'#### ACCEPT/REDO/CANCEL for LASSO, DRAW, PicAColor ############################

Private Sub cmdAccRedoCancel_Click(Index As Integer)
   
If aLasso Then
   Select Case Index
   Case 0   ' Accept
      ' Lasso accept
      
      If NumOfSLines > 1 Then
         Extract_Lasso_Area
      End If
      
      Clear_LassoLines
      
      picDisplay.MousePointer = vbDefault
      fraAccRedoCancel.Visible = False
      fraInstructions.Visible = False
      fraThumbBar.Visible = True
   Case 1   ' Redo
      Clear_LassoLines
      aLasso = True
      
   Case 2   ' Cancel
      picDisplay.MousePointer = vbDefault
      Clear_LassoLines
      fraAccRedoCancel.Visible = False
      fraInstructions.Visible = False
      fraThumbBar.Visible = True
   End Select
End If

If aPicAColor Then
   Select Case Index
   Case 0   ' Accept
      DrawColor = picColor(1).BackColor
      aPicAColor = False
      fraAccRedoCancel.Visible = False
      Screen.MousePointer = vbDefault
   
   Case 1   ' Redo
   Case 2   'Cancel
      picColor(1).BackColor = DrawColor
      aPicAColor = False
      fraAccRedoCancel.Visible = False
      Screen.MousePointer = vbDefault
   
   End Select
End If

If aClipON Then

   Select Case Index
   Case 0   ' Accept
      
      SR.Visible = False
      
      If ixwidth < 8 Or iyheight < 8 Then
         fraAccRedoCancel.Visible = False
         Exit Sub
      End If
      
      If Not aMerged Then
         
         W = ixwidth
         H = iyheight
         
         If PicNum = 0 Then
            maxw = ixwidth
            maxh = iyheight
         End If
         
         ' Resize picFull(PicNum)
         With picFull(PicNum)
            .Picture = LoadPicture
            .Width = ixwidth
            .Height = iyheight
            .Refresh
         End With
         
         ' Blit new rectangle from picDisplay to location on picFull(PicNum)
         BitBlt picFull(PicNum).hdc, 0, 0, ixwidth, iyheight, _
            picDisplay.hdc, ixTL, iyTL, vbSrcCopy
         
         picFull(PicNum).Refresh
         
         ' Resize picDisplay
         With picDisplay
            .Picture = LoadPicture
            .Width = ixwidth
            .Height = iyheight
            .Refresh
         End With
         
         ' Blit new rectangle from picFull(PicNum) to picDisplay
         BitBlt picDisplay.hdc, 0, 0, ixwidth, iyheight, _
            picFull(PicNum).hdc, 0, 0, vbSrcCopy
         
         picDisplay.Refresh
         
         aClipON = False
         
         FixScrollbars picFrame, picDisplay, HS, VS
         
         ' Store clipped picture
         
         CompleteStoring

         picDisplay.MousePointer = vbDefault
         
         fraInstructions.Visible = False
         fraThumbBar.Visible = True
      
      Else  ' Merge Clip
      
         ww = ixwidth
         hh = iyheight
         
         With picTemp
            .Width = ww
            .Height = hh
            .Picture = LoadPicture
         End With

         BitBlt picTemp.hdc, 0, 0, ww, hh, _
         picDisplay.hdc, ixTL, iyTL, &H42

         BitBlt picTemp.hdc, 0, 0, ww, hh, _
         picDisplay.hdc, ixTL, iyTL, vbSrcCopy

         With picDisplay
            .Width = ww
            .Height = hh
            .Picture = LoadPicture
         End With

         BitBlt picDisplay.hdc, 0, 0, ww, hh, _
         picTemp.hdc, 0, 0, &H42

         BitBlt picDisplay.hdc, 0, 0, ww, hh, _
         picTemp.hdc, 0, 0, vbSrcCopy
         
         aClipON = False
         'aClipped = True
         picDisplay.MousePointer = vbDefault
         
         FixScrollbars picFrame, picDisplay, HS, VS
      
      End If
      
      fraInstructions.Visible = False
      fraThumbBar.Visible = True
      fraAccRedoCancel.Visible = False
   
   Case 1   ' Redo clip rectangle
      
      SR.Visible = False
      'fraAccRedoCancel.Visible = False
   
   Case 2  ' Cancel No clipping
      
      SR.Visible = False
      
      aClipON = False
      picDisplay.MousePointer = vbDefault
   
      fraInstructions.Visible = False
      fraThumbBar.Visible = True
      fraAccRedoCancel.Visible = False
        
      aClipON = False
      'aClipped = False
   
   End Select  ' Select Case response
End If

End Sub
'#### END ACCEPT/REDO/CANCEL for LASSO, DRAW ############################


'#### MOVE MERGED PICS  ###################################

Private Sub MOVE_MERGE_START(X As Single, Y As Single)
      
      Dim R As RECT

      SetCursor LoadCursor(0, 32649&)  ' Quick hand cursor

      For j = NumOfStoredPics - 1 To 1 Step -1
      
         R.Left = picFull(j).Left
         R.Top = picFull(j).Top
         R.Right = R.Left + picFull(j).Width
         R.Bottom = R.Top + picFull(j).Height
         
         If PtInRect(R, X, Y) <> 0 Then
            
            picFullX = X - R.Left
            picFullY = Y - R.Top
            Exit For
         
         End If
      
      Next j
      
      jpic = j ' The select picFull(jpic)
      
      If jpic > 0 Then
         maxw = picFull(0).Width
         maxh = picFull(0).Height
         ReDim memBack(1 To maxw, 1 To maxh)
         memBack(1, 1) = TransColor
         memBack(maxw, maxh) = TransColor
         ptrmemBack = VarPtr(memBack(1, 1))
         FillFixedASMalpha
      End If

End Sub

Private Sub MOVE_MERGE_MOVE(X As Single, Y As Single)

      If jpic > 0 Then
      
         ' Done @ MouseDown
         'maxw = picFull(0).Width
         'maxh = picFull(0).Height
         
         ' Move hidden picFull(jpic)
         picFull(jpic).Left = X - picFullX
         picFull(jpic).Top = Y - picFullY
         
            
         '''''''   Calling MERGE    ' ASMON
         '''''''   Slow & Stack overflow
                        
         ' Done @ MouseDown
          ReDim memBack(1 To maxw, 1 To maxh)
          memBack(1, 1) = TransColor
          memBack(maxw, maxh) = TransColor
         ' NB Needs to be (1 To maxw) ??. (maxw) gives a streaky output ??!!

         ' Background Pic To memBack
         ' memback will contain final merged picture
         GETDIBS picFull(0).Image, 0

         ' Done @ MouseDown
         ptrmemBack = VarPtr(memBack(1, 1))
         FillFixedASMalpha
         
         For i = 1 To NumOfStoredPics - 1
            ' Following pictures
            ' Coords
            W = picFull(i).Width
            H = picFull(i).Height
            T = picFull(i).Top
            L = picFull(i).Left
            
            ReDim memPic(1 To W, 1 To H)
            'Get pic To memPic
            GETDIBS picFull(i).Image, i
            
            iza = PicAlpha(i)    ' 0   -> 256
            
            ptrmemPic = VarPtr(memPic(1, 1))
            
            FillVaryingASMalpha
            
            ' Merge current pic in memPic with background memBack
            ' where overlaps & taking account of PicAlpha (ie iza)
            response = CallWindowProc(ptrMC, ptrStruc, 2&, 3&, 4&)
                  
         Next i
         
         SetStretchBltMode picDisplay.hdc, HALFTONE
         
         StretchDIBits picDisplay.hdc, _
            0&, 0&, maxw, maxh, _
            0&, 0&, maxw, maxh, _
            memBack(1, 1), bmbac, 0, vbSrcCopy
         
         picDisplay.Refresh
         
         ' Done @ picDisplay_MouseUp
         'Erase memBack()
         'Erase memPic()
      
      End If
End Sub


'#### Moving fully visible pictures ##################################################

Private Sub picFull_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   If aFileOps = True Then Exit Sub

   If Index > 0 Then
      
      SetCursor LoadCursor(0, 32649&)  ' Quick hand cursor
      
      picFullX = X
      picFullY = Y

   End If
   
End Sub

Private Sub picFull_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

   If aFileOps = True Then Exit Sub
   
   If Index > 0 Then
      
      SetCursor LoadCursor(0, 32649&)  ' Quick hand cursor
   
      i = Index
      
      If Button <> 0 Then
         
         picFull(i).Left = picFull(i).Left + (X - picFullX)
         picFull(i).Top = picFull(i).Top + (Y - picFullY)
      
         LabXY = Str$(picFull(i).Left) & Str$(picFull(i).Top)
      
      End If
   
   End If

End Sub
'#### END Moving fully visible pictures ##################################################

'#### LOOP DELAY - SETTING TLIM FOR Sleep API ##############

Private Sub mnuSetLoopDelay_Click()
   
   txtTLIM.Text = Str$(TLIM)
   fraSetTLIM.Visible = True

End Sub

Private Sub txtTLIM_Change()
   
   'CLEAR_Incomplete_Actions
   
   If Not IsNumeric(txtTLIM.Text) Then txtTLIM.Text = "0"
   
   If Len(txtTLIM.Text) = 0 Then
      TLIM = 0
      txtTLIM.Text = "0"
   Else
      TLIM = Val(txtTLIM.Text)
   End If
   
   If TLIM < 0 Then
      TLIM = 0
      txtTLIM.Text = "0"
   End If
End Sub

Private Sub cmdTLIM_Click()
   
   fraSetTLIM.Visible = False
   
End Sub
'#### END LOOP DELAY - SETTING TLIM FOR Sleep API ##############


'#### SCROLL BARS #########################################

Private Sub HS_Change()
   picDisplay.Left = -HS.Value
End Sub

Private Sub HS_Scroll()
   picDisplay.Left = -HS.Value
End Sub

Private Sub picFrame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   LabXY = Str$(X) & Str$(Y)
End Sub

Private Sub VS_Change()
   picDisplay.Top = -VS.Value
End Sub


Private Sub VS_Scroll()
   picDisplay.Top = -VS.Value
End Sub


' #### ARRANGE THUMBBAR ########################################

Private Sub LineUpThumbBarElements()
'Dim DownSpec$

   fraThumbBar.Visible = False
   
   picThumbContainer.Width = (54 * MaxNumOfPics + 12) * STX
   
   
   'MaxNumOfPics = 21  ' optSelect(0 - 20) Set at Form_Load
   
   For i = 1 To MaxNumOfPics - 1
      Load picThumb(i): picThumb(i).Visible = True
      picThumb(i).Print "  pic(" & LTrim$(Str$(i)) & ")"
      
      Load optSelect(i): optSelect(i).Visible = True
      optSelect(i).Caption = " pic" & LTrim$(Str$(i))
      optSelect(i).Enabled = False
      
      Load cmdSwapPics(i): cmdSwapPics(i).Visible = True
      cmdSwapPics(i).Enabled = False
      
      Load cmdMerge(i): cmdMerge(i).Visible = True
      cmdMerge(i).Enabled = False
      
      Load picFull(i)
      picFull(i).ZOrder
   Next i
   
   For i = 0 To MaxNumOfPics - 1
      picThumb(i).Top = picThumb(0).Top
      picThumb(i).Width = 48
      picThumb(i).Height = 48
      picThumb(i).Left = 4 + 54 * i
      picThumb(i).Refresh
   Next i
   picThumb(0).Print "pic(0)"

   For i = 0 To MaxNumOfPics - 1
      optSelect(i).Top = optSelect(0).Top
      optSelect(i).Left = picThumb(i).Left
      optSelect(i).Width = 48
      If i > 0 Then
         optSelect(i).DownPicture = optSelect(i - 1).DownPicture
      End If
      If i < MaxNumOfPics - 1 Then
         cmdSwapPics(i).Left = picThumb(i).Left + 41
         cmdSwapPics(i).Top = cmdSwapPics(0).Top
         cmdSwapPics(i).ZOrder
         
         cmdMerge(i).Left = picThumb(i).Left + 41
         cmdMerge(i).Top = cmdMerge(0).Top
         cmdMerge(i).ZOrder
      End If
   Next i
   
   LeftPicNum = 0
   
   'fraThumbBar.Top = Form1.Height / STY - fraThumbBar.Height
   
   fraThumbBar.Visible = True
'   Me.Show

End Sub
' #### END ARRANGE THUMBBAR ########################################


'### HELP ######################################################

Private Sub mnuHelp_Click()

a$ = PathSpec$ & "LayersHelp.txt"
If Len(Dir$(a$)) = 0 Then
   MsgBox "LayersHelp.txt missing ", , "Layers - Help"
   Exit Sub
Else
   aHelp = True
   'frmHelp.Hide  ' Allows vbModal disabling other forms
   frmHelp.Show vbModeless ''vbModal
End If
End Sub
'### END HELP ######################################################

'#### MAGNIFIER #######################################

Private Sub chkMagnifier_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

   chkMagnifier.Value = 0
   
   If aMagON = False Then
      aMagON = True
      MagForm.Show
   Else
      MagForm.Hide
      aMagON = False
   End If

End Sub

'#### QUITTING STUFF #########################################

Private Sub mnuExit_Click()
   
   Form_QueryUnload 1, 0

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Form As Form

If aHelp Then Unload frmHelp

If UnloadMode = 0 Then    'Close on Form1 pressed
      
   response = MsgBox("", vbQuestion + vbYesNo, "Quit Application ?")
   If response = vbNo Then
      Cancel = True
   Else  'response= Yes
      Cancel = False
      
      Set OS = Nothing
      
      If aClipBoardUsed Then
         response = MsgBox(" Clipborad used - CLEAR?", vbYesNo + vbQuestion, "Layers - Quitting")
         If response = vbYes Then Clipboard.Clear
      End If
      
      ' Make sure all forms cleared
      For Each Form In Forms
         Unload Form
         Set Form = Nothing
      Next Form
      End
   
   End If
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim Form As Form

   Set OS = Nothing
   
   Screen.MousePointer = vbDefault
   
   ' Make sure all forms cleared
   For Each Form In Forms
      Unload Form
      Set Form = Nothing
   Next Form

   End

End Sub


Private Sub Command1_Click(Index As Integer)
' TESTS
Select Case Index
' Toggle picFull(picNum) visibility
Case 0  ' Red
   picFull(PicNum).Visible = True
Case 1     ' Green
   picFull(PicNum).Visible = False

' Toggle picTemp visibility
Case 2
   picTemp.Visible = True
Case 3
   picTemp.Visible = False
End Select
End Sub

