VERSION 5.00
Begin VB.Form frmText 
   BackColor       =   &H00008080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "   TEXT"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   381
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   543
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdTX 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   5160
      TabIndex        =   6
      Top             =   1890
      Width           =   1095
   End
   Begin VB.CommandButton cmdTX 
      Caption         =   "Text color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   5175
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdTX 
      Caption         =   "Font"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   5190
      TabIndex        =   4
      Top             =   705
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   2340
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "frmText.frx":0000
      Top             =   105
      Width           =   4740
   End
   Begin VB.CommandButton cmdTX 
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   5190
      TabIndex        =   0
      Top             =   165
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00008000&
      Caption         =   "Label2"
      ForeColor       =   &H00C0FFFF&
      Height          =   330
      Left            =   6705
      TabIndex        =   3
      Top             =   2100
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label1"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   2550
      Width           =   480
   End
End
Attribute VB_Name = "frmText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmText  frmText.frm

Option Explicit

Private CurFont As New StdFont

Private Sub Form_Load()
Me.Caption = " TEXT  - written on to a new Transparent Layer"

Text1.Text = vbNullString
Label1.BackColor = TransColor
Label1.ForeColor = TextColor

Set CurFont = New StdFont
End Sub


Private Sub cmdTX_Click(Index As Integer)
Dim cc As CFDialog

   Select Case Index
   Case 0   ' Return
      
      'Transfer TheText$ to a New Transparent Layer
      ' Gather info for Return
      W = Label1.Width + 2
      H = Label1.Height + 2
      PicNum = NumOfStoredPics '- 1
      TheText$ = Text1.Text
      TextFont.fntName = CurFont.Name
      TextFont.fntSize = CurFont.Size
      TextFont.fntItalic = CurFont.Italic
      TextFont.fntBold = CurFont.Bold
      
      Set CurFont = Nothing
      Unload Me
   
   Case 1   ' Font
      
      Set cc = New CFDialog
      ' Initial Font
      With CurFont
         .Size = 8
         .Weight = 1
         .Italic = False
         .Bold = False
         .Strikethrough = False
         .Underline = False
         .Name = "MS Sans Serif"
      End With
      
      With Label1
         .FontName = CurFont.Name
         .FontSize = CurFont.Size
         .FontItalic = CurFont.Italic
         .FontBold = CurFont.Bold
         .ForeColor = TextColor
         '.FontStrikethru = False
         '.FontUnderline = False
      End With

      If cc.VBChooseFont(CurFont, , Me.hwnd) Then
         With Label1
            .FontName = CurFont.Name
            .FontSize = CurFont.Size
            .FontItalic = CurFont.Italic
            .FontBold = CurFont.Bold
            .FontStrikethru = False
            .FontUnderline = False
         End With
         ' Show size
         Label2.Caption = Str$(Label1.Width) & Str$(Label1.Height)
      End If
   
      Set cc = Nothing
   
   Case 2   ' TextColor
   
      Dim CF As CFDialog
      Dim TheColor As Long
      
      Set CF = New CFDialog
   
      If CF.VBChooseColor(TheColor, , , , Me.hwnd) Then
         
         TextColor = TheColor
         Label1.ForeColor = TextColor
      
      End If
      
      Set CF = Nothing
   
   Case 3   ' Cancel
      
      TheText$ = vbNullString
      Set CurFont = Nothing
      Unload frmText
   
   End Select

End Sub


Private Sub Form_Unload(Cancel As Integer)

   Set CurFont = Nothing
   Unload frmText

End Sub

Private Sub Text1_Change()
   ' Show text
   Label1.Caption = Text1.Text
   ' Show size
   Label2.Caption = Str$(Label1.Width) & Str$(Label1.Height)
End Sub
