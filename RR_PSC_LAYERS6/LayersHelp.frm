VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "  Layers Demo Help"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   270
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   6540
      IntegralHeight  =   0   'False
      Left            =   2850
      TabIndex        =   2
      Top             =   135
      Width           =   5295
   End
   Begin VB.ListBox List1 
      Height          =   5820
      IntegralHeight  =   0   'False
      Left            =   135
      TabIndex        =   1
      Top             =   840
      Width           =   2640
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CLOSE"
      Height          =   330
      Left            =   270
      TabIndex        =   0
      Top             =   195
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "Contents"
      Height          =   195
      Left            =   180
      TabIndex        =   3
      Top             =   585
      Width           =   735
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmHelp (LayersHelp.frm)

Option Explicit

Option Base 1


Private Sub Command1_Click()
aHelp = False
Unload frmHelp
End Sub

Private Sub Form_Load()
Dim F4H As Long
Dim F4W As Long
Dim F4L As Long
Dim Contents As Long

' Size & make form stay on top
F4H = 0.5 * Form1.Height / STX
F4W = frmHelp.Width / STY
F4L = (Form1.Left + Form1.Width - frmHelp.Width) / STX

response = SetWindowPos(Me.hwnd, HWND_TOPMOST, _
       F4L, 60, F4W, F4H, wflags2)   ' X,Y,W,H

Form_Resize
   
Show
DoEvents

Open PathSpec$ & "LayersHelp.txt" For Input As #1
Screen.MousePointer = vbHourglass
Input #1, Contents
For i = 1 To Contents    ' Number of FVHelp Contents' items
   Line Input #1, a$
   List1.AddItem a$
Next i

Do Until EOF(1)
   
   Line Input #1, a$
   List2.AddItem a$
Loop

Close

Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
frmHelp.Left = Form1.Left + Form1.Width - frmHelp.Width
List2.Top = Command1.Top
List2.Height = frmHelp.Height - List2.Top - 650
List1.Top = Command1.Top + 600
List1.Height = frmHelp.Height - List1.Top - 650
End Sub

Private Sub Form_Unload(Cancel As Integer)
aHelp = False
Unload Me
End Sub

Private Sub List1_Click()
'Select item
i = List1.ListIndex
a$ = List1.List(i) & Chr$(0)
If Len(a$) <> 0 Then
   'Search List2 for Text$ & place at top
   response = SendMessageLong(List2.hwnd, LB_FINDSTRINGEXACT, -1&, _
   ByVal a$)
   
   List2.ListIndex = response
   If List2.ListIndex > 0 Then
      List2.TopIndex = List2.ListIndex - 1
   End If
End If
End Sub

