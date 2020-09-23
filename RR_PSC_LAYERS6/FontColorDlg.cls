VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' CFDialog  FontColorDlg.cls   Color Font Dialog
' From vbAccelerator.com

Option Explicit


'Public Enum EErrorCommonDialog
'    eeBaseCommonDialog = 13450  ' CommonDialog
'End Enum

Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
'    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Sub CopyMemoryStr Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, ByVal lpvSource As String, ByVal cbCopy As Long)

Private Type TCHOOSECOLOR
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long
End Type

Private Declare Function ChooseColor Lib "comdlg32.dll" Alias _
"ChooseColorA" (pChoosecolor As TCHOOSECOLOR) As Long

Public Enum EChooseColor
    CC_RGBInit = &H1
    CC_FullOpen = &H2
    CC_PreventFullOpen = &H4
    CC_ColorShowHelp = &H8
' Win95 only
    CC_SolidColor = &H80
    CC_AnyColor = &H100
' End Win95 only
    CC_ENABLEHOOK = &H10
    CC_ENABLETEMPLATE = &H20
    CC_EnableTemplateHandle = &H40
End Enum
Private Declare Function GetSysColor Lib "USER32" (ByVal nIndex As Long) As Long

Private Type TCHOOSEFONT
    lStructSize As Long         ' Filled with UDT size
    hWndOwner As Long           ' Caller's window handle
    hdc As Long                 ' Printer DC/IC or NULL
    lpLogFont As Long           ' Pointer to LOGFONT
    iPointSize As Long          ' 10 * size in points of font
    flags As Long               ' Type flags
    rgbColors As Long           ' Returned text color
    lCustData As Long           ' Data passed to hook function
    lpfnHook As Long            ' Pointer to hook function
    lpTemplateName As Long      ' Custom template name
    hInstance As Long           ' Instance handle for template
    lpszStyle As String         ' Return style field
    nFontType As Integer        ' Font type bits
    iAlign As Integer           ' Filler
    nSizeMin As Long            ' Minimum point size allowed
    nSizeMax As Long            ' Maximum point size allowed
End Type
Private Declare Function ChooseFont Lib "COMDLG32" _
    Alias "ChooseFontA" (chfont As TCHOOSEFONT) As Long

Private Const LF_FACESIZE = 32
Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Public Enum EChooseFont
    CF_ScreenFonts = &H1
    CF_PrinterFonts = &H2
    CF_BOTH = &H3
    CF_FontShowHelp = &H4
    CF_UseStyle = &H80
    CF_EFFECTS = &H100
    CF_AnsiOnly = &H400
    CF_NoVectorFonts = &H800
    CF_NoOemFonts = CF_NoVectorFonts
    CF_NoSimulations = &H1000
    CF_LimitSize = &H2000
    CF_FixedPitchOnly = &H4000
    CF_WYSIWYG = &H8000&  ' Must also have ScreenFonts And PrinterFonts
    CF_ForceFontExist = &H10000
    CF_ScalableOnly = &H20000
    CF_TTOnly = &H40000
    CF_NoFaceSel = &H80000
    CF_NoStyleSel = &H100000
    CF_NoSizeSel = &H200000
    ' Win95 only
    CF_SelectScript = &H400000
    CF_NoScriptSel = &H800000
    CF_NoVertFonts = &H1000000

    CF_InitToLogFontStruct = &H40
    CF_Apply = &H200
    CF_EnableHook = &H8
    CF_EnableTemplate = &H10
    CF_EnableTemplateHandle = &H20
    CF_FontNotSupported = &H238
End Enum

' These are extra nFontType bits that are added to what is returned to the
' EnumFonts callback routine

Public Enum EFontType
    Simulated_FontType = &H8000&
    Printer_FontType = &H4000
    Screen_FontType = &H2000
    Bold_FontType = &H100
    Italic_FontType = &H200
    Regular_FontType = &H400
End Enum



Private Declare Function CommDlgExtendedError Lib "COMDLG32" () As Long

' Array of custom colors lasts for life of app
Private alCustom(0 To 15) As Long, fNotFirst As Boolean

Private m_lApiReturn As Long
Private m_lExtendedError As Long

'Public Property Get APIReturn() As Long
'    'return object's APIReturn property
'    APIReturn = m_lApiReturn
'End Property
'Public Property Get ExtendedError() As Long
'    'return object's ExtendedError property
'    ExtendedError = m_lExtendedError
'End Property

'#If fComponent Then
Private Sub Class_Initialize()
    InitColors
End Sub
'#End If

'' ChooseColor wrapper
Function VBChooseColor(Color As Long, _
                       Optional AnyColor As Boolean = True, _
                       Optional FullOpen As Boolean = False, _
                       Optional DisableFullOpen As Boolean = False, _
                       Optional owner As Long = -1, _
                       Optional flags As Long) As Boolean

    Dim chclr As TCHOOSECOLOR
    chclr.lStructSize = Len(chclr)

    ' Color must get reference variable to receive result
    ' Flags can get reference variable or constant with bit flags
    ' Owner can take handle of owning window
    If owner <> -1 Then chclr.hWndOwner = owner

    ' Assign color (default uninitialized value of zero is good default)
    chclr.rgbResult = Color

    ' Mask out unwanted bits
    Dim afMask As Long
    afMask = CLng(Not (CC_ENABLEHOOK Or _
                       CC_ENABLETEMPLATE))
    ' Pass in flags
    chclr.flags = afMask And (CC_RGBInit Or _
                  IIf(AnyColor, CC_AnyColor, CC_SolidColor) Or _
                  (-FullOpen * CC_FullOpen) Or _
                  (-DisableFullOpen * CC_PreventFullOpen))

    ' If first time, initialize to white
    If fNotFirst = False Then InitColors

    chclr.lpCustColors = VarPtr(alCustom(0))
    ' All other fields zero

    m_lApiReturn = ChooseColor(chclr)
    
    Select Case m_lApiReturn
    Case 1
        ' Success
        VBChooseColor = True
        Color = chclr.rgbResult
    Case 0
        ' Cancelled
        VBChooseColor = False
        Color = -1
    Case Else
        ' Extended error
        m_lExtendedError = CommDlgExtendedError()
        VBChooseColor = False
        Color = -1
    End Select

End Function

Private Sub InitColors()
    Dim i As Long
    ' Initialize with first 16 system interface colors
    For i = 0 To 15
        alCustom(i) = GetSysColor(i)
    Next
    fNotFirst = True
End Sub

' ChooseFont wrapper
Function VBChooseFont(CurFont As Font, _
                      Optional PrinterDC As Long = -1, _
                      Optional owner As Long = -1, _
                      Optional Color As Long = vbBlack, _
                      Optional MinSize As Long = 0, _
                      Optional MaxSize As Long = 0, _
                      Optional flags As Long = 0) As Boolean

    m_lApiReturn = 0
    m_lExtendedError = 0

    ' Unwanted Flags bits
    Const CF_FontNotSupported = CF_Apply Or CF_EnableHook Or CF_EnableTemplate
    
    ' Flags can get reference variable or constant with bit flags
    ' PrinterDC can take printer DC
    If PrinterDC = -1 Then
        PrinterDC = 0
        If flags And CF_PrinterFonts Then PrinterDC = Printer.hdc
    Else
        flags = flags Or CF_PrinterFonts
    End If
    ' Must have some fonts
    If (flags And CF_PrinterFonts) = 0 Then flags = flags Or CF_ScreenFonts
    ' Color can take initial color, receive chosen color
    'If Color <> vbBlack Then flags = flags Or CF_EFFECTS
    'flags = flags Or CF_EFFECTS
    ' MinSize can be minimum size accepted
    If MinSize Then flags = flags Or CF_LimitSize
    ' MaxSize can be maximum size accepted
    If MaxSize Then flags = flags Or CF_LimitSize

    ' Put in required internal flags and remove unsupported
    flags = (flags Or CF_InitToLogFontStruct) And Not CF_FontNotSupported
    
    ' Initialize LOGFONT variable
    Dim fnt As LOGFONT
    Const PointsPerTwip = 1440 / 72
    fnt.lfHeight = -(CurFont.Size * (PointsPerTwip / Screen.TwipsPerPixelY))
    fnt.lfWeight = CurFont.Weight
    fnt.lfItalic = CurFont.Italic
    fnt.lfUnderline = CurFont.Underline
    fnt.lfStrikeOut = CurFont.Strikethrough
    ' Other fields zero
    StrToBytes fnt.lfFaceName, CurFont.Name

    ' Initialize TCHOOSEFONT variable
    Dim CF As TCHOOSEFONT
    CF.lStructSize = Len(CF)
    If owner <> -1 Then CF.hWndOwner = owner
    CF.hdc = PrinterDC
    CF.lpLogFont = VarPtr(fnt)
    CF.iPointSize = CurFont.Size * 10
    CF.flags = flags
    CF.rgbColors = Color
    CF.nSizeMin = MinSize
    CF.nSizeMax = MaxSize
    
    ' All other fields zero
    m_lApiReturn = ChooseFont(CF)
    Select Case m_lApiReturn
    Case 1
        ' Success
        VBChooseFont = True
        flags = CF.flags
        Color = CF.rgbColors
        CurFont.Bold = CF.nFontType And Bold_FontType
        'CurFont.Italic = cf.nFontType And Italic_FontType
        CurFont.Italic = fnt.lfItalic
        CurFont.Strikethrough = fnt.lfStrikeOut
        CurFont.Underline = fnt.lfUnderline
        CurFont.Weight = fnt.lfWeight
        CurFont.Size = CF.iPointSize / 10
        CurFont.Name = BytesToStr(fnt.lfFaceName)
    Case 0
        ' Cancelled
        VBChooseFont = False
    Case Else
        ' Extended error
        m_lExtendedError = CommDlgExtendedError()
        VBChooseFont = False
    End Select
        
End Function

Private Sub StrToBytes(ab() As Byte, S As String)
    If IsArrayEmpty(ab) Then
        ' Assign to empty array
        ab = StrConv(S, vbFromUnicode)
    Else
        Dim cab As Long
        ' Copy to existing array, padding or truncating if necessary
        cab = UBound(ab) - LBound(ab) + 1
        If Len(S) < cab Then S = S & String$(cab - Len(S), 0)
        'If UnicodeTypeLib Then
        '    Dim st As String
        '    st = StrConv(s, vbFromUnicode)
        '    CopyMemoryStr ab(LBound(ab)), st, cab
        'Else
            CopyMemoryStr ab(LBound(ab)), S, cab
        'End If
    End If
End Sub


Private Function BytesToStr(ab() As Byte) As String
    BytesToStr = StrConv(ab, vbUnicode)
End Function

Private Function IsArrayEmpty(va As Variant) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = va(LBound(va))
    IsArrayEmpty = (Err <> 0)
End Function





