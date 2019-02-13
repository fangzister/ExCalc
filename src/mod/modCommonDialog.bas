Attribute VB_Name = "modCommonDialog"
Option Explicit

Private Declare Function ChooseFont _
                Lib "comdlg32.dll" _
                Alias "ChooseFontA" (ByRef pChoosefont As TypeFont) As Long

Private Const LF_FACESIZE = 32

Private Const CF_PRINTERFONTS = &H2

Private Const CF_SCREENFONTS = &H1

Private Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)

Private Const CF_EFFECTS = &H100&

Private Const CF_FORCEFONTEXIST = &H10000

Private Const CF_INITTOLOGFONTSTRUCT = &H40&

Private Const CF_LIMITSIZE = &H2000&

Private Const REGULAR_FONTTYPE = &H400

'charset Constants

Private Const ANSI_CHARSET = 0

Private Const ARABIC_CHARSET = 178

Private Const BALTIC_CHARSET = 186

Private Const CHINESEBIG5_CHARSET = 136

Private Const DEFAULT_CHARSET = 1

Private Const EASTEUROPE_CHARSET = 238

Private Const GB2312_CHARSET = 134

Private Const GREEK_CHARSET = 161

Private Const HANGEUL_CHARSET = 129

Private Const HEBREW_CHARSET = 177

Private Const JOHAB_CHARSET = 130

Private Const MAC_CHARSET = 77

Private Const OEM_CHARSET = 255

Private Const RUSSIAN_CHARSET = 204

Private Const SHIFTJIS_CHARSET = 128

Private Const SYMBOL_CHARSET = 2

Private Const THAI_CHARSET = 222

Private Const TURKISH_CHARSET = 162

Private Type LogFont

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
    lfFaceName As String * 31

End Type

Private Type TypeFont

    lStructSize As Long
    hwndOwner As Long ' caller's window handle
    hDC As Long ' printer DC/IC or NULL
    lpLogFont As Long ' ptr. to a LOGFONT struct
    iPointSize As Long ' 10 * size in points of selected font
    Flags As Long ' enum. type flags
    rgbColors As Long ' returned text color
    lCustData As Long ' data passed to hook fn.
    lpfnHook As Long ' ptr. to hook function
    lpTemplateName As String ' custom template name
    hInstance As Long ' instance handle of.EXE that
    ' contains cust. dlg. template
    lpszStyle As String ' return the style field here
    ' must be LF_FACESIZE or bigger
    nFontType As Integer ' same value reported to the EnumFonts
    ' call back with the extra FONTTYPE_
    ' bits added
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long ' minimum pt size allowed &
    nSizeMax As Long ' max pt size allowed if

    ' CF_LIMITSIZE is used
End Type

Public Sub SetFont(obj As Control)

    Dim ret      As Long

    Dim cf       As TypeFont

    Dim lfont    As LogFont

    Dim FontName As String
    
    cf.Flags = CF_BOTH Or CF_EFFECTS Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT Or CF_LIMITSIZE
    cf.lpLogFont = VarPtr(lfont)
    cf.lStructSize = LenB(cf)
    'cf.lStructSize = Len(cf) ' size of structure
    cf.hwndOwner = obj.Parent.hwnd ' window Form1 is opening this dialog box
    cf.hDC = Printer.hDC ' device context of default printer (using VB's mechanism)
    cf.rgbColors = RGB(0, 0, 0) ' black
    cf.nFontType = REGULAR_FONTTYPE ' regular font type i.e. not bold or anything
    cf.nSizeMin = 10 ' minimum point size
    cf.nSizeMax = 72 ' maximum point size
    ret = ChooseFont(cf) 'brings up the font dialog
 
    If ret <> 0 Then ' success
        FontName = StrConv(lfont.lfFaceName, vbUnicode, &H804) 'Retrieve chinese font name in english version os
        FontName = Left$(FontName, InStr(1, FontName, vbNullChar) - 1)

        'Assign the font properties to text1
        With obj.Font
            .Charset = lfont.lfCharSet 'assign charset to font
            .Name = FontName
            .Size = cf.iPointSize / 10 'assign point size
        End With

    End If

End Sub

Public Sub SetFonts(Objs() As Control)

End Sub

