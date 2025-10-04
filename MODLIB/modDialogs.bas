Attribute VB_Name = "modDialogs"
Option Explicit
Option Base 0

Public Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type

Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" _
    (pOpenfilename As OPENFILENAME) As Long

Public Const CC_ANYCOLOR = &H100
Public Const CC_ENABLEHOOK = &H10
Public Const CC_ENABLETEMPLATE = &H20
Public Const CC_ENABLETEMPLATEHANDLE = &H40
Public Const CC_FULLOPEN = &H2
Public Const CC_PREVENTFULLOPEN = &H4
Public Const CC_RGBINIT = &H1
Public Const CC_SHOWHELP = &H8
Public Const CC_SOLIDCOLOR = &H80

Public Type CHOOSECOLORSTRUCT
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        rgbResult As Long
        lpCustColors As Long
        flags As Long
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type

Public Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" _
    (pChoosecolor As CHOOSECOLORSTRUCT) As Long



Public Const CF_APPLY = &H200&
Public Const CF_ANSIONLY = &H400&
Public Const CF_TTONLY = &H40000
Public Const CF_EFFECTS = &H100&
Public Const CF_ENABLEHOOK = &H8&
Public Const CF_ENABLETEMPLATE = &H10&
Public Const CF_ENABLETEMPLATEHANDLE = &H20&
Public Const CF_FIXEDPITCHONLY = &H4000&
Public Const CF_FORCEFONTEXIST = &H10000
Public Const CF_INITTOLOGFONTSTRUCT = &H40&
Public Const CF_LIMITSIZE = &H2000&
Public Const CF_NOFACESEL = &H80000
Public Const CF_NOSCRIPTSEL = &H800000
Public Const CF_NOSTYLESEL = &H100000
Public Const CF_NOSIZESEL = &H200000
Public Const CF_NOSIMULATIONS = &H1000&
Public Const CF_NOVECTORFONTS = &H800&
Public Const CF_NOOEMFONTS = CF_NOVECTORFONTS
Public Const CF_NOVERTFONTS = &H1000000
Public Const CF_PRINTERFONTS = &H2
Public Const CF_SCALABLEONLY = &H20000
Public Const CF_SCREENFONTS = &H1
Public Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Public Const CF_SCRIPTSONLY = CF_ANSIONLY
Public Const CF_SELECTSCRIPT = &H400000
Public Const CF_SHOWHELP = &H4&
Public Const CF_USESTYLE = &H80&
Public Const CF_WYSIWYG = &H8000

Public Const BOLD_FONTTYPE = &H100
Public Const ITALIC_FONTTYPE = &H200
Public Const PRINTER_FONTTYPE = &H4000
Public Const SCREEN_FONTTYPE = &H2000
Public Const SIMULATED_FONTTYPE = &H8000

Public Type CHOOSEFONTSTRUCT
        lStructSize As Long
        hwndOwner As Long           '  caller's window handle
        hdc As Long                 '  printer DC/IC or NULL
        lpLogFont As Long
        iPointSize As Long          '  10 * size in points of selected font
        flags As Long               '  enum. type flags
        rgbColors As Long           '  returned text color
        lCustData As Long           '  data passed to hook fn.
        lpfnHook As Long            '  ptr. to hook function
        lpTemplateName As String    '  custom template name
        hInstance As Long           '  instance handle of.EXE that
                                    '    contains cust. dlg. template
        lpszStyle As String         '  return the style field here
                                    '  must be LF_FACESIZE or bigger
        nFontType As Integer        '  same value reported to the EnumFonts
                                    '    call back with the extra FONTTYPE_
                                    '    bits added
        MISSING_ALIGNMENT As Integer
        nSizeMin As Long            '  minimum pt size allowed &
        nSizeMax As Long            '  max pt size allowed if    CF_LIMITSIZE is used
End Type

Public Const LF_FACESIZE = 31
Public Type LOGFONT
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
        lfFaceName(0 To LF_FACESIZE) As Byte
End Type


Public Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" _
    (pChoosefont As CHOOSEFONTSTRUCT) As Long

Public aCustomColors(16) As Long




Public Const GMEM_MOVEABLE = &H2, GMEM_ZEROINIT = &H40

Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

'Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, ByVal lpInitData As Any) As Long
'Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

