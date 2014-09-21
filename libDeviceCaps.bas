Attribute VB_Name = "libDeviceCaps"
'libDeviceCaps - libDeviceCaps.bas
'   GetDeviceCaps Interface...
'   Copyright © 2000, SunGard Investor Accounting Systems
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Problem:    Programmer:     Description:
'   02/07/00    None        Ken Clark       Documentation said that color depth is defined as:
'                                           2 ^ (GetDeviceCaps(hDC, BITSPIXEL) * GetDeviceCaps(hDC, PLANES))
'                                           ... if this is correct, then "Color Depth" isn't really what I intended, so now I'm
'                                           simply returning 2 ^ GetDeviceCaps(hDC, BITSPIXEL), because that seems to give me the
'                                           numbers I intended;
'   02/05/00    None        Ken Clark       Created;
'=================================================================================================================================
Option Explicit
Const DRIVERVERSION = 0
Const TECHNOLOGY = 2
Const HORZSIZE = 4
Const VERTSIZE = 6
Const HORZRES = 8
Const VERTRES = 10
Const BITSPIXEL = 12
Const PLANES = 14
Const NUMBRUSHES = 16
Const NUMPENS = 18
Const NUMMARKERS = 20
Const NUMFONTS = 22
Const NUMCOLORS = 24
Const PDEVICESIZE = 26
Const CURVECAPS = 28
Const LINECAPS = 30
Const POLYGONALCAPS = 32
Const TEXTCAPS = 34
Const CLIPCAPS = 36
Const RASTERCAPS = 38
Const ASPECTX = 40
Const ASPECTY = 42
Const ASPECTXY = 44
Const LOGPIXELSX = 88
Const LOGPIXELSY = 90
Const SIZEPALETTE = 104
Const NUMRESERVED = 106
Const COLORRES = 108
Const DT_PLOTTER = 0
Const DT_RASDISPLAY = 1
Const DT_RASPRINTER = 2
Const DT_RASCAMERA = 3
Const DT_CHARSTREAM = 4
Const DT_METAFILE = 5
Const DT_DISPFILE = 6
Const CP_NONE = 0
Const CP_RECTANGLE = 1
Const RC_BITBLT = 1
Const RC_BANDING = 2
Const RC_SCALING = 4
Const RC_BITMAP64 = 8
Const RC_GDI20_OUTPUT = &H10
Const RC_DI_BITMAP = &H80
Const RC_PALETTE = &H100
Const RC_DIBTODEV = &H200
Const RC_BIGFONT = &H400
Const RC_STRETCHBLT = &H800
Const RC_FLOODFILL = &H1000
Const RC_STRETCHDIB = &H2000

'Const CCHDEVICENAME = 32
'Const CCHFORMNAME = 32
'Type DEVMODE
'    dmDeviceName As String * CCHDEVICENAME
'    dmSpecVersion As Integer
'    dmDriverVersion As Integer
'    dmSize As Integer
'    dmDriverExtra As Integer
'    dmFields As Long
'    dmOrientation As Integer
'    dmPaperSize As Integer
'    dmPaperLength As Integer
'    dmPaperWidth As Integer
'    dmScale As Integer
'    dmCopies As Integer
'    dmDefaultSource As Integer
'    dmPrintQuality As Integer
'    dmColor As Integer
'    dmDuplex As Integer
'    dmYResolution As Integer
'    dmTTOption As Integer
'    dmCollate As Integer
'    dmFormName As String * CCHFORMNAME
'    dmUnusedPadding As Integer
'    dmBitsPerPel As Long
'    dmPelsWidth As Long
'    dmPelsHeight As Long
'    dmDisplayFlags As Long
'    dmDisplayFrequency As Long
'End Type
'Private Declare Function CreateIC Lib "gdi32" Alias "CreateICA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As DEVMODE) As Long
Private Declare Function CreateIC Lib "gdi32" Alias "CreateICA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Sub Get_Device_Information(hDC As Long)
    Debug.Print "(HORZSIZE)" & vbTab & "Width in millimeters:", GetDeviceCaps(hDC, HORZSIZE)
    Debug.Print "(VERTSIZE)" & vbTab & "Height in millimeters:", GetDeviceCaps(hDC, VERTSIZE)
    Debug.Print "(HORZRES)" & vbTab & "Width in Pixels:", GetDeviceCaps(hDC, HORZRES)
    Debug.Print "(VERTREZ)" & vbTab & "Height in raster Lines:", GetDeviceCaps(hDC, VERTRES)
    Debug.Print "(BITSPIXEL)" & vbTab & "Color bits per Pixel:", GetDeviceCaps(hDC, BITSPIXEL)
    Debug.Print "(PLANES)" & vbTab & "Number of Color Planes:", GetDeviceCaps(hDC, PLANES)
    Debug.Print "(NUMBRUSHES)" & vbTab & "Number of device brushes:", GetDeviceCaps(hDC, NUMBRUSHES)
    Debug.Print "(NUMPENS)" & vbTab & "Number of device pens:", GetDeviceCaps(hDC, NUMPENS)
    Debug.Print "(NUMMARKERS)" & vbTab & "Number of device markers:", GetDeviceCaps(hDC, NUMMARKERS)
    Debug.Print "(NUMFONTS)" & vbTab & "Number of device fonts:", GetDeviceCaps(hDC, NUMFONTS)
    Debug.Print "(NUMCOLORS)" & vbTab & "Number of device colors:", GetDeviceCaps(hDC, NUMCOLORS)
    Debug.Print "(PDEVICESIZE)" & vbTab & "Size of device structure:", GetDeviceCaps(hDC, PDEVICESIZE)
    Debug.Print "(ASPECTX)" & vbTab & "Relative width of pixel:", GetDeviceCaps(hDC, ASPECTX)
    Debug.Print "(ASPECTY)" & vbTab & "Relative height of pixel:", GetDeviceCaps(hDC, ASPECTY)
    Debug.Print "(ASPECTXY)" & vbTab & "Relative diagonal of pixel:", GetDeviceCaps(hDC, ASPECTXY)
    Debug.Print "(LOGPIXELSX)" & vbTab & "Horizontal dots per inch:", GetDeviceCaps(hDC, LOGPIXELSX)
    Debug.Print "(LOGPIXELSY)" & vbTab & "Vertical dots per inch:", GetDeviceCaps(hDC, LOGPIXELSY)
    Debug.Print "(SIZEPALETTE)" & vbTab & "Number of palette entries:", GetDeviceCaps(hDC, SIZEPALETTE)
    Debug.Print "(NUMRESERVED)" & vbTab & "Reserved palette entries:", GetDeviceCaps(hDC, NUMRESERVED)
    Debug.Print "(SIZEPALETTE)" & vbTab & "Actual color resolution:", GetDeviceCaps(hDC, SIZEPALETTE)
End Sub
Public Function GetColorDepth() As Double
    Dim hDC As Long
    hDC = CreateIC("DISPLAY", "", "", 0&)
    GetColorDepth = 2 ^ (GetDeviceCaps(hDC, BITSPIXEL)) ' * GetDeviceCaps(hDC, PLANES))
    Call DeleteDC(hDC)
End Function
Public Function GetXRes() As Double
    Dim hDC As Long
    hDC = CreateIC("DISPLAY", "", "", 0&)
    GetXRes = GetDeviceCaps(hDC, HORZRES)
    Call DeleteDC(hDC)
End Function
Public Function GetYRes() As Double
    Dim hDC As Long
    hDC = CreateIC("DISPLAY", "", "", 0&)
    GetYRes = GetDeviceCaps(hDC, VERTRES)
    Call DeleteDC(hDC)
End Function
'This one doesn't have anything to do with the GetDeviceCaps API function,
'but it's handy, and this is as good a place as any for it to live...
Function TranslateColor(ByVal clr As OLE_COLOR, Optional hPal As Long = 0) As Long
    If OleTranslateColor(clr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function
