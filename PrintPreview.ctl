VERSION 5.00
Begin VB.UserControl PrintPreview 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1812
      Left            =   600
      ScaleHeight     =   1812
      ScaleWidth      =   2052
      TabIndex        =   0
      Top             =   600
      Width           =   2052
   End
End
Attribute VB_Name = "PrintPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'--------------------------
'Submitted by Paolo Bonzini
'--------------------------

'Default Property Values:
Const m_def_Pages = 1
Const m_def_Page = 1

'Property Variables:
Dim m_ScaleMode As ScaleModeConstants
Dim m_Page As Integer
Dim m_Pages As Integer
Dim m_Ratio As Single

'Other Variables:
Dim m_Initialized As Boolean
Dim m_Target As Object
Dim m_PageWidth As Single
Dim m_PageHeight As Single

'User events:
Event Click()
Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event Resize()

'Printing events:
Event DrawPage(Page As Integer, Printing As Boolean)
Event EndPrinting(Canceled As Boolean)
Event StatusUpdate(CurrentPage As Integer, TotalPages As Integer, TotalCurrentPage As Long, Cancel As Boolean)

Private Declare Function GetTextExtentPoint Lib "gdi32" Alias "GetTextExtentPointA" (ByVal hdc As Long, ByVal lpszString As String, ByVal cbString As Long, lpSize As size) As Long
Private Declare Function SetTextJustification Lib "gdi32" (ByVal hdc As Long, ByVal nBreakExtra As Long, ByVal nBreakCount As Long) As Long

Private Type size
    x As Long
    Y As Long
End Type
Public Sub DrawCircle(x As Single, Y As Single, _
    Radius As Single, Optional Color As Long = -1, _
    Optional StartPos, Optional EndPos, _
    Optional Aspect As Single = 1)

    'Fail if they did not set the size of the page
    If Not m_Initialized Then Err.Raise 5

    'Set the color to a default value if it was not provided
    If Color = -1 Then Color = m_Target.ForeColor

    If IsMissing(StartPos) Then
        'raise argument not optional errors if endpos found
        'else draw a full circle
        If Not IsMissing(EndPos) Then Err.Raise 449
        m_Target.Circle (x, Y), Radius, Color, , , Aspect
    Else
        'raise argument not optional errors if endpos missing
        'else draw an arc or a pie.
        If IsMissing(EndPos) Then Err.Raise 449
        m_Target.Circle (x, Y), Radius, Color, _
            StartPos, EndPos, Aspect
    End If
End Sub
Public Sub Line(Flags As Integer, X1 As Single, Y1 As Single, _
    X2 As Single, Y2 As Single, Color As Long)

    'Fail if they did not set the size of the page
    If Not m_Initialized Then Err.Raise 5

    'Note that, for some strange reason, Visual Basic rejects
    ' m_Target.[Line] Flags, X1, Y1, X2, Y2, Color

    If m_Target Is picDraw Then
        picDraw.[Line] Flags, X1, Y1, X2, Y2, Color
    Else
        Printer.[Line] Flags, X1, Y1, X2, Y2, Color
    End If
End Sub
Public Sub PaintPicture(Picture As IPictureDisp, _
    X1 As Single, Y1 As Single, Optional Width1 As Variant, _
    Optional Height1 As Variant, Optional X2 As Variant, _
    Optional Y2 As Variant, Optional Width2 As Variant, _
    Optional Height2 As Variant, Optional Opcode As Variant)

    'Fail if they did not set the size of the page
    If Not m_Initialized Then Err.Raise 5

    picDraw.PaintPicture Picture, X1, Y1, Width1, Height1, _
        X2, Y2, Width2, Height2, Opcode
End Sub
Public Function Point(x As Single, Y As Single) As Long
    'Fail if they did not set the size of the page. Warning:
    'using this method is dangerous. Pairs of coordinates that
    'on screen refer to the same pixel might not do the same
    'on printer, and vice versa.
    If Not m_Initialized Then Err.Raise 5
    Point = picDraw.Point(x, Y)
End Function
Public Sub PrintDoc(Optional ByVal FirstPage As Integer = 1, _
    Optional ByVal LastPage As Integer, _
    Optional ByVal Copies As Integer = 1, _
    Optional ByVal Collate As Boolean = True)

    Dim i As Integer, j As Long, Kill As Boolean

    'Fail if they did not set the size of the page
    If Not m_Initialized Then Err.Raise 5
    With Printer
        'get the hdc for the printer. This is necessary for the
        'values that we give to the properties to be retained.
        'If we would not use this dummy Print statement, Visual
        'Basic would discard any assignment to the Printer
        'object's graphic properties.
        Printer.Width = m_PageWidth
        Printer.Height = m_PageHeight
        Printer.Print ""
        Kill = False
    
        If LastPage = 0 Then LastPage = Pages
    
        'The properties are transferred so that the DrawPage event
        'can assume know that properties that are never changed by
        'the program retain the same value that they had in design
        'mode. The user must not make any assumption on the value
        'of any other property, though
        Printer.ScaleLeft = picDraw.ScaleLeft
        Printer.ScaleTop = picDraw.ScaleTop
        Printer.ScaleWidth = picDraw.ScaleWidth
        Printer.ScaleHeight = picDraw.ScaleHeight
        Printer.ScaleMode = m_ScaleMode
        Printer.DrawMode = picDraw.DrawMode
        Printer.DrawWidth = picDraw.DrawWidth
        Printer.DrawStyle = picDraw.DrawStyle
        Printer.FillColor = picDraw.FillColor
        Printer.FillStyle = picDraw.FillStyle
        Printer.ForeColor = picDraw.ForeColor
        Printer.FontTransparent = picDraw.FontTransparent
        Set m_Target = Printer
    
        i = FirstPage
        For j = 1 To (LastPage - FirstPage + 1) * Copies
            RaiseEvent StatusUpdate(i - FirstPage + 1, _
                LastPage - FirstPage + 1, (j), Kill)
            If Kill Then Exit For
    
            If j <> 1 Then .NewPage
            If Collate Then
                'We are collating copies. This means that we output
                'an entire copy of the document at a time. Notice
                'the parentheses around i, needed to pass the
                'parameter by value.
                RaiseEvent DrawPage((i), True)
                If i = LastPage Then i = FirstPage Else i = i + 1
            Else
                'We are not collating copies. This means that we output
                'many copies of the same page before switching to
                'the next. Again, notice the parentheses.
                RaiseEvent DrawPage((i), True)
                If (j Mod Copies) = 0 Then i = i + 1
            End If
        Next j
        RaiseEvent EndPrinting((Kill))
    
        Set m_Target = picDraw
    
        If Kill Then .KillDoc Else .EndDoc
    End With
End Sub
Public Sub PrintText(Text As String)
    'Fail if they did not set the size of the page
    If Not m_Initialized Then Err.Raise 5

    If m_Target Is Printer Then
        m_Target.Print Text
        Exit Sub
    End If

    With picDraw
        Dim OldSize As Single
        Dim szBig As size, szSmall As size
        Dim Posi As Long, nBreak As Long

        'The text width in pixel must be exactly
        'm_Ratio * szBig, where szBig is the space used
        'by the text when using the same font that will
        'be used when printing.
        'We use SetTextJustification to add the
        'correct number of extra pixels
        GetTextExtentPoint hdc, Text, Len(Text), szBig
        OldSize = .Font.size
        .Font.size = .Font.size * m_Ratio
        GetTextExtentPoint .hdc, Text, Len(Text), szSmall

        Posi = InStr(Text, " ")
        Do Until Posi = 0
            nBreak = nBreak + 1
            Posi = InStr(Posi + 1, Text, " ")
        Loop
        SetTextJustification .hdc, szBig.x * m_Ratio - szSmall.x, nBreak
        m_Target.Print Text;
        SetTextJustification .hdc, 0, 0
        .Font.size = OldSize
    End With
End Sub
Public Sub Refresh()
    picDraw.Refresh
End Sub
Public Property Get ScaleMode() As ScaleModeConstants
    ScaleMode = m_ScaleMode
End Property
Public Property Let ScaleMode(ByVal _
    New_ScaleMode As ScaleModeConstants)

    m_ScaleMode = New_ScaleMode
    If Ambient.UserMode = False Or Not m_Initialized Then
        'Either we are in design mode, or they did not tell
        'us anything about the size of the page, so we are
        'not print previewing anything.
        m_Target.ScaleMode() = New_ScaleMode
    ElseIf m_Target Is Printer Then
        'We are speaking to the printer, and the printer is
        'already using the correct page size, so we just set
        'the Printer object’s ScaleMode to the one they asked
        'for.
        m_Target.ScaleMode() = New_ScaleMode
    ElseIf New_ScaleMode = vbUser Then
        'In User mode, we just use the ScaleWidth/ScaleHeight
        'supplied by the control container.
        m_Target.ScaleMode() = New_ScaleMode
    Else
        picDraw.ScaleMode = New_ScaleMode
        picDraw.Scale (0, 0)-( _
            picDraw.ScaleX(m_PageWidth, vbTwips), _
            picDraw.ScaleY(m_PageHeight, vbTwips))
    End If

    PropertyChanged "ScaleMode"
    PropertyChanged "ScaleHeight"
    PropertyChanged "ScaleWidth"
    PropertyChanged "ScaleLeft"
    PropertyChanged "ScaleTop"
End Property
Public Sub SetPaperSize(PaperWidth As Single, _
    PaperHeight As Single, _
    Optional PaperScale As ScaleModeConstants = vbInches)

    ' Set the physical page size:
    m_PageWidth = UserControl.ScaleX(PaperWidth, PaperScale)
    m_PageHeight = UserControl.ScaleY(PaperHeight, PaperScale)

    m_Initialized = True
    ScaleMode = ScaleMode
    UserControl_Resize
End Sub
Public Sub SetPixel(x As Single, Y As Single, Optional Color As Long = -1)
    'Fail if they did not set the size of the page
    If Not m_Initialized Then Err.Raise 5

    If Color = -1 Then Color = picDraw.ForeColor
    picDraw.PSet (x, Y), Color
End Sub
Public Sub SetScale(Optional X1 As Variant, _
    Optional Y1 As Variant, Optional X2 As Variant, _
    Optional Y2 As Variant)

    If IsMissing(X1) Then
        If Not IsMissing(X2) Then Err.Raise 449
        If Not IsMissing(Y1) Then Err.Raise 449
        If Not IsMissing(Y2) Then Err.Raise 449
        ScaleMode = vbTwips
    Else
        If IsMissing(X2) Then Err.Raise 449
        If IsMissing(Y1) Then Err.Raise 449
        If IsMissing(Y2) Then Err.Raise 449
        m_Target.Scale (X1, Y1)-(X2, Y2)
        m_ScaleMode = vbUser
    End If
End Sub
Public Function TextHeight(str As String) As Single

    If Not m_Initialized Then Err.Raise 5
    
    Dim sz As size
    GetTextExtentPoint hdc, str, Len(str), sz
    TextHeight = ScaleY((sz.Y), vbPixels)
    If m_Target Is picDraw Then TextHeight = TextHeight * m_Ratio
End Function
Public Function TextWidth(str As String) As Single

    If Not m_Initialized Then Err.Raise 5
    
    Dim sz As size
    GetTextExtentPoint hdc, str, Len(str), sz
    TextWidth = ScaleX((sz.x), vbPixels)
    If m_Target Is picDraw Then TextWidth = TextWidth * m_Ratio
End Function
Private Sub picDraw_Paint()
    Dim OldTarget As Object

    If Ambient.UserMode And m_Initialized Then
        Set OldTarget = m_Target
        Set m_Target = picDraw
        RaiseEvent DrawPage((m_Page), False)
        Set m_Target = OldTarget
    End If
End Sub
Private Sub UserControl_Resize()
    If m_Initialized Then
        m_Ratio = UserControl.ScaleHeight / m_PageHeight

        'Don't set width unless it is absolutely necessary
        If Abs((m_PageWidth * m_Ratio) - Width) > _
            Screen.TwipsPerPixelX Then

            Width = m_PageWidth * m_Ratio
        End If
    End If
    picDraw.Move UserControl.ScaleLeft, UserControl.ScaleTop, _
        UserControl.ScaleWidth, UserControl.ScaleHeight
    ScaleMode = ScaleMode
    picDraw.Refresh
    RaiseEvent Resize
End Sub
'Start of wizard-generated code
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Misc"
    ForeColor = m_Target.ForeColor
End Property
Public Property Let ForeColor(ByVal NewForeColor As OLE_COLOR)
    m_Target.ForeColor() = NewForeColor
    PropertyChanged "ForeColor"
End Property
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property
Public Property Get Font() As IFontDisp
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
    Set Font = m_Target.Font
End Property
Public Property Set Font(ByVal New_Font As IFontDisp)
    Set m_Target.Font = New_Font
    PropertyChanged "Font"
End Property
Public Property Get CurrentY() As Single
Attribute CurrentY.VB_MemberFlags = "400"
    CurrentY = m_Target.CurrentY
End Property
Public Property Let CurrentY(ByVal New_CurrentY As Single)
    m_Target.CurrentY() = New_CurrentY
    PropertyChanged "CurrentY"
End Property
Public Property Get CurrentX() As Single
Attribute CurrentX.VB_MemberFlags = "400"
    CurrentX = m_Target.CurrentX
End Property
Public Property Let CurrentX(ByVal New_CurrentX As Single)
    m_Target.CurrentX() = New_CurrentX
    PropertyChanged "CurrentX"
End Property
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_MemberFlags = "400"
    FontUnderline = m_Target.FontUnderline
End Property
Public Property Let FontUnderline( _
    ByVal New_FontUnderline As Boolean)
    m_Target.FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
End Property
Public Property Get FontTransparent() As Boolean
    FontTransparent = m_Target.FontTransparent
End Property
Public Property Let FontTransparent( _
    ByVal New_FontTransparent As Boolean)
    m_Target.FontTransparent() = New_FontTransparent
    PropertyChanged "FontTransparent"
End Property
Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_MemberFlags = "400"
    FontStrikethru = m_Target.FontStrikethru
End Property
Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    m_Target.FontStrikethru() = New_FontStrikethru
    PropertyChanged "FontStrikethru"
End Property
Public Property Get FontSize() As Single
Attribute FontSize.VB_MemberFlags = "400"
    FontSize = m_Target.FontSize
End Property
Public Property Let FontSize(ByVal New_FontSize As Single)
    m_Target.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property
Public Property Get FontName() As String
Attribute FontName.VB_MemberFlags = "400"
    FontName = m_Target.FontName
End Property
Public Property Let FontName(ByVal New_FontName As String)
    m_Target.FontName() = New_FontName
    PropertyChanged "FontName"
End Property
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_MemberFlags = "400"
    FontItalic = m_Target.FontItalic
End Property
Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    m_Target.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
End Property
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_MemberFlags = "400"
    FontBold = m_Target.FontBold
End Property
Public Property Let FontBold(ByVal New_FontBold As Boolean)
    m_Target.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property
Public Property Get FillStyle() As FillStyleConstants
Attribute FillStyle.VB_ProcData.VB_Invoke_Property = ";Misc"
    FillStyle = m_Target.FillStyle
End Property
Public Property Let FillStyle(ByVal New_FillStyle As FillStyleConstants)
    m_Target.FillStyle() = New_FillStyle
    PropertyChanged "FillStyle"
End Property
Public Property Get FillColor() As OLE_COLOR
Attribute FillColor.VB_ProcData.VB_Invoke_Property = ";Misc"
    FillColor = m_Target.FillColor
End Property
Public Property Let FillColor(ByVal NewFillColor As OLE_COLOR)
    m_Target.FillColor() = NewFillColor
    PropertyChanged "FillColor"
End Property
Public Property Get DrawWidth() As Integer
Attribute DrawWidth.VB_ProcData.VB_Invoke_Property = ";Misc"
    DrawWidth = m_Target.DrawWidth
End Property
Public Property Let DrawWidth(ByVal New_DrawWidth As Integer)
    m_Target.DrawWidth() = New_DrawWidth
    PropertyChanged "DrawWidth"
End Property
Public Property Get DrawStyle() As DrawStyleConstants
Attribute DrawStyle.VB_ProcData.VB_Invoke_Property = ";Misc"
    DrawStyle = m_Target.DrawStyle
End Property
Public Property Let DrawStyle(ByVal New_DrawStyle As DrawStyleConstants)
    m_Target.DrawStyle() = New_DrawStyle
    PropertyChanged "DrawStyle"
End Property
Public Property Get DrawMode() As DrawModeConstants
Attribute DrawMode.VB_ProcData.VB_Invoke_Property = ";Misc"
    DrawMode = m_Target.DrawMode
End Property
Public Property Let DrawMode(ByVal New_DrawMode As DrawModeConstants)
    m_Target.DrawMode() = New_DrawMode
    PropertyChanged "DrawMode"
End Property
Public Property Get hdc() As Long
Attribute hdc.VB_MemberFlags = "400"
    hdc = m_Target.hdc
End Property
Public Property Get hwnd() As Long
Attribute hwnd.VB_MemberFlags = "400"
    hwnd = UserControl.hwnd()
End Property
Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property
Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property
Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property
Public Property Let MousePointer( _
    ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property
Public Function ScaleY(Height As Single, _
    Optional FromScale As Variant, _
    Optional ToScale As Variant) As Single
    If Not m_Initialized Then Err.Raise 5
    If IsMissing(FromScale) Then FromScale = m_Target.ScaleMode
    If IsMissing(ToScale) Then ToScale = m_Target.ScaleMode
    ScaleY = m_Target.ScaleY(Height, FromScale, ToScale)
End Function
Public Function ScaleX(Width As Single, _
    Optional FromScale As Variant, _
    Optional ToScale As Variant) As Single
    If Not m_Initialized Then Err.Raise 5
    If IsMissing(FromScale) Then FromScale = m_Target.ScaleMode
    If IsMissing(ToScale) Then ToScale = m_Target.ScaleMode
    ScaleX = m_Target.ScaleX(Width, FromScale, ToScale)
End Function
Public Property Get ScaleWidth() As Single
Attribute ScaleWidth.VB_ProcData.VB_Invoke_Property = ";Scale"
    ScaleWidth = m_Target.ScaleWidth
End Property
Public Property Let ScaleWidth(ByVal New_ScaleWidth As Single)
    m_Target.ScaleWidth() = New_ScaleWidth
    m_ScaleMode = vbUser
    PropertyChanged "ScaleWidth"
    PropertyChanged "ScaleMode"
End Property
Public Property Get ScaleTop() As Single
Attribute ScaleTop.VB_ProcData.VB_Invoke_Property = ";Scale"
    ScaleTop = m_Target.ScaleTop
End Property
Public Property Let ScaleTop(ByVal New_ScaleTop As Single)
    m_Target.ScaleTop() = New_ScaleTop
    m_ScaleMode = vbUser
    PropertyChanged "ScaleTop"
    PropertyChanged "ScaleMode"
End Property
Public Property Get ScaleLeft() As Single
Attribute ScaleLeft.VB_ProcData.VB_Invoke_Property = ";Scale"
    ScaleLeft = m_Target.ScaleLeft
End Property
Public Property Let ScaleLeft(ByVal New_ScaleLeft As Single)
    m_Target.ScaleLeft() = New_ScaleLeft
    m_ScaleMode = vbUser
    PropertyChanged "ScaleLeft"
    PropertyChanged "ScaleMode"
End Property
Public Property Get ScaleHeight() As Single
Attribute ScaleHeight.VB_ProcData.VB_Invoke_Property = ";Scale"
    ScaleHeight = m_Target.ScaleHeight
End Property
Public Property Let ScaleHeight(ByVal New_ScaleHeight As Single)
    m_Target.ScaleHeight() = New_ScaleHeight
    m_ScaleMode = vbUser
    PropertyChanged "ScaleHeight"
    PropertyChanged "ScaleMode"
End Property
Public Property Get Page() As Integer
Attribute Page.VB_MemberFlags = "400"
    Page = m_Page
End Property
Public Property Let Page(ByVal New_Page As Integer)
    m_Page = New_Page
    PropertyChanged "Page"
    If Ambient.UserMode Then Refresh
End Property
Public Property Get Pages() As Integer
Attribute Pages.VB_MemberFlags = "400"
    Pages = m_Pages
End Property
Public Property Let Pages(ByVal New_Pages As Integer)
    m_Pages = New_Pages
    PropertyChanged "Pages"
End Property
Public Property Get Ratio() As Single
Attribute Ratio.VB_MemberFlags = "400"
    If Not m_Initialized Then Err.Raise 5
    Ratio = m_Ratio
End Property
Private Sub picDraw_Click()
    RaiseEvent Click
End Sub
Private Sub picDraw_DblClick()
    RaiseEvent DblClick
End Sub
Private Sub picDraw_MouseDown(Button As Integer, _
    Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, x, Y)
End Sub
Private Sub picDraw_MouseMove(Button As Integer, _
    Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, x, Y)
End Sub
Private Sub picDraw_MouseUp(Button As Integer, _
    Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, x, Y)
End Sub
Private Sub UserControl_Initialize()
    Set m_Target = picDraw
    m_Page = m_def_Page
End Sub
'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set Font = Ambient.Font
    Pages = m_def_Pages
End Sub
'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next

    With PropBag
        MousePointer = .ReadProperty("MousePointer", 0)
        Set MouseIcon = .ReadProperty("MouseIcon", Nothing)
        ForeColor = .ReadProperty("ForeColor", &H80000012)
        Enabled = .ReadProperty("Enabled", True)
        Font = .ReadProperty("Font", Ambient.Font)
        FillStyle = .ReadProperty("FillStyle", 1)
        FillColor = .ReadProperty("FillColor", &H0&)
        DrawWidth = .ReadProperty("DrawWidth", 1)
        DrawStyle = .ReadProperty("DrawStyle", 0)
        DrawMode = .ReadProperty("DrawMode", 13)
        Pages = .ReadProperty("Pages", m_def_Pages)
        ScaleWidth = .ReadProperty("ScaleWidth", 4800)
        ScaleTop = .ReadProperty("ScaleTop", 0)
        ScaleLeft = .ReadProperty("ScaleLeft", 0)
        ScaleHeight = .ReadProperty("ScaleHeight", 3600)
        ScaleMode = .ReadProperty("ScaleMode", 1)
    End With
End Sub
'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        .WriteProperty "ForeColor", ForeColor, &H80000012
        .WriteProperty "Enabled", Enabled, True
        .WriteProperty "Font", Font, Ambient.Font
        .WriteProperty "FillStyle", FillStyle, 1
        .WriteProperty "FillColor", FillColor, &H0&
        .WriteProperty "Pages", Pages, m_def_Pages
        .WriteProperty "DrawWidth", DrawWidth, 1
        .WriteProperty "DrawStyle", DrawStyle, 0
        .WriteProperty "DrawMode", DrawMode, 13
        .WriteProperty "ScaleWidth", ScaleWidth, Null
        .WriteProperty "ScaleTop", ScaleTop, 0
        .WriteProperty "ScaleLeft", ScaleLeft, 0
        .WriteProperty "ScaleHeight", ScaleHeight, Null
        .WriteProperty "ScaleMode", ScaleMode, 1
        .WriteProperty "MousePointer", MousePointer, 0
        .WriteProperty "MouseIcon", MouseIcon, Nothing
    End With
End Sub
