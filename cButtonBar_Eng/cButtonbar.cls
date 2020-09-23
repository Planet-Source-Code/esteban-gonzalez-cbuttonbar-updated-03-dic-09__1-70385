VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cButtonBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Clase:       cButtonBar
'Autor:       Esteban Gonzalez
'Fecha:       05 de diciembre de 2007
'Modificado:  03 de diciembre de 2009
'Descripcion: Class that emulates an horizontal bar of
'             buttons.
'             It allows the use of skins for the buttons.

Option Explicit

Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal Bytelen As Long)
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Const VK_TAB = &H9
Private Const VK_RETURN As Long = &HD
Private Const VK_SPACE As Long = &H20

'Functions to find out if mouse is over the picturebox.
Private Declare Function GetFocusAPI Lib "user32.dll" Alias "GetFocus" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long


Private Declare Function GetStockObject Lib "gdi32.dll" (ByVal nIndex As Long) As Long
Private Const BLACK_BRUSH As Long = 4
Private Const DKGRAY_BRUSH As Long = 3
Private Const WHITE_BRUSH As Long = 0
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Const COLOR_BTNFACE As Long = 15
Private Const COLOR_3DFACE As Long = COLOR_BTNFACE
Private Const COLOR_BTNHIGHLIGHT As Long = 20
Private Const COLOR_3DHILIGHT As Long = COLOR_BTNHIGHLIGHT
Private Const COLOR_3DHIGHLIGHT As Long = COLOR_BTNHIGHLIGHT
Private Const COLOR_BTNSHADOW As Long = 16
Private Const COLOR_3DSHADOW As Long = COLOR_BTNSHADOW

Private Declare Function GetCurrentObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal uObjectType As Long) As Long
Private Const OBJ_BITMAP As Long = 7
Private Const OBJ_BRUSH As Long = 2
Private Const OBJ_FONT As Long = 6
Private Const OBJ_PEN As Long = 1

Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal x1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function OffsetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function PtInRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long

Private Declare Function BeginPath Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function EndPath Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SelectClipPath Lib "gdi32.dll" (ByVal hdc As Long, ByVal iMode As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Const RGN_AND As Long = 1
Private Const RGN_COPY As Long = 5


'Api Functions for DC Operations
Private Declare Function GetBkMode Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetBkColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function GetBkColor Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function GetTextColor Lib "gdi32.dll" (ByVal hdc As Long) As Long

Private Const OPAQUE As Long = 2
Private Const TRANSPARENT As Long = 1
Private Const PS_SOLID As Long = 0
Private Const PS_DASH As Long = 1
Private Const PS_DOT As Long = 2


'Api Text Functions
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextEx Lib "user32.dll" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, ByRef lpRect As RECT, ByVal un As Long, ByRef lpDrawTextParams As Any) As Long

Private Const DT_LEFT As Long = &H0&
Private Const DT_RIGHT As Long = &H2&
Private Const DT_TOP As Long = &H0&
Private Const DT_VCENTER As Long = &H4&
Private Const DT_CENTER As Long = &H1&
Private Const DT_BOTTOM As Long = &H8&
Private Const DT_CALCRECT As Long = &H400&
Private Const DT_EXPANDTABS As Long = &H40&
Private Const DT_SINGLELINE As Long = &H20&
Private Const DT_MULTILINE As Long = (&H1&)
Private Const DT_END_ELLIPSIS As Long = &H8000&
Private Const DT_NOPREFIX As Long = &H800&
Private Const DT_WORD_ELLIPSIS As Long = &H40000

'Api Functions for Drawing the buttons with no skin
Private Declare Function DrawFocusRectAPI Lib "user32.dll" Alias "DrawFocusRect" (ByVal hdc As Long, ByRef lpRect As RECT) As Long
Private Declare Function DrawFrameControl Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Const DFC_BUTTON As Long = 4
Private Const DFCS_BUTTONPUSH As Long = &H10
Private Const DFCS_FLAT As Long = &H4000
Private Const DFCS_HOT As Long = &H1000
Private Const DFCS_INACTIVE As Long = &H100
Private Const DFCS_PUSHED As Long = &H200
Private Const DFCS_CHECKED As Long = &H400
Private Const DFCS_ADJUSTRECT As Long = &H2000

Private Declare Function DrawEdge Lib "user32.dll" (ByVal hdc As Long, ByRef qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Const BDR_INNER As Long = &HC
Private Const BDR_OUTER As Long = &H3
Private Const BDR_RAISED As Long = &H5
Private Const BDR_RAISEDINNER As Long = &H4
Private Const BDR_RAISEDOUTER As Long = &H1
Private Const BDR_SUNKEN As Long = &HA
Private Const BDR_SUNKENINNER As Long = &H8
Private Const BDR_SUNKENOUTER As Long = &H2
Private Const EDGE_BUMP As Long = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Private Const EDGE_ETCHED As Long = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const EDGE_RAISED As Long = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN As Long = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Const EDGE_FLAT As Long = BDR_RAISEDINNER
Private Const EDGE_FLAT_DOWN As Long = BDR_SUNKENOUTER
Private Const BF_ADJUST As Long = &H2000
Private Const BF_BOTTOM As Long = &H8
Private Const BF_FLAT As Long = &H4000
Private Const BF_LEFT As Long = &H1
Private Const BF_RIGHT As Long = &H4
Private Const BF_TOP As Long = &H2
Private Const BF_SOFT As Long = &H1000
Private Const BF_RECT As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)


'Api Bitmap Functions
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32.dll" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, ByRef lpBits As Any) As Long
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hdc As Long, ByRef pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function DrawState Lib "user32.dll" Alias "DrawStateA" (ByVal hdc As Long, ByVal HBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long
Private Declare Function CreateDIBPatternBrushPt Lib "gdi32.dll" (ByRef lpPackedDIB As Any, ByVal iUsage As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long

Private Const DST_BITMAP As Long = &H4
Private Const DST_ICON As Long = &H3
Private Const DST_TEXT As Long = &H1
Private Const DSS_DISABLED As Long = &H20
Private Const DSS_NORMAL As Long = &H0
Private Const DSS_MONO As Long = &H80


Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Const BI_RGB As Long = 0&
Private Const DIB_RGB_COLORS As Long = 0

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type RGBStruct
 Red As Byte
 Green As Byte
 Blue As Byte
 AlphaChannel As Byte
End Type

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(4) As Long
End Type

'Api functions for Drawing on a DC or Bitmap
Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, ByRef lpBits As Any, ByRef lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function FrameRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal HBrush As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal HBrush As Long) As Long
Private Declare Function Rectangle Lib "gdi32.dll" (ByVal hdc As Long, ByVal x1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function Ellipse Lib "gdi32.dll" (ByVal hdc As Long, ByVal x1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function Polygon Lib "gdi32.dll" (ByVal hdc As Long, ByVal lpPoint As Long, ByVal nCount As Long) As Long
Private Declare Function Polyline Lib "gdi32.dll" (ByVal hdc As Long, ByRef lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function APIGetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Private Declare Function ExtFloodFill Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Private Const FLOODFILLBORDER As Long = 0
Private Const FLOODFILLSURFACE As Long = 1

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
 X As Long
 Y As Long
End Type

Private Type SafeArrayBound
    cElements As Long
    lLbound As Long
End Type

Private Type SAFEARRAY1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 0) As SafeArrayBound
End Type

Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SafeArrayBound
End Type

'Functions for working with Assembler in VB. NOT USED!!
Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


'Constants for the ButtonBar.
Private Const fMASKCOLOR = 1
Private Const fAUTOSIZE = 2
Private Const fDRAWFOCUSRECT = 4

'Constants for identifying the buttons.
Private Const BUTTON_PREVIOUS = -2
Private Const BUTTON_NEXT = -1
Private Const NO_BUTTON = -3

'Constant that defines the margin between text and icon
'and borders.
Private Const MARGINSIZE = 3

'Constants that define button internal states.
Private Const BTN_VISIBLE = 1
Private Const BTN_SHOWABLE = 2
Private Const BTN_FULLVISIBLE = BTN_VISIBLE Or BTN_SHOWABLE

'Constant that defines the color of the icon shadow.
Private Const SHADOW_COLOR = &HC0C0C0

'Constantes that represents the posibles alignments
'of text and icon.
Public Enum eAlignment
 eLeft = 1
 eRight = 2
 eHCenter = 4
 eTop = 8
 eBottom = 16
 eVCenter = 32
End Enum

'Constants that represents the types of buttons for the bar.
Public Enum eButtonTypes
 eSeparator = 1
 eButton = 2
 eCheck = 4
 'eCheck = 0100 in binary
 'eOption = 1100 in binary
 'Implies that eOption type shares a bit with the
 'eCheck type, but not with eSeparator or eButton.
 eOption = 12
End Enum

'Constants that represents the button styles.
Public Enum eButtonStyles
 eFlat = 1
 eHot = 2
 e3D = 3
 eOwnerDrawn = 4
 eSkinned = 5
End Enum


'Constants that reprensents the button states.
Public Enum eButtonStates
 eNormal = 0
 eFocused = 2
 eOver = 1
 eDown = 4
 eDisabled = 8
End Enum

'Constants for the skin offsets according to
'the button state.
Private Const BTN_SKN_NORMAL  As Long = 0
Private Const BTN_SKN_FOCUSED  As Long = 1
Private Const BTN_SKN_OVER  As Long = 2
Private Const BTN_SKN_DOWN  As Long = 3
Private Const BTN_SKN_DISABLED  As Long = 4

'Constants for the filling type of the button.
Public Enum eFillMode
 eVerticalGradient
 eHorizontalGradient
 eTransparent
 eSolid
 ePatternFill
 eBlit
 eStretchBlit
End Enum

'Structure that contains info about a button.
Private Type tButton
 Text As String
 Icon As StdPicture
 Style As eButtonStyles
 State As eButtonStates
 ButtonType As eButtonTypes
 TextAlign As eAlignment
 IconAlign As eAlignment
 Tooltip As String
 Left As Long
 Top As Long
 Width As Long
 Height As Long
 'It's used just like the general skin, it must have the 5
 'different states within the image.
 Skin As StdPicture
 'Skin and Icon Maskcolor. NOT USED!!
 MaskColor As Long
 Flags As Long
End Type

Private Type tDCData
 DC As Long
 OldBmp As Long
 OldPen As Long
 oldBrush As Long
 OldFont As Long
 CurBmp As Long
 CurPen As Long
 CurBrush As Long
 CurFont As Long
 HasBmp As Boolean
End Type

Private WithEvents m_Container As PictureBox
Attribute m_Container.VB_VarHelpID = -1
Private WithEvents m_Font As StdFont
Attribute m_Font.VB_VarHelpID = -1
'Skin DC.
Private m_SkinDC As tDCData
'Picturebox DC (just for simplicity)
Private m_DC As tDCData
'SkinMask DC.
Private m_SkinMaskDC As tDCData
'BackBuffer DC
Private m_BackBuffer As tDCData
'Contains a DC with the picturebox size to temporary
'allocate the particular button skins.
Private m_ButtonSkinDC As tDCData
'General Skin.
Private m_Skin As StdPicture
'Skin Width.
Private m_SkinWidth As Long
'Skin Height.
Private m_SkinHeight As Long
'Width and Height of Skin Borders.
Private m_SkinBorderWidth(5) As Long
Private m_SkinBorderHeight(5) As Long
'Array that contains the info of the buttons.
Private m_Buttons() As tButton
'Quantity of buttons in the bar.
Private m_nButtons As Long
'Bar background color.
Private m_BackColor As Long
'First background color of the button for the following states:
'Normal, Focused, Disabled, Over y Down.
Private m_BackColor1(5) As Long
'Second button background color for the following states:
'Normal, Focused, Disabled, Over y Down.
Private m_BackColor2(5) As Long
'First button text color for the following states:
'Normal, Focused, Disabled, Over y Down.
Private m_TextColor1(5) As Long
'Second button background color for the following states:
'Normal, Focused, Disabled, Over y Down.
Private m_TextColor2(5) As Long
'Button border color for the following states:
'Normal, Focused, Disabled, Over y Down.
Private m_BorderColor(5) As Long
'Border Width.
Private m_BorderWidth As Long
'Bar flags..
Private m_BarFlags As Long
'Skin and icon transparent color.
Private m_MaskColor As Long
'Index of button that is currently pressed.
Private m_ButtonDown As Long
'Index of button that is currently mouse over.
Private m_ButtonOver As Long
'Index of button that is currently focused.
Private m_ButtonFocus As Long
'Variables that contains the Previous and Next buttons.
Private m_ButtonPrevious As tButton
Private m_ButtonNext As tButton
'Button background Filling mode.
Private m_FillMode As Long
'Default Width and Height of buttons.
Private m_DefaultButtonWidth As Long
Private m_DefaultButtonHeight As Long
'Horizontal space between two consecutives buttons.
Private m_HorizontalButtonGap As Long
'Vertical space between the button and the picturebox.
Private m_VerticalButtonGap As Long
'Background pattern for checked buttons with no skin.
Private m_CheckPatternBrush As Long
'Indicates if the repaint call was generated from the
'RedrawBar Sub.
Private m_RedrawAllButtons As Boolean
'Indicates if the bar is redrawn or not.
Private m_LockUpdate As Boolean
'Index First visible button.
Private m_FirstVisibleButton As Long
'Quantity of visible buttons.
Private m_nVisibleButtons As Long

Public Event Click(Index As Long)
Public Event MouseDown(Index As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Index As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Index As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseEnter(Index As Long)
Public Event MouseLeave(Index As Long)
Public Event GotFocus(Index As Long)
Public Event LostFocus(Index As Long, Desc As String)
Public Event Paint() '(Index As Long)
Public Event Resize() '(Index As Long)
Public Event KeyDown(Index As Long, KeyCode As Integer, Shift As Integer)
Public Event KeyUp(Index As Long, KeyCode As Integer, Shift As Integer)
Public Event KeyPressed(Index As Long, KeyAscii As Integer)


'=============================================
'Public Methods Definitions
'=============================================

'Creates a new button and add it to the bar.
'Returns the button index.
Public Function AddButton(BtnType As eButtonTypes, Text As String, Optional Style As eButtonStyles = eFlat, Optional Tooltip As String = "", Optional Img As StdPicture, Optional TextAlignment As eAlignment = eRight Or eVCenter, Optional ImgAlignment As eAlignment = eLeft Or eVCenter, Optional State As eButtonStates = eNormal) As Long
 Dim dx As Long, dy As Long
 
 'If there's no picturebox asociated to the bar, don't
 'create the button.
 If (m_Container Is Nothing) Then
  AddButton = -1
  Exit Function
 End If
 ReDim Preserve m_Buttons(m_nButtons) As tButton
 With m_Buttons(m_nButtons)
  .Text = Text
  .Tooltip = Tooltip
  .TextAlign = TextAlignment
  .IconAlign = ImgAlignment
'  Debug.Print "Icon Alignment = " & ImgAlignment
  .State = State
  .Style = Style
  .ButtonType = BtnType
  If (m_BarFlags And fAUTOSIZE) <> fAUTOSIZE Then
   .Height = m_DefaultButtonHeight
   .Width = m_DefaultButtonWidth
  Else
   If Not (Img Is Nothing) Then
    dx = m_Container.ScaleX(Img.Width)
    dy = m_Container.ScaleY(Img.Height)
   End If
   dx = dx + m_Container.TextWidth(.Text) + MARGINSIZE * 2
   dy = dy + m_Container.TextHeight(.Text) + MARGINSIZE * 2
   If Style = eSkinned Then
    dx = dx + m_SkinBorderWidth(0) * 2
    dy = dy + m_SkinBorderHeight(0) * 2
   Else
    dx = dx + m_BorderWidth * 2
    dy = dy + m_BorderWidth * 2
   End If
   .Width = dx
   .Height = dy
  End If
  .Top = m_VerticalButtonGap
  .Left = m_HorizontalButtonGap
  If m_nButtons > 0 Then
   .Left = .Left + m_Buttons(m_nButtons - 1).Left + m_Buttons(m_nButtons - 1).Width
  End If
  .Flags = BTN_FULLVISIBLE
  If BtnType = eSeparator Then
   Set .Skin = Img
   .Width = m_Container.ScaleX(.Skin.Width)
   .Height = m_Container.ScaleY(.Skin.Height)
  Else
   Set .Icon = Img
  End If
 End With
 AddButton = m_nButtons
' Debug.Print "AddButton " & m_nButtons & " - m_FirstVisibleButton = " & m_FirstVisibleButton
 m_nButtons = m_nButtons + 1
'' Debug.Print "RedrawBar from AddButton"
'' RedrawBar
 CheckVisibleButtons
End Function

'Deletes a button or a separator from the bar.
'Returns the number of buttons and separators in the bar.
Public Function RemoveButton(Index As Long) As Long
 Dim i As Long
 
 If Index >= 0 And Index < m_nButtons Then
  If Index < m_nButtons Then
   'Reorder the array.
'   CopyMemory m_Buttons(Index), m_Buttons(Index + 1), (m_nButtons - Index) * Len(m_Buttons(Index))
   For i = Index To m_nButtons - 2
    m_Buttons(i) = m_Buttons(i + 1)
   Next i
  End If
  m_nButtons = m_nButtons - 1
  If m_nButtons > 1 Then
   ReDim Preserve m_Buttons(m_nButtons) As tButton
  End If
  RemoveButton = m_nButtons
'  Debug.Print "RedrawBar de RemoveButton"
'  Debug.Print "RemoveButton - M_FIRSTVISIBLEBUTTON = " & m_FirstVisibleButton
'  Debug.Print "RemoveButton - Indice = " & Indice
  If Index < m_FirstVisibleButton Then
   m_FirstVisibleButton = m_FirstVisibleButton - 1
   UpdateButtonPos Index, 0, False
  Else
   If Index = m_FirstVisibleButton Then
    m_Buttons(Index).Left = m_HorizontalButtonGap
    UpdateButtonPos Index + 1, m_nButtons - 1, True
   Else
    UpdateButtonPos Index, m_nButtons - 1, True
   End If
  End If
  CheckVisibleButtons
  RedrawBar
 End If
End Function

'Change the button position.
'Returns the number of buttons in the bar.
Public Function MoveButton(Index As Long, Position As Long) As Long
 Dim Temp As tButton, i As Long
 
 'Is Valid index?
 If Index >= 0 And Index < m_nButtons Then
  'Is Valid Position?
  If Position >= 0 And Position < m_nButtons Then
   MoveButton = m_nButtons
   'If there's no movement, exit.
   If Index = Position Then
    Exit Function
   'If position is located to the right of the button
   ElseIf Index < Position And Index < (m_nButtons - 1) Then
    'Move the button to the right
    Temp = m_Buttons(Index)
    For i = Index + 1 To Position
     m_Buttons(i - 1) = m_Buttons(i)
    Next i
    m_Buttons(Position) = Temp
    If Position >= m_FirstVisibleButton And Index < m_FirstVisibleButton Then
     m_Buttons(m_FirstVisibleButton).Left = m_HorizontalButtonGap
     If m_FirstVisibleButton > 0 Then UpdateButtonPos m_FirstVisibleButton - 1, 0, False
     If Position >= m_FirstVisibleButton Then UpdateButtonPos m_FirstVisibleButton + 1, m_nButtons - 1, True
    Else
     m_Buttons(Index).Left = m_Buttons(Position).Left
    UpdateButtonPos Index + 1, Position, (Position > Index)
    End If
   'If position is located to the left of the button
   ElseIf Index > Position Then
    'Move the button to the left
    Temp = m_Buttons(Index)
    For i = (Index - 1) To Position Step -1
     m_Buttons(i + 1) = m_Buttons(i)
    Next i
    m_Buttons(Position) = Temp
'    Debug.Print "FirstVisibleButton = " & m_Buttons(m_FirstVisibleButton).Text
    If Position < m_FirstVisibleButton And Index >= m_FirstVisibleButton Then
    'If we change the button order, the first visible button
    'does not remain the same.
     m_Buttons(m_FirstVisibleButton).Left = m_HorizontalButtonGap
     If m_FirstVisibleButton <= Index Then UpdateButtonPos m_FirstVisibleButton + 1, m_nButtons - 1, True
     If m_FirstVisibleButton > 0 Then UpdateButtonPos m_FirstVisibleButton - 1, 0, False
    Else
    'If Position or Index are both previous or later to
    'the first visible button, then:
     m_Buttons(Position).Left = m_Buttons(Position + 1).Left
     UpdateButtonPos Position + 1, Index, True '(Position < Index)
    End If
   End If
'   Debug.Print "RedrawBar from MoveButton"
'   Debug.Print "MoveButtons New Coords. (" & Index & ", " & Position & ") - FirstVisible = " & m_FirstVisibleButton
   For i = 0 To m_nButtons - 1
    Debug.Print m_Buttons(i).Text & " - m_Buttons(" & i & ").Left = " & m_Buttons(i).Left
   Next i
   Debug.Print
   CheckVisibleButtons
   RedrawBar
  End If
 End If
End Function

'Asociates the picturebox to the class, and creates the bar.
Public Function Create(ContainerControl As PictureBox) As Long
 Dim i As Long, SysColor As Long
 Dim Result As Boolean
 Dim mbi As BITMAPINFO
 Dim Color As RGBQUAD
 
 Set m_Container = ContainerControl
 m_Container.ScaleMode = vbPixels
 Create = m_Container.hwnd
 'Create a DC for BackBuffer.
 Result = CloneDC(m_Container.hdc, m_Container.ScaleWidth, m_Container.ScaleHeight, m_BackBuffer)
 'Create a DC for the skin.
 Result = CloneDC(0, 0, 0, m_SkinDC)
 'Create a DC for the skin mask.
 Result = CloneDC(0, 0, 0, m_SkinMaskDC)
 'Create a DC with the picturebox size to temporay allocate
 'the particular button skins.
 Result = CloneDC(m_Container.hdc, m_Container.ScaleWidth, m_Container.ScaleHeight, m_ButtonSkinDC)
 
 m_DC.DC = m_Container.hdc
 If (m_Container.BackColor And &H80000000) = &H80000000 Then
 'Is a system color
  SysColor = (m_Container.BackColor And &H7FFFFFFF)
  m_BackColor = GetSysColor(SysColor)
 Else
  m_BackColor = m_Container.BackColor
 End If
 For i = 0 To 5
  m_BackColor1(i) = m_BackColor
  m_TextColor1(i) = 0&
  m_TextColor2(i) = &HFFFFFF
 Next i
 'Create a brush pattern for check buttons with no skin.
 With mbi
  .bmiHeader.biSize = LenB(.bmiHeader)
  .bmiHeader.biBitCount = 32
  .bmiHeader.biHeight = 2
  .bmiHeader.biWidth = 2
  .bmiHeader.biPlanes = 1
  .bmiHeader.biCompression = BI_RGB
  .bmiColors(0) = m_BackColor
  .bmiColors(1) = RGB(255, 255, 255)
  .bmiColors(2) = RGB(255, 255, 255)
  .bmiColors(3) = m_BackColor
 End With
 m_CheckPatternBrush = CreateDIBPatternBrushPt(mbi, DIB_RGB_COLORS)
' Debug.Print "HBrush = " & HBrush

'Set the coordinates of navigation buttons (Previous and Next).
 m_ButtonNext.Width = m_Container.TextWidth("<<") + (m_BorderWidth * 2) + (MARGINSIZE * 2)
 m_ButtonNext.Left = m_Container.ScaleWidth - m_HorizontalButtonGap - m_ButtonNext.Width
 m_ButtonNext.Top = m_VerticalButtonGap
 m_ButtonNext.Height = m_Container.ScaleHeight - (m_VerticalButtonGap * 2)
 'Set text and alignment.
 m_ButtonNext.TextAlign = eLeft Or eVCenter
 m_ButtonNext.Text = ">>"
 'Set button type.
 m_ButtonNext.ButtonType = eButton
 m_ButtonNext.Style = e3D
 m_ButtonNext.State = eNormal
 
 m_ButtonPrevious.Left = m_ButtonNext.Left - m_HorizontalButtonGap - m_ButtonNext.Width
 m_ButtonPrevious.Width = m_ButtonNext.Width
 m_ButtonPrevious.Top = m_VerticalButtonGap
 m_ButtonPrevious.Height = m_Container.ScaleHeight - (m_VerticalButtonGap * 2)
 'Set text and alignment.
 m_ButtonPrevious.TextAlign = eLeft Or eVCenter
 m_ButtonPrevious.Text = "<<"
 'Set button type.
 m_ButtonPrevious.ButtonType = eButton
 m_ButtonPrevious.Style = e3D
 m_ButtonPrevious.State = eNormal
End Function

'Release every object and resource asociated with the bar
'and destroys it.
Public Sub Destroy()
 Dim i As Long
 
 'Exit if there's no bar.
 If m_Container Is Nothing Then Exit Sub

 DestroyDC m_SkinDC
 DestroyDC m_SkinMaskDC
 DestroyDC m_BackBuffer
 DestroyDC m_ButtonSkinDC
 DeleteObject m_CheckPatternBrush
 Set m_Skin = Nothing
 Set m_Font = Nothing
 For i = 0 To m_nButtons - 1
  Set m_Buttons(i).Skin = Nothing
  Set m_Buttons(i).Icon = Nothing
 Next i
 m_nButtons = 0
 Erase m_Buttons
 Set m_Container = Nothing
End Sub

'Redraw a given button or the whole bar.
Public Sub Refresh(Optional ByVal BtnIndex As Long = NO_BUTTON)
 If BtnIndex = NO_BUTTON Then
'  Debug.Print "RedrawBar from Refresh"
  RedrawBar
 ElseIf BtnIndex >= 0 And BtnIndex < m_nButtons Then
  DrawButton BtnIndex
 End If
End Sub

'Moves or resize the bar.
Public Sub Move(Left As Long, Top As Long, Width As Long, Height As Long)
 'Exit if there's no bar.
 If m_Container Is Nothing Then Exit Sub
 m_Container.Move Left, Top, Width, Height
End Sub

'Asigns the general skin to the bar.
Public Function SetSkin(ByVal Img As StdPicture) As Boolean
 Dim Result As Long
 
 'Exit if there's no bar.
 If m_Container Is Nothing Then Exit Function
 
 With m_SkinDC
  If .DC = 0 Then
   SetSkin = False
  Else
   If Not (Img Is Nothing) Then
    Result = SelectObject(.DC, Img.handle)
    If .OldBmp = 0 Then .OldBmp = Result
    .CurBmp = Img.handle
    m_SkinWidth = m_Container.ScaleX(Img.Width)
    m_SkinHeight = m_Container.ScaleY(Img.Height)
    If CreateMaskImage(m_SkinDC, m_SkinMaskDC, m_SkinWidth, m_SkinHeight, m_MaskColor) Then
    Else
     SetSkin = False
     Exit Function
    End If
'    Debug.Print "RedrawBar from SetSkin"
    RedrawBar
   End If
   Set m_Skin = Img
   SetSkin = True
  End If
 End With
End Function

Public Function GetButtonBackColor(ByRef FirstColor As Long, ByRef SecondColor As Long, BtnState As eButtonStates) As Long
 FirstColor = m_BackColor1(BtnState)
 SecondColor = m_BackColor2(BtnState)
End Function

Public Function SetButtonBackColor(ByRef FirstColor As Long, ByRef SecondColor As Long, BtnState As eButtonStates) As Long
 Dim SysColor As Long
 Dim State As Long
 
 State = GetButtonStateSkin(BtnState)
 
 If (FirstColor And &H80000000) = &H80000000 Then
 'System Color
  SysColor = (FirstColor And &H7FFFFFFF)
  FirstColor = GetSysColor(SysColor)
 End If
 m_BackColor1(State) = FirstColor
 
 If (SecondColor And &H80000000) = &H80000000 Then
 'System Color
  SysColor = (SecondColor And &H7FFFFFFF)
  SecondColor = GetSysColor(SysColor)
 End If
 m_BackColor2(State) = SecondColor
' Debug.Print "RedrawBar from SetButtonBackColor"
 RedrawBar
End Function

Public Function GetButtonTextColor(ByRef FirstColor As Long, ByRef SecondColor As Long, BtnState As eButtonStates) As Long
 FirstColor = m_TextColor1(BtnState)
 SecondColor = m_TextColor2(BtnState)
End Function

Public Function SetButtonTextColor(ByRef FirstColor As Long, ByRef SecondColor As Long, BtnState As eButtonStates) As Long
 Dim SysColor As Long
 Dim State As Long
 
 State = GetButtonStateSkin(BtnState)
 If (FirstColor And &H80000000) = &H80000000 Then
 'System Color
  SysColor = (FirstColor And &H7FFFFFFF)
  FirstColor = GetSysColor(SysColor)
 End If
 m_TextColor1(State) = FirstColor
 
 If (SecondColor And &H80000000) = &H80000000 Then
 'System Color
  SysColor = (SecondColor And &H7FFFFFFF)
  SecondColor = GetSysColor(SysColor)
 End If
 m_TextColor2(State) = SecondColor
' Debug.Print "RedrawBar from SetButtonTextColor"
 RedrawBar
End Function

Public Function GetButtonBorderColor(ByRef Color As Long, BtnState As eButtonStates) As Long
 Color = m_BorderColor(BtnState)
 GetButtonBorderColor = m_BorderColor(BtnState)
End Function

Public Function SetButtonBorderColor(ByRef Color As Long, BtnState As eButtonStates) As Long
 Dim SysColor As Long
 Dim State As Long
 
 State = GetButtonStateSkin(BtnState)
 
 If (Color And &H80000000) = &H80000000 Then
 'System Color
  SysColor = (Color And &H7FFFFFFF)
  Color = GetSysColor(SysColor)
 End If
 m_BorderColor(State) = Color
' Debug.Print "RedrawBar from SetButtonBorderColor"
 RedrawBar
End Function

'Modifies button states.
Friend Function SetButtonState(ByVal Index As Long, State As eButtonStates, SetState As Boolean) As eButtonStates
 Dim i As Long
 
 If Index >= 0 And Index < m_nButtons Then
  'If an Option button is being set its eDown state to True,
  'search the current Option button from the group that
  'is currently pressed and set its eDown state to False.
  If (m_Buttons(Index).ButtonType = eOption) And (State = eDown) Then
   If SetState Then
    'Check to the left of the button
    For i = Index To 0 Step -1
     With m_Buttons(i)
      If .ButtonType = eOption Then
       'Modify the state if the Option button is Down
       If (.State And eDown) Then
        .State = .State And (Not eDown)
        DrawButton i
        GoTo SetCurState
       End If
      Else
       Exit For
      End If
     End With
    Next i
    'Check to the right of the button
    For i = Index To m_nButtons - 1
     With m_Buttons(i)
      If .ButtonType = eOption Then
       'Modify the state if the Option button is Down
       If (.State And eDown) Then
        .State = .State And (Not eDown)
        DrawButton i
        GoTo SetCurState
       End If
      Else
       Exit For
      End If
     End With
    Next i
   Else
    Exit Function
   End If
  End If

SetCurState:

  If SetState Then
   m_Buttons(Index).State = m_Buttons(Index).State Or State
  Else
   m_Buttons(Index).State = m_Buttons(Index).State And (Not State)
  End If
'  Debug.Print "SetButtonState (&H" & Hex(State) & ") = &H" & Hex(m_Buttons(Index).State)
  SetButtonState = m_Buttons(Index).State
 ElseIf Index = BUTTON_PREVIOUS Then
  If SetState Then
   m_ButtonPrevious.State = m_ButtonPrevious.State Or State
  Else
   m_ButtonPrevious.State = m_ButtonPrevious.State And (Not State)
  End If
'  Debug.Print "SetButtonState (&H" & Hex(State) & ") = &H" & Hex(m_ButtonPrevious.State)
  SetButtonState = m_ButtonPrevious.State
 ElseIf Index = BUTTON_NEXT Then
  If SetState Then
   m_ButtonNext.State = m_ButtonNext.State Or State
  Else
   m_ButtonNext.State = m_ButtonNext.State And (Not State)
  End If
'  Debug.Print "SetButtonState (&H" & Hex(State) & ") = &H" & Hex(m_ButtonNext.State)
  SetButtonState = m_ButtonNext.State
 Else
  SetButtonState = -1
  Exit Function
 End If
End Function

'Set button size.
'Returns whether button is visible or not.
Public Function SetButtonSize(ByVal Index As Long, Optional ByVal W As Long = 0, Optional ByVal H As Long = 0) As Boolean
 If Index >= 0 And Index < m_nButtons Then
  If W > 0 Then m_Buttons(Index).Width = W
  If H > 0 Then m_Buttons(Index).Height = H
  UpdateButtonPos Index, m_nButtons - 1, True
  CheckVisibleButtons
  SetButtonSize = (m_Buttons(Index).Flags And BTN_FULLVISIBLE)
'  Debug.Print "RedrawBar en SetButtonSize"
'  RedrawBar
 End If
End Function

'Sets a particular skin for a given button.
Public Function SetButtonSkin(ByVal Index As Long, ByVal BtnSkin As StdPicture, Optional ByVal SkinMaskColor As Long = -1) As Boolean
 If Index >= 0 And Index < m_nButtons Then
  Set m_Buttons(Index).Skin = BtnSkin
  If SkinMaskColor = -1 Then SkinMaskColor = m_MaskColor
  m_Buttons(Index).MaskColor = SkinMaskColor
  DrawButton Index
 End If
End Function

Public Sub SetFocus(Optional ByVal Index As Long = NO_BUTTON)
 m_Container.SetFocus
 If Index > NO_BUTTON Then
  Me.SetButtonState m_ButtonFocus, eFocused, False
  DrawButton m_ButtonFocus
  m_ButtonFocus = Index
  Me.SetButtonState m_ButtonFocus, eFocused, True
 End If
 DrawButton m_ButtonFocus
End Sub

'=============================================
'Private Events Implementation
'=============================================

Private Sub Class_Initialize()
 m_BarFlags = fMASKCOLOR Or fDRAWFOCUSRECT
 m_FillMode = eSolid
 m_BorderWidth = 1
 m_ButtonOver = NO_BUTTON
 m_ButtonDown = NO_BUTTON
 m_ButtonFocus = NO_BUTTON
 m_DefaultButtonWidth = 80
 m_DefaultButtonHeight = 40
 m_VerticalButtonGap = 3
 m_HorizontalButtonGap = 3
 m_MaskColor = -1
End Sub

Private Sub Class_Terminate()
 Me.Destroy
End Sub

Private Sub m_Container_GotFocus()
' Debug.Print "m_ButtonFocus in Got_Focus = " & m_ButtonFocus & " Has the Focus? = " & CBool(Me.ButtonState(m_ButtonFocus) And eFocused)
 If (Me.ButtonState(m_ButtonFocus) And eDisabled) Then Exit Sub
 If m_ButtonFocus >= 0 And m_ButtonFocus < m_nButtons Then
  'Draw Focus Rect
  DrawButton m_ButtonFocus
'  RaiseEvent GotFocus(m_ButtonFocus)
 End If
End Sub

Private Sub m_Container_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim NextButton As Long
'THERE'S A PROBLEM WHEN A COMMAND BUTTON HAS ITS DEFAULT
'PROPERTY SET TO TRUE, THE PICTUREBOX DOES NOT RECEIVE
'THE KEYDOWN EVENT WITH THE VBKEYRETURN KEYCODE.
' Debug.Print "Enter KeyDown"
' Debug.Print "m_ButtonFocus = " & m_ButtonFocus
 If KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then
  If (Me.ButtonState(m_ButtonFocus) And eDisabled) Then Exit Sub
   'If a button is already pressed, do nothing.
  If m_ButtonDown > NO_BUTTON Then Exit Sub
  If m_ButtonFocus > NO_BUTTON Then
   m_ButtonDown = m_ButtonFocus
   If (Me.ButtonType(m_ButtonDown) And eCheck) <> eCheck Then
    Me.SetButtonState m_ButtonDown, eDown, True
   Else
    'Invert the state of Check button
    Me.SetButtonState m_ButtonDown, eDown, Not CBool(Me.ButtonState(m_ButtonDown) And eDown)
   End If
   DrawButton m_ButtonDown
   'Keyboard event.
   RaiseEvent KeyDown(m_ButtonDown, KeyCode, Shift)
  End If
 ElseIf KeyCode = vbKeyLeft Then
  If m_ButtonFocus < 0 Then Exit Sub
  NextButton = NextAvailableButton(m_ButtonFocus - 1, 0)
  If NextButton = NO_BUTTON Then Exit Sub
  'Remove the focus from the current button.
  Me.SetButtonState m_ButtonFocus, eFocused, False
  DrawButton m_ButtonFocus
  RaiseEvent KeyDown(m_ButtonFocus, KeyCode, Shift)
  RaiseEvent LostFocus(m_ButtonFocus, "KeyDown")
  m_ButtonFocus = NextButton
  Me.SetButtonState m_ButtonFocus, eFocused, True
  DrawButton m_ButtonFocus
  RaiseEvent GotFocus(m_ButtonFocus)
 ElseIf KeyCode = vbKeyRight Then
  If m_ButtonFocus < 0 Then Exit Sub
  NextButton = NextAvailableButton(m_ButtonFocus + 1, m_nButtons - 1)
  If NextButton = NO_BUTTON Then Exit Sub
  'Remove the focus from the current button.
  Me.SetButtonState m_ButtonFocus, eFocused, False
  DrawButton m_ButtonFocus
  RaiseEvent KeyDown(m_ButtonFocus, KeyCode, Shift)
  RaiseEvent LostFocus(m_ButtonFocus, "KeyDown")
  m_ButtonFocus = NextButton
  Me.SetButtonState m_ButtonFocus, eFocused, True
  DrawButton m_ButtonFocus
  RaiseEvent GotFocus(m_ButtonFocus)
 End If

End Sub

Private Sub m_Container_KeyPress(KeyAscii As Integer)
 If m_ButtonFocus >= 0 Then
  RaiseEvent KeyPressed(m_ButtonFocus, KeyAscii)
 End If
End Sub

Private Sub m_Container_KeyUp(KeyCode As Integer, Shift As Integer)
 If (Me.ButtonState(m_ButtonFocus) And eDisabled) Then Exit Sub
 If KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then
  If m_ButtonDown > NO_BUTTON Then
   If (Me.ButtonType(m_ButtonDown) And eCheck) <> eCheck Then
    Me.SetButtonState m_ButtonDown, eDown, False
   End If
   m_ButtonFocus = m_ButtonDown
   DrawButton m_ButtonDown
   RaiseEvent Click(m_ButtonDown)
   'Add the functionality of navigation buttons.
   If m_ButtonDown = BUTTON_PREVIOUS Then
    m_FirstVisibleButton = m_FirstVisibleButton - 1
    m_Buttons(m_FirstVisibleButton).Left = m_HorizontalButtonGap 'm_Buttons(m_FirstVisibleButton).Left + m_Buttons(m_FirstVisibleButton).Width
    UpdateButtonPos m_FirstVisibleButton + 1, m_nButtons - 1, True
    UpdateButtonPos m_FirstVisibleButton, 0, False
    CheckVisibleButtons
   ElseIf m_ButtonDown = BUTTON_NEXT Then
    m_FirstVisibleButton = m_FirstVisibleButton + 1
    m_Buttons(m_FirstVisibleButton).Left = m_HorizontalButtonGap
    UpdateButtonPos m_FirstVisibleButton - 1, 0, False
    UpdateButtonPos m_FirstVisibleButton, m_nButtons - 1, True
    CheckVisibleButtons
   End If
  'Keyboard event.
   RaiseEvent KeyUp(m_ButtonDown, KeyCode, Shift)
   m_ButtonDown = NO_BUTTON
  End If
 End If

End Sub

Private Sub m_Container_LostFocus()
 Dim NextBtnFocus As Long
 Dim i As Long, Result As Long
 
 If IsKeyDown(vbKeyTab) Then
  If m_ButtonFocus < 0 Then Exit Sub
  NextBtnFocus = NextAvailableButton(m_ButtonFocus + 1, m_nButtons - 1)
  If NextBtnFocus <> NO_BUTTON Then
   'Remove Focus Rect from current button.
   Me.SetButtonState m_ButtonFocus, eFocused, False
   DrawButton m_ButtonFocus
   RaiseEvent LostFocus(m_ButtonFocus, "LostFocus")
'   Debug.Print "m_ButtonFocus in Lost_Focus = " & m_ButtonFocus & " Has the Focus? = " & CBool(Me.ButtonState(m_ButtonFocus) And eFocused)
   m_ButtonFocus = NextBtnFocus
   Me.SetButtonState m_ButtonFocus, eFocused, True
   m_Container.SetFocus
   RaiseEvent GotFocus(m_ButtonFocus)
  Else
   'Si m_Container does not have the focus, draw the button
   'with no focus.
   DrawButton m_ButtonFocus
'   RaiseEvent LostFocus(m_ButtonFocus, "LostFocus")
  End If
 End If
End Sub

Private Sub m_Container_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim TempBtn As Long
 
' Debug.Print "mOUSE DOWN"
 TempBtn = GetButtonCapture(CLng(X), CLng(Y))
 If Button = vbLeftButton Then
  If TempBtn > NO_BUTTON Then
'    If TempBtn = BUTTON_PREVIOUS Then
'     Debug.Print "m_ButtonDown = BUTTON_PREVIOUS is disabled? " & CBool(Me.ButtonState(TempBtn) And eDisabled)
'    ElseIf TempBtn = BUTTON_NEXT Then
'     Debug.Print "m_ButtonDown = BUTTON_NEXT is disabled? " & CBool(Me.ButtonState(TempBtn) And eDisabled)
'    End If
   
   If (Me.ButtonState(TempBtn) And eDisabled) Then Exit Sub
    m_ButtonDown = TempBtn
'   Debug.Print "Button Pressed = " & m_ButtonDown
   If (Me.ButtonType(m_ButtonDown) And eCheck) <> eCheck Then
    Me.SetButtonState m_ButtonDown, eDown, True
   Else
    'Invert state of check button
'    Debug.Print "Before changing, Check Button State is Down? = " & CBool(Me.ButtonState(m_ButtonDown) And eDown)
    Me.SetButtonState m_ButtonDown, eDown, Not CBool(Me.ButtonState(m_ButtonDown) And eDown)
'    Debug.Print "After changing, Check Button State is Down? = " & CBool(Me.ButtonState(m_ButtonDown) And eDown)
   End If
    'Remove the focus of the current button.
    If m_ButtonFocus > NO_BUTTON Then
     Me.SetButtonState m_ButtonFocus, eFocused, False
     DrawButton m_ButtonFocus
     RaiseEvent LostFocus(m_ButtonFocus, "MouseDown")
    End If
    'Set the focus into the new button.
    Me.SetButtonState m_ButtonDown, eFocused, True
    m_ButtonFocus = m_ButtonDown
'    Debug.Print "m_ButtonFocus en Mouse_Down = " & m_ButtonFocus
    DrawButton m_ButtonDown
    RaiseEvent GotFocus(m_ButtonFocus)
'   Debug.Print "MouseDown in x = " & x & ", y = " & y
   'Mouse Left button event.
   RaiseEvent MouseDown(m_ButtonDown, Button, Shift, X, Y)
  End If
 Else
  If (Me.ButtonState(TempBtn) And eDisabled) Then Exit Sub
  'Mouse Right or Middle button event.
  RaiseEvent MouseDown(m_ButtonDown, Button, Shift, X, Y)
 End If
End Sub

Private Sub m_Container_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim ButtonIndex As Long
 
 'Obtain button Index that has the Mouse capture.
  ButtonIndex = GetButtonCapture(CLng(X), CLng(Y))
 'Check for Mouse Over
 If m_ButtonOver = ButtonIndex Then
  Exit Sub
 End If

 If (Me.ButtonState(ButtonIndex) And eDisabled) Then Exit Sub

 'If Mouse button is pressed
 If Button = vbLeftButton Then
  'If there's a bar button currently pressed
  If m_ButtonDown > NO_BUTTON Then
   'If pressed button is different from the one with
   'mouse capture, remove the Down State.
   If m_ButtonDown <> ButtonIndex Then
    'm_ButtonOver <> ButtonIndex? Yes.
    Me.SetButtonState m_ButtonDown, eDown, False
    m_ButtonDown = NO_BUTTON
   Else
   'If the button with the mouse capture is the button
   'with the focus, set the Down State.
    If m_ButtonFocus = ButtonIndex Then
     Me.SetButtonState m_ButtonDown, eDown, True
     m_ButtonDown = m_ButtonFocus
    End If
   End If
  End If
 End If
 
 'If there's a button with the mouse capture, removes it.
 If m_ButtonOver > NO_BUTTON Then
  Me.SetButtonState m_ButtonOver, eOver, False
  DrawButton m_ButtonOver
  'This events are only raised for buttons that the
  'user has created.
  If m_ButtonOver >= 0 Then
'   Debug.Print "MouseLeave " & m_ButtonOver & " in x = " & x & ", y = " & y
   RaiseEvent MouseLeave(m_ButtonOver)
  End If
 End If
 'If a button got the mouse capture, adds it.
 If ButtonIndex > NO_BUTTON Then
  Me.SetButtonState ButtonIndex, eOver, True
  DrawButton ButtonIndex
  'In MouseEnter show the tooltip.
  If ButtonIndex >= 0 Then
   m_Container.ToolTipText = m_Buttons(ButtonIndex).Tooltip
  End If
'  Debug.Print "MouseOver Boton " & ButtonIndex
  'This events are only raised for buttons that the
  'user has created.
  If ButtonIndex >= 0 Then
   RaiseEvent MouseEnter(ButtonIndex)
'  Debug.Print "MouseEnter " & ButtonIndex & " in x = " & x & ", y = " & y
  End If
 Else
  'If no button has the mouse capture, remove the
  'picturebox tooltip.
  m_Container.ToolTipText = ""
 End If
 m_ButtonOver = ButtonIndex
End Sub

Private Sub m_Container_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim ButtonIndex As Long
 
 ButtonIndex = GetButtonCapture(CLng(X), CLng(Y))
 If Button = vbLeftButton Then
 'If is a user button.
  If ButtonIndex > NO_BUTTON Then
   If (Me.ButtonState(ButtonIndex) And eDisabled) Then Exit Sub
   'Mouse_Down Event button = Mouse_Up Event Button?
   If m_ButtonDown = ButtonIndex Then
    If (Me.ButtonType(m_ButtonDown) And eCheck) <> eCheck Then
     Me.SetButtonState m_ButtonDown, eDown, False
    End If
    DrawButton m_ButtonDown
    RaiseEvent Click(ButtonIndex)
   End If
   'Add navigation buttons functionality.
   If m_ButtonDown = BUTTON_PREVIOUS Then
    m_FirstVisibleButton = m_FirstVisibleButton - 1
    m_Buttons(m_FirstVisibleButton).Left = m_HorizontalButtonGap 'm_Buttons(m_FirstVisibleButton).Left + m_Buttons(m_FirstVisibleButton).Width
    UpdateButtonPos m_FirstVisibleButton + 1, m_nButtons - 1, True
    UpdateButtonPos m_FirstVisibleButton, 0, False
    CheckVisibleButtons
   ElseIf m_ButtonDown = BUTTON_NEXT Then
'    m_Buttons(m_FirstVisibleButton).Left = m_Buttons(m_FirstVisibleButton).Left - m_Buttons(m_FirstVisibleButton).Width
    m_FirstVisibleButton = m_FirstVisibleButton + 1
    m_Buttons(m_FirstVisibleButton).Left = m_HorizontalButtonGap
    UpdateButtonPos m_FirstVisibleButton - 1, 0, False
    UpdateButtonPos m_FirstVisibleButton, m_nButtons - 1, True
    CheckVisibleButtons
   End If
   'Mouse Left button event.
   RaiseEvent MouseUp(ButtonIndex, Button, Shift, X, Y)
   m_ButtonDown = NO_BUTTON
  End If
 Else
  If (Me.ButtonState(ButtonIndex) And eDisabled) Then Exit Sub
  'Mouse Right or middle button event.
  RaiseEvent MouseUp(ButtonIndex, Button, Shift, X, Y)
 End If
End Sub

Private Sub m_Container_Paint()
' Debug.Print "m_Container_Paint event"
' Debug.Print "RedrawBar from Paint"
 RedrawBar
 RaiseEvent Paint
End Sub

Private Sub m_Container_Resize()
 ResizeDC m_BackBuffer, m_Container.ScaleWidth, m_Container.ScaleHeight
 ResizeDC m_ButtonSkinDC, m_Container.ScaleWidth, m_Container.ScaleHeight
'Set the coords of navigation buttons.
 m_ButtonNext.Left = m_Container.ScaleWidth - m_HorizontalButtonGap - m_ButtonNext.Width
 m_ButtonNext.Top = m_VerticalButtonGap
 m_ButtonNext.Height = m_Container.ScaleHeight - (m_VerticalButtonGap * 2)
 
 m_ButtonPrevious.Left = m_ButtonNext.Left - m_HorizontalButtonGap - m_ButtonNext.Width
 m_ButtonPrevious.Top = m_VerticalButtonGap
 m_ButtonPrevious.Height = m_Container.ScaleHeight - (m_VerticalButtonGap * 2)

' RedrawBar
 RaiseEvent Resize
End Sub

Private Sub m_Font_FontChanged(ByVal PropertyName As String)
 Set m_Container.Font = m_Font
' Debug.Print "RedrawBar from FontChanged"
 RedrawBar
End Sub

'=============================================
'Public Properties implementation
'=============================================

'Update the bar?
Public Property Get LockUpdate() As Boolean
 LockUpdate = m_LockUpdate
End Property

'Update the bar?
Public Property Let LockUpdate(ByVal NuevoValor As Boolean)
 m_LockUpdate = NuevoValor
 If Not NuevoValor Then
'  Debug.Print "RedrawBar from LockUpdate"
  RedrawBar
 End If
End Property

'Space between buttons (in pixels).
Public Property Get HorizontalButtonGap() As Long
 HorizontalButtonGap = m_HorizontalButtonGap
End Property

'Space between buttons (in pixels).
Public Property Let HorizontalButtonGap(ByVal NuevoValor As Long)
 m_HorizontalButtonGap = NuevoValor
' Debug.Print "RedrawBar from HorizontalButtonGap"
 RedrawBar
End Property

'Space between buttons and picturebox (in pixels).
Public Property Get VerticalButtonGap() As Long
 VerticalButtonGap = m_VerticalButtonGap
End Property

'Space between buttons and picturebox (in pixels).
Public Property Let VerticalButtonGap(ByVal NuevoValor As Long)
 m_VerticalButtonGap = NuevoValor
' Debug.Print "RedrawBar from VerticalButtonGap"
 RedrawBar
End Property

'Button Background filling mode.
Public Property Get FillMode() As eFillMode
 FillMode = m_FillMode
End Property

'Button Background filling mode.
Public Property Let FillMode(ByVal NuevoValor As eFillMode)
 m_FillMode = NuevoValor
' Debug.Print "RedrawBar from FillMode"
 RedrawBar
End Property

'Bar background color.
Public Property Get BackColor() As OLE_COLOR
 BackColor = m_Container.BackColor
End Property

'Bar background color.
Public Property Let BackColor(ByVal NuevoValor As OLE_COLOR)
 Dim SysColor As Long
 
 If (NuevoValor And &H80000000) = &H80000000 Then
 'System color
  SysColor = (NuevoValor And &H7FFFFFFF)
  NuevoValor = GetSysColor(SysColor)
 End If
 m_BackColor = NuevoValor
 
End Property

'Button Border Width.
Public Property Get BorderWidth() As Long
 BorderWidth = m_BorderWidth
End Property

'Button Border Width.
Public Property Let BorderWidth(ByVal NuevoValor As Long)
 m_BorderWidth = NuevoValor
' Debug.Print "RedrawBar from BorderWidth"
 RedrawBar
End Property

'Left coord. of the bar.
Public Property Get Left() As Long
 Left = m_Container.Left
End Property

'Left coord. of the bar.
Public Property Let Left(ByVal NuevoValor As Long)
 m_Container.Left = NuevoValor
End Property

'Top coord. of the bar.
Public Property Get Top() As Long
 Top = m_Container.Top
End Property

'Top coord. of the bar.
Public Property Let Top(ByVal NuevoValor As Long)
 m_Container.Top = NuevoValor
End Property

'Width of the bar.
Public Property Get Width() As Long
 Width = m_Container.Width
End Property

'Width of the bar.
Public Property Let Width(ByVal NuevoValor As Long)
 m_Container.Width = NuevoValor
End Property

'Height of the bar.
Public Property Get Height() As Long
 Height = m_Container.Height
End Property

'Height of the bar.
Public Property Let Height(ByVal NuevoValor As Long)
 m_Container.Height = NuevoValor
End Property

'Window handle of the bar (picturebox window handle).
Public Property Get hwnd() As Long
 hwnd = m_Container.hwnd
End Property

'DC handle of the bar (picturebox DC handle).
Public Property Get hdc() As Long
 hdc = m_Container.hdc
End Property

'Skin and icons transparent color.
Public Property Get MaskColor() As Long
 MaskColor = m_MaskColor
End Property

'Skin and icons transparent color.
Public Property Let MaskColor(ByVal NuevoValor As Long)
 m_MaskColor = NuevoValor
 If Me.UseMaskColor Then
'  Debug.Print "RedrawBar from MaskColor"
  RedrawBar
 End If
End Property

'Transparent Color is used?.
Public Property Get UseMaskColor() As Boolean
 UseMaskColor = (m_BarFlags And fMASKCOLOR) <> 0
End Property

'Transparent Color is used?.
Public Property Let UseMaskColor(ByVal NuevoValor As Boolean)
 If NuevoValor Then
  m_BarFlags = m_BarFlags Or fMASKCOLOR
 Else
  m_BarFlags = m_BarFlags And (Not fMASKCOLOR)
 End If
' Debug.Print "RedrawBar from UseMaskColor"
 RedrawBar
End Property

'Indicates that the size of button is calculated according
'to the text and icon size.
Public Property Get AutoSizeButtons() As Boolean
 AutoSizeButtons = (m_BarFlags And fAUTOSIZE) <> 0
End Property

'Indicates that the size of button is calculated according
'to the text and icon size.
Public Property Let AutoSizeButtons(ByVal NuevoValor As Boolean)
 If NuevoValor Then
  m_BarFlags = m_BarFlags Or fAUTOSIZE
 Else
  m_BarFlags = m_BarFlags And (Not fAUTOSIZE)
 End If
' Debug.Print "RedrawBar from AutosizeButtons"
 RedrawBar
End Property

'Draw Focus rect?
Public Property Get DrawFocusRect() As Boolean
 DrawFocusRect = (m_BarFlags And fDRAWFOCUSRECT) <> 0
End Property

'Draw Focus rect?
Public Property Let DrawFocusRect(ByVal NuevoValor As Boolean)
 If NuevoValor Then
  m_BarFlags = m_BarFlags Or fDRAWFOCUSRECT
 Else
  m_BarFlags = m_BarFlags And (Not fDRAWFOCUSRECT)
 End If
' Debug.Print "RedrawBar from DrawFocusRect"
 RedrawBar
End Property

'Number of buttons in the bar.
Public Property Get nButtons() As Long
 nButtons = m_nButtons
End Property

'Text of given button.
Public Property Get ButtonText(ByVal Index As Long) As String
 If Index >= 0 And Index < m_nButtons Then
  ButtonText = m_Buttons(Index).Text
 End If
End Property

'Text of given button.
Public Property Let ButtonText(ByVal Index As Long, ByVal NuevoValor As String)
 If Index >= 0 And Index < m_nButtons Then
  m_Buttons(Index).Text = NuevoValor
  DrawButton Index
 End If
End Property

'Tooltip of given button.
Public Property Get ButtonTooltip(ByVal Index As Long) As String
 If Index >= 0 And Index < m_nButtons Then
  ButtonTooltip = m_Buttons(Index).Tooltip
 End If
End Property

'Tooltip of given button.
Public Property Let ButtonTooltip(ByVal Index As Long, ByVal NuevoValor As String)
 If Index >= 0 And Index < m_nButtons Then
   m_Buttons(Index).Tooltip = NuevoValor
   DrawButton Index
 End If
End Property

'State of given button.
Public Property Get ButtonState(ByVal Index As Long) As eButtonStates
 If Index >= 0 And Index < m_nButtons Then
  ButtonState = m_Buttons(Index).State
 ElseIf Index = BUTTON_PREVIOUS Then
  ButtonState = m_ButtonPrevious.State
 ElseIf Index = BUTTON_NEXT Then
  ButtonState = m_ButtonNext.State
 End If
End Property

'State of given button.
'Public Property Let ButtonState(ByVal Index As Long, ByVal NuevoValor As eButtonStates)
' If Index >= 0 And Index < m_nButtons Then
'  m_Buttons(Index).State = NuevoValor
'  DrawButton Index
' ElseIf Index = BUTTON_PREVIOUS Then
'  m_ButtonPrevious.State = NuevoValor
'  DrawButton Index
' ElseIf Index = BUTTON_NEXT Then
'  m_ButtonNext.State = NuevoValor
'  DrawButton Index
' End If
'End Property

'Style of given button.
Public Property Get ButtonStyle(ByVal Index As Long) As eButtonStyles
 If Index >= 0 And Index < m_nButtons Then
  ButtonStyle = m_Buttons(Index).Style
 ElseIf Index = BUTTON_PREVIOUS Then
  ButtonStyle = m_ButtonPrevious.Style
 ElseIf Index = BUTTON_NEXT Then
  ButtonStyle = m_ButtonNext.Style
 End If
End Property

'Style of given button.
Public Property Let ButtonStyle(ByVal Index As Long, ByVal NuevoValor As eButtonStyles)
 If Index >= 0 And Index < m_nButtons Then
  m_Buttons(Index).Style = NuevoValor
  DrawButton Index
 ElseIf Index = BUTTON_NEXT Or Index = BUTTON_PREVIOUS Then
  m_ButtonPrevious.Style = NuevoValor
  m_ButtonNext.Style = NuevoValor
  DrawButton BUTTON_PREVIOUS
  DrawButton BUTTON_NEXT
 End If
End Property

'Type of given button.
Public Property Get ButtonType(ByVal Index As Long) As eButtonTypes
 If Index >= 0 And Index < m_nButtons Then
  ButtonType = m_Buttons(Index).ButtonType
 End If
End Property

'Type of given button.
Public Property Let ButtonType(ByVal Index As Long, ByVal NuevoValor As eButtonTypes)
 If Index >= 0 And Index < m_nButtons Then
  m_Buttons(Index).ButtonType = NuevoValor
  DrawButton Index
 End If
End Property

'Icon of given button.
Public Property Get ButtonIcon(ByVal Index As Long) As StdPicture
 If Index >= 0 And Index < m_nButtons Then
  Set ButtonIcon = m_Buttons(Index).Icon
 End If
End Property

'Icon of given button.
Public Property Set ButtonIcon(ByVal Index As Long, ByVal NuevoValor As StdPicture)
 If Index >= 0 And Index < m_nButtons Then
  Set m_Buttons(Index).Icon = NuevoValor
  DrawButton Index
 End If
End Property

'Text alignment of given button.
Public Property Get ButtonTextAlignment(ByVal Index As Long) As eAlignment
 If Index >= 0 And Index < m_nButtons Then
  ButtonTextAlignment = m_Buttons(Index).TextAlign
 End If
End Property

'Text alignment of given button.
Public Property Let ButtonTextAlignment(ByVal Index As Long, ByVal NuevoValor As eAlignment)
 If Index >= 0 And Index < m_nButtons Then
  m_Buttons(Index).TextAlign = NuevoValor
  DrawButton Index
 End If
End Property

'Icon alignment of given button.
Public Property Get ButtonIconAlignment(ByVal Index As Long) As eAlignment
 If Index >= 0 And Index < m_nButtons Then
  ButtonIconAlignment = m_Buttons(Index).IconAlign
 End If
End Property

'Icon alignment of given button.
Public Property Let ButtonIconAlignment(ByVal Index As Long, ByVal NuevoValor As eAlignment)
 If Index >= 0 And Index < m_nButtons Then
  m_Buttons(Index).IconAlign = NuevoValor
 End If
End Property

'Width of given button (READ ONLY).
Public Property Get ButtonWidth(ByVal Index As Long) As Long
 If Index >= 0 And Index < m_nButtons Then
  ButtonWidth = m_Buttons(Index).Width
 End If
End Property

'Width of given button.
'Public Property Let ButtonWidth(ByVal Index As Long, ByVal NuevoValor As Long)
' If Index >= 0 And Index < m_nButtons Then
'  m_Buttons(Index).Width = NuevoValor
'  UpdateButtonPos Index, m_nButtons
'  DrawButton Index
' End If
'End Property

'Height of given button (READ ONLY).
Public Property Get ButtonHeight(ByVal Index As Long) As Long
 If Index >= 0 And Index < m_nButtons Then
  ButtonHeight = m_Buttons(Index).Height
 End If
End Property

'Height of given button.
'Public Property Let ButtonHeight(ByVal Index As Long, ByVal NuevoValor As Long)
' If Index >= 0 And Index < m_nButtons Then
'  m_Buttons(Index).Height = NuevoValor
'  DrawButton Index
' End If
'End Property

''Indica la skin particular del boton dado.
'Public Property Get ButtonSkin(ByVal Index As Long) As StdPicture
' If Index >= 0 And Index < m_nButtons Then
'  Set ButtonSkin = m_Buttons(Index).Skin
'  DrawButton Index
' End If
'End Property

'Default Width of buttons.
Public Property Get DefaultButtonWidth() As Long
 DefaultButtonWidth = m_DefaultButtonWidth
End Property

'Default Width of buttons.
Public Property Let DefaultButtonWidth(ByVal NuevoValor As Long)
 m_DefaultButtonWidth = NuevoValor
' Debug.Print "RedrawBar from DefaultButtonWidth"
' RedrawBar
End Property

'Default Height of buttons.
Public Property Get DefaultButtonHeight() As Long
 DefaultButtonHeight = m_DefaultButtonHeight
End Property

'Default Height of buttons.
Public Property Let DefaultButtonHeight(ByVal NuevoValor As Long)
 m_DefaultButtonHeight = NuevoValor
' Debug.Print "RedrawBar from DefaultButtonHeight"
' RedrawBar
End Property

'Bar enabled?.
Public Property Get Enabled() As Boolean
 Enabled = m_Container.Enabled
End Property

'Bar enabled?.
Public Property Let Enabled(ByVal NuevoValor As Boolean)
 m_Container.Enabled = NuevoValor
End Property

'Bar visible?.
Public Property Get Visible() As Boolean
 Visible = m_Container.Visible
End Property

'Bar visible?.
Public Property Let Visible(ByVal NuevoValor As Boolean)
 m_Container.Visible = NuevoValor
End Property

'Global button font (picturebox font).
Public Property Get Font() As StdFont
 Font = m_Font
End Property

'Global button font (picturebox font).
Public Property Set Font(ByVal NuevoValor As StdFont)
 Set m_Font = NuevoValor
 Set m_Container.Font = NuevoValor
' Debug.Print "RedrawBar from Font"
 RedrawBar
End Property

'Width of skin borders.
Public Property Get SkinBorderWidth(ByVal BtnState As Long) As Long
 SkinBorderWidth = m_SkinBorderWidth(BtnState)
End Property

'Width of skin borders.
Public Property Let SkinBorderWidth(ByVal BtnState As Long, ByVal NuevoValor As Long)
 If NuevoValor >= 0 Then
  m_SkinBorderWidth(BtnState) = NuevoValor
'  Debug.Print "RedrawBar from SkinBorderWidth"
  RedrawBar
 End If
End Property

'Height of skin borders.
Public Property Get SkinBorderHeight(ByVal BtnState As Long) As Long
 SkinBorderHeight = m_SkinBorderHeight(BtnState)
End Property

'Height of skin borders.
Public Property Let SkinBorderHeight(ByVal BtnState As Long, ByVal NuevoValor As Long)
 If NuevoValor >= 0 Then
  m_SkinBorderHeight(BtnState) = NuevoValor
'  Debug.Print "RedrawBar from SkinBorderWidth"
  RedrawBar
 End If
End Property

'=============================================
'Private Methods Implementation
'=============================================

'Redraw the whole bar.
Private Sub RedrawBar()
 Dim i As Long
 Dim BackBrush As Long, rc As RECT
' Static Cont As Long
 
 If m_nButtons <= 0 Then Exit Sub
 If m_LockUpdate Then Exit Sub
' m_Container.Cls
 BackBrush = CreateSolidBrush(GetBkColor(m_Container.hdc))
 SetRect rc, 0, 0, m_Container.ScaleWidth, m_Container.ScaleHeight
 FillRect m_BackBuffer.DC, rc, BackBrush
 DeleteObject BackBrush
 m_RedrawAllButtons = True
' Debug.Print Cont & " Entrada de RedrawBar"
 For i = 0 To m_nButtons
  DrawButton i
 Next i
 DrawButton BUTTON_PREVIOUS
 DrawButton BUTTON_NEXT
 m_RedrawAllButtons = False
' Debug.Print Cont & " Salida de RedrawBar"
' Debug.Print
' Cont = Cont + 1
 BitBlt m_DC.DC, 0, 0, m_Container.ScaleWidth, m_Container.ScaleHeight, m_BackBuffer.DC, 0, 0, vbSrcCopy
End Sub

'Draw the entire given button.
Private Function DrawButton(Index As Long) As Boolean
 Dim ThisButton As tButton
 Dim rcButton As RECT, rcClient As RECT
 Dim Brush1 As Long, Result As Boolean
 Dim pts(4) As POINTAPI
 
 If Index = BUTTON_PREVIOUS Then
  ThisButton = m_ButtonPrevious
 ElseIf Index = BUTTON_NEXT Then
  ThisButton = m_ButtonNext
 ElseIf Index >= 0 And Index < m_nButtons Then
  ThisButton = m_Buttons(Index)
 Else
  Exit Function
 End If
 
 With ThisButton
  'If button is visible
  Result = GetButtonRect(Index, rcButton)
  If Result Then
   If Not m_RedrawAllButtons Then
    'Erase button background
    Brush1 = CreateSolidBrush(m_BackColor)
    FillRect m_BackBuffer.DC, rcButton, Brush1
    DeleteObject Brush1
   End If
   If .ButtonType = eSeparator Then
    'm_ButtonSkinDC.HasBmp = True always, so .CurBmp contains
    'the current bitmap (.HasBmp), therefore we just have
    'to select the button skin bitmap, blit it, and select
    'the .CurBmp bitmap again to the DC.
    If Not (.Skin Is Nothing) Then
     SelectObject m_ButtonSkinDC.DC, .Skin.handle
     BlitPicture m_BackBuffer, m_ButtonSkinDC, .Skin, rcButton, eVCenter, 0, 0, m_MaskColor
     SelectObject m_ButtonSkinDC.DC, m_ButtonSkinDC.CurBmp
    End If
   Else
    GetButtonClientRect Index, rcButton, rcClient
    DrawButtonFrame ThisButton, rcButton, rcClient
    DrawButtonIcon ThisButton, rcButton, rcClient
    DrawButtonText ThisButton, rcButton, rcClient
    DrawFocus Index
   End If
   
   If Not m_RedrawAllButtons Then
'    Debug.Print "Not RedrawAllButtons"
    BitBlt m_DC.DC, .Left, .Top, .Width, .Height, m_BackBuffer.DC, .Left, .Top, vbSrcCopy
'   Else
'    Debug.Print "RedrawAllButtons"
   End If
  End If
 End With
 DrawButton = Result
End Function

'Update the coords of given range of buttons.
Private Sub UpdateButtonPos(ByVal FromIndex As Long, ByVal ToIndex As Long, ByVal MoveRight As Boolean)
 Dim i As Long
 Dim dx As Long
 
 'If FromIndex isn't valid, exit.
 If FromIndex < 0 Or FromIndex >= m_nButtons Then Exit Sub
 'If ToIndex isn't valid, exit.
 If ToIndex < 0 Or ToIndex >= m_nButtons Then Exit Sub
 
 'Update from left to right.
 If MoveRight Then
  If FromIndex > 0 Then
   dx = m_Buttons(FromIndex - 1).Left + m_Buttons(FromIndex - 1).Width
  End If
  dx = dx + m_HorizontalButtonGap
'  Debug.Print "From " & FromIndex & " to " & ToIndex
  For i = FromIndex To ToIndex
   m_Buttons(i).Left = dx
'   Debug.Print m_Buttons(i).Text & ".Left = " & dx
   dx = dx + m_Buttons(i).Width + m_HorizontalButtonGap
  Next i
'  Debug.Print
 Else
  'Update from right to left.
  If FromIndex < m_nButtons - 1 Then
   dx = m_Buttons(FromIndex + 1).Left
  End If
  dx = dx - m_HorizontalButtonGap
'  Debug.Print "From " & FromIndex & " to " & ToIndex
  For i = FromIndex To ToIndex Step -1
   dx = dx - m_Buttons(i).Width
   m_Buttons(i).Left = dx
'   Debug.Print m_Buttons(i).Text & ".Left = " & dx
   dx = dx - m_HorizontalButtonGap
  Next i
'  Debug.Print
 End If
End Sub

'Returns the index of the button with the mouse capture,
'if any, else NO_BUTTON
Private Function GetButtonCapture(X As Long, Y As Long) As Long
 Dim rc As RECT
 Dim VisibleFlag As Boolean
 Dim i As Long
 
 'Check if mouse over picturebox
 If Not IsOverhWnd(m_Container.hwnd, CSng(X), CSng(Y)) Then
  GetButtonCapture = NO_BUTTON
  Exit Function
 End If
  
 VisibleFlag = GetButtonRect(BUTTON_PREVIOUS, rc)
 If VisibleFlag And PtInRect(rc, X, Y) <> 0 Then
  GetButtonCapture = BUTTON_PREVIOUS
  Exit Function
 End If
 VisibleFlag = GetButtonRect(BUTTON_NEXT, rc)
 If VisibleFlag And PtInRect(rc, X, Y) <> 0 Then
  GetButtonCapture = BUTTON_NEXT
  Exit Function
 End If
 
 For i = 0 To m_nButtons - 1
  VisibleFlag = GetButtonRect(i, rc)
  If VisibleFlag And PointInRect(rc, X, Y) Then
   'Separators and disabled buttons can't get mouse capture.
   If m_Buttons(i).ButtonType = eSeparator Then Exit For
   If m_Buttons(i).State = eDisabled Then Exit For
   GetButtonCapture = i
'   Debug.Print "Button " & i & " has mouse capture"
   Exit Function
  End If
 Next i
 GetButtonCapture = NO_BUTTON
End Function

'Returns the coords of the given button.
Private Function GetButtonRect(Index As Long, ByRef rcButton As RECT) As Boolean
 If Index = BUTTON_PREVIOUS Then
  With m_ButtonPrevious
   rcButton.Left = .Left
   rcButton.Top = .Top
   rcButton.Right = rcButton.Left + .Width
   rcButton.Bottom = rcButton.Top + .Height
   GetButtonRect = (.Flags = BTN_FULLVISIBLE)
'   Debug.Print "BUTTON_PREVIOUS.Visible = " & GetButtonRect
  End With
 ElseIf Index = BUTTON_NEXT Then
  With m_ButtonNext
   rcButton.Left = .Left
   rcButton.Top = .Top
   rcButton.Right = rcButton.Left + .Width
   rcButton.Bottom = rcButton.Top + .Height
   GetButtonRect = (.Flags = BTN_FULLVISIBLE)
'   Debug.Print "BUTTON_NEXT.Visible = " & GetButtonRect
  End With
 ElseIf Index >= 0 And Index < m_nButtons Then
  With m_Buttons(Index)
   rcButton.Left = .Left
   rcButton.Top = .Top
   rcButton.Right = rcButton.Left + .Width
   rcButton.Bottom = rcButton.Top + .Height
   GetButtonRect = (.Flags = BTN_FULLVISIBLE)
  End With
 End If
End Function

'Returns the client coords. of the given button.
Private Function GetButtonClientRect(Index As Long, ByRef rcButton As RECT, ByRef rcClient As RECT) As Boolean
 Dim Bdr_Height As Long, Bdr_Width As Long
 Dim StateSkin As Long
 
 If Index <= NO_BUTTON And Index >= m_nButtons Then Exit Function
 If Index >= 0 And Index < m_nButtons Then
  StateSkin = GetButtonStateSkin(m_Buttons(Index).State)
 ElseIf Index = BUTTON_NEXT Then
  StateSkin = GetButtonStateSkin(m_ButtonNext.State)
 ElseIf Index = BUTTON_PREVIOUS Then
  StateSkin = GetButtonStateSkin(m_ButtonPrevious.State)
 Else
  Exit Function
 End If
 If m_Skin Is Nothing Then
  Bdr_Height = m_BorderWidth
  Bdr_Width = m_BorderWidth
 Else
  Bdr_Height = m_SkinBorderHeight(StateSkin)
  Bdr_Width = m_SkinBorderWidth(StateSkin)
 End If
 With rcButton
  rcClient.Left = .Left + Bdr_Width '+ MARGINSIZE
  rcClient.Top = .Top + Bdr_Height '+ MARGINSIZE
  rcClient.Right = .Right - Bdr_Width '- MARGINSIZE
  rcClient.Bottom = .Bottom - Bdr_Height '- MARGINSIZE
 End With
 GetButtonClientRect = True
End Function

Private Function Max(Num1 As Long, Num2 As Long) As Long
 If Num1 >= Num2 Then
  Max = Num1
 Else
  Max = Num2
 End If
End Function

Private Function Min(Num1 As Long, Num2 As Long) As Long
 If Num1 <= Num2 Then
  Min = Num1
 Else
  Min = Num2
 End If
End Function

'Draws an horizontal or vertical gradient rectangle.
Private Function GradientFill(DC As Long, rc As RECT, ColorFrom As Long, ColorTo As Long, VerticalGradient As Boolean) As Boolean
 Dim C1 As RGBStruct, C2 As RGBStruct
 Dim R As Single, G As Single, B As Single
 Dim dR As Single, dB As Single, dG As Single
 Dim Rango As Long
 Dim p(2) As POINTAPI
 Dim HPen As Long, hOldPen As Long
 Dim i As Long
 
 CopyMemory C1, ColorFrom, 4&
 CopyMemory C2, ColorTo, 4&
 
 If VerticalGradient Then
  Rango = rc.Bottom - rc.Top
 Else
  Rango = rc.Right - rc.Left
 End If
 
 dR = CSng(C2.Red) - CSng(C1.Red) / Rango
 dG = CSng(C2.Green) - CSng(C1.Green) / Rango
 dB = CSng(C2.Blue) - CSng(C1.Blue) / Rango
 
 R = C1.Red: G = C1.Green: B = C1.Blue
 p(0).X = rc.Left: p(0).Y = rc.Top
 
 If VerticalGradient Then
  p(1).X = rc.Right
  For i = rc.Top To rc.Bottom
   p(0).Y = i: p(1).Y = i
   HPen = CreatePen(PS_SOLID, 1, RGB(CInt(R), CInt(G), CInt(B)))
   hOldPen = SelectObject(DC, HPen)
   Polyline DC, p(0), 2
   SelectObject DC, hOldPen
   DeleteObject HPen
   R = R + dR
   G = G + dG
   B = B + dB
  Next i
 Else
  p(1).Y = rc.Bottom
  For i = rc.Left To rc.Right
   p(0).X = i: p(1).X = i
   HPen = CreatePen(PS_SOLID, 1, RGB(CInt(R), CInt(G), CInt(B)))
   hOldPen = SelectObject(DC, HPen)
   Polyline DC, p(0), 2
   SelectObject DC, hOldPen
   DeleteObject HPen
   R = R + dR
   G = G + dG
   B = B + dB
  Next i
 End If
End Function

'Draws a tile rect from the skin into the picturebox.
Private Function TileSkinRect(rcDest As RECT, rcSrc As RECT) As Boolean
 Dim DestW As Long, DestH As Long
 Dim SrcW As Long, SrcH As Long
 Dim i As Long
 
 On Error Resume Next
 
' Debug.Print "rcButton = (" & rcDest.Left & ", " & rcDest.Top & ") - (" & rcDest.Right & ", " & rcDest.Bottom & ")"
' Debug.Print "rcSkin = (" & rcSrc.Left & ", " & rcSrc.Top & ") - (" & rcSrc.Right & ", " & rcSrc.Bottom & ")"
 With rcSrc
  SrcH = .Bottom - .Top
  SrcW = .Right - .Left
 End With
 With rcDest
  DestH = .Bottom - .Top
  DestW = .Right - .Left
 End With
 For i = 0 To (DestW \ SrcW) - 1
'  Debug.Print i & " TileBlit Horizontal"
  BitBlt m_BackBuffer.DC, rcDest.Left + (i * SrcW), rcDest.Top, SrcW, SrcH, m_SkinDC.DC, rcSrc.Left, rcSrc.Top, vbSrcCopy
 Next i
 If (DestW Mod SrcW) > 0 Then
  BitBlt m_BackBuffer.DC, rcDest.Right - (DestW Mod SrcW), rcDest.Top, (DestW Mod SrcW), SrcH, m_SkinDC.DC, rcSrc.Left, rcSrc.Top, vbSrcCopy
 End If
 For i = 1 To (DestH \ SrcH) - 1
'  Debug.Print i & " TileBlit Vertical"
  BitBlt m_BackBuffer.DC, rcDest.Left, rcDest.Top + (i * SrcH), DestW, SrcH, m_BackBuffer.DC, rcDest.Left, rcDest.Top, vbSrcCopy
 Next i
 If (DestH Mod SrcH) > 0 Then
  BitBlt m_BackBuffer.DC, rcDest.Left, rcDest.Bottom - (DestH Mod SrcH), DestW, (DestH Mod SrcH), m_BackBuffer.DC, rcDest.Left, rcDest.Top, vbSrcCopy
 End If
End Function

'Checks if coords. are inside the rectangle.
Private Function PointInRect(rc As RECT, X As Long, Y As Long) As Boolean
 If rc.Left <= X And rc.Right > X Then
  If rc.Top <= Y And rc.Bottom > Y Then
   PointInRect = True
  End If
 End If
End Function

'Creates a DC with the given size with the same attributes
'as the original DC, that can have a particular bitmap
'(DCWidth = 0, DCHeight = 0).
Private Function CloneDC(DCSrc As Long, DCWidth As Long, DCHeight As Long, DCDest As tDCData) As Boolean
 
 With DCDest
  .DC = CreateCompatibleDC(DCSrc)
'  Debug.Print "DC = " & .DC
  If DCWidth <> 0 And DCHeight <> 0 Then
   'Creates a bitmap of the given size.
    .CurBmp = CreateCompatibleBitmap(DCSrc, DCWidth, DCHeight)
   If .CurBmp <> 0 Then
    .HasBmp = True
    .OldBmp = SelectObject(.DC, .CurBmp)
   End If
  End If
  'Asign the same attributes of the original DC.
  .CurBrush = GetCurrentObject(DCSrc, OBJ_BRUSH)
  .oldBrush = SelectObject(.DC, .CurBrush)
  .CurFont = GetCurrentObject(DCSrc, OBJ_FONT)
  .OldFont = SelectObject(.DC, .CurFont)
  .CurPen = GetCurrentObject(DCSrc, OBJ_PEN)
  .OldPen = SelectObject(.DC, .CurPen)
  SetBkColor .DC, GetBkColor(DCSrc)
  SetTextColor .DC, GetTextColor(DCSrc)
  SetBkMode .DC, GetBkMode(DCSrc)
  CloneDC = .DC <> 0
 End With
End Function

'Release every object asigned to the DC and destroy it.
Private Function DestroyDC(DCData As tDCData) As Boolean
 
 With DCData
  If .DC = 0 Then Exit Function
  SelectObject .DC, .OldBmp
  If .HasBmp Then
   DeleteObject .CurBmp
  End If
  SelectObject .DC, .oldBrush
  SelectObject .DC, .OldFont
  SelectObject .DC, .OldPen
  DeleteDC .DC
  FillMemory DCData, LenB(DCData), 0
 End With
End Function

'Draws the button background and border.
Private Function DrawButtonFrame(ThisButton As tButton, rcButton As RECT, rcClient As RECT) As Long
 Dim rcSkin As RECT, rcBorder As RECT
 Dim Brush1 As Long
 Dim SkinButtonWidth As Long
 Dim ThisBtnBorderWidth As Long, ThisBtnBorderHeight As Long
 Dim m_Flag As Long, m_Flag2 As Long, Result As Long
 Dim StateSkin As Long
 Dim OldTextColor As Long, OldBkColor As Long
 
 With ThisButton
  If Not (.Skin Is Nothing) Then
   SkinButtonWidth = m_Container.ScaleX(.Skin.Width) \ 5
   'm_ButtonSkinDC.HasBmp = True always, so .CurBmp contains
   'the current bitmap (.HasBmp), therefore we just have
   'to select the button skin bitmap, blit it, and select
   'the .CurBmp bitmap again to the DC.
    Result = SkinButtonWidth * GetButtonStateSkin(.State)
'    Debug.Print "Result From SkinButton = " & Result
   SelectObject m_ButtonSkinDC.DC, .Skin.handle
   BlitPicture m_BackBuffer, m_ButtonSkinDC, .Skin, rcButton, eVCenter, Result, 0, m_MaskColor
   SelectObject m_ButtonSkinDC.DC, m_ButtonSkinDC.CurBmp
  Else
   'BUTTON STATES
   'eNormal = 0, eFocused = 1, eOver = 2, eDown = 3, eDisabled = 4
   StateSkin = GetButtonStateSkin(ThisButton.State)
'   Debug.Print "StateSkin = " & StateSkin
'   Debug.Print .Text & " - Estado = &H" & Hex(.State)
   SkinButtonWidth = m_SkinWidth / 5
   rcSkin.Left = (SkinButtonWidth * StateSkin) + m_SkinBorderWidth(StateSkin) + 1
   rcSkin.Top = m_SkinBorderHeight(StateSkin) + 1
   rcSkin.Right = (SkinButtonWidth * (StateSkin + 1)) - m_SkinBorderWidth(StateSkin) - 1
   rcSkin.Bottom = m_SkinHeight - m_SkinBorderHeight(StateSkin) - 1
'   Debug.Print "rcSkin = (" & rcSkin.Left & ", " & rcSkin.Top & ") - (" & rcSkin.Right & ", " & rcSkin.Bottom & ")"
   If m_FillMode = eSolid Then
    'If the button has no skin, its type is check and
    'its state is down, we draw it with the pattern
    'brush, else draw it normal.
    If (.ButtonType And eCheck) And (.State And eDown) Then
     Result = FillRect(m_BackBuffer.DC, rcClient, m_CheckPatternBrush)
    Else
     Brush1 = CreateSolidBrush(m_BackColor1(StateSkin))
     Result = FillRect(m_BackBuffer.DC, rcClient, Brush1)
     DeleteObject Brush1
    End If
       
   ElseIf m_FillMode = eVerticalGradient Then
    GradientFill m_BackBuffer.DC, rcClient, m_BackColor1(StateSkin), m_BackColor2(StateSkin), True
   ElseIf m_FillMode = eHorizontalGradient Then
    GradientFill m_BackBuffer.DC, rcClient, m_BackColor1(StateSkin), m_BackColor2(StateSkin), False
   'If general skin is used, select filling mode (Blit, Stretch o Tile).
   ElseIf Not (m_Skin Is Nothing) Then
    If m_FillMode = eBlit Then
     BitBlt m_BackBuffer.DC, rcClient.Left, rcClient.Top, (rcClient.Right - rcClient.Left), (rcClient.Bottom - rcClient.Top), m_SkinDC.DC, rcSkin.Left, rcSkin.Top, vbSrcCopy
    ElseIf m_FillMode = eStretchBlit Then
     StretchBlt m_BackBuffer.DC, rcClient.Left, rcClient.Top, (rcClient.Right - rcClient.Left), (rcClient.Bottom - rcClient.Top), m_SkinDC.DC, rcSkin.Left, rcSkin.Top, (rcSkin.Right - rcSkin.Left), (rcSkin.Bottom - rcSkin.Top), vbSrcCopy
    ElseIf m_FillMode = ePatternFill Then
'     Debug.Print "Tile del Fondo"
     TileSkinRect rcClient, rcSkin
    End If
   'If there's no skin but FillMode is Blit, Stretch or Tile,
   'replace it by Solid.
   Else
    'If the button has no skin, its type is check and
    'its state is down, we draw it with the pattern
    'brush, else draw it normal.
    If .ButtonType And eCheck And (.State And eDown) Then
     Result = FillRect(m_BackBuffer.DC, rcClient, m_CheckPatternBrush)
    Else
     Brush1 = CreateSolidBrush(m_BackColor1(StateSkin))
     Result = FillRect(m_BackBuffer.DC, rcClient, Brush1)
     DeleteObject Brush1
    End If
   End If
     
   If m_Skin Is Nothing Then
   'Si no hay skin para el boton, lo dibujamos tipo Windows.
   'If button has no skin, draw it like a Windows 98 button.
    If (.State And eDown) Then
'     Debug.Print "State is Down"
     m_Flag = m_Flag Or DFCS_PUSHED
     If .ButtonType And eCheck Then
      m_Flag = m_Flag Or DFCS_CHECKED
     End If
    End If
      
    m_Flag2 = BF_ADJUST Or BF_RECT
    If .Style <> e3D Then
     If (.State And eDown) Then
'      Debug.Print "State is Down"
      m_Flag = EDGE_FLAT_DOWN
     ElseIf (.State And eOver) Then
'      Debug.Print "State is Over"
      m_Flag = EDGE_FLAT
      If .Style = eHot Then m_Flag2 = BF_RECT
     Else
      m_Flag = EDGE_FLAT
      If .Style = eHot Then m_Flag2 = 0
     End If
     Result = DrawEdge(m_BackBuffer.DC, rcButton, m_Flag, m_Flag2)
    Else
     If (.State And eDown) Then
 '     Debug.Print "State is Down"
      m_Flag = EDGE_SUNKEN
     Else
      m_Flag = EDGE_RAISED
     End If
     Result = DrawEdge(m_BackBuffer.DC, rcButton, m_Flag, m_Flag2)
    End If

   ElseIf .Style = eSkinned Then
   'Draw upper border
   'Calculate upper button border rect
    SetRect rcBorder, rcButton.Left + m_SkinBorderWidth(StateSkin), rcButton.Top, rcButton.Right - m_SkinBorderWidth(StateSkin), m_SkinBorderHeight(StateSkin)
   'Calculate rect of the skin that is tiled as the border.
    SetRect rcSkin, (SkinButtonWidth * StateSkin) + m_SkinBorderWidth(StateSkin), 0, (SkinButtonWidth * (StateSkin + 1)) - m_SkinBorderWidth(StateSkin), m_SkinBorderHeight(StateSkin)
    TileSkinRect rcBorder, rcSkin
    
   'Draw lower border
   'Calculate lower button border rect
    SetRect rcBorder, rcButton.Left + m_SkinBorderWidth(StateSkin), rcButton.Bottom - m_SkinBorderHeight(StateSkin), rcButton.Right - m_SkinBorderWidth(StateSkin), rcButton.Bottom
   'Calculate rect of the skin that is tiled as the border.
    SetRect rcSkin, (SkinButtonWidth * StateSkin) + m_SkinBorderWidth(StateSkin), m_SkinHeight - m_SkinBorderHeight(StateSkin), (SkinButtonWidth * (StateSkin + 1)) - m_SkinBorderWidth(StateSkin), m_SkinHeight
    TileSkinRect rcBorder, rcSkin
   
   'Draw left border
   'Calculate left button border rect
    SetRect rcBorder, rcButton.Left, rcButton.Top + m_SkinBorderHeight(StateSkin), rcButton.Left + m_SkinBorderWidth(StateSkin), rcButton.Bottom - m_SkinBorderHeight(StateSkin)
   'Calculate rect of the skin that is tiled as the border.
    SetRect rcSkin, (SkinButtonWidth * StateSkin), m_SkinBorderHeight(StateSkin), (SkinButtonWidth * StateSkin) + m_SkinBorderWidth(StateSkin), m_SkinHeight - m_SkinBorderHeight(StateSkin)
    TileSkinRect rcBorder, rcSkin
   
   'Draw right border
   'Calculate right button border rect
    SetRect rcBorder, rcButton.Right - m_SkinBorderWidth(StateSkin), rcButton.Top + m_SkinBorderHeight(StateSkin), rcButton.Right, rcButton.Bottom - m_SkinBorderHeight(StateSkin)
   'Calculate rect of the skin that is tiled as the border.
    SetRect rcSkin, (SkinButtonWidth * (StateSkin + 1)) - m_SkinBorderWidth(StateSkin), m_SkinBorderHeight(StateSkin), (SkinButtonWidth * (StateSkin + 1)), m_SkinHeight - m_SkinBorderHeight(StateSkin)
    TileSkinRect rcBorder, rcSkin
     
'   Debug.Print "----------------------------"
   
    OldBkColor = SetBkColor(m_BackBuffer.DC, &HFFFFFF)
    OldTextColor = SetTextColor(m_BackBuffer.DC, &H0&)
   'Blit upper left corner
    SetRect rcSkin, rcButton.Left, rcButton.Top, rcButton.Left + m_SkinBorderWidth(StateSkin), rcButton.Top + m_SkinBorderHeight(StateSkin)
    BitBlt m_BackBuffer.DC, rcSkin.Left, rcSkin.Top, m_SkinBorderWidth(StateSkin), m_SkinBorderHeight(StateSkin), m_SkinDC.DC, (SkinButtonWidth * StateSkin), 0, vbSrcInvert
    BitBlt m_BackBuffer.DC, rcSkin.Left, rcSkin.Top, m_SkinBorderWidth(StateSkin), m_SkinBorderHeight(StateSkin), m_SkinMaskDC.DC, (SkinButtonWidth * StateSkin), 0, vbSrcAnd
    BitBlt m_BackBuffer.DC, rcSkin.Left, rcSkin.Top, m_SkinBorderWidth(StateSkin), m_SkinBorderHeight(StateSkin), m_SkinDC.DC, (SkinButtonWidth * StateSkin), 0, vbSrcInvert
   
   'Blit upper right corner
    SetRect rcSkin, rcButton.Right - m_SkinBorderWidth(StateSkin), rcButton.Top, rcButton.Right, rcButton.Top + m_SkinBorderHeight(StateSkin)
    BitBlt m_BackBuffer.DC, rcSkin.Left, rcSkin.Top, m_SkinBorderWidth(StateSkin), m_SkinBorderHeight(StateSkin), m_SkinDC.DC, (SkinButtonWidth * (StateSkin + 1)) - m_SkinBorderWidth(StateSkin), 0, vbSrcInvert
    BitBlt m_BackBuffer.DC, rcSkin.Left, rcSkin.Top, m_SkinBorderWidth(StateSkin), m_SkinBorderHeight(StateSkin), m_SkinMaskDC.DC, (SkinButtonWidth * (StateSkin + 1)) - m_SkinBorderWidth(StateSkin), 0, vbSrcAnd
    BitBlt m_BackBuffer.DC, rcSkin.Left, rcSkin.Top, m_SkinBorderWidth(StateSkin), m_SkinBorderHeight(StateSkin), m_SkinDC.DC, (SkinButtonWidth * (StateSkin + 1)) - m_SkinBorderWidth(StateSkin), 0, vbSrcInvert
   
   'Blit lower left corner
    SetRect rcSkin, rcButton.Left, rcButton.Bottom - m_SkinBorderHeight(StateSkin), rcButton.Left + m_SkinBorderWidth(StateSkin), rcButton.Bottom
    BitBlt m_BackBuffer.DC, rcSkin.Left, rcSkin.Top, m_SkinBorderWidth(StateSkin), m_SkinBorderHeight(StateSkin), m_SkinDC.DC, (SkinButtonWidth * StateSkin), m_SkinHeight - m_SkinBorderHeight(StateSkin), vbSrcInvert
    BitBlt m_BackBuffer.DC, rcSkin.Left, rcSkin.Top, m_SkinBorderWidth(StateSkin), m_SkinBorderHeight(StateSkin), m_SkinMaskDC.DC, (SkinButtonWidth * StateSkin), m_SkinHeight - m_SkinBorderHeight(StateSkin), vbSrcAnd
    BitBlt m_BackBuffer.DC, rcSkin.Left, rcSkin.Top, m_SkinBorderWidth(StateSkin), m_SkinBorderHeight(StateSkin), m_SkinDC.DC, (SkinButtonWidth * StateSkin), m_SkinHeight - m_SkinBorderHeight(StateSkin), vbSrcInvert
   
   'Blit lower right corner
    SetRect rcSkin, rcButton.Right - m_SkinBorderWidth(StateSkin), rcButton.Bottom - m_SkinBorderHeight(StateSkin), rcButton.Right, rcButton.Bottom
    BitBlt m_BackBuffer.DC, rcSkin.Left, rcSkin.Top, m_SkinBorderWidth(StateSkin), m_SkinBorderHeight(StateSkin), m_SkinDC.DC, (SkinButtonWidth * (StateSkin + 1)) - m_SkinBorderWidth(StateSkin), m_SkinHeight - m_SkinBorderHeight(StateSkin), vbSrcInvert
    BitBlt m_BackBuffer.DC, rcSkin.Left, rcSkin.Top, m_SkinBorderWidth(StateSkin), m_SkinBorderHeight(StateSkin), m_SkinMaskDC.DC, (SkinButtonWidth * (StateSkin + 1)) - m_SkinBorderWidth(StateSkin), m_SkinHeight - m_SkinBorderHeight(StateSkin), vbSrcAnd
    BitBlt m_BackBuffer.DC, rcSkin.Left, rcSkin.Top, m_SkinBorderWidth(StateSkin), m_SkinBorderHeight(StateSkin), m_SkinDC.DC, (SkinButtonWidth * (StateSkin + 1)) - m_SkinBorderWidth(StateSkin), m_SkinHeight - m_SkinBorderHeight(StateSkin), vbSrcInvert
   
    SetTextColor m_BackBuffer.DC, OldTextColor
    SetBkColor m_BackBuffer.DC, OldBkColor
   End If
     
  End If
 End With
End Function

'Draw button text.
Private Function DrawButtonText(ThisButton As tButton, rcButton As RECT, rcClient As RECT) As Long
 Dim rcText As RECT
 Dim m_Flag As Long, m_Flag2 As Long, Result As Long
 Dim BkMode As Long
 Dim StateSkin As Long
 
 
 With ThisButton
   StateSkin = GetButtonStateSkin(ThisButton.State)
 
  rcText.Left = rcClient.Left + MARGINSIZE
  rcText.Top = rcClient.Top + MARGINSIZE
  rcText.Right = rcClient.Right - MARGINSIZE
  rcText.Bottom = rcClient.Bottom - MARGINSIZE
  If .Text = "" Then Exit Function
  If (.TextAlign And eRight) = eRight Then
   m_Flag = DT_RIGHT
  ElseIf (.TextAlign And eHCenter) = eHCenter Then
   m_Flag = DT_CENTER
  Else
   'do nothing, DT_LEFT by default
  End If
  If (.TextAlign And eBottom) = eBottom Then
   m_Flag = m_Flag Or DT_BOTTOM
  ElseIf (.TextAlign And eVCenter) = eVCenter Then
   m_Flag = m_Flag Or DT_VCENTER
  Else
   'do nothing, DT_TOP by default.
  End If
  
   m_Flag = m_Flag Or DT_SINGLELINE Or DT_END_ELLIPSIS Or DT_WORD_ELLIPSIS
  
  If (.State And eDown) Then OffsetRect rcText, 1, 1
      
  BkMode = SetBkMode(m_BackBuffer.DC, TRANSPARENT)
  If m_TextColor1(StateSkin) <> m_TextColor2(StateSkin) Then
   SetTextColor m_BackBuffer.DC, m_TextColor2(StateSkin)
   Result = DrawTextEx(m_BackBuffer.DC, .Text, Len(.Text), rcText, m_Flag, ByVal 0&)
   If Result = 0 Then Debug.Print "DrawTextEx fallo con el texto " & .Text
   OffsetRect rcText, -1, -1
  End If
  SetTextColor m_BackBuffer.DC, m_TextColor1(StateSkin)
  Result = DrawTextEx(m_BackBuffer.DC, .Text, Len(.Text), rcText, m_Flag, ByVal 0&)
  If Result = 0 Then Debug.Print "DrawTextEx fallo con el texto " & .Text
  SetBkMode m_BackBuffer.DC, BkMode
 End With
End Function

'Draw button icon.
Private Function DrawButtonIcon(ThisButton As tButton, rcButton As RECT, rcClient As RECT) As RECT
 Dim rcIcon As RECT
 Dim Brush1 As Long
 Dim OldBkColor As Long, OldTextColor As Long
 Dim Icon_Width As Long, Icon_Height As Long
 Dim m_Flag As Long, m_Flag2 As Long, Result As Long
 Dim m_IconDC As tDCData, m_IconMaskDC As tDCData
 
 With ThisButton
  If .Icon Is Nothing Then Exit Function
  
  Icon_Height = m_Container.ScaleY(.Icon.Height) ', 8, vbPixels)
  Icon_Width = m_Container.ScaleX(.Icon.Width) ', 8, vbPixels)
  If (.IconAlign And eLeft) = eLeft Then
   rcIcon.Left = rcClient.Left + MARGINSIZE
  ElseIf (.IconAlign And eHCenter) = eHCenter Then
   rcIcon.Left = rcClient.Left + MARGINSIZE + (rcClient.Right - rcClient.Left - Icon_Width - MARGINSIZE * 2) \ 2
  Else 'If (.IconAlign And eRight) = eRight Then
   rcIcon.Left = rcClient.Right - MARGINSIZE - Icon_Width
  End If
  rcIcon.Right = rcIcon.Left + Icon_Width
    
  If (.IconAlign And eTop) = eTop Then
   rcIcon.Top = rcClient.Top + MARGINSIZE
  ElseIf (.IconAlign And eVCenter) = eVCenter Then
   rcIcon.Top = rcClient.Top + MARGINSIZE + ((rcClient.Bottom - rcClient.Top - Icon_Height - MARGINSIZE * 2) \ 2)
  Else
   rcIcon.Top = rcClient.Bottom - MARGINSIZE - Icon_Height
  End If
  rcIcon.Bottom = rcIcon.Top + Icon_Height
    
  'Check if button image is an icon or bitmap.
  If .Icon.Type = vbPicTypeIcon Then
   m_Flag = DST_ICON
  ElseIf .Icon.Type = vbPicTypeBitmap Then
   m_Flag = DST_BITMAP
  Else
   Exit Function
  End If
    
  CloneDC 0, 0, 0, m_IconDC
  CloneDC 0, 0, 0, m_IconMaskDC
  m_IconDC.OldBmp = SelectObject(m_IconDC.DC, .Icon.handle)
  CreateMaskImage m_IconDC, m_IconMaskDC, Icon_Width, Icon_Height, m_MaskColor
  
  If (.State And eDown) Then
   OffsetRect rcIcon, 1, 1
  ElseIf (.State And eDisabled) Then
   DrawDisabledBitmap m_BackBuffer, m_IconDC, m_IconMaskDC, rcIcon.Left, rcIcon.Top, Icon_Width, Icon_Height
   GoTo ReleaseIcon
  ElseIf (.State And eOver) Then
  'Creates a shadow icon for the button. NOT IMPLEMENTED
   If (.ButtonType And eCheck) And (.State And eDown) Then
    'do nothing
   Else
    'Brush1 = CreateSolidBrush(SHADOW_COLOR)
    'result = DrawState(m_BackBuffer.DC, Brush1, 0&, .Icon.handle, 0, rcIcon.Left + 1, rcIcon.Top + 1, Icon_Width, Icon_Height, m_Flag Or DSS_MONO)
    'DeleteObject Brush1
'    OldBkColor = SetBkColor(m_BackBuffer.DC, m_MaskColor)
'    OldTextColor = SetTextColor(m_BackBuffer.DC, SHADOW_COLOR)
'    BitBlt m_BackBuffer.DC, rcIcon.Left + 1, rcIcon.Top + 1, Icon_Width, Icon_Height, m_IconDC.DC, 0, 0, vbSrcInvert
'    BitBlt m_BackBuffer.DC, rcIcon.Left + 1, rcIcon.Top + 1, Icon_Width, Icon_Height, m_IconMaskDC.DC, 0, 0, vbSrcAnd
'    BitBlt m_BackBuffer.DC, rcIcon.Left + 1, rcIcon.Top + 1, Icon_Width, Icon_Height, m_IconMaskDC.DC, 0, 0, vbSrcInvert
   End If
  End If

  If .Icon.Type = vbPicTypeIcon Then
   Result = DrawState(m_BackBuffer.DC, 0&, 0&, .Icon.handle, 0, rcIcon.Left, rcIcon.Top, Icon_Width, Icon_Height, m_Flag)
  ElseIf .Icon.Type = vbPicTypeBitmap Then
'   If CloneDC(0, 0, 0, m_IconDC) Then
'    Debug.Print "DrawButtonIcon - BlitPicture"
    BlitPicture m_BackBuffer, m_IconDC, .Icon, rcIcon, .IconAlign, 0, 0, m_MaskColor
'   End If
  End If
  
ReleaseIcon:
  SelectObject m_IconDC.DC, m_IconDC.OldBmp
  DestroyDC m_IconDC
  DestroyDC m_IconMaskDC
 End With
 DrawButtonIcon = rcIcon
End Function

'Draws an image into the given DC with the given coords
'and transparency option.
Private Function BlitPicture(DestDC As tDCData, SrcDC As tDCData, Img As StdPicture, rcDest As RECT, ImgAlignment As eAlignment, SrcX As Long, SrcY As Long, Optional ByVal MaskColor As Long = -1) As Long
 Dim MaskDC As tDCData
 Dim Img_Height As Long, Img_Width As Long
 Dim OldBkColor As Long, OldTextColor As Long
 
 Img_Height = Min(m_Container.ScaleY(Img.Height), rcDest.Bottom - rcDest.Top)
 Img_Width = Min(m_Container.ScaleX(Img.Width), rcDest.Right - rcDest.Left)
 
 If (ImgAlignment And eLeft) = eLeft Then
  rcDest.Left = rcDest.Left
 ElseIf (ImgAlignment And eHCenter) = eHCenter Then
  rcDest.Left = rcDest.Left + (rcDest.Right - rcDest.Left - Img_Width) \ 2
 Else
  rcDest.Left = rcDest.Right - Img_Width
 End If
 rcDest.Right = rcDest.Left + Img_Width
    
 If (ImgAlignment And eTop) = eTop Then
  rcDest.Top = rcDest.Top
 ElseIf (ImgAlignment And eVCenter) = eVCenter Then
  rcDest.Top = rcDest.Top + ((rcDest.Bottom - rcDest.Top - Img_Height) \ 2)
 Else
  rcDest.Top = rcDest.Bottom - Img_Height
 End If
 rcDest.Bottom = rcDest.Top + Img_Height
 
 If MaskColor < 0 Then
  'Blit directly
  SelectObject SrcDC.DC, Img.handle
  BitBlt DestDC.DC, rcDest.Left, rcDest.Top, Img_Width, Img_Height, SrcDC.DC, SrcX, SrcY, vbSrcCopy
  SelectObject SrcDC.DC, SrcDC.CurBmp
 Else
  'Do a transparent blit
  'Create a new DC with the default bitmap, that will
  'hold the image mask.
  CloneDC SrcDC.DC, 0, 0, MaskDC
  'Create the mascara in MaskDC.
  If CreateMaskImage(SrcDC, MaskDC, m_Container.ScaleX(Img.Width), m_Container.ScaleY(Img.Height), m_MaskColor) Then
   'Para el Transparent Blit, utilizamos la tecnica del
   '"True Mask Method", de Ron Gery (Bitmaps with Transparency)
   OldBkColor = SetBkColor(DestDC.DC, RGB(255, 255, 255))
   OldTextColor = SetTextColor(DestDC.DC, RGB(0, 0, 0))
   BitBlt DestDC.DC, rcDest.Left, rcDest.Top, Img_Width, Img_Height, SrcDC.DC, SrcX, SrcY, vbSrcInvert
   BitBlt DestDC.DC, rcDest.Left, rcDest.Top, Img_Width, Img_Height, MaskDC.DC, SrcX, SrcY, vbSrcAnd
   BitBlt DestDC.DC, rcDest.Left, rcDest.Top, Img_Width, Img_Height, SrcDC.DC, SrcX, SrcY, vbSrcInvert
'   BitBlt frmTest.hdc, 200, 85, Img_Width, Img_Height, MaskDC.DC, 0, 0, vbSrcCopy
   SetBkColor DestDC.DC, OldBkColor
   SetTextColor DestDC.DC, OldTextColor
  Else
   'If mask could not be created, do a SrcCopy.
'   Debug.Print "No se pudo crear la mascara en BlitPicture"
   BitBlt DestDC.DC, rcDest.Left, rcDest.Top, Img_Width, Img_Height, SrcDC.DC, 0, 0, vbSrcCopy
  End If
  DestroyDC MaskDC
 End If
End Function

'Checks if given key was pressed.
Private Function IsKeyDown(VbKey As KeyCodeConstants) As Boolean
 IsKeyDown = False
 If GetKeyState(VbKey) And -256 Then
  IsKeyDown = True
 End If
End Function

'Draw the focus rectangle on the given button.
Private Sub DrawFocus(Index As Long)
 Dim rcButton As RECT, rcClient As RECT
 Dim Btn As tButton, hwnd As Long
 
 hwnd = GetFocusAPI()
' Debug.Print "Hwnd with Focus = " & hwnd
' Debug.Print "m_Container.hwnd = " & m_Container.hwnd
 If hwnd <> m_Container.hwnd Then Exit Sub
 If Index >= 0 And Index < m_nButtons Then
  Btn = m_Buttons(Index)
 End If
 With Btn
  If .ButtonType = eSeparator Then Exit Sub
  If (m_BarFlags And fDRAWFOCUSRECT) = fDRAWFOCUSRECT Then
   If (.State And eFocused) Then
    If GetButtonRect(Index, rcButton) Then
     GetButtonClientRect Index, rcButton, rcClient
     If .Style = e3D Then
      With rcClient
       .Left = .Left + 2
       .Top = .Top + 2
       .Right = .Right - 2
       .Bottom = .Bottom - 2
      End With
     ElseIf (.Style = eHot) And (.State And eNormal) Then
      Debug.Print "Hot Style = " & (.State And eNormal) 'Then
      Exit Sub
     Else
      With rcClient
       .Left = .Left + 1
       .Top = .Top + 1
       .Right = .Right - 1
       .Bottom = .Bottom - 1
      End With
     End If
     DrawFocusRectAPI m_BackBuffer.DC, rcClient
    End If
   End If
  End If
 End With
End Sub

'Returns skin index that match the given button state.
Private Function GetButtonStateSkin(State As eButtonStates) As Long
 If (State And eDisabled) Then
  GetButtonStateSkin = BTN_SKN_DISABLED
 ElseIf (State And eDown) Then
  GetButtonStateSkin = BTN_SKN_DOWN
 ElseIf (State And eOver) Then
  GetButtonStateSkin = BTN_SKN_OVER
 ElseIf (State And eFocused) And (m_Container.hwnd = GetFocusAPI) Then
  GetButtonStateSkin = BTN_SKN_FOCUSED
 Else
  GetButtonStateSkin = BTN_SKN_NORMAL
 End If
End Function

'Creates a mask image from a given image.
Private Function CreateMaskImage(ByRef SrcDC As tDCData, ByRef MaskDC As tDCData, SknW As Long, SknH As Long, ByVal lTransparentColor As Long) As Boolean
 Dim OldBkColor As Long
 Dim CurBmp As Long
 
 If (MaskDC.DC <> 0) Then
  If (SrcDC.DC <> 0) Then
   'If there's a bitmap (mask) on the DC, delete it to
   'create the new one.
   CurBmp = CreateBitmap(SknW, SknH, 1, 1, ByVal 0&)
   If CurBmp <> 0 Then
    If MaskDC.HasBmp Then
     SelectObject MaskDC.DC, CurBmp
     DeleteObject MaskDC.CurBmp
    Else
     MaskDC.OldBmp = SelectObject(MaskDC.DC, CurBmp)
    End If
    MaskDC.CurBmp = CurBmp
    MaskDC.HasBmp = True
   Else
    CreateMaskImage = False
    Exit Function
   End If
   'Create the mask.
   OldBkColor = SetBkColor(SrcDC.DC, lTransparentColor)
   BitBlt MaskDC.DC, 0, 0, SknW, SknH, SrcDC.DC, 0, 0, vbSrcCopy
'   BitBlt frmTest.hdc, 200, 85, SknW, SknH, MaskDC.DC, 0, 0, vbSrcCopy
   SetBkColor SrcDC.DC, OldBkColor
  Else
  End If
 Else
 End If
 CreateMaskImage = MaskDC.HasBmp
End Function

'Checks the number of visible buttons, and if necessary
'shows the navigation buttons or not, and redraws the bar.
Private Function CheckVisibleButtons() As Long
 Dim i As Long
 Dim rcMove As RECT
 Dim FirstNotShowable As Long, NotShowable As Boolean

 Me.LockUpdate = True
 m_nVisibleButtons = 0
 FirstNotShowable = m_nButtons
 SetRect rcMove, m_ButtonPrevious.Left, 0, m_Container.ScaleWidth, m_Container.ScaleHeight
 For i = m_FirstVisibleButton To m_nButtons - 1
  With m_Buttons(i)
   'Any button that does not intersect rcMove is directly
   'visible.
   If (.Left + .Width) < rcMove.Left Then
    'Button is "Showable"
    m_nVisibleButtons = m_nVisibleButtons + 1
    .Flags = .Flags Or BTN_SHOWABLE
   Else
    'If part of the button is outside the picturebox,
    'the button is directly not visible, that's why we
    'must show the navigation buttons.
    If (.Left + .Width) > m_Container.ScaleWidth Then
'     Debug.Print "ScaleWidth - " & .Text & " NO ES VISIBLE"
     m_ButtonPrevious.Flags = BTN_FULLVISIBLE
     m_ButtonNext.Flags = BTN_FULLVISIBLE
     NotShowable = True
     If FirstNotShowable = m_nButtons Then FirstNotShowable = i
     Exit For
    ElseIf (.Left + .Width) > rcMove.Left Then
    'Check if button intersect rcMove, but does not go outside
    'the picturebox, then the button is likely to be not visible.
     If FirstNotShowable = m_nButtons Then FirstNotShowable = i
     'Button can be shown.
     m_nVisibleButtons = m_nVisibleButtons + 1
     .Flags = .Flags Or BTN_SHOWABLE
    End If
   End If
  End With
  
 Next i
 
 If NotShowable Or (m_FirstVisibleButton > 0) Then
  'Hide the buttons that can't be shown.
'  Debug.Print "NotShowable From " & FirstNotShowable & " to " & m_nButtons - 1
   For i = FirstNotShowable To m_nButtons - 1
    m_Buttons(i).Flags = m_Buttons(i).Flags And (Not BTN_SHOWABLE)
   Next i
   m_nVisibleButtons = FirstNotShowable - m_FirstVisibleButton
'  Debug.Print "BUTTON PREVIOS.Disabled = " & (m_FirstVisibleButton = 0)
'  Debug.Print "BUTTON NEXT.Disabled = " & ((m_FirstVisibleButton + m_nVisibleButtons) = m_nButtons)
  Me.SetButtonState BUTTON_PREVIOUS, eDisabled, (m_FirstVisibleButton = 0)
  Me.SetButtonState BUTTON_NEXT, eDisabled, ((m_FirstVisibleButton + m_nVisibleButtons) > m_nButtons - 1)
 Else
  'Hide the navigation buttons, make them visible when
  'necessary.
  m_ButtonPrevious.Flags = m_ButtonPrevious.Flags And (Not BTN_VISIBLE)
  m_ButtonNext.Flags = m_ButtonNext.Flags And (Not BTN_VISIBLE)
 End If
 
 
 Me.LockUpdate = False
' Debug.Print
' RedrawBar
End Function

'Creates a DIB from the given image and modifies it such
'that looks like a 3D grayed disabled image, and draws
'it on the given DC at the given position.
Private Sub DrawDisabledBitmap(DestDC As tDCData, ImgDC As tDCData, MaskDC As tDCData, DestX As Long, DestY As Long, ImgW As Long, ImgH As Long, Optional ByVal Factor As Long = 655)
Dim sa As SAFEARRAY2D
Dim TempDC As Long
Dim DibBmp As Long
Dim PtrBits As Long
Dim bmi As BITMAPINFO
Dim BmpBits() As Byte
Dim OldBmp As Long
Dim Color_Pixel As Long 'RGBQUAD
Dim X As Long, Y As Long
Dim DarkColor As Long, LightColor As Long
Dim OldBkColor As Long, OldTextColor As Long
Dim GreyColor As Long

 If ImgDC.DC = 0 Then Exit Sub
 
 With bmi.bmiHeader
  .biBitCount = 24
  .biHeight = ImgH
  .biWidth = ImgW
  .biSize = LenB(bmi.bmiHeader)
  .biPlanes = 1
 End With
 
 DibBmp = CreateDIBSection(Me.hdc, bmi, DIB_RGB_COLORS, PtrBits, 0&, 0&)
 If DibBmp <> 0 Then
  With sa
   .cDims = 2
   .cbElements = 1
   .Bounds(0).lLbound = 0
   .Bounds(0).cElements = ImgH
   .Bounds(1).lLbound = 0
   .Bounds(1).cElements = (ImgW * 3 + 3) And &HFFFFFFFC
   .pvData = PtrBits
  End With
  CopyMemory ByVal VarPtrArray(BmpBits()), VarPtr(sa), 4&
  TempDC = CreateCompatibleDC(Me.hdc)
  If TempDC <> 0 Then
   OldBmp = SelectObject(TempDC, DibBmp)
   BitBlt TempDC, 0, 0, ImgW, ImgH, ImgDC.DC, 0, 0, vbSrcCopy
   DarkColor = m_TextColor1(BTN_SKN_DISABLED)
   LightColor = m_TextColor2(BTN_SKN_DISABLED)
   For Y = 0 To bmi.bmiHeader.biHeight - 1
    For X = 0 To bmi.bmiHeader.biWidth - 1
     'Get the pixel (x,y)
     Color_Pixel = ((BmpBits(X * 3, Y)) * &H10000) + ((BmpBits(X * 3 + 1, Y)) * &H100&) + BmpBits(X * 3 + 2, Y)
     'If transparency color, skip those pixels.
     If m_MaskColor >= 0 And (Color_Pixel = m_MaskColor) Then
      'do nothing
     Else
      GreyColor = (CLng(BmpBits(X * 3, Y)) + CLng(BmpBits(X * 3 + 1, Y)) + CLng(BmpBits(X * 3 + 2, Y))) '\ 3
      If GreyColor > Factor Then
       BmpBits(X * 3, Y) = (LightColor \ &H10000) And &HFF
       BmpBits(X * 3 + 1, Y) = (LightColor \ &H100) And &HFF
       BmpBits(X * 3 + 2, Y) = LightColor And &HFF
      Else
       BmpBits(X * 3, Y) = (DarkColor \ &H10000) And &HFF
       BmpBits(X * 3 + 1, Y) = (DarkColor \ &H100) And &HFF
       BmpBits(X * 3 + 2, Y) = DarkColor And &HFF
      End If
     End If
    Next X
   Next Y
   
ReleaseAll:
   'For Transparent Blit, use the "True Mask Method" tecnique
   'from Ron Gery (Bitmaps with Transparency)
   OldBkColor = SetBkColor(DestDC.DC, RGB(255, 255, 255))
   OldTextColor = SetTextColor(DestDC.DC, RGB(0, 0, 0))
   BitBlt DestDC.DC, DestX, DestY, ImgW, ImgH, TempDC, 0, 0, vbSrcInvert
   BitBlt DestDC.DC, DestX, DestY, ImgW, ImgH, MaskDC.DC, 0, 0, vbSrcAnd
   BitBlt DestDC.DC, DestX, DestY, ImgW, ImgH, TempDC, 0, 0, vbSrcInvert
   SetBkColor DestDC.DC, OldBkColor
   SetTextColor DestDC.DC, OldTextColor
   SelectObject TempDC, OldBmp
   DeleteDC TempDC
  End If
  CopyMemory ByVal VarPtrArray(BmpBits()), 0&, 4&
  DeleteObject DibBmp
 End If
End Sub

'Resize the bitmap of the given memory DC.
Private Sub ResizeDC(SrcDC As tDCData, DCWidth As Long, DCHeight As Long)
 Dim ThisBmp As Long
 
 With SrcDC
  If DCWidth <> 0 And DCHeight <> 0 Then
   
   'Creates a bitmap of the given size.
   ThisBmp = CreateCompatibleBitmap(SrcDC.DC, DCWidth, DCHeight)
   If .HasBmp Then
    DeleteObject (SelectObject(.DC, ThisBmp))
   End If
   .CurBmp = ThisBmp
  End If
 End With
End Sub

'Returns the first available button (enabled, visible and
'not separator type button) within the given range, if any.
Private Function NextAvailableButton(FromIndex As Long, ToIndex As Long) As Long
 Dim i As Long
 
 NextAvailableButton = NO_BUTTON
 
 'If Index "From" isn't valid, exit.
 If FromIndex < 0 Or FromIndex >= m_nButtons Then Exit Function
 'If Index "To" isn't valid, exit.
 If ToIndex < 0 Or ToIndex >= m_nButtons Then Exit Function
 
 If FromIndex <= ToIndex Then
  For i = FromIndex To ToIndex
   With m_Buttons(i)
    If (.ButtonType <> eSeparator) And (.State <> eDisabled) And (.Flags = BTN_FULLVISIBLE) Then
     NextAvailableButton = i
     Exit Function
    End If
   End With
  Next i
 ElseIf FromIndex > ToIndex Then
  For i = FromIndex To ToIndex Step -1
   With m_Buttons(i)
    If (.ButtonType <> eSeparator) And (.State <> eDisabled) And (.Flags = BTN_FULLVISIBLE) Then
     NextAvailableButton = i
     Exit Function
    End If
   End With
  Next i
 
 End If
End Function

'Rutine taken from the following code:
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=62187&lngWId=1
Private Function IsOverhWnd(hwnd As Long, X As Single, Y As Single) As Boolean
Dim Rec As RECT

    'get control position within the desktop
    If GetWindowRect(hwnd, Rec) = 0 Then Exit Function
    
'    'x & y are currently in twips, so convert them to pixels
'    X = X / Screen.TwipsPerPixelX
'    Y = Y / Screen.TwipsPerPixelY
    
    'check if cursor is over the control
    If (X < 0) Or (Y < 0) Or (X > Rec.Right - Rec.Left) Or (Y > Rec.Bottom - Rec.Top) Then
        ReleaseCapture 'stop capturing the mouse
        IsOverhWnd = False
       Else
        SetCapture hwnd 'capture the mouse leaving the control
        IsOverhWnd = True
    End If
    
End Function
