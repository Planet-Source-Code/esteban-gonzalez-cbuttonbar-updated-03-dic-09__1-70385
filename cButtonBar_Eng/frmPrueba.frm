VERSION 5.00
Begin VB.Form frmPrueba 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mis Pruebas"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6570
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   305
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   438
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Draw Disabled"
      Height          =   540
      Left            =   5100
      TabIndex        =   7
      Top             =   2025
      Width           =   1290
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CreateMask"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   5100
      TabIndex        =   6
      Top             =   1350
      Width           =   1290
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test Booleano"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5100
      TabIndex        =   5
      Top             =   750
      Width           =   1290
   End
   Begin VB.PictureBox picTest 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2040
      Index           =   0
      Left            =   75
      ScaleHeight     =   132
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   292
      TabIndex        =   2
      Top             =   2175
      Width           =   4440
   End
   Begin VB.CommandButton btnTileBlit 
      Caption         =   "Tile Blit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5100
      TabIndex        =   1
      Top             =   150
      Width           =   1290
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   4575
      Picture         =   "frmPrueba.frx":0000
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   0
      Top             =   300
      Width           =   450
   End
   Begin VB.PictureBox picTest 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2040
      Index           =   1
      Left            =   75
      ScaleHeight     =   132
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   292
      TabIndex        =   3
      Top             =   150
      Width           =   4440
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   4575
      Picture         =   "frmPrueba.frx":1022
      Top             =   1350
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Left            =   4650
      Picture         =   "frmPrueba.frx":1C64
      Top             =   750
      Width           =   240
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3420
      Left            =   4650
      TabIndex        =   4
      Top             =   1050
      Width           =   1770
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmPrueba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32.dll" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, ByRef lpBits As Any) As Long
Private Declare Function SetBkColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Const COLOR_BTNSHADOW As Long = 16
Private Const COLOR_BTNHIGHLIGHT As Long = 20
Private Const COLOR_3DHILIGHT As Long = COLOR_BTNHIGHLIGHT
Private Const COLOR_3DSHADOW As Long = COLOR_BTNSHADOW
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long

Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


'Copia en mosaico un rectangulo perteneciente a la skin
'en un rectangulo perteneciente a la barra.
Private Function TileRect() As Boolean
 Dim DestW As Long, DestH As Long
 Dim SrcW As Long, SrcH As Long
 Dim i As Long
 
 For i = 0 To (picTest(1).ScaleWidth \ picIcon.ScaleWidth)
  BitBlt picTest(1).hdc, (i * picIcon.ScaleWidth), 0, picIcon.ScaleWidth, picIcon.ScaleHeight, picIcon.hdc, 0, 0, vbSrcCopy
 Next i
 For i = 1 To (picTest(1).ScaleHeight \ picIcon.ScaleHeight)
  BitBlt picTest(1).hdc, 0, (i * picIcon.ScaleHeight), picTest(1).ScaleWidth, picIcon.ScaleHeight, picTest(1).hdc, 0, 0, vbSrcCopy
 Next i
 BitBlt picTest(0).hdc, 0, 0, picTest(1).ScaleWidth, picTest(1).ScaleHeight, picTest(1).hdc, 0, 0, vbSrcCopy
End Function

Private Sub btnTileBlit_Click()
 Dim rcIcon As RECT, rcTile As RECT
 'TileRect
 With rcIcon
  .Left = 0
  .Top = 3
  .Right = 3
  .Bottom = 27
 End With
 With rcTile
  .Left = 0
  .Top = 3
  .Right = 3
  .Bottom = 129
 End With
 TileSkinRect rcTile, rcIcon
End Sub

'Copia en mosaico un rectangulo perteneciente a la skin
'en un rectangulo perteneciente a la barra.
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
' SelectClipRgn m_BackBuffer.DC, 0&
' BeginPath m_BackBuffer.DC
' Rectangle m_BackBuffer.DC, rcDest.Left, rcDest.Top, rcDest.Right, rcDest.Bottom
' EndPath m_BackBuffer.DC
' SelectClipPath m_BackBuffer.DC, RGN_AND
 lblInfo.Caption = ""
 For i = 1 To (DestW \ SrcW)
  lblInfo.Caption = lblInfo.Caption & vbCrLf & i & " TileBlit Horizontal"
  lblInfo.Refresh
  BitBlt picTest(1).hdc, rcDest.Left + (i * SrcW), rcDest.Top, SrcW, SrcH, picIcon.hdc, rcSrc.Left, rcSrc.Top, vbSrcCopy
  Sleep 500
 Next i
 lblInfo.Caption = lblInfo.Caption & vbCrLf & "ACA BLITEAMOS HORIZONTALMENTE" & vbCrLf
 lblInfo.Refresh
 Sleep 3000
 For i = 1 To (DestH \ SrcH)
  lblInfo.Caption = lblInfo.Caption & vbCrLf & i & " TileBlit Vertical"
  lblInfo.Refresh
  BitBlt picTest(0).hdc, rcDest.Left, rcDest.Top + (i * SrcH), DestW, DestH, picTest(1).hdc, rcDest.Left, rcDest.Top, vbSrcCopy
  Sleep 500
 Next i
' SelectClipRgn m_BackBuffer.DC, 0&
End Function

Private Sub Command1_Click()
 Dim s As String
 
  s = "12 (1100) And 8 (1000) = " & (12 And 8) & " = " & CBool(12 And 8)
  s = s & vbCrLf & "12 (1100) And 3 (0011) = " & (12 And 3) & " = " & CBool(12 And 3)
  MsgBox s
End Sub

Private Sub Command2_Click()
 Dim ImgDC As Long
 Dim MaskDC As Long
 Dim MaskBmp As Long, OldBmp As Long, OldMask As Long
 
 ImgDC = CreateCompatibleDC(0)
 MaskDC = CreateCompatibleDC(0)
 OldBmp = SelectObject(ImgDC, imgIcon.Picture.handle)
 MaskBmp = CreateBitmap(Me.ScaleX(imgIcon.Picture.Width), Me.ScaleY(imgIcon.Picture.Height), 1, 1, 0&)
 OldMask = SelectObject(MaskDC, MaskBmp)
 SetBkColor ImgDC, RGB(255, 0, 255)
 BitBlt MaskDC, 0, 0, Me.ScaleX(imgIcon.Picture.Width), Me.ScaleY(imgIcon.Picture.Height), ImgDC, 0, 0, vbSrcCopy
 BitBlt Me.hdc, 310, 120, Me.ScaleX(imgIcon.Picture.Width), Me.ScaleY(imgIcon.Picture.Height), MaskDC, 0, 0, vbSrcCopy
End Sub

Private Sub Command3_Click()
 Dim hDib As Long, hDibDC As Long, OldDibBmp As Long
 Dim OldBmp As Long, hdc As Long
 Dim OldBkColor As Long, OldTextColor As Long
 
 hdc = CreateCompatibleDC(Me.hdc)
 If hdc <> 0 Then
  hDibDC = CreateCompatibleDC(Me.hdc)
  If hDibDC <> 0 Then
'   hDib = CreateBitmap(32, 32, 1, 1, 0&)
'    hdib = create
   If hDib <> 0 Then
    OldDibBmp = SelectObject(hDibDC, hDib)
    OldBmp = SelectObject(hdc, Image1.Picture.handle)
    BitBlt hDibDC, 0, 0, 32, 32, hdc, 0, 0, vbSrcCopy
'    OldBkColor = SetBkColor(Me.hdc, GetSysColor(COLOR_3DHILIGHT))
'    OldTextColor = SetBkColor(Me.hdc, GetSysColor(COLOR_3DSHADOW))
    BitBlt Me.hdc, Image1.Left, Command3.Top, 32, 32, hDibDC, 0, 0, vbSrcCopy
'    SetBkColor Me.hdc, OldBkColor
'    SetTextColor Me.hdc, OldTextColor
    DeleteObject SelectObject(hDibDC, OldDibBmp)
    SelectObject hdc, OldBmp
   End If
   DeleteDC hDibDC
  End If
  DeleteDC hdc
 End If
End Sub

Private Const COLOR_SCROLLBAR As Long = 0
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal Bytelen As Long)
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal HBrush As Long) As Long
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, ByRef lpBits As Any) As Long

Private Declare Function CreateDIBPatternBrushPt Lib "gdi32.dll" (ByRef lpPackedDIB As Any, ByVal iUsage As Long) As Long

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

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(9) As Long
End Type

Private Const DIB_RGB_COLORS As Long = 0
Private Const BI_RGB As Long = 0&

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Private Sub PatternBlit()
 Dim rc As RECT
 Dim HBrush As Long
 Dim Result As Long
 Dim BkColor As Long, SysColor As Long
 Dim mbi As BITMAPINFO
 Dim Color As RGBQUAD
 
 With rc
  .Left = 140
  .Top = 80
  .Right = .Left + 151
  .Bottom = .Top + 51
 End With
 
 If (Me.BackColor And &H80000000) = &H80000000 Then
 'Es un color del sistema
  SysColor = (Me.BackColor And &H7FFFFFFF)
  CopyMemory Color, GetSysColor(SysColor), 4&
  BkColor = RGB(Color.rgbRed, Color.rgbGreen, Color.rgbBlue)
  Debug.Print "BackColor &H" & Hex(BkColor) & ", is SysColor " & SysColor
 Else
  BkColor = Me.BackColor
  Debug.Print "BackColor is RGB"
 End If
 
 'Creamos el patron del brush para los botones Check sin Skin.
 With mbi
  .bmiHeader.biSize = LenB(.bmiHeader)
  .bmiHeader.biBitCount = 32
  .bmiHeader.biHeight = 2
  .bmiHeader.biWidth = 2
  .bmiHeader.biPlanes = 1
  .bmiHeader.biCompression = BI_RGB
  .bmiColors(0) = BkColor
  .bmiColors(1) = RGB(255, 255, 255)
  .bmiColors(2) = RGB(255, 255, 255)
  .bmiColors(3) = BkColor
 End With
 HBrush = CreateDIBPatternBrushPt(mbi, DIB_RGB_COLORS)
' Debug.Print "HBrush = " & HBrush
 Result = FillRect(Me.hdc, rc, HBrush)
' Debug.Print "Result = " & Result
 DeleteObject HBrush
End Sub


