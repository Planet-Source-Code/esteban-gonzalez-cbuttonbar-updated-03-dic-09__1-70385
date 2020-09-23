VERSION 5.00
Begin VB.Form frmTest2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test de otra skin"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   391
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   460
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picContents 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4890
      Index           =   0
      Left            =   75
      ScaleHeight     =   322
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   447
      TabIndex        =   1
      Top             =   900
      Width           =   6765
      Begin VB.TextBox txtIntro 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4740
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   32
         Top             =   0
         Width           =   6690
      End
   End
   Begin VB.CheckBox chkSkinnedTabs 
      Caption         =   "Usar skins en las pestañas"
      Height          =   240
      Left            =   75
      TabIndex        =   38
      Top             =   75
      Width           =   2790
   End
   Begin VB.PictureBox picTabs 
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   75
      ScaleHeight     =   450
      ScaleWidth      =   6765
      TabIndex        =   0
      Top             =   450
      Width           =   6765
   End
   Begin VB.PictureBox picContents 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4890
      Index           =   3
      Left            =   75
      ScaleHeight     =   322
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   447
      TabIndex        =   33
      Top             =   900
      Width           =   6765
      Begin VB.PictureBox picButtons2 
         Height          =   840
         Left            =   150
         ScaleHeight     =   780
         ScaleWidth      =   6405
         TabIndex        =   35
         Top             =   1875
         Width           =   6465
      End
      Begin VB.PictureBox picButtons1 
         Height          =   840
         Left            =   150
         ScaleHeight     =   780
         ScaleWidth      =   2580
         TabIndex        =   34
         Top             =   375
         Width           =   2640
      End
      Begin VB.Image imgXP2 
         Height          =   315
         Left            =   4125
         Picture         =   "frmTest2.frx":0000
         Top             =   600
         Width           =   2250
      End
      Begin VB.Image imgButtonSetSkin 
         Height          =   360
         Index           =   2
         Left            =   150
         Picture         =   "frmTest2.frx":2556
         Top             =   4275
         Width           =   7500
      End
      Begin VB.Image imgButtonSetSkin 
         Height          =   360
         Index           =   1
         Left            =   150
         Picture         =   "frmTest2.frx":B238
         Top             =   3525
         Width           =   7500
      End
      Begin VB.Image imgButtonSetSkin 
         Height          =   360
         Index           =   0
         Left            =   150
         Picture         =   "frmTest2.frx":13F1A
         Top             =   2775
         Width           =   7500
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo de Botones con skins particulares"
         Height          =   195
         Left            =   150
         TabIndex        =   37
         Top             =   1575
         Width           =   2865
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo de Botones con skin General"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   36
         Top             =   75
         Width           =   2535
      End
   End
   Begin VB.PictureBox picContents 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4890
      Index           =   2
      Left            =   75
      ScaleHeight     =   322
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   447
      TabIndex        =   3
      Top             =   900
      Width           =   6765
      Begin VB.ComboBox cboBotones 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2850
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   3300
         Width           =   3240
      End
      Begin VB.CommandButton btnEliminar 
         Caption         =   "Eliminar Botón"
         Height          =   465
         Left            =   5175
         TabIndex        =   21
         Top             =   4200
         Width           =   1365
      End
      Begin VB.CommandButton btnModificar 
         Caption         =   "Modificar Botón"
         Height          =   465
         Left            =   3825
         TabIndex        =   20
         Top             =   4200
         Width           =   1365
      End
      Begin VB.PictureBox picToolbarLarge 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Left            =   45
         ScaleHeight     =   930
         ScaleWidth      =   6540
         TabIndex        =   18
         Top             =   1350
         Width           =   6600
      End
      Begin VB.ComboBox cboIcono 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmTest2.frx":1CBFC
         Left            =   150
         List            =   "frmTest2.frx":1CC15
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   4425
         Width           =   1665
      End
      Begin VB.TextBox txtTexto 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         TabIndex        =   14
         Text            =   "Nuevo Botón"
         ToolTipText     =   "Texto del Botón"
         Top             =   3750
         Width           =   1665
      End
      Begin VB.CommandButton btnAgregar 
         Caption         =   "Agregar Botón"
         Height          =   465
         Left            =   2475
         TabIndex        =   13
         Top             =   4200
         Width           =   1365
      End
      Begin VB.PictureBox picToolbarSmall 
         Height          =   540
         Left            =   45
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   437
         TabIndex        =   4
         Top             =   375
         Width           =   6615
      End
      Begin VB.Image imgSepLarge 
         Height          =   720
         Left            =   2475
         Picture         =   "frmTest2.frx":1CC5A
         Top             =   2850
         Width           =   150
      End
      Begin VB.Image imgSepSmall 
         Height          =   360
         Left            =   2100
         Picture         =   "frmTest2.frx":1D29C
         Top             =   3000
         Width           =   150
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccione un botón para modificar o eliminar"
         Height          =   195
         Left            =   2850
         TabIndex        =   23
         Top             =   3000
         Width           =   3210
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Barra de Herramientas con botones grandes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   75
         TabIndex        =   19
         Top             =   1050
         Width           =   3180
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Icono del Botón"
         Height          =   240
         Left            =   150
         TabIndex        =   17
         Top             =   4200
         Width           =   1665
      End
      Begin VB.Label Label5 
         Caption         =   "Texto del Botón"
         Height          =   240
         Left            =   150
         TabIndex        =   15
         Top             =   3525
         Width           =   1665
      End
      Begin VB.Image ImgToolbar 
         Height          =   225
         Index           =   0
         Left            =   225
         Picture         =   "frmTest2.frx":1D5DE
         Top             =   3075
         Width           =   240
      End
      Begin VB.Image ImgToolbar 
         Height          =   240
         Index           =   1
         Left            =   525
         Picture         =   "frmTest2.frx":1D8F0
         Top             =   3075
         Width           =   240
      End
      Begin VB.Image ImgToolbar 
         Height          =   225
         Index           =   2
         Left            =   825
         Picture         =   "frmTest2.frx":1DC32
         Top             =   3075
         Width           =   240
      End
      Begin VB.Image ImgToolbar 
         Height          =   225
         Index           =   3
         Left            =   1125
         Picture         =   "frmTest2.frx":1DF44
         Top             =   3075
         Width           =   240
      End
      Begin VB.Image ImgToolbar 
         Height          =   240
         Index           =   4
         Left            =   1425
         Picture         =   "frmTest2.frx":1E256
         Top             =   3075
         Width           =   240
      End
      Begin VB.Image ImgToolbar 
         Height          =   225
         Index           =   5
         Left            =   1725
         Picture         =   "frmTest2.frx":1E598
         Top             =   3075
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   6
         Left            =   225
         TabIndex        =   12
         Top             =   2850
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   1
         Left            =   525
         TabIndex        =   11
         Top             =   2850
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   2
         Left            =   825
         TabIndex        =   10
         Top             =   2850
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   3
         Left            =   1125
         TabIndex        =   9
         Top             =   2850
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   4
         Left            =   1425
         TabIndex        =   8
         Top             =   2850
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   5
         Left            =   1725
         TabIndex        =   7
         Top             =   2850
         Width           =   240
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Iconos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   225
         TabIndex        =   6
         Top             =   2625
         Width           =   1665
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Barra de Herramientas con botones pequeños"
         Height          =   195
         Left            =   75
         TabIndex        =   5
         Top             =   150
         Width           =   3255
      End
   End
   Begin VB.PictureBox picContents 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4890
      Index           =   1
      Left            =   75
      ScaleHeight     =   322
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   447
      TabIndex        =   2
      Top             =   900
      Width           =   6765
      Begin VB.Frame Frame2 
         Caption         =   "Uso de Skins Generales"
         Height          =   2265
         Left            =   375
         TabIndex        =   26
         Top             =   1950
         Width           =   5940
         Begin VB.OptionButton optSkinned 
            Caption         =   "Estilo XP (Borde 3x3):"
            Height          =   240
            Index           =   0
            Left            =   225
            TabIndex        =   31
            Top             =   375
            Width           =   2940
         End
         Begin VB.OptionButton optSkinned 
            Caption         =   "Estilo W98 (Borde 3x3):"
            Height          =   240
            Index           =   1
            Left            =   225
            TabIndex        =   30
            Top             =   750
            Width           =   2940
         End
         Begin VB.OptionButton optSkinned 
            Caption         =   "Estilo Flat  (Borde 3x3):"
            Height          =   240
            Index           =   2
            Left            =   225
            TabIndex        =   29
            Top             =   1125
            Width           =   2940
         End
         Begin VB.OptionButton optSkinned 
            Caption         =   "Estilo Custom  (Borde 3x3):"
            Height          =   240
            Index           =   3
            Left            =   225
            TabIndex        =   28
            Top             =   1500
            Width           =   2940
         End
         Begin VB.OptionButton optSkinned 
            Caption         =   "Estilo Original (Sin Skin):"
            Height          =   240
            Index           =   4
            Left            =   225
            TabIndex        =   27
            Top             =   1875
            Width           =   2940
         End
         Begin VB.Image imgSolid 
            Height          =   300
            Left            =   3450
            Picture         =   "frmTest2.frx":1E8AA
            Top             =   1425
            Width           =   2250
         End
         Begin VB.Image imgFlat 
            Height          =   300
            Left            =   3450
            Picture         =   "frmTest2.frx":1F8CC
            Top             =   1050
            Width           =   2250
         End
         Begin VB.Image imgW98 
            Height          =   240
            Left            =   3450
            Picture         =   "frmTest2.frx":21C5E
            Top             =   675
            Width           =   1200
         End
         Begin VB.Image imgXP 
            Height          =   315
            Left            =   3450
            Picture         =   "frmTest2.frx":22BA0
            Top             =   300
            Width           =   2250
         End
      End
      Begin VB.PictureBox picButtonBar 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Left            =   75
         ScaleHeight     =   62
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   432
         TabIndex        =   24
         Top             =   375
         Width           =   6540
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Diferentes Tipos de botones disponibles"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   75
         TabIndex        =   25
         Top             =   150
         Width           =   2850
      End
   End
   Begin VB.Image imgTabs 
      Height          =   420
      Left            =   150
      Picture         =   "frmTest2.frx":250F6
      Top             =   450
      Width           =   3825
   End
End
Attribute VB_Name = "frmTest2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function DrawTextEx Lib "user32.dll" Alias "DrawTextExA" (ByVal hDC As Long, ByVal lpsz As String, ByVal n As Long, ByRef lpRect As RECT, ByVal un As Long, ByRef lpDrawTextParams As Long) As Long
Private Const DT_CENTER As Long = &H1
Private Const DT_CALCRECT As Long = &H400
Private Const DT_MULTILINE As Long = (&H1)
Private Const DT_VCENTER As Long = &H4

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Dim WithEvents m_Tabs As cButtonBar
Attribute m_Tabs.VB_VarHelpID = -1
Dim WithEvents m_ButtonBar As cButtonBar
Attribute m_ButtonBar.VB_VarHelpID = -1
Dim WithEvents m_ToolbarSmall As cButtonBar
Attribute m_ToolbarSmall.VB_VarHelpID = -1
Dim WithEvents m_ToolbarLarge As cButtonBar
Attribute m_ToolbarLarge.VB_VarHelpID = -1
Dim WithEvents m_Buttons1 As cButtonBar
Attribute m_Buttons1.VB_VarHelpID = -1
Dim WithEvents m_Buttons2 As cButtonBar
Attribute m_Buttons2.VB_VarHelpID = -1

Private Sub btnAgregar_Click()
 Dim IconIndex As Long
 Dim BtnIcon As StdPicture
 
 If cboIcono.ListIndex >= 0 And cboIcono.ListIndex < 6 Then
  Set BtnIcon = ImgToolbar(cboIcono.ListIndex).Picture
 End If
 If cboIcono.ListIndex = 6 Then
  m_ToolbarSmall.AddButton eSeparator, "", eFlat, , imgSepSmall.Picture, , eHCenter Or eVCenter, eNormal
  m_ToolbarLarge.AddButton eSeparator, "", eHot, , imgSepLarge.Picture, eBottom Or eHCenter, eHCenter Or eTop, eNormal
  cboBotones.AddItem "Separator"
 Else
  m_ToolbarSmall.AddButton eButton, "", eFlat, txtTexto.Text, BtnIcon, , eHCenter Or eVCenter, eNormal
  m_ToolbarLarge.AddButton eButton, txtTexto.Text, eHot, txtTexto.Text, BtnIcon, eBottom Or eHCenter, eHCenter Or eTop, eNormal
  If txtTexto.Text <> "" Then
   cboBotones.AddItem txtTexto.Text
  Else
   cboBotones.AddItem "Boton sin Texto"
  End If
 End If
End Sub

Private Sub btnEliminar_Click()

 If cboBotones.ListIndex < 0 Then Exit Sub
 m_ToolbarSmall.RemoveButton cboBotones.ListIndex
 m_ToolbarLarge.RemoveButton cboBotones.ListIndex
 cboBotones.RemoveItem cboBotones.ListIndex
End Sub

Private Sub btnModificar_Click()
 
 If cboBotones.ListIndex >= 0 Then
  If cboIcono.ListIndex >= 0 Then
   Set m_ToolbarSmall.ButtonIcon(cboBotones.ListIndex) = ImgToolbar(cboIcono.ListIndex).Picture
   Set m_ToolbarLarge.ButtonIcon(cboBotones.ListIndex) = ImgToolbar(cboIcono.ListIndex).Picture
  End If
  If txtTexto.Text <> "" Then
   m_ToolbarLarge.ButtonText(cboBotones.ListIndex) = txtTexto.Text
   m_ToolbarLarge.ButtonTooltip(cboBotones.ListIndex) = txtTexto.Text
   m_ToolbarSmall.ButtonTooltip(cboBotones.ListIndex) = txtTexto.Text
  End If
  cboIcono.ListIndex = -1
  cboBotones.ListIndex = -1
 End If
End Sub

Private Sub chkSkinnedTabs_Click()
 Dim i As Long
 
 If chkSkinnedTabs.Value = 1 Then
  With m_Tabs
   .LockUpdate = True
   .SetSkin imgTabs.Picture
   .FillMode = eStretchBlit
   For i = -2 To .nButtons - 1
    .ButtonStyle(i) = eSkinned
   Next i
   .LockUpdate = False
  End With
 Else
  With m_Tabs
   .LockUpdate = True
   .SetSkin Nothing
   .FillMode = eSolid
   .BorderWidth = 2
   For i = -2 To .nButtons - 1
    .ButtonStyle(i) = eHot
   Next i
   .LockUpdate = False
  End With
 End If
End Sub

Private Sub Form_Load()
 picContents(0).ZOrder 0
 CreateTabs
 CreateButtonBar
 CreateToolbar_Small
 CreateToolbar_Large
 CreateButtonsSet
 LoadIntro
End Sub

Private Sub Form_Unload(Cancel As Integer)
 m_Tabs.Destroy
 m_ButtonBar.Destroy
 m_ToolbarSmall.Destroy
 m_ToolbarLarge.Destroy
 m_Buttons1.Destroy
 m_Buttons2.Destroy
End Sub

Private Sub m_Tabs_Click(Index As Long)
 If Index < 0 Then Exit Sub
 picContents(Index).ZOrder 0
End Sub

Private Sub CreateTabs()
 Dim i As Long
 
 Set m_Tabs = New cButtonBar
 With m_Tabs
  .Create picTabs
  .LockUpdate = True
  .DefaultButtonHeight = 28
'  .DefaultButtonWidth = 80
  .MaskColor = RGB(255, 0, 255)
  For i = 0 To 4
   .SkinBorderHeight(i) = 5
   .SkinBorderWidth(i) = 5
  Next i
  'Aca lo utilizamos para que coincida con los bordes
  'de la skin.
  .BorderWidth = 5
  .DrawFocusRect = True
  .HorizontalButtonGap = 0
  .VerticalButtonGap = 1
  .AutoSizeButtons = True
  
  .FillMode = eSolid
  .AddButton eOption, "Introducción", eHot, "Introducción", , eHCenter Or eVCenter, , eDown
  .AddButton eOption, "Barra de Botones", eHot, "Barra de Botones", , eHCenter Or eVCenter, , eNormal
  .AddButton eOption, "Barra de Herramientas", eHot, "Barra de Herramientas", , eHCenter Or eVCenter, , eNormal
  .AddButton eOption, "Conjunto de Botones", eHot, "Conj. de Botones", , eHCenter Or eVCenter, , eNormal
  .ButtonStyle(-2) = eSkinned
  .ButtonStyle(-1) = eSkinned
  
  .BorderWidth = 2
  .LockUpdate = False
 End With
End Sub

Private Sub CreateButtonBar()
 Dim i As Long
 
 Set m_ButtonBar = New cButtonBar
 With m_ButtonBar
  .Create picButtonBar
  .LockUpdate = True
  .MaskColor = RGB(255, 0, 255)
  .BorderWidth = 3
  .AddButton eButton, "Boton 0", eFlat, "Boton 0", ImgToolbar(0).Picture, eBottom Or eHCenter, eTop Or eHCenter, eNormal
  .SetButtonSize 0, 70, 50
  .AddButton eButton, "Boton 1", e3D, "Boton 1", ImgToolbar(1).Picture, eLeft Or eVCenter, eRight Or eVCenter, eNormal
  .SetButtonSize 1, 80, 30
  .AddButton eButton, "Boton 2", eHot, "Boton 2", ImgToolbar(4).Picture, eRight Or eVCenter, eLeft Or eVCenter, eNormal
  .SetButtonSize 2, 80, 30
  .AddButton eCheck, "Boton 3", eFlat, "Boton 3", ImgToolbar(2).Picture, eTop Or eHCenter, eBottom Or eHCenter, eNormal
  .SetButtonSize 3, 70, 50
  .AddButton eButton, "", eHot, "Boton 4", ImgToolbar(5).Picture, eLeft Or eBottom, eHCenter Or eVCenter, eDisabled
  .SetButtonSize 4, 24, 24
  .AddButton eButton, "5", e3D, "5", ImgToolbar(2).Picture, eBottom Or eHCenter, eHCenter Or eTop, eNormal
  .SetButtonSize 5, 24, 50
  .SetButtonTextColor RGB(128, 128, 128), .BackColor, eDisabled
  .LockUpdate = False
 End With
End Sub

Private Sub optSkinned_Click(Index As Integer)
 Select Case Index
  Case 0 'XP
'   Debug.Print "Skin XP"
   SetSkinnedButtons imgXP.Picture
  Case 1 'W98
'   Debug.Print "Skin W98"
   SetSkinnedButtons imgW98.Picture
  Case 2 'FLAT
'   Debug.Print "Skin Flat"
   SetSkinnedButtons imgFlat.Picture
  Case 3 'CUSTOM
'   Debug.Print "Skin Solid"
   SetSkinnedButtons imgSolid.Picture
  Case 4 'RESTORE
'   Debug.Print "Restore Buttons"
   RestoreButtons
 End Select
End Sub

Private Sub RestoreButtons()

 With m_ButtonBar
  .LockUpdate = True
  .DrawFocusRect = True
  .SetSkin Nothing
  .BorderWidth = 3
  .ButtonStyle(0) = eFlat
  .ButtonStyle(1) = e3D
  .ButtonStyle(2) = eHot
  .ButtonStyle(3) = eFlat
  .ButtonStyle(4) = eFlat
  .ButtonStyle(5) = e3D
  .ButtonStyle(-1) = e3D
  .ButtonStyle(-2) = e3D
  .FillMode = ePatternFill '= eHorizontalGradient
  .LockUpdate = False
 End With
End Sub

Private Sub SetSkinnedButtons(Img As StdPicture)
 Dim i As Long
 
 With m_ButtonBar
  .LockUpdate = True
  If Img = imgFlat.Picture Then
   .FillMode = ePatternFill
  Else
   .FillMode = eStretchBlit
  End If
  .DrawFocusRect = (Img <> imgXP.Picture)
  .MaskColor = RGB(255, 0, 255)
  For i = 0 To 4
   .SkinBorderHeight(i) = 3
   .SkinBorderWidth(i) = 3
  Next i
  For i = -2 To .nButtons - 1
   .ButtonStyle(i) = eSkinned
  Next i
  .SetSkin Img
  .LockUpdate = False
 End With
End Sub

Private Sub LoadIntro()
 Dim m_Handle As Long
 Dim s As String
 
'Habilitamos el manejo de errores
 On Error Resume Next
 
 m_Handle = FreeFile
 Open App.Path & "\readme.txt" For Binary Access Read As #m_Handle
 If Err.Number = 0 Then
  s = Space(LOF(m_Handle)) '\ 2)
  Get #m_Handle, 1, s
  txtIntro.Text = s
 End If
 Close #m_Handle
 
End Sub

Private Sub CreateToolbar_Small()
 
 Set m_ToolbarSmall = New cButtonBar
 With m_ToolbarSmall
  .Create picToolbarSmall
  .LockUpdate = True
  .MaskColor = RGB(255, 0, 255)
  .BorderWidth = 3
  .DefaultButtonHeight = 24
  .DefaultButtonWidth = 24
  .AutoSizeButtons = False
  .DrawFocusRect = False
  .AddButton eButton, "", eFlat, "Boton 0", ImgToolbar(0).Picture, eBottom Or eHCenter, eVCenter Or eHCenter, eNormal
  .AddButton eButton, "", eFlat, "Boton 1", ImgToolbar(1).Picture, eLeft Or eVCenter, eVCenter Or eHCenter, eNormal
  .AddButton eButton, "", eFlat, "Boton 2", ImgToolbar(4).Picture, eRight Or eVCenter, eVCenter Or eHCenter, eNormal
  .AddButton eSeparator, "", eSkinned, , imgSepSmall.Picture
  .AddButton eCheck, "", eFlat, "Boton 4", ImgToolbar(2).Picture, eTop Or eHCenter, eVCenter Or eHCenter, eNormal
  .AddButton eSeparator, "", eSkinned, , imgSepSmall.Picture
  .AddButton eButton, "", eFlat, "Boton 6", ImgToolbar(5).Picture, eLeft Or eBottom, eVCenter Or eHCenter, eDisabled
  .AddButton eButton, "", eFlat, "Boton 7", ImgToolbar(2).Picture, eBottom Or eHCenter, eVCenter Or eHCenter, eNormal
  
  .SetButtonTextColor RGB(128, 128, 128), .BackColor, eDisabled
  .LockUpdate = False
 End With
 
 cboBotones.AddItem "Boton 0"
 cboBotones.AddItem "Boton 1"
 cboBotones.AddItem "Boton 2"
 cboBotones.AddItem "Separador"
 cboBotones.AddItem "Boton 4"
 cboBotones.AddItem "Separador"
 cboBotones.AddItem "Boton 6"
 cboBotones.AddItem "Boton 7"
 
End Sub

Private Sub CreateToolbar_Large()
 
 Set m_ToolbarLarge = New cButtonBar
 With m_ToolbarLarge
  .Create picToolbarLarge
  .LockUpdate = True
  .MaskColor = RGB(255, 0, 255)
  .BorderWidth = 3
  .DefaultButtonHeight = 48
  .AutoSizeButtons = True
  .AddButton eButton, "Boton 0", eHot, "Boton 0", ImgToolbar(0).Picture, eBottom Or eHCenter, eHCenter Or eTop, eNormal
  .AddButton eButton, "Boton 1", eHot, "Boton 1", ImgToolbar(1).Picture, eBottom Or eHCenter, eHCenter Or eTop, eNormal
  .AddButton eButton, "Boton 2", eHot, "Boton 2", ImgToolbar(4).Picture, eBottom Or eHCenter, eHCenter Or eTop, eNormal
  .AddButton eSeparator, "", eSkinned, , imgSepLarge.Picture
  .AddButton eCheck, "Boton 4", eHot, "Boton 4", ImgToolbar(2).Picture, eBottom Or eHCenter, eHCenter Or eTop, eNormal
  .AddButton eSeparator, "", eSkinned, , imgSepLarge.Picture
  .AddButton eButton, "Boton 6", eHot, "Boton 6", ImgToolbar(5).Picture, eBottom Or eHCenter, eHCenter Or eTop, eDisabled
  .AddButton eButton, "Boton 7", eHot, "Boton 7", ImgToolbar(2).Picture, eBottom Or eHCenter, eHCenter Or eTop, eNormal
  
  .SetButtonTextColor RGB(128, 128, 128), .BackColor, eDisabled
  .LockUpdate = False
 End With
End Sub

Private Sub CreateButtonsSet()
 Dim i As Long
 
 Set m_Buttons1 = New cButtonBar
 With m_Buttons1
  .Create picButtons1
  .LockUpdate = True
  .MaskColor = RGB(255, 0, 255)
'  .BorderWidth = 3
  .DefaultButtonHeight = 40
  .DefaultButtonWidth = 60
  .AutoSizeButtons = True
  For i = 0 To 4
   .SkinBorderHeight(i) = 3
   .SkinBorderWidth(i) = 3
  Next i
  .FillMode = eStretchBlit
  .AddButton eButton, "Aceptar", eSkinned, "Aceptar", ImgToolbar(0).Picture, eBottom Or eHCenter, eHCenter Or eTop, eNormal
  .AddButton eButton, "Cancelar", eSkinned, "Cancelar", ImgToolbar(1).Picture, eBottom Or eHCenter, eHCenter Or eTop, eNormal
  
  .SetSkin imgXP2.Picture
  .SetButtonTextColor RGB(128, 128, 128), .BackColor, eDisabled
  .LockUpdate = False
 End With
 
 Set m_Buttons2 = New cButtonBar
 With m_Buttons2
  .Create picButtons2
  .LockUpdate = True
  .MaskColor = RGB(255, 0, 255)
  .BorderWidth = 3
  .DefaultButtonHeight = 24
  .DefaultButtonWidth = 100
  .DrawFocusRect = False
  .AddButton eButton, "", eOwnerDrawn, "Aceptar", , eBottom Or eHCenter, eHCenter Or eTop, eNormal
  .SetButtonSkin 0, imgButtonSetSkin(0).Picture
  .AddButton eButton, "", eOwnerDrawn, "Cancelar", , eBottom Or eHCenter, eHCenter Or eTop, eNormal
  .SetButtonSkin 1, imgButtonSetSkin(1).Picture
  .AddButton eButton, "", eOwnerDrawn, "Cancelar", , eBottom Or eHCenter, eHCenter Or eTop, eNormal
  .SetButtonSkin 2, imgButtonSetSkin(2).Picture
  .LockUpdate = False
 End With

End Sub

