VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "cButtonBar Test"
   ClientHeight    =   7785
   ClientLeft      =   2865
   ClientTop       =   330
   ClientWidth     =   9105
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   519
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   607
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDebug 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7290
      Left            =   6150
      MultiLine       =   -1  'True
      TabIndex        =   44
      Top             =   450
      Width           =   2865
   End
   Begin VB.Frame Frame2 
      Caption         =   "Skinned Button"
      Height          =   2265
      Left            =   75
      TabIndex        =   34
      Top             =   5475
      Width           =   5940
      Begin VB.OptionButton optSkinned 
         Caption         =   "Original Style (No Border):"
         Height          =   240
         Index           =   4
         Left            =   225
         TabIndex        =   15
         Top             =   1875
         Width           =   2940
      End
      Begin VB.OptionButton optSkinned 
         Caption         =   "Custom Style  (3x3 Border):"
         Height          =   240
         Index           =   3
         Left            =   225
         TabIndex        =   14
         Top             =   1500
         Width           =   2940
      End
      Begin VB.OptionButton optSkinned 
         Caption         =   "Flat Style  (3x3 Border):"
         Height          =   240
         Index           =   2
         Left            =   225
         TabIndex        =   13
         Top             =   1125
         Width           =   2940
      End
      Begin VB.OptionButton optSkinned 
         Caption         =   "W98 Style  (3x3 Border):"
         Height          =   240
         Index           =   1
         Left            =   225
         TabIndex        =   12
         Top             =   750
         Width           =   2940
      End
      Begin VB.OptionButton optSkinned 
         Caption         =   "XP Style  (3x3 Border):"
         Height          =   240
         Index           =   0
         Left            =   225
         TabIndex        =   11
         Top             =   375
         Width           =   2940
      End
      Begin VB.Image imgXP 
         Height          =   315
         Left            =   3450
         Picture         =   "frmTest.frx":0000
         Top             =   300
         Width           =   2250
      End
      Begin VB.Image imgW98 
         Height          =   240
         Left            =   3450
         Picture         =   "frmTest.frx":2556
         Top             =   675
         Width           =   1200
      End
      Begin VB.Image imgFlat 
         Height          =   300
         Left            =   3450
         Picture         =   "frmTest.frx":3498
         Top             =   1050
         Width           =   2250
      End
      Begin VB.Image imgSolid 
         Height          =   300
         Left            =   3450
         Picture         =   "frmTest.frx":582A
         Top             =   1425
         Width           =   2250
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "New Button"
      Height          =   3240
      Left            =   75
      TabIndex        =   24
      Top             =   2175
      Width           =   5940
      Begin VB.CheckBox chkText 
         Caption         =   "Left"
         Height          =   240
         Index           =   0
         Left            =   2325
         TabIndex        =   43
         Top             =   750
         Width           =   1290
      End
      Begin VB.CheckBox chkText 
         Caption         =   "Right"
         Height          =   240
         Index           =   1
         Left            =   2325
         TabIndex        =   42
         Top             =   1050
         Value           =   1  'Checked
         Width           =   1290
      End
      Begin VB.CheckBox chkText 
         Caption         =   "HCenter"
         Height          =   240
         Index           =   2
         Left            =   2325
         TabIndex        =   41
         Top             =   1350
         Width           =   1290
      End
      Begin VB.CheckBox chkText 
         Caption         =   "Top"
         Height          =   240
         Index           =   3
         Left            =   2325
         TabIndex        =   40
         Top             =   1650
         Width           =   1290
      End
      Begin VB.CheckBox chkText 
         Caption         =   "Bottom"
         Height          =   240
         Index           =   4
         Left            =   2325
         TabIndex        =   39
         Top             =   1950
         Width           =   1290
      End
      Begin VB.CheckBox chkText 
         Caption         =   "VCenter"
         Height          =   240
         Index           =   5
         Left            =   2325
         TabIndex        =   38
         Top             =   2250
         Value           =   1  'Checked
         Width           =   1290
      End
      Begin VB.CommandButton btnMove 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   3375
         TabIndex        =   36
         Top             =   2700
         Width           =   390
      End
      Begin VB.CommandButton btnMove 
         Caption         =   "<<"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   3000
         TabIndex        =   35
         Top             =   2700
         Width           =   390
      End
      Begin VB.CommandButton btnAdd 
         Caption         =   "Add Button"
         Height          =   465
         Left            =   4425
         TabIndex        =   10
         Top             =   2700
         Width           =   1365
      End
      Begin VB.CheckBox chkIcon 
         Caption         =   "VCenter"
         Height          =   240
         Index           =   5
         Left            =   4350
         TabIndex        =   9
         Top             =   2250
         Value           =   1  'Checked
         Width           =   1290
      End
      Begin VB.CheckBox chkIcon 
         Caption         =   "Bottom"
         Height          =   240
         Index           =   4
         Left            =   4350
         TabIndex        =   8
         Top             =   1950
         Width           =   1290
      End
      Begin VB.CheckBox chkIcon 
         Caption         =   "Top"
         Height          =   240
         Index           =   3
         Left            =   4350
         TabIndex        =   7
         Top             =   1650
         Width           =   1290
      End
      Begin VB.CheckBox chkIcon 
         Caption         =   "HCenter"
         Height          =   240
         Index           =   2
         Left            =   4350
         TabIndex        =   6
         Top             =   1350
         Width           =   1290
      End
      Begin VB.CheckBox chkIcon 
         Caption         =   "Right"
         Height          =   240
         Index           =   1
         Left            =   4350
         TabIndex        =   5
         Top             =   1050
         Width           =   1290
      End
      Begin VB.CheckBox chkIcon 
         Caption         =   "Left"
         Height          =   240
         Index           =   0
         Left            =   4350
         TabIndex        =   4
         Top             =   750
         Value           =   1  'Checked
         Width           =   1290
      End
      Begin VB.ComboBox cboStyle 
         Height          =   315
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1950
         Width           =   1665
      End
      Begin VB.TextBox txtText 
         Height          =   315
         Left            =   150
         TabIndex        =   1
         Text            =   "Boton 4"
         ToolTipText     =   "Texto del Botón"
         Top             =   600
         Width           =   1665
      End
      Begin VB.ComboBox cboIcon 
         Height          =   315
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1275
         Width           =   1665
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Move the Button with Text ""Button 0""."
         Height          =   195
         Left            =   150
         TabIndex        =   37
         Top             =   2775
         Width           =   2715
      End
      Begin VB.Label Label7 
         Caption         =   "Text Alignment"
         Height          =   240
         Left            =   2250
         TabIndex        =   29
         Top             =   375
         Width           =   1740
      End
      Begin VB.Label Label6 
         Caption         =   "Icon Alignment"
         Height          =   240
         Left            =   4200
         TabIndex        =   28
         Top             =   375
         Width           =   1515
      End
      Begin VB.Label Label5 
         Caption         =   "Button Style"
         Height          =   240
         Left            =   150
         TabIndex        =   27
         Top             =   1725
         Width           =   1665
      End
      Begin VB.Label Label4 
         Caption         =   "Button Text"
         Height          =   240
         Left            =   150
         TabIndex        =   26
         Top             =   375
         Width           =   1665
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Button Icon"
         Height          =   240
         Left            =   150
         TabIndex        =   25
         Top             =   1050
         Width           =   1665
      End
   End
   Begin VB.CommandButton BtnRefresh 
      Caption         =   "Refresh"
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
      Left            =   4725
      TabIndex        =   0
      ToolTipText     =   "Set Focus"
      Top             =   1425
      Width           =   1215
   End
   Begin VB.PictureBox picButtonBar 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   75
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   392
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   150
      Width           =   5940
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Separator Bitmap"
      Height          =   195
      Left            =   2775
      TabIndex        =   46
      Top             =   1125
      Width           =   1215
   End
   Begin VB.Image imgSep 
      Height          =   855
      Left            =   3300
      Picture         =   "frmTest.frx":684C
      Top             =   1350
      Width           =   240
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C00000&
      Caption         =   "Debug Window"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6150
      TabIndex        =   45
      Top             =   150
      Width           =   2820
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Borde para SolidButton: 3 pixeles x 3 pixeles"
      Height          =   195
      Left            =   75
      TabIndex        =   33
      Top             =   6750
      Width           =   3120
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Borde para FlatButton: 3 pixeles x 3 pixeles"
      Height          =   195
      Left            =   75
      TabIndex        =   32
      Top             =   6375
      Width           =   3030
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Borde para W98Button: 3 pixeles x 3 pixeles"
      Height          =   195
      Left            =   75
      TabIndex        =   31
      Top             =   6000
      Width           =   3120
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Borde para XPButton: 3 pixeles x 3 pixeles"
      Height          =   195
      Left            =   75
      TabIndex        =   30
      Top             =   5625
      Width           =   2985
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Icons"
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
      Left            =   75
      TabIndex        =   23
      Top             =   1275
      Width           =   1665
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
      Left            =   1575
      TabIndex        =   22
      Top             =   1500
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
      Left            =   1275
      TabIndex        =   21
      Top             =   1500
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
      Left            =   975
      TabIndex        =   20
      Top             =   1500
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
      Left            =   675
      TabIndex        =   19
      Top             =   1500
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
      Left            =   375
      TabIndex        =   18
      Top             =   1500
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
      Index           =   0
      Left            =   75
      TabIndex        =   17
      Top             =   1500
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   225
      Index           =   5
      Left            =   1575
      Picture         =   "frmTest.frx":733E
      Top             =   1725
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   4
      Left            =   1275
      Picture         =   "frmTest.frx":7650
      Top             =   1725
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   225
      Index           =   3
      Left            =   975
      Picture         =   "frmTest.frx":7992
      Top             =   1725
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   225
      Index           =   2
      Left            =   675
      Picture         =   "frmTest.frx":7CA4
      Top             =   1725
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   1
      Left            =   375
      Picture         =   "frmTest.frx":7FB6
      Top             =   1725
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   225
      Index           =   0
      Left            =   75
      Picture         =   "frmTest.frx":82F8
      Top             =   1725
      Width           =   240
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents m_Bar As cButtonBar
Attribute m_Bar.VB_VarHelpID = -1

'Clicks counter for chkText controls. I use this
'because when the Value Property is modified a Click
'event is Raised..
Dim ContClicksText As Long
'The same for chkIcon controls.
Dim ContClicksIcon As Long

Private Sub btnAdd_Click()
 Dim m_IconAlign As eAlignment
 Dim m_TextAlign As eAlignment
 
 If chkText(0) = 1 Then m_TextAlign = eLeft
 If chkText(1) = 1 Then m_TextAlign = eRight
 If chkText(2) = 1 Then m_TextAlign = eHCenter
 If chkText(3) = 1 Then m_TextAlign = m_TextAlign Or eTop
 If chkText(4) = 1 Then m_TextAlign = m_TextAlign Or eBottom
 If chkText(5) = 1 Then m_TextAlign = m_TextAlign Or eVCenter
 
 If chkIcon(0) = 1 Then m_IconAlign = eLeft
 If chkIcon(1) = 1 Then m_IconAlign = eRight
 If chkIcon(2) = 1 Then m_IconAlign = eHCenter
 If chkIcon(3) = 1 Then m_IconAlign = m_IconAlign Or eTop
 If chkIcon(4) = 1 Then m_IconAlign = m_IconAlign Or eBottom
 If chkIcon(5) = 1 Then m_IconAlign = m_IconAlign Or eVCenter
 
 m_Bar.AddButton eButton, txtText.Text, cboStyle.ListIndex + 1, txtText.Text, Image1(cboIcon.ListIndex).Picture, m_TextAlign, m_IconAlign, eNormal
End Sub

Private Sub btnMove_Click(Index As Integer)
 Static ThisButton As Long
 
 If Index = 0 Then
  Debug.Print "Moving to the left"
  m_Bar.MoveButton ThisButton, ThisButton - 1
  If btnMove(1).Enabled = False Then btnMove(1).Enabled = True
  If ThisButton > 0 Then ThisButton = ThisButton - 1
  If ThisButton = 0 Then btnMove(0).Enabled = False
 Else
  Debug.Print "Moving to the right"
  m_Bar.MoveButton ThisButton, ThisButton + 1
  If btnMove(0).Enabled = False Then btnMove(0).Enabled = True
  If ThisButton < m_Bar.nButtons - 1 Then ThisButton = ThisButton + 1
  If ThisButton = m_Bar.nButtons - 1 Then btnMove(1).Enabled = False
 End If
End Sub

Private Sub chkIcon_Click(Index As Integer)
 
 ContClicksIcon = ContClicksIcon + 1
 CheckIconAlign Index
End Sub

Private Sub chkText_Click(Index As Integer)
 ContClicksText = ContClicksText + 1
 CheckTextAlign Index
End Sub

Private Sub BtnRefresh_Click()
 m_Bar.Refresh
End Sub

Private Sub Form_Load()
 Dim i As Long
 
 Set m_Bar = New cButtonBar
 m_Bar.Create picButtonBar
 m_Bar.MaskColor = RGB(255, 0, 255)
' m_Bar.DefaultButtonHeight = 40
' m_Bar.DefaultButtonWidth = 90
 For i = 0 To 5
  cboIcon.AddItem "Icon " & i
 Next i
 cboIcon.ListIndex = 0
 cboStyle.AddItem "Flat"
 cboStyle.AddItem "Hot"
 cboStyle.AddItem "3D"
 cboStyle.AddItem "OwnerDrawn"
 cboStyle.AddItem "Skinned"
 cboStyle.ListIndex = 0
 
' Debug.Print "AddButton 0 and SetButtonSize"
 m_Bar.AddButton eButton, "Button 0", eFlat, "Button 0", Image1(0).Picture, eBottom Or eHCenter, eTop Or eHCenter, eNormal
 m_Bar.SetButtonSize 0, 70, 50
' Debug.Print "AddButton 1 and SetButtonSize"
 m_Bar.AddButton eButton, "Button 1", e3D, "Button 1", Image1(1).Picture, eLeft Or eVCenter, eRight Or eVCenter, eNormal
 m_Bar.SetButtonSize 1, 80, 30
' Debug.Print "AddButton 2 and SetButtonSize"
 m_Bar.AddButton eButton, "Button 2", eHot, "Button 2", Image1(4).Picture, eRight Or eVCenter, eLeft Or eVCenter, eNormal
 m_Bar.SetButtonSize 2, 80, 30
' Debug.Print "AddButton 3 and SetButtonSize"
 m_Bar.AddButton eCheck, "Button 3", eFlat, "Button 3", Image1(2).Picture, eTop Or eHCenter, eBottom Or eHCenter, eNormal
 m_Bar.SetButtonSize 3, 70, 50
' Debug.Print "AddButton 4 and SetButtonSize"
 m_Bar.AddButton eButton, "", eHot, "Button 4", Image1(5).Picture, eLeft Or eBottom, eHCenter Or eVCenter, eNormal
 m_Bar.SetButtonSize 4, 24, 24
 m_Bar.AddButton eSeparator, "", e3D, "5", imgSep.Picture, eBottom Or eHCenter, eHCenter Or eTop, eNormal
 m_Bar.SetButtonSize 5, 10, 57
' Debug.Print "AddButton 5 and SetButtonSize"
 m_Bar.AddButton eButton, "5", e3D, "5", Image1(2).Picture, eBottom Or eHCenter, eHCenter Or eTop, eNormal
 m_Bar.SetButtonSize 6, 24, 50
 
' m_Bar.SetButtonTextColor RGB(128, 128, 128), RGB(192, 192, 192), eDisabled
 m_Bar.SetButtonTextColor RGB(128, 128, 128), m_Bar.BackColor, eDisabled
' For i = 0 To 5
'  m_Bar.SetButtonBackColor RGB(192, 192, 192), RGB(225, 225, 225), i
' Next i
' m_Bar.FillMode = eHorizontalGradient
 txtText.Text = "Button " & m_Bar.nButtons
End Sub

Private Sub Form_Unload(Cancel As Integer)
 m_Bar.Destroy
End Sub

Private Sub m_Bar_Click(Index As Long)
 DebugPrint "Button " & Index & "generated Click event"
End Sub

Private Sub m_Bar_GotFocus(Index As Long)
 DebugPrint "Button " & Index & " generated GotFocus event"
End Sub

Private Sub m_Bar_KeyDown(Index As Long, KeyCode As Integer, Shift As Integer)
 DebugPrint "Button " & Index & " generated KeyDown event"
End Sub

Private Sub m_Bar_KeyPressed(Index As Long, KeyAscii As Integer)
 DebugPrint "Button " & Index & " generated KeyPressed event"
End Sub

Private Sub m_Bar_KeyUp(Index As Long, KeyCode As Integer, Shift As Integer)
 DebugPrint "Button " & Index & " generated KeyUp event"
End Sub

Private Sub m_Bar_LostFocus(Index As Long, Desc As String)
 DebugPrint "Button " & Index & " generated LostFocus event in " & Desc
End Sub

Private Sub m_Bar_MouseDown(Index As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
 DebugPrint "Button " & Index & " generated MouseDown event"
End Sub

Private Sub m_Bar_MouseEnter(Index As Long)
 DebugPrint "Button " & Index & " generated MouseEnter event"
End Sub

Private Sub m_Bar_MouseLeave(Index As Long)
 DebugPrint "Button " & Index & " generated MouseLeave event"
End Sub

Private Sub m_Bar_MouseUp(Index As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
 DebugPrint "Button " & Index & " generated MouseUp event"
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

 With m_Bar
  .LockUpdate = True
  .DrawFocusRect = True
  .SetSkin Nothing
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
 
 With m_Bar
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
  For i = -2 To m_Bar.nButtons - 1
   .ButtonStyle(i) = eSkinned
  Next i
  .SetSkin Img
  .LockUpdate = False
 End With
End Sub

Private Sub CheckTextAlign(ByVal Index As Integer)
 Dim i As Long
 Dim SetValue As Long, RemoveValue As Long
 
 If ContClicksText > 1 Then
  GoTo ErrHandler 'Exit Sub
 End If
 If chkText(Index).Value = 0 Then
  'There must be only one horizontal value selected and
  'one vertical value selected for alignment.
  chkText(Index).Value = 1
  GoTo ErrHandler
 Else
  SetValue = 0
  RemoveValue = 1
 End If
 
 If Index < 3 Then
  For i = 0 To 2
   chkText(i).Value = SetValue
  Next i
  chkText(Index).Value = RemoveValue
 Else
  For i = 3 To 5
   chkText(i).Value = SetValue
  Next i
  chkText(Index).Value = RemoveValue
 End If
 
ErrHandler:
 ContClicksText = ContClicksText - 1
End Sub

Private Sub CheckIconAlign(ByVal Index As Integer)
 Dim i As Long
 Dim SetValue As Long, RemoveValue As Long
 
 If ContClicksIcon > 1 Then
  GoTo ErrHandler 'Exit Sub
 End If
 If chkIcon(Index).Value = 0 Then
  'There must be only one horizontal value selected and
  'one vertical value selected for alignment.
  chkIcon(Index).Value = 1
  GoTo ErrHandler
 Else
  SetValue = 0
  RemoveValue = 1
 End If
 
 If Index < 3 Then
  For i = 0 To 2
   chkIcon(i).Value = SetValue
  Next i
  chkIcon(Index).Value = RemoveValue
 Else
  For i = 3 To 5
   chkIcon(i).Value = SetValue
  Next i
  chkIcon(Index).Value = RemoveValue
 End If
 
ErrHandler:
 ContClicksIcon = ContClicksIcon - 1
End Sub

Private Sub DebugPrint(TextStr As String)
 txtDebug.Text = txtDebug.Text & TextStr & vbCrLf
End Sub
