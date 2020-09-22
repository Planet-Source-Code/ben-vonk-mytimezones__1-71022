VERSION 5.00
Begin VB.UserControl FlatButton 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   ClientHeight    =   492
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   768
   ScaleHeight     =   41
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   64
   ToolboxBitmap   =   "FlatButton.ctx":0000
   Begin VB.PictureBox picPicture 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   252
      Left            =   120
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   492
   End
End
Attribute VB_Name = "FlatButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'FlatButton Control
'
'Author Ben Vonk
'10-09-2005 First version
'26-03-2007 Second version, fixed some bugs and add the OnlyIconClick property

Option Explicit

' Public Events
Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

' Public Enumerations
'Public Enum BackStyles
'   Transparent
'   Opaque
'End Enum

Public Enum OLEDropModes
   None
   Manual
End Enum

' Private Type
'Private Type Rect
'   Left                 As Long
'   Top                  As Long
'   Right                As Long
'   Bottom               As Long
'End Type

' Private Variables
Private m_BackStyle     As BackStyles   ' sets the flatbutton backstyle
Private HasBackImage    As Boolean      ' checked if backimage is set
Private IsClicked       As Boolean      ' checked if checkbox is clicked
Private m_OnlyIconClick As Boolean      ' moves icon only when button is clicked
Private MouseIn         As Boolean      ' checked if mouse is in the flatbutton
Private IconSize        As Integer      ' hold size of the icon
Private m_IconX         As Long         ' sets the Left of the icon
Private m_IconY         As Long         ' sets the top of the icon
Private ButtonRect      As Rect         ' holds the size of the flatbutton used for API PtInRect
Private m_Icon          As StdPicture   ' holds the button icon

' Private API's
'Private Declare Function BitBlt Lib "GDI32" (ByVal hDCDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal dwRop As Long) As Long
'Private Declare Function DrawIconEx Lib "User32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyHeight As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
'Private Declare Function PtInRect Lib "User32" (Rect As Rect, ByVal lPtX As Long, ByVal lPtY As Long) As Integer

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphcs in an object."

   BackColor = picPicture.BackColor

End Property

Public Property Let BackColor(ByVal NewBackColor As OLE_COLOR)

   picPicture.BackColor = NewBackColor
   PropertyChanged "BackColor"
   HasBackImage = False
   
   Call Refresh

End Property

Public Property Get BackStyle() As BackStyles
Attribute BackStyle.VB_Description = "Returns/sets the border style for an object."

   BackStyle = m_BackStyle

End Property

Public Property Let BackStyle(ByVal NewBackStyle As BackStyles)

   m_BackStyle = NewBackStyle
   PropertyChanged "BackStyle"
   HasBackImage = False
   
   Call Refresh

End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."

   Enabled = UserControl.Enabled

End Property

Public Property Let Enabled(ByVal NewEnabled As Boolean)

   UserControl.Enabled = NewEnabled
   PropertyChanged "Enabled"

End Property

Public Property Get Icon() As StdPicture
Attribute Icon.VB_Description = "Returns/sets a icon to be displayed in the button. Before use, first set the properties IconX and IconY!"

   Set Icon = m_Icon

End Property

Public Property Let Icon(ByVal NewIcon As StdPicture)

   Set Icon = NewIcon

End Property

Public Property Set Icon(ByVal NewIcon As StdPicture)

   Set m_Icon = NewIcon
   Set NewIcon = Nothing
   PropertyChanged "Icon"
   HasBackImage = False
   
   Call Refresh

End Property

Public Property Get IconX() As Long
Attribute IconX.VB_Description = "Returns/sets the icon X position."

   IconX = m_IconX

End Property

Public Property Let IconX(ByVal NewIconX As Long)

   m_IconX = NewIconX
   PropertyChanged "IconX"
   
   Call Refresh

End Property

Public Property Get IconY() As Long
Attribute IconY.VB_Description = "Returns/sets the icon Y position."

   IconY = m_IconY

End Property

Public Property Let IconY(ByVal NewIconY As Long)

   m_IconY = NewIconY
   PropertyChanged "IconY"
   
   Call Refresh

End Property

Public Property Get OLEDropMode() As OLEDropModes
Attribute OLEDropMode.VB_Description = "Returns/sets whether this object can act as as OLE drop target."

   OLEDropMode = UserControl.OLEDropMode

End Property

Public Property Let OLEDropMode(ByVal NewOLEDropMode As OLEDropModes)

   UserControl.OLEDropMode = NewOLEDropMode
   PropertyChanged "OLEDropMode"

End Property

Public Property Get OnlyIconClick() As Boolean
Attribute OnlyIconClick.VB_Description = "Returns/sets the click style of the control. If is set only the icon will be moved when the control is clicked."

   OnlyIconClick = m_OnlyIconClick

End Property

Public Property Let OnlyIconClick(ByVal NewOnlyIconClick As Boolean)

   m_OnlyIconClick = NewOnlyIconClick
   PropertyChanged "OnlyIconClick"
   
   Call GetButtonSize

End Property

Public Property Get Picture() As StdPicture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."

   Set Picture = Image

End Property

Public Property Let Picture(ByRef NewPicture As StdPicture)

   Set Picture = NewPicture

End Property

Public Property Set Picture(ByRef NewPicture As StdPicture)

   picPicture.Picture = NewPicture
   Set NewPicture = Nothing
   PropertyChanged "Picture"
   HasBackImage = False
   
   Call Refresh

End Property

' refresh the flatbutton
Public Sub Refresh()

   If IsClicked Then Exit Sub
   
   With picPicture
      .Width = ScaleWidth
      .Height = ScaleHeight
      .AutoSize = True
   End With
   
   Call GetButtonSize
   Call DrawBackground
   Call DrawBorder
   Call DrawIcon

End Sub

Private Function CheckMouseIn(ByVal X As Single, ByVal Y As Single) As Boolean

   With UserControl
      If PtInRect(ButtonRect, X, Y) Then
         CheckMouseIn = True
         
      Else
         CheckMouseIn = False
         
         Call DrawBorder
      End If
   End With

End Function

' draw button background
Private Sub DrawBackground()

Dim blnRedraw(1)    As Boolean
Dim intScaleMode(1) As Integer
Dim lngLeft         As Long
Dim lngTop          As Long

   If HasBackImage Then Exit Sub
   
   ' set true so this sub will only be done if there is something changed
   HasBackImage = True
   
   ' if the backstyle is opaque show the stored background picture
   If m_BackStyle = Opaque Then
      BitBlt hDC, 0, 0, ScaleWidth, ScaleHeight, picPicture.hDC, 0, 0, vbSrcCopy
      
   Else
      ' if the backstyle is transparent
      lngTop = Extender.Top
      lngLeft = Extender.Left
      
      With Extender.Parent
         ' store and set the parent properties
         blnRedraw(0) = .AutoRedraw
         intScaleMode(0) = .ScaleMode
         .AutoRedraw = True
         .ScaleMode = vbPixels
         
         If .Name <> Extender.Container.Name Then
            ' store and set the container properties
            ' show the container background part
            With Extender.Container
               blnRedraw(1) = .AutoRedraw
               intScaleMode(1) = .ScaleMode
               .AutoRedraw = True
               .ScaleMode = vbPixels
               BitBlt hDC, 0, 0, ScaleWidth, ScaleHeight, .hDC, lngLeft, lngTop, vbSrcCopy
               ' restore the container settings
               .ScaleMode = intScaleMode(0)
               .AutoRedraw = blnRedraw(0)
            End With
            
         Else
            ' show the parent background part
            BitBlt hDC, 0, 0, ScaleWidth, ScaleHeight, .hDC, lngLeft, lngTop, vbSrcCopy
         End If
         
         ' restore the parent settings
         .ScaleMode = intScaleMode(0)
         .AutoRedraw = blnRedraw(0)
      End With
   End If
   
   UserControl.Picture = Image
   Erase blnRedraw, intScaleMode

End Sub

' draw button borderstyle
Private Sub DrawBorder(Optional ByVal State As Boolean)

   If State Then
      If m_OnlyIconClick Then
         Call DrawIcon(True)
         
      Else
         UserControl.BorderStyle = vbFixedSingle
      End If
      
   ElseIf m_OnlyIconClick Then
      Call DrawIcon
      
   Else
      UserControl.BorderStyle = vbBSNone
   End If
   
   UserControl.Refresh

End Sub

' draw button icon
Private Sub DrawIcon(Optional ByVal MoveIcon As Boolean)

'Const DI_NORMAL As Long = &H3

   If m_Icon Is Nothing Then Exit Sub
   
   Cls
   DrawIconEx hDC, m_IconX + (1 And MoveIcon), m_IconY + (1 And MoveIcon), m_Icon.Handle, IconSize, IconSize, 0, 0, DI_NORMAL

End Sub

' fills the button rectangle
Private Sub GetButtonSize()

   With ButtonRect
      If m_OnlyIconClick And Not picPicture Is Nothing Then
         If Not m_Icon Is Nothing Then IconSize = m_Icon.Width / Screen.TwipsPerPixelX * 0.6
         
         .Top = m_IconY
         .Left = m_IconX
         .Right = .Left + IconSize
         .Bottom = .Top + IconSize
         
      Else
         .Top = 0
         .Left = 0
         .Right = ScaleWidth
         .Bottom = ScaleHeight
      End If
   End With

End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)

   If (PropertyName <> "BackColor") Or (m_BackStyle = Transparent) Then Exit Sub
   If Parent.Visible Then BackColor = Parent.BackColor

End Sub

Private Sub UserControl_Initialize()

   m_BackStyle = Opaque
   
   Call Refresh

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   RaiseEvent MouseDown(Button, Shift, X, Y)
   
   If (Button = vbLeftButton) And CheckMouseIn(X, Y) Then
      IsClicked = True
      MouseIn = True
      
      Call DrawBorder(True)
   End If

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   RaiseEvent MouseMove(Button, Shift, X, Y)
   
   If Button <> vbLeftButton Then Exit Sub
   
   MouseIn = CheckMouseIn(X, Y)
   
   If MouseIn Then Call DrawBorder(IsClicked)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   RaiseEvent MouseUp(Button, Shift, X, Y)
   
   If Button <> vbLeftButton Then Exit Sub
   
   IsClicked = False
   
   Call DrawBorder
   
   If MouseIn Then
      MouseIn = False
      RaiseEvent Click
   End If

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

   With PropBag
      picPicture.BackColor = .ReadProperty("BackColor", vbButtonFace)
      m_BackStyle = .ReadProperty("BackStyle", Opaque)
      UserControl.Enabled = .ReadProperty("Enabled", True)
      Set m_Icon = .ReadProperty("Icon", Nothing)
      m_IconX = .ReadProperty("IconX", 0)
      m_IconY = .ReadProperty("IconY", 0)
      UserControl.OLEDropMode = .ReadProperty("OLEDropMode", vbOLEDropNone)
      m_OnlyIconClick = .ReadProperty("OnlyIconClick", False)
      picPicture.Picture = .ReadProperty("Picture", Nothing)
   End With
   
   HasBackImage = False
   
   Call Refresh

End Sub

Private Sub UserControl_Resize()

   HasBackImage = False
   
   Call Refresh

End Sub

Private Sub UserControl_Terminate()

   Set m_Icon = Nothing

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   With PropBag
      .WriteProperty "BackColor", picPicture.BackColor, vbButtonFace
      .WriteProperty "BackStyle", m_BackStyle, Opaque
      .WriteProperty "Enabled", UserControl.Enabled, True
      .WriteProperty "Icon", m_Icon, Nothing
      .WriteProperty "IconX", m_IconX, 0
      .WriteProperty "IconY", m_IconY, 0
      .WriteProperty "OLEDropMode", UserControl.OLEDropMode, vbOLEDropNone
      .WriteProperty "OnlyIconClick", m_OnlyIconClick, False
      .WriteProperty "Picture", picPicture.Picture, Nothing
   End With

End Sub

