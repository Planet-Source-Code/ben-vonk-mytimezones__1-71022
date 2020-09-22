VERSION 5.00
Begin VB.UserControl CheckBox 
   AutoRedraw      =   -1  'True
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1872
   KeyPreview      =   -1  'True
   ScaleHeight     =   30
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   156
   ToolboxBitmap   =   "CheckBox.ctx":0000
   Begin VB.PictureBox picPicture 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   384
      Index           =   3
      Left            =   1440
      Picture         =   "CheckBox.ctx":0312
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.PictureBox picPicture 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   384
      Index           =   2
      Left            =   960
      Picture         =   "CheckBox.ctx":0BDC
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.PictureBox picPicture 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   384
      Index           =   1
      Left            =   480
      Picture         =   "CheckBox.ctx":14A6
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.PictureBox picPicture 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   384
      Index           =   0
      Left            =   0
      Picture         =   "CheckBox.ctx":1D70
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   384
   End
End
Attribute VB_Name = "CheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'CheckBox Control
'
'Author Ben Vonk
'20-08-2005 First version
'25-10-2005 Second version, some bugfixes and updated with option for transparency

Option Explicit

' Public Events
Public Event Click()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

' Public Enumerations
Public Enum Alignments
   [Left Justify]
   [Right Justify]
End Enum

'Public Enum BackStyles
'   Transparent
'   Opaque
'End Enum

Public Enum IconSizes
   [16x16]
   [32x32]
   [48x48]
   [64x64]
End Enum

Public Enum Values
   Unchecked
   Checked
   Grayed
End Enum

' Private Types
Private Type Icons
   Icon(3)                   As StdPicture   ' holds the 4 icons
   Size                      As Integer      ' size of the icon
   Left                      As Long         ' X position of the icon
End Type

'Private Type Rect
'   Left                     As Long
'   Top                      As Long
'   Right                    As Long
'   Bottom                   As Long
'End Type

' Private Variables
Private m_BackStyle          As BackStyles   ' sets the checkbox backstyle
Private HasBackImage         As Boolean      ' checked if backimage is set
Private IsClicked            As Boolean      ' checked if checkbox is clicked
Private IsInitControl        As Boolean      ' checked if the control properties are initialise and if so then exit the AmbientChanged sub
Private IsResizing           As Boolean      ' checked if checkbox will be resized in the Resize procedure
Private MouseIn              As Boolean      ' checked if mouse is in the checkbox
Private m_AutoSize           As Boolean      ' sets the autosize
Private IconProps            As Icons        ' holds the icon properties
Private m_AccessKey          As Integer      ' holds the checkbox accesskey
Private m_Alignment          As Integer      ' sets the alignment
Private m_AlignCaption       As Integer      ' sets the alignment for the caption
Private m_Value              As Integer      ' sets the checkbox value
Private SizeIcon             As Integer      ' holds the size of the icons
Private CaptionLeft          As Long         ' sets the X position of the caption
Private m_BackColor          As Long         ' sets the backcolor
Private m_ForeColor          As Long         ' sets the forecolor
Private ButtonRect           As Rect         ' holds the size of the checkbox used for API PtInRect
Private m_Font               As StdFont      ' sets the font
Private m_Picture            As StdPicture   ' sets the background picture
Private CaptionText          As String       ' holds the caption text to display
Private m_Caption            As String       ' sets the caption

' Private API's
'Private Declare Function BitBlt Lib "GDI32" (ByVal hDCDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal dwRop As Long) As Long
'Private Declare Function DrawIconEx Lib "User32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
'Private Declare Function PtInRect Lib "User32" (Rect As Rect, ByVal lPtX As Long, ByVal lPtY As Long) As Integer

Public Property Get AlignCaption() As Alignments
Attribute AlignCaption.VB_Description = "Returns/sets the allignment of a CheckBox or OptionButton, or a control's text."

   AlignCaption = m_AlignCaption

End Property

Public Property Let AlignCaption(ByVal NewAlignCaption As Alignments)

   m_AlignCaption = NewAlignCaption
   PropertyChanged "AlignCaption"
   
   Call Resize
   Call DrawButton

End Property

Public Property Get Alignment() As Alignments
Attribute Alignment.VB_Description = "Returns/sets the allignment of a CheckBox or OptionButton, or a control's text."

   Alignment = m_Alignment

End Property

Public Property Let Alignment(ByVal NewAlignment As Alignments)

   m_Alignment = NewAlignment
   PropertyChanged "Alignment"
   
   Call Resize
   Call DrawButton

End Property

Public Property Get AutoSize() As Boolean
Attribute AutoSize.VB_Description = "Returns/sets wheter a control is automatically resized to display its entire contents."

   AutoSize = m_AutoSize

End Property

Public Property Let AutoSize(ByVal NewAutoSize As Boolean)

   m_AutoSize = NewAutoSize
   PropertyChanged "AutoSize"
   
   Call Resize
   Call DrawBackground
   Call DrawButton

End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."

   BackColor = m_BackColor

End Property

Public Property Let BackColor(ByVal NewBackColor As OLE_COLOR)

   m_BackColor = NewBackColor
   PropertyChanged "BackColor"
   
   Call Resize
   Call DrawButton

End Property

Public Property Get BackStyle() As BackStyles
Attribute BackStyle.VB_Description = "Returns/sets the border style for an object."

   BackStyle = m_BackStyle

End Property

Public Property Let BackStyle(ByVal NewBackStyle As BackStyles)

   m_BackStyle = NewBackStyle
   PropertyChanged "BackStyle"
   ' set false to repaint the background
   HasBackImage = False
   
   Call DrawBackground
   Call DrawButton

End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."

   Caption = m_Caption

End Property

Public Property Let Caption(ByVal NewCaption As String)

   m_Caption = NewCaption
   PropertyChanged "Caption"
   
   Call GetAccessKey
   Call Resize
   Call DrawButton

End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."

   Enabled = UserControl.Enabled

End Property

Public Property Let Enabled(ByVal NewEnabled As Boolean)

   UserControl.Enabled = NewEnabled
   PropertyChanged "Enabled"
   
   Call DrawButton

End Property

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Returns a Font object."

   Set Font = m_Font

End Property

Public Property Let Font(ByVal NewFont As StdFont)

   Set Font = NewFont

End Property

Public Property Set Font(ByVal NewFont As StdFont)

   Set m_Font = NewFont
   Set NewFont = Nothing
   PropertyChanged "Font"
   
   Call Resize
   Call DrawButton

End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."

   ForeColor = m_ForeColor

End Property

Public Property Let ForeColor(ByVal NewForeColor As OLE_COLOR)

   m_ForeColor = NewForeColor
   PropertyChanged "ForeColor"
   
   Call Resize
   Call DrawButton

End Property

Public Property Get IconChecked() As StdPicture
Attribute IconChecked.VB_Description = "Returns/sets a icon to be displayed when the CheckBox is checked."

   Set IconChecked = IconProps.Icon(1)

End Property

Public Property Let IconChecked(ByRef NewIconChecked As StdPicture)

   Set IconChecked = NewIconChecked

End Property

Public Property Set IconChecked(ByRef NewIconChecked As StdPicture)

   Call SetNewIcon(1, "IconChecked", NewIconChecked)

End Property

Public Property Get IconCheckedGrayed() As StdPicture
Attribute IconCheckedGrayed.VB_Description = "Returns/sets a icon to be displayed when the CheckBox is checked and grayed."

   Set IconCheckedGrayed = IconProps.Icon(3)

End Property

Public Property Let IconCheckedGrayed(ByRef NewIconCheckedGrayed As StdPicture)

   Set IconCheckedGrayed = NewIconCheckedGrayed

End Property

Public Property Set IconCheckedGrayed(ByRef NewIconCheckedGrayed As StdPicture)

   Call SetNewIcon(3, "IconCheckedGrayed", NewIconCheckedGrayed)

End Property

Public Property Get IconSize() As IconSizes
Attribute IconSize.VB_Description = "Returns/sets the size of an icon object."

   IconSize = IconProps.Size

End Property

Public Property Let IconSize(ByVal NewIconSize As IconSizes)

   IconProps.Size = NewIconSize
   PropertyChanged "IconSize"
   
   Call Resize
   Call DrawBackground
   Call DrawButton

End Property

Public Property Get IconUnchecked() As StdPicture
Attribute IconUnchecked.VB_Description = "Returns/sets a icon to be displayed when the CheckBox is unchecked."

   Set IconUnchecked = IconProps.Icon(0)

End Property

Public Property Let IconUnchecked(ByRef NewIconUnchecked As StdPicture)

   Set IconUnchecked = NewIconUnchecked

End Property

Public Property Set IconUnchecked(ByRef NewIconUnchecked As StdPicture)

   Call SetNewIcon(0, "IconUnchecked", NewIconUnchecked)

End Property

Public Property Get IconUncheckedGrayed() As StdPicture
Attribute IconUncheckedGrayed.VB_Description = "Returns/sets a icon to be displayed when the CheckBox is unchecked and grayed."

   Set IconUncheckedGrayed = IconProps.Icon(2)

End Property

Public Property Let IconUncheckedGrayed(ByRef NewIconUncheckedGrayed As StdPicture)

   Set IconUncheckedGrayed = NewIconUncheckedGrayed

End Property

Public Property Set IconUncheckedGrayed(ByRef NewIconUncheckedGrayed As StdPicture)

   Call SetNewIcon(2, "IconUncheckedGrayed", NewIconUncheckedGrayed)

End Property

Public Property Get Picture() As StdPicture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a CommandButton, OptionButton or CheckBox."

   Set Picture = UserControl.Image

End Property

Public Property Let Picture(ByRef NewPicture As StdPicture)

   Set Picture = NewPicture

End Property

Public Property Set Picture(ByRef NewPicture As StdPicture)

   Set m_Picture = NewPicture
   Set NewPicture = Nothing
   PropertyChanged "Picture"
   ' set false to repaint the background
   HasBackImage = False
   
   Call DrawBackground
   Call DrawButton

End Property

Public Property Get Value() As Values
Attribute Value.VB_Description = "Returns/sets the value of a object."

   If m_Value > Grayed Then
      Value = Grayed
      
   Else
      Value = m_Value
   End If

End Property

Public Property Let Value(ByVal NewValue As Values)

   If (NewValue = Grayed) And (m_Value = Checked) Then
      m_Value = NewValue + 1
      
   Else
      m_Value = NewValue
   End If
   
   PropertyChanged "Value"
   
   Call DrawButton

End Property

Private Function CheckMouseIn(ByVal X As Single, ByVal Y As Single) As Boolean

   With UserControl
      If PtInRect(ButtonRect, X, Y) Then
         CheckMouseIn = True
         
      Else
         CheckMouseIn = False
         
         Call DrawButton
      End If
   End With

End Function

Private Sub DrawBackground()

Dim blnRedraw    As Boolean
Dim intScaleMode As Integer
Dim lngLeft      As Long
Dim lngTop       As Long

   If HasBackImage Then Exit Sub
   
   On Local Error Resume Next
   ' set true so this job will only be done if there is something changed
   HasBackImage = True
   ' get and store the parent properties
   intScaleMode = Parent.ScaleMode
   blnRedraw = Parent.AutoRedraw
   Parent.ScaleMode = vbPixels
   Parent.AutoRedraw = True
   
   If m_BackStyle = Opaque Then
      If Not m_Picture Is Nothing Then
         With m_Picture
            .Render hDC, 0, 0, ScaleWidth, ScaleHeight, 0, .Height, .Width, -.Height, ByVal 0&
            UserControl.Picture = UserControl.Image
         End With
         
      Else
         UserControl.Picture = Nothing
      End If
      
   Else
      lngTop = Extender.Top
      lngLeft = Extender.Left
      
      If InStr(Extender.Tag, ",") Then
         lngTop = Val(Split(Extender.Tag, ",")(0))
         lngLeft = Val(Split(Extender.Tag, ",", 2)(1))
      End If
      
      BitBlt hDC, 0, 0, ScaleWidth, ScaleHeight, Parent.hDC, lngLeft, lngTop, vbSrcCopy
      UserControl.Picture = UserControl.Image
   End If
   
   ' restore the parent settings
   Parent.ScaleMode = intScaleMode
   Parent.AutoRedraw = blnRedraw
   On Local Error GoTo 0

End Sub

Private Sub DrawButton(Optional ByVal Sunken As Boolean)

Dim intCount As Integer

   With UserControl
      If Sunken Then intCount = 1
      
      .Cls
      DrawIconEx .hDC, IconProps.Left + (1 And Sunken), intCount, IconProps.Icon(Abs(m_Value) + (2 And Not .Enabled)).Handle, SizeIcon - (2 And Sunken), SizeIcon - (2 And Sunken), 0, 0, &H3
      .CurrentY = (SizeIcon - TextHeight("X")) \ 2
      .CurrentX = CaptionLeft
      
      For intCount = 1 To Len(CaptionText)
         .FontUnderline = (intCount = m_AccessKey)
         UserControl.Print Mid(CaptionText, intCount, 1);
      Next 'intCount
   End With

End Sub

' get the checkbox accesskey
Private Sub GetAccessKey()

Dim intCount As Integer

   m_AccessKey = 0
   CaptionText = m_Caption
   
   For intCount = Len(CaptionText) To 1 Step -1
      If Mid(CaptionText, intCount, 1) = "&" Then
         If intCount > 1 Then
            If Mid(CaptionText, intCount - 1, 1) = "&" Then
               intCount = intCount - 1
               
            Else
               Exit For
            End If
            
         Else
            Exit For
         End If
      End If
   Next 'intCount
   
   If intCount Then
      m_AccessKey = intCount
      AccessKeys = Mid(CaptionText, intCount + 1, 1)
      CaptionText = Left(CaptionText, intCount - 1) & Mid(CaptionText, intCount + 1)
   End If

End Sub

' resize the checkbox
Private Sub Resize()

Dim lngTextWidth As Long
Dim lngWidth     As Long

   With UserControl
      .BackColor = m_BackColor
      Set .Font = m_Font
      .ForeColor = m_ForeColor
   End With
   
   With Screen
      ' ignore all errors
      On Local Error Resume Next
      ' do so to exit the UserControl_Resize event
      IsResizing = True
      SizeIcon = (IconProps.Size + 1) * 16
      Height = SizeIcon * .TwipsPerPixelY
      lngTextWidth = TextWidth(CaptionText)
      lngWidth = SizeIcon + lngTextWidth + (3 And lngTextWidth)
      
      If m_AutoSize Or (ScaleWidth < lngWidth) Then Width = (lngWidth + 3) * .TwipsPerPixelX
      
      If m_Alignment = [Left Justify] Then
         If m_AlignCaption = [Left Justify] Then
            CaptionLeft = SizeIcon + 3
            
         Else
            CaptionLeft = ScaleWidth - lngTextWidth
         End If
         
      Else
         If m_AlignCaption = [Left Justify] Then
            CaptionLeft = 0
            
         Else
            CaptionLeft = ScaleWidth - lngTextWidth - SizeIcon - 3
         End If
         
         If m_AutoSize Then
            Extender.Left = Extender.Left + Extender.Width - ScaleX(Width, vbTwips, Parent.ScaleMode)
            Width = (lngWidth + 3) * .TwipsPerPixelX
         End If
      End If
   End With
   
   If m_Alignment = [Left Justify] Then
      IconProps.Left = 0
      
   Else
      IconProps.Left = ScaleWidth - SizeIcon
   End If
   
   ' restore it to handle user resizing
   IsResizing = False
   On Local Error GoTo 0
   
   ' fill the checkbox rectangle
   With ButtonRect
      .Top = 0
      .Left = 0
      .Right = ScaleWidth
      .Bottom = ScaleHeight
   End With
   
End Sub

Private Sub SetNewIcon(ByVal Index As Integer, ByVal ChangedProperty As String, ByRef NewIcon As StdPicture)

   If NewIcon Is Nothing Then Set NewIcon = picPicture.Item(Index).Picture
   
   Set IconProps.Icon(Index) = NewIcon
   Set NewIcon = Nothing
   PropertyChanged ChangedProperty
   
   Call DrawButton

End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)

   Value = Abs(Not CBool(Value))
   RaiseEvent Click

End Sub

' if ambient is changed repaint the background
Private Sub UserControl_AmbientChanged(PropertyName As String)

   If IsInitControl Or Not Parent.Visible Then Exit Sub
   
   ' set false to repaint the background
   HasBackImage = False
   
   Call DrawBackground
   Call DrawButton

End Sub

Private Sub UserControl_InitProperties()

   IsInitControl = True
   
   With IconProps
      .Size = [32x32]
      Set .Icon(0) = picPicture.Item(0).Picture
      Set .Icon(1) = picPicture.Item(1).Picture
      Set .Icon(2) = picPicture.Item(2).Picture
      Set .Icon(3) = picPicture.Item(3).Picture
   End With
   
   m_BackColor = Ambient.BackColor
   m_BackStyle = Opaque
   m_Caption = Ambient.DisplayName
   Set Font = Ambient.Font
   m_ForeColor = Ambient.ForeColor
   m_Value = Unchecked
   
   Call GetAccessKey
   Call Resize
   Call DrawBackground
   Call DrawButton

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

   RaiseEvent KeyDown(KeyCode, Shift)

End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)

   RaiseEvent KeyPress(KeyAscii)

End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

   RaiseEvent KeyUp(KeyCode, Shift)

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   RaiseEvent MouseDown(Button, Shift, X, Y)
   
   If Button = vbLeftButton Then
      MouseIn = True
      IsClicked = True
      
      Call DrawButton(True)
   End If

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   RaiseEvent MouseMove(Button, Shift, X, Y)
   MouseIn = CheckMouseIn(X, Y)
   
   If (Button = vbLeftButton) And IsClicked Then If MouseIn Then Call DrawButton(True)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

   RaiseEvent MouseUp(Button, Shift, X, Y)
   
   If MouseIn And (Button = vbLeftButton) Then
      IsClicked = False
      Value = Abs(Not CBool(Value))
      
      RaiseEvent Click
   End If

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

   IsInitControl = True
   
   With PropBag
      m_AlignCaption = .ReadProperty("AlignCaption", [Left Justify])
      m_Alignment = .ReadProperty("Alignment", [Left Justify])
      m_AutoSize = .ReadProperty("AutoSize", False)
      m_BackColor = .ReadProperty("BackColor", Ambient.BackColor)
      m_BackStyle = .ReadProperty("BackStyle", Opaque)
      m_Caption = .ReadProperty("Caption", "")
      UserControl.Enabled = .ReadProperty("Enabled", True)
      Set m_Font = .ReadProperty("Font", Ambient.Font)
      m_ForeColor = .ReadProperty("ForeColor", Ambient.ForeColor)
      Set IconProps.Icon(1) = .ReadProperty("IconChecked", picPicture.Item(1).Picture)
      Set IconProps.Icon(3) = .ReadProperty("IconCheckedGrayed", picPicture.Item(3).Picture)
      IconProps.Size = .ReadProperty("IconSize", [32x32])
      Set IconProps.Icon(0) = .ReadProperty("IconUnchecked", picPicture.Item(0).Picture)
      Set IconProps.Icon(2) = .ReadProperty("IconUncheckedGrayed", picPicture.Item(2).Picture)
      Set m_Picture = .ReadProperty("Picture", Nothing)
      m_Value = .ReadProperty("Value", Unchecked)
   End With
   
   Call GetAccessKey
   Call Resize
   Call DrawBackground
   Call DrawButton
   
   IsInitControl = False

End Sub

Private Sub UserControl_Resize()

   If IsResizing Then Exit Sub
   
   ' if ambient is changed repaint the background
   HasBackImage = False
   IsInitControl = True
   
   Call Resize
   Call DrawBackground
   Call DrawButton
   
   IsInitControl = True

End Sub

Private Sub UserControl_Terminate()

Dim intCount As Integer

   Set m_Font = Nothing
   Set m_Picture = Nothing
   
   For intCount = 0 To 3
      Set IconProps.Icon(intCount) = Nothing
   Next 'intCount

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   With PropBag
      .WriteProperty "AlignCaption", m_AlignCaption, [Left Justify]
      .WriteProperty "Alignment", m_Alignment, [Left Justify]
      .WriteProperty "AutoSize", m_AutoSize, False
      .WriteProperty "BackColor", m_BackColor, Ambient.BackColor
      .WriteProperty "BackStyle", m_BackStyle, Opaque
      .WriteProperty "Caption", m_Caption, ""
      .WriteProperty "Enabled", UserControl.Enabled, True
      .WriteProperty "Font", m_Font, Ambient.Font
      .WriteProperty "ForeColor", m_ForeColor, Ambient.ForeColor
      .WriteProperty "IconChecked", IconProps.Icon(1), picPicture.Item(1).Picture
      .WriteProperty "IconCheckedGrayed", IconProps.Icon(3), picPicture.Item(2).Picture
      .WriteProperty "IconSize", IconProps.Size, [32x32]
      .WriteProperty "IconUnchecked", IconProps.Icon(0), picPicture.Item(0).Picture
      .WriteProperty "IconUncheckedGrayed", IconProps.Icon(2), picPicture.Item(2).Picture
      .WriteProperty "Picture", m_Picture, Nothing
      .WriteProperty "Value", m_Value, Unchecked
   End With

End Sub
