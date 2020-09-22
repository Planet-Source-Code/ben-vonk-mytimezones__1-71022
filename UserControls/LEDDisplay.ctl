VERSION 5.00
Begin VB.UserControl LEDDisplay 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   336
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   336
   ScaleHeight     =   28
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   28
   ToolboxBitmap   =   "LEDDisplay.ctx":0000
   Begin VB.Timer tmrDisplay 
      Enabled         =   0   'False
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "LEDDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'LEDDisplay Control
'
'Author Ben Vonk
'20-05-2007 First version
'30-06-2007 Second version, add Size property and fixed some minor bugs

Option Explicit

' Public Events
Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event TurnComplete()

' Private Constants
Private Const CHAR_HEIGHT As Long = 8
Private Const CHAR_WIDTH  As Long = 6

' Public Enumerations
'Public Enum BorderStyles
'   [None]
'   [Fixed Single]
'End Enum

Public Enum Sizes
   Thiny
   Small
   Medium
   Large
   ExtraLarge
   Big
   Extreem
End Enum

'Public Enum Speeds
'   Slow
'   Default
'   Fast
'End Enum

' Private Variables
Private m_NoTextScrolling As Boolean
Private m_BorderStyle     As BorderStyles
Private FontData()        As Byte
Private DisplayBitmap     As Long
Private DisplayDC         As Long
Private DisplayHeight     As Long
Private FontBitmap        As Long
Private FontDC            As Long
Private FontSize          As Long
Private SizeLED           As Long
Private TextLength        As Long
Private TextPosition      As Long
Private WidthLED          As Long
Private m_DisplayColor    As OLE_COLOR
Private m_ForeColor       As OLE_COLOR
Private m_Size            As Sizes
Private m_Speed           As Speeds
Private FilterChars(1)    As String
Private m_Text            As String
Private WorkingText       As String

' Private API's
'Private Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'Private Declare Function CreateCompatibleBitmap Lib "GDI32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'Private Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hDC As Long) As Long
'Private Declare Function DeleteDC Lib "GDI32" (ByVal hDC As Long) As Long
'Private Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
'Private Declare Function SelectObject Lib "GDI32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetPixel Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Public Property Get Active() As Boolean
Attribute Active.VB_Description = "Returns/sets activation of a LEDDisplay control."

   Active = tmrDisplay.Enabled

End Property

Public Property Let Active(ByVal NewActive As Boolean)

   If m_Text = "" Then NewActive = False
   If NewActive Then Call CreateText
   
   tmrDisplay.Enabled = NewActive
   PropertyChanged "Active"

End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."

   BackColor = UserControl.BackColor

End Property

Public Property Let BackColor(ByVal NewBackColor As OLE_COLOR)

   UserControl.BackColor = NewBackColor
   PropertyChanged "BackColor"
   
   Call CreateDisplay

End Property

Public Property Get BorderStyle() As BorderStyles
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."

   BorderStyle = m_BorderStyle

End Property

Public Property Let BorderStyle(ByVal NewBorderStyle As BorderStyles)

Dim blnChange As Boolean

   blnChange = (UserControl.BorderStyle <> NewBorderStyle)
   m_BorderStyle = NewBorderStyle
   PropertyChanged "BorderStyle"
   
   If blnChange Then Call UserControl_Resize

End Property

Public Property Get DisplayColor() As OLE_COLOR
Attribute DisplayColor.VB_Description = "Returns/sets the background color used to display text in the display."

   DisplayColor = m_DisplayColor

End Property

Public Property Let DisplayColor(ByVal NewDisplayColor As OLE_COLOR)

   m_DisplayColor = NewDisplayColor
   PropertyChanged "DisplayColor"
   
   Call CreateDisplay

End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text in the display."

   ForeColor = m_ForeColor

End Property

Public Property Let ForeColor(ByVal NewForeColor As OLE_COLOR)

   m_ForeColor = NewForeColor
   PropertyChanged "ForeColor"
   
   Call CreateDisplay

End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."

   hWnd = UserControl.hWnd

End Property

Public Property Get NoTextScrolling() As Boolean
Attribute NoTextScrolling.VB_Description = "Determines if text must be displayed static or scrolled."

   NoTextScrolling = m_NoTextScrolling

End Property

Public Property Let NoTextScrolling(ByVal NewNoScrollingText As Boolean)

   m_NoTextScrolling = NewNoScrollingText
   PropertyChanged "NoTextScrolling"

End Property

Public Property Get Size() As Sizes
Attribute Size.VB_Description = "Returns/sets the LED size used in a LEDDisplay object."

   Size = m_Size

End Property

Public Property Let Size(ByVal NewSize As Sizes)

   If NewSize < Thiny Then NewSize = Thiny
   If NewSize > Extreem Then NewSize = Extreem
   
   m_Size = NewSize
   PropertyChanged "Size"
   
   Call CreateFontBitmap
   Call CreateDisplay
   Call UserControl_Resize

End Property

Public Property Get Speed() As Speeds
Attribute Speed.VB_Description = "Returns/sets the speed of a LEDDisplay control."

   Speed = m_Speed

End Property

Public Property Let Speed(ByVal NewSpeed As Speeds)

   m_Speed = NewSpeed
   PropertyChanged "Speed"
   
   Call SetSpeed

End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."

   Text = m_Text

End Property

Public Property Let Text(ByVal NewText As String)

   m_Text = NewText
   PropertyChanged "Text"
   
   Call CreateText
   
   If Trim(WorkingText) = "" Then Call ClearAll

End Property

Public Sub ClearAll()

   Active = False
   TextPosition = 0
   
   Call ClearDisplay
   
   Refresh

End Sub

Private Function CreateDisplayText(ByVal Text As String) As Long

Static lngTextLenght As Long

Dim intAscii         As Integer
Dim lngPosition      As Long

   If lngTextLenght <> (Len(Text) * CHAR_WIDTH) Then
      ' first delete created device
      If DisplayBitmap Then
         DeleteObject DisplayBitmap
         DeleteDC DisplayDC
      End If
      
      ' create a device context for display text
      lngTextLenght = Len(Text) * CHAR_WIDTH
      DisplayBitmap = MakeBitmap(DisplayDC, lngTextLenght)
   End If
   
   For lngPosition = 0 To Len(Text) - 1
      intAscii = FilterCharacter(Mid(Text, lngPosition + 1, 1))
      BitBlt DisplayDC, lngPosition * (WidthLED + SizeLED), 1, WidthLED, DisplayHeight - 1, FontDC, 0 + ((intAscii - 32) * WidthLED And ((intAscii > 31) And (intAscii < 127))), 0, vbSrcCopy
      BitBlt DisplayDC, lngPosition * (WidthLED + SizeLED) - SizeLED, 1, SizeLED, DisplayHeight - 1, FontDC, 0, 0, vbSrcCopy
   Next 'lngPosition
   
   BitBlt DisplayDC, lngPosition * (WidthLED + SizeLED) - SizeLED, 1, SizeLED, DisplayHeight - 1, FontDC, 0, 0, vbSrcCopy
   CreateDisplayText = lngTextLenght * SizeLED

End Function

Private Function FilterCharacter(ByVal Character As String) As Integer

Dim intCharacter As Integer

   intCharacter = InStr(FilterChars(0), Character)
   
   If intCharacter Then
      FilterCharacter = Asc(Mid(FilterChars(1), intCharacter \ 7 + 1, 1))
      
   Else
      FilterCharacter = Asc(Character)
   End If

End Function

Private Function MakeBitmap(ByRef hDC As Long, ByVal Width As Long) As Long

Dim lngBitmap As Long

   ' create a device context, compatible with the screen
   hDC = CreateCompatibleDC(UserControl.hDC)
   ' create a bitmap, compatible with the screen
   lngBitmap = CreateCompatibleBitmap(UserControl.hDC, Width * SizeLED - 1, CHAR_HEIGHT * SizeLED - 1)
   ' select the bitmap into the device context
   SelectObject hDC, lngBitmap
   MakeBitmap = lngBitmap

End Function

Private Sub ClearDisplay()

   CreateDisplayText Space(ScaleWidth / SizeLED / CHAR_WIDTH + 1)
   BitBlt hDC, 1, 1, ScaleWidth * CHAR_WIDTH, DisplayHeight, DisplayDC, 0, 1, vbSrcCopy

End Sub

Private Sub CreateDisplay()

   Call CreateFont
   Call ClearDisplay
   
   If Len(m_Text) Then Call CreateText
   
   Refresh

End Sub

Private Sub CreateFont()

Dim intSizeX As Integer
Dim intSizeY As Integer
Dim intX     As Integer
Dim intY     As Integer
Dim lngColor As Long
Dim lngY     As Long
Dim lngX     As Long

   DisplayHeight = CHAR_HEIGHT * SizeLED
   WidthLED = (CHAR_WIDTH - 1) * SizeLED
   
   Call ClearDisplay
   
   For intX = 0 To FontSize - 1
      lngX = intX * SizeLED
      
      For intY = 0 To CHAR_HEIGHT - 2
         If FontData(intX) \ (2 ^ (7 - intY)) Mod 2 Then
            lngColor = m_ForeColor
            
         Else
            lngColor = m_DisplayColor
         End If
         
         lngY = intY * SizeLED
         
         For intSizeX = 0 To SizeLED - 2
            For intSizeY = 0 To SizeLED - 2
               SetPixel FontDC, lngX + intSizeX, lngY + intSizeY, lngColor
            Next 'intSizeY
         Next 'intSizeX
         
         lngColor = UserControl.BackColor
         
         For intSizeX = lngX To lngX + SizeLED - 1
            SetPixel FontDC, intSizeX, lngY + SizeLED - 1, lngColor
         Next 'intSizeX
         
         For intSizeY = lngY To lngY + SizeLED - 2
            SetPixel FontDC, lngX + SizeLED - 1, intSizeY, lngColor
         Next 'intSizeY
      Next 'intY
   Next 'intX

End Sub

Private Sub CreateFontBitmap()

   If FontBitmap Then
      DeleteObject FontBitmap
      DeleteDC FontDC
   End If
   
   SizeLED = m_Size + 2
   ' create a device context for font set
   FontBitmap = MakeBitmap(FontDC, FontSize + FontSize \ 5)

End Sub

Private Sub CreateText()

Dim lngSpace As Long
Dim strSpace As String

   If m_NoTextScrolling Then
      lngSpace = ((ScaleWidth - (Len(m_Text) * SizeLED * CHAR_WIDTH)) / 2) / SizeLED / CHAR_WIDTH
      
      If lngSpace > 0 Then strSpace = Space(lngSpace)
      
   Else
      strSpace = Space(ScaleWidth / SizeLED / CHAR_WIDTH + 1)
   End If
   
   WorkingText = strSpace & m_Text
   TextLength = CreateDisplayText(WorkingText)

End Sub

Private Sub SetSpeed()

   tmrDisplay.Interval = Choose(m_Speed + 1, 40, 25, 10)

End Sub

Private Sub tmrDisplay_Timer()

   TextPosition = TextPosition - SizeLED
   
   If TextPosition < TextLength Then TextPosition = TextPosition + TextLength
   
   TextPosition = TextPosition Mod TextLength
   BitBlt hDC, 1, 1, TextPosition + SizeLED, DisplayHeight, DisplayDC, TextLength - TextPosition - SizeLED, 1, vbSrcCopy
   Refresh
   DoEvents
   
   If TextPosition < SizeLED Then RaiseEvent TurnComplete
   If NoTextScrolling Then TextPosition = 0

End Sub

Private Sub UserControl_Click()

   RaiseEvent Click

End Sub

Private Sub UserControl_DblClick()

   RaiseEvent DblClick

End Sub

Private Sub UserControl_Initialize()

Dim intCount As Integer
Dim strChars As String

   ' load resource font to buffer
   FontData = LoadResData(1, "BIN")
   FontSize = UBound(FontData) + 1
   ' set filter characters
   FilterChars(1) = "ACEIDNOUYaceidnouy"
   
   For intCount = 1 To 18
      strChars = Choose(intCount, "ÀÁÂÃÄÅÆ", "Ç", "ÈÉÊË", "ÌÍÎÏ", "Ð", "Ñ", "ÒÓÔÕÖ", "ÙÚÛÜ", "Ý", "àáâãäåæ", "ç", "éêëè", "ìíîï", "ð", "ñ", "òóôõö", "ùúûü", "ýÿ")
      FilterChars(0) = FilterChars(0) & strChars & String(7 - Len(strChars), vbNullChar)
   Next 'intCount

End Sub

Private Sub UserControl_InitProperties()

   UserControl.BackColor = vbBlack
   m_DisplayColor = &H808000
   m_ForeColor = vbCyan
   m_Size = Medium
   m_Speed = Default
   
   Call SetSpeed
   Call CreateFontBitmap
   Call CreateDisplay

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

   RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

   With PropBag
      tmrDisplay.Enabled = .ReadProperty("Active", False)
      UserControl.BackColor = .ReadProperty("BackColor", vbBlack)
      m_BorderStyle = .ReadProperty("BorderStyle", [Fixed Single])
      m_DisplayColor = .ReadProperty("DisplayColor", &H808000)
      m_ForeColor = .ReadProperty("ForeColor", vbCyan)
      m_NoTextScrolling = .ReadProperty("NoTextScrolling", False)
      m_Size = .ReadProperty("Size", Medium)
      m_Speed = .ReadProperty("Speed", Default)
      m_Text = .ReadProperty("Text", "")
   End With
   
   Call SetSpeed
   Call CreateFontBitmap
   Call CreateFont
   Call CreateText
   Call UserControl_Resize

End Sub

Private Sub UserControl_Resize()

Static blnBusy As Boolean

Dim intSize    As Integer
Dim sngLEDs    As Single

   If blnBusy Then Exit Sub
   
   With Screen
      blnBusy = True
      UserControl.BorderStyle = m_BorderStyle
      Height = DisplayHeight * .TwipsPerPixelY - (m_Size + 1 - (4 And (m_BorderStyle = [Fixed Single]))) * .TwipsPerPixelY
      sngLEDs = (ScaleWidth - 1) / SizeLED
      
      Call ClearDisplay
      
      If sngLEDs <> Int(sngLEDs) Then
         intSize = Width / .TwipsPerPixelX - ScaleWidth + (1 And (m_BorderStyle = vbBSNone))
         
         If m_BorderStyle = [Fixed Single] Then intSize = intSize + m_Size - ((m_Size - Small) And (m_Size > Small))
         
         Width = (Int(sngLEDs) * SizeLED + intSize) * .TwipsPerPixelX
      End If
      
      blnBusy = False
   End With
   
   Call ClearDisplay
   
   Refresh

End Sub

Private Sub UserControl_Terminate()

   DeleteObject FontBitmap
   DeleteDC FontDC
   DeleteObject DisplayBitmap
   DeleteDC DisplayDC
   Erase FontData, FilterChars

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   With PropBag
      .WriteProperty "Active", tmrDisplay.Enabled, False
      .WriteProperty "BackColor", UserControl.BackColor, vbBlack
      .WriteProperty "BorderStyle", UserControl.BorderStyle, [Fixed Single]
      .WriteProperty "DisplayColor", m_DisplayColor, &H808000
      .WriteProperty "ForeColor", m_ForeColor, vbCyan
      .WriteProperty "NoTextScrolling", m_NoTextScrolling, False
      .WriteProperty "Size", m_Size, Medium
      .WriteProperty "Speed", m_Speed, Default
      .WriteProperty "Text", m_Text, ""
   End With

End Sub
