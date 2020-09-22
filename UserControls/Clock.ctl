VERSION 5.00
Begin VB.UserControl Clock 
   ClientHeight    =   2172
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4344
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   KeyPreview      =   -1  'True
   ScaleHeight     =   181
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   362
   ToolboxBitmap   =   "Clock.ctx":0000
   Begin VB.Timer tmrClock 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   3840
      Top             =   1800
   End
   Begin VB.PictureBox picClock 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   1932
      Left            =   120
      ScaleHeight     =   161
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   1932
      Begin VB.PictureBox picAlarm 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   192
         Left            =   840
         Picture         =   "Clock.ctx":0312
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   840
         Visible         =   0   'False
         Width           =   192
      End
   End
   Begin VB.PictureBox picBackGround 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1932
      Left            =   2280
      ScaleHeight     =   161
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   1932
   End
   Begin VB.Image imgToolTipText 
      Height          =   2172
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2172
   End
End
Attribute VB_Name = "Clock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Clock Control
'
'Author Ben Vonk
'15-06-2004 First version (based on Ghislain Chabot's 'Analog O Digital Clock ScreenSaver' at http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=25704&lngWId=1)
'06-07-2005 Some BugFixes and Updated with option for gradientstyles
'04-11-2005 Some BugFixes

Option Explicit

' Public Events
Public Event Alarm()
Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

' Private Constants
Private Const FLOODFILLBORDER             As Long = 0
Private Const CLOCK_HOURSCOLOR            As Long = &H800000
Private Const CLOCKPLATE_BORDERCOLOR      As Long = &HC00000
Private Const CLOCKPLATE_COLOR            As Long = &HE0E0E0
Private Const DATE_COLOR                  As Long = &H800000
Private Const DIGITS_BACKCOLOR            As Long = &H808000
Private Const DIGITS_FORECOLOR            As Long = &HFFFF00
Private Const GRADIENT_COLOR              As Long = &HFFFFFF
Private Const HAND_SECONDCOLOR            As Long = &HFF&
Private Const HANDS_HOURMINUTEBORDERCOLOR As Long = &HD94600
Private Const HANDS_HOURMINUTECOLOR       As Long = &H800000
Private Const MARKS_HOURCOLOR             As Long = &HFF&
Private Const MARKS_MINUTECOLOR           As Long = &HC00000
Private Const NAME_BACKCOLOR              As Long = &H808000
Private Const NAME_FORECOLOR              As Long = &HFFFF00
'Private Const PI_PART                     As Single = 3.14159265358979 / 180

' Public Enumerations
'Public Enum BackStyles
'   Transparent
'   Opaque
'End Enum

Public Enum ClockBorderStyles
   NoBorder
   Raised
   Sunken
   Edged
End Enum

Public Enum ClockDateStyles
   NoDate
   Up
   Down
   Switch
End Enum

Public Enum ClockPlateGradientPositions
   TopLeft
   TopCenter
   TopRight
   CenterLeft
   Center
   CenterRight
   BottomLeft
   BottomCenter
   BottomRight
End Enum

Public Enum ClockPlateGradientStyles
   NoGradient
   OutToIn
   InToOut
End Enum

Public Enum ClockHandsHourMinuteStyles
   Bold
   Flat
   Thin
   Hairline
End Enum

Public Enum ClockHandsMoves
   ByTime
   Smooth
End Enum

Public Enum ClockLabelPositions
   Off
   High
   Low
End Enum

Public Enum ClockMarksShows
   HoursMinutes
   Hours
   Quarters
End Enum

Public Enum ClockMarksStyles
   NoMarks
   CircleMarks
   SquareMarks
   LineMarks
End Enum

Public Enum ClockMirrorStyles
   NoMirror
   HorizontalMirror
   VerticalMirror
   HorizontalVerticalMirror
End Enum

Public Enum ClockStyles
   CircleClock
   SquareClock
End Enum

' Private Types
'Private Type Rect
'    Left                                  As Long
'    Top                                   As Long
'    Right                                 As Long
'    Bottom                                As Long
'End Type

Private Type TimeNameBoxRect
   BoxX                                   As Single
   BoxY                                   As Single
   TextX                                  As Single
   TextY                                  As Single
End Type

' Private Variables
Private DatePosition                      As Boolean
Private GiveAlarm                         As Boolean
Private GiveBell                          As Boolean
Private m_Active                          As Boolean
Private m_ClockBold                       As Boolean
Private m_ClockHourBell                   As Boolean
Private m_ClockHours                      As Boolean
Private m_ClockPlateBorder                As Boolean
Private m_ClockPlateBorderBold            As Boolean
Private m_DateBold                        As Boolean
Private m_DigitsAmPm                      As Boolean
Private m_DigitsBold                      As Boolean
Private m_Locked                          As Boolean
Private m_NameBold                        As Boolean
Private m_BackStyle                       As BackStyles
Private m_ClockPlateBackStyle             As BackStyles
Private m_DigitsBackStyle                 As BackStyles
Private m_NameBackStyle                   As BackStyles
Private m_BorderStyle                     As ClockBorderStyles
Private m_DigitsBorderStyle               As ClockBorderStyles
Private m_NameBorderStyle                 As ClockBorderStyles
Private m_DateStyle                       As ClockDateStyles
Private m_ClockPlateGradientPosition      As ClockPlateGradientPositions
Private m_ClockPlateGradientStyle         As ClockPlateGradientStyles
Private m_HandSecondScroll                As ClockHandsMoves
Private m_HandsHourMinuteScroll           As ClockHandsMoves
Private m_HandsHourMinuteStyle            As ClockHandsHourMinuteStyles
Private m_DigitsPosition                  As ClockLabelPositions
Private m_NamePosition                    As ClockLabelPositions
Private m_MarksShows                      As ClockMarksShows
Private m_MarksStyle                      As ClockMarksStyles
Private m_ClockMirrored                   As ClockMirrorStyles
Private m_ClockPlateStyle                 As ClockStyles
Private ClockSize                         As Integer
Private CountBells                        As Integer
Private Counter                           As Integer
Private LengthHour                        As Integer
Private LengthMinute                      As Integer
Private LengthSecond                      As Integer
Private Lengths(5)                        As Integer
Private MoveHour                          As Integer
Private MoveMinute                        As Integer
Private TimeHour                          As Integer
Private TimeMinute                        As Integer
Private TimeSecond                        As Integer
Private TempValue                         As Integer
Private TimeValue                         As Long
Private m_TimeZoneValue                   As Long
Private m_BackColor                       As OLE_COLOR
Private m_ClockHoursColor                 As OLE_COLOR
Private m_ClockPlateBorderColor           As OLE_COLOR
Private m_ClockPlateColor                 As OLE_COLOR
Private m_ClockPlateGradientColor         As OLE_COLOR
Private m_DateColor                       As OLE_COLOR
Private m_DigitsBackColor                 As OLE_COLOR
Private m_DigitsForeColor                 As OLE_COLOR
Private m_HandSecondColor                 As OLE_COLOR
Private m_HandsHourMinuteColor            As OLE_COLOR
Private m_HandsHourMinuteBorderColor      As OLE_COLOR
Private m_MarksHourColor                  As OLE_COLOR
Private m_MarksMinuteColor                As OLE_COLOR
Private m_NameBackColor                   As OLE_COLOR
Private m_NameForeColor                   As OLE_COLOR
Private FrameRect                         As Rect
Private m_Font                            As StdFont
Private m_Picture                         As StdPicture
Private ClockHour                         As Single
Private ClockMinute                       As Single
Private ClockSecond                       As Single
Private ClockTimer                        As Single
Private ClockWidth                        As Single
Private ClockX(3)                         As Single
Private ClockY(3)                         As Single
Private Parts(3)                          As Single
Private Buffer(1)                         As String
Private DigitsFormat                      As String
Private m_AlarmTime                       As String
Private m_AlarmToolTipText                As String
Private m_DateTime                        As String
Private m_NameClock                       As String
Private m_SetDate                         As String
Private m_SetTime                         As String
Private NextBell                          As String
Private TimeNameBox(1)                    As TimeNameBoxRect

' Private API's
'Private Declare Function BitBlt Lib "GDI32" (ByVal hDCDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateEllipticRgn Lib "GDI32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'Private Declare Function CreateRectRgn Lib "GDI32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function ExtFloodFill Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
'Private Declare Function StretchBlt Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
'Private Declare Function DrawEdge Lib "User32" (ByVal hDC As Long, qrc As Rect, ByVal edge As Long, ByVal grfFlags As Long) As Long
'Private Declare Function SetWindowRgn Lib "User32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Public Property Get Active() As Boolean
Attribute Active.VB_Description = "Returns/sets activation of a clock control."

   Active = m_Active

End Property

Public Property Let Active(ByVal NewActive As Boolean)

   m_Active = NewActive
   PropertyChanged ("Active")
   
   Call InitialiseClock
   
   tmrClock.Enabled = m_Active

End Property

Public Property Get AlarmTime() As String
Attribute AlarmTime.VB_Description = "Returns/sets the alarmtime of a clock control."

   AlarmTime = m_AlarmTime

End Property

Public Property Let AlarmTime(ByVal NewAlarmTime As String)

   m_AlarmTime = Format(CheckTime(NewAlarmTime), "hh:mm")
   PropertyChanged ("AlarmTime")
   
   Call InitialiseClock

End Property

Public Property Get AlarmToolTipText() As String
Attribute AlarmToolTipText.VB_Description = "Returns/sets the alarm tooltiptext of a Clock control. A # in the text will indicates the place for time."

   AlarmToolTipText = m_AlarmToolTipText

End Property

Public Property Let AlarmToolTipText(ByVal NewAlarmToolTipText As String)

   m_AlarmToolTipText = NewAlarmToolTipText
   PropertyChanged ("AlarmToolTipText")
   
   Call SetToolTipText

End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."

   BackColor = m_BackColor

End Property

Public Property Let BackColor(ByVal NewBackColor As OLE_COLOR)

   m_BackColor = NewBackColor
   PropertyChanged ("BackColor")
   
   Call InitialiseClock

End Property

Public Property Get BackStyle() As BackStyles
Attribute BackStyle.VB_Description = "Indicates the backstyle of the clock control is transparent or opaque."

   BackStyle = m_BackStyle

End Property

Public Property Let BackStyle(ByVal NewBackStyle As BackStyles)

   m_BackStyle = NewBackStyle
   PropertyChanged ("BackStyle")
   
   Call MakeTransparent
   Call InitialiseClock

End Property

Public Property Get ClockBorderStyles() As ClockBorderStyles
Attribute ClockBorderStyles.VB_Description = "Returns/sets the border style of the clock control."

   ClockBorderStyles = m_BorderStyle

End Property

Public Property Let ClockBorderStyles(ByVal NewBorderStyle As ClockBorderStyles)

   m_BorderStyle = NewBorderStyle
   PropertyChanged ("ClockBorderStyles")
   
   Call InitialiseClock

End Property

Public Property Get ClockBold() As Boolean
Attribute ClockBold.VB_Description = "Returns/sets the boldsize of the clock control."

   ClockBold = m_ClockBold

End Property

Public Property Let ClockBold(ByVal NewClockBold As Boolean)

   m_ClockBold = NewClockBold
   PropertyChanged ("ClockBold")
   
   Call InitialiseClock

End Property

Public Property Get ClockHourBell() As Boolean
Attribute ClockHourBell.VB_Description = "Returns/sets the clockbell that rings every hour."

   ClockHourBell = m_ClockHourBell

End Property

Public Property Let ClockHourBell(ByVal NewClockHourBell As Boolean)

   m_ClockHourBell = NewClockHourBell
   PropertyChanged ("ClockHourBell")
   
   Call InitialiseClock

End Property

Public Property Get ClockHours() As Boolean
Attribute ClockHours.VB_Description = "Returns/sets to switch the hours of the clock control on/off."

   ClockHours = m_ClockHours

End Property

Public Property Let ClockHours(ByVal NewClockHours As Boolean)

   m_ClockHours = NewClockHours
   PropertyChanged ("ClockHours")
   
   Call InitialiseClock

End Property

Public Property Get ClockHoursColor() As OLE_COLOR
Attribute ClockHoursColor.VB_Description = "Returns/sets the color of the hour marks of the clock control."

   ClockHoursColor = m_ClockHoursColor

End Property

Public Property Let ClockHoursColor(ByVal NewClockHoursColor As OLE_COLOR)

   m_ClockHoursColor = NewClockHoursColor
   PropertyChanged ("ClockHoursColor")
   
   Call InitialiseClock

End Property

Public Property Get ClockMirrored() As ClockMirrorStyles
Attribute ClockMirrored.VB_Description = "Returns/sets to show the clock in a mirrored style or normal."

   ClockMirrored = m_ClockMirrored

End Property

Public Property Let ClockMirrored(ByVal NewClockMirrored As ClockMirrorStyles)

   m_ClockMirrored = NewClockMirrored
   PropertyChanged ("ClockMirrored")
   
   Call InitialiseClock

End Property

Public Property Get ClockPlateBackStyle() As BackStyles
Attribute ClockPlateBackStyle.VB_Description = "Returns/sets the backstyle of the plate of the clock control."

   ClockPlateBackStyle = m_ClockPlateBackStyle

End Property

Public Property Let ClockPlateBackStyle(ByVal NewClockPlateBackStyle As BackStyles)

   m_ClockPlateBackStyle = NewClockPlateBackStyle
   PropertyChanged ("ClockPlateBackStyle")
   
   Call InitialiseClock

End Property

Public Property Get ClockPlateBorder() As Boolean
Attribute ClockPlateBorder.VB_Description = "Returns/sets to switch the border of the clockplate on/off."

   ClockPlateBorder = m_ClockPlateBorder

End Property

Public Property Let ClockPlateBorder(ByVal NewClockPlateBorder As Boolean)

   m_ClockPlateBorder = NewClockPlateBorder
   PropertyChanged ("ClockPlateBorder")
   
   Call InitialiseClock

End Property

Public Property Get ClockPlateBorderBold() As Boolean
Attribute ClockPlateBorderBold.VB_Description = "Returns/sets to switch the bold of the clockplate border on/off."

   ClockPlateBorderBold = m_ClockPlateBorderBold

End Property

Public Property Let ClockPlateBorderBold(ByVal NewClockPlateBorderBold As Boolean)

   m_ClockPlateBorderBold = NewClockPlateBorderBold
   PropertyChanged ("ClockPlateBorderBold")
   
   Call InitialiseClock

End Property

Public Property Get ClockPlateBorderColor() As OLE_COLOR
Attribute ClockPlateBorderColor.VB_Description = "Returns/sets the border color of the clockplate."

   ClockPlateBorderColor = m_ClockPlateBorderColor

End Property

Public Property Let ClockPlateBorderColor(ByVal NewClockPlateBorderColor As OLE_COLOR)

   m_ClockPlateBorderColor = NewClockPlateBorderColor
   PropertyChanged ("ClockPlateBorderColor")
   
   Call InitialiseClock

End Property

Public Property Get ClockPlateColor() As OLE_COLOR
Attribute ClockPlateColor.VB_Description = "Returns/sets the clockplate color."

   ClockPlateColor = m_ClockPlateColor

End Property

Public Property Let ClockPlateColor(ByVal NewClockPlateColor As OLE_COLOR)

   m_ClockPlateColor = NewClockPlateColor
   PropertyChanged ("ClockPlateColor")
   
   Call InitialiseClock

End Property

Public Property Get ClockPlateGradientColor() As OLE_COLOR
Attribute ClockPlateGradientColor.VB_Description = "Returns/sets the clockplate gradient color."

   ClockPlateGradientColor = m_ClockPlateGradientColor

End Property

Public Property Let ClockPlateGradientColor(ByVal NewClockPlateGradientColor As OLE_COLOR)

   m_ClockPlateGradientColor = NewClockPlateGradientColor
   PropertyChanged ("ClockPlateGradientColor")
   
   If m_ClockPlateGradientStyle Then Call InitialiseClock

End Property

Public Property Get ClockPlateGradientPosition() As ClockPlateGradientPositions
Attribute ClockPlateGradientPosition.VB_Description = "Returns/sets the view position of the clockplate gradient."

   ClockPlateGradientPosition = m_ClockPlateGradientPosition

End Property

Public Property Let ClockPlateGradientPosition(ByVal NewClockPlateGradientPosition As ClockPlateGradientPositions)

   If NewClockPlateGradientPosition < TopLeft Or NewClockPlateGradientPosition > BottomRight Then Exit Property
   
   m_ClockPlateGradientPosition = NewClockPlateGradientPosition
   PropertyChanged ("ClockPlateGradientPosition")
   
   Call InitialiseClock

End Property

Public Property Get ClockPlateGradientStyle() As ClockPlateGradientStyles
Attribute ClockPlateGradientStyle.VB_Description = "Returns/sets the view style used to display the clockplate gradient."

   ClockPlateGradientStyle = m_ClockPlateGradientStyle

End Property

Public Property Let ClockPlateGradientStyle(ByVal NewClockPlateGradientStyle As ClockPlateGradientStyles)

   If m_ClockPlateGradientStyle And (NewClockPlateGradientStyle = NoGradient) Then picBackGround.Picture = Nothing
   
   m_ClockPlateGradientStyle = NewClockPlateGradientStyle
   PropertyChanged ("ClockPlateGradientStyle")
   
   Call InitialiseClock

End Property

Public Property Get ClockPlateStyle() As ClockStyles
Attribute ClockPlateStyle.VB_Description = "Returns/sets the style of the clockplate."

   ClockPlateStyle = m_ClockPlateStyle

End Property

Public Property Let ClockPlateStyle(ByVal NewClockPlateStyle As ClockStyles)

   m_ClockPlateStyle = NewClockPlateStyle
   PropertyChanged ("ClockPlateStyle")
   
   Call MakeTransparent
   Call InitialiseClock

End Property

Public Property Get DateBold() As Boolean
Attribute DateBold.VB_Description = "Returns/sets to switch the fontbold of the clock date on/off."

   DateBold = m_DateBold

End Property

Public Property Let DateBold(ByVal NewDateBold As Boolean)

   m_DateBold = NewDateBold
   PropertyChanged ("DateBold")
   
   Call InitialiseClock

End Property

Public Property Get DateColor() As OLE_COLOR
Attribute DateColor.VB_Description = "Returns/sets the foreground color of the clock date."

   DateColor = m_DateColor

End Property

Public Property Let DateColor(ByVal NewDateColor As OLE_COLOR)

   m_DateColor = NewDateColor
   PropertyChanged ("DateColor")
   
   Call InitialiseClock

End Property

Public Property Get DateStyle() As ClockDateStyles
Attribute DateStyle.VB_Description = "Returns/sets the style of the clock date."

   DateStyle = m_DateStyle

End Property

Public Property Let DateStyle(ByVal NewDateStyle As ClockDateStyles)

   m_DateStyle = NewDateStyle
   PropertyChanged ("DateStyle")
   
   Call InitialiseClock

End Property

Public Property Get DateTime() As String
Attribute DateTime.VB_Description = "Returns the actual date and time of the clock control."

   DateTime = m_DateTime

End Property

Public Property Get DigitsAmPm() As Boolean
Attribute DigitsAmPm.VB_Description = "Returns/sets to switch the Am/Pm of the clock control on/off."

   DigitsAmPm = m_DigitsAmPm

End Property

Public Property Let DigitsAmPm(ByVal NewDigitsAmPm As Boolean)

   m_DigitsAmPm = NewDigitsAmPm
   PropertyChanged ("DigitsAmPm")
   DigitsFormat = GetDigitsFormat
   
   Call InitialiseClock

End Property

Public Property Get DigitsBackColor() As OLE_COLOR
Attribute DigitsBackColor.VB_Description = "Returns/sets the background color for the digits of the clock control."

   DigitsBackColor = m_DigitsBackColor

End Property

Public Property Let DigitsBackColor(ByVal NewDigitsBackColor As OLE_COLOR)

   m_DigitsBackColor = NewDigitsBackColor
   PropertyChanged ("DigitsBackColor")
   
   Call InitialiseClock

End Property

Public Property Get DigitsBackStyle() As BackStyles
Attribute DigitsBackStyle.VB_Description = "Indicates the digits backstyle of the clock control is transparent or opaque."

   DigitsBackStyle = m_DigitsBackStyle

End Property

Public Property Let DigitsBackStyle(ByVal NewDigitsBackStyle As BackStyles)

   m_DigitsBackStyle = NewDigitsBackStyle
   PropertyChanged ("DigitsBackStyle")
   
   Call InitialiseClock

End Property

Public Property Get DigitsBold() As Boolean
Attribute DigitsBold.VB_Description = "Returns/sets to switch the digits fontbold of the clock date on/off."

   DigitsBold = m_DigitsBold

End Property

Public Property Let DigitsBold(ByVal NewDigitsBold As Boolean)

   m_DigitsBold = NewDigitsBold
   PropertyChanged ("DigitsBold")
   
   Call InitialiseClock

End Property

Public Property Get DigitsBorderStyle() As ClockBorderStyles
Attribute DigitsBorderStyle.VB_Description = "Returns/sets the digits border style of the clock control."

   DigitsBorderStyle = m_DigitsBorderStyle

End Property

Public Property Let DigitsBorderStyle(ByVal NewDigitsBorderStyle As ClockBorderStyles)

   m_DigitsBorderStyle = NewDigitsBorderStyle
   PropertyChanged ("DigitsBorderStyle")
   
   Call InitialiseClock

End Property

Public Property Get DigitsForeColor() As OLE_COLOR
Attribute DigitsForeColor.VB_Description = "Returns/sets the foreground color for the digits of the clock control."

   DigitsForeColor = m_DigitsForeColor

End Property

Public Property Let DigitsForeColor(ByVal NewDigitsForeColor As OLE_COLOR)

   m_DigitsForeColor = NewDigitsForeColor
   PropertyChanged ("DigitsForeColor")
   
   Call InitialiseClock

End Property

Public Property Get DigitsPosition() As ClockLabelPositions
Attribute DigitsPosition.VB_Description = "Returns/sets the position for the digits of the clock control."

   DigitsPosition = m_DigitsPosition

End Property

Public Property Let DigitsPosition(ByVal NewDigitsPosition As ClockLabelPositions)

   m_DigitsPosition = NewDigitsPosition
   PropertyChanged ("DigitsPosition")
   
   Call InitialiseClock

End Property

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Returns a Font object."

   Set Font = m_Font

End Property

Public Property Let Font(ByRef NewFont As StdFont)

   Set Font = NewFont

End Property

Public Property Set Font(ByRef NewFont As StdFont)

   Set m_Font = NewFont
   Set NewFont = Nothing
   
   With m_Font
      .Bold = False
      .Italic = False
      .Strikethrough = False
      .Underline = False
   End With
   
   PropertyChanged ("Font")
   picClock.Font = m_Font
   
   Call InitialiseClock

End Property

Public Property Get HandSecondColor() As OLE_COLOR
Attribute HandSecondColor.VB_Description = "Returns/sets the color for the second hand of the clock control."

   HandSecondColor = m_HandSecondColor

End Property

Public Property Let HandSecondColor(ByVal NewHandSecondColor As OLE_COLOR)

   m_HandSecondColor = NewHandSecondColor
   PropertyChanged ("HandSecondColor")
   
   Call InitialiseClock

End Property

Public Property Get HandSecondScroll() As ClockHandsMoves
Attribute HandSecondScroll.VB_Description = "Returns/sets the scrollstyle for the second hand of the clock control."

   HandSecondScroll = m_HandSecondScroll

End Property

Public Property Let HandSecondScroll(ByVal NewHandSecondScroll As ClockHandsMoves)

   m_HandSecondScroll = NewHandSecondScroll
   PropertyChanged ("HandSecondScroll")
   
   Call InitialiseClock

End Property

Public Property Get HandsHourMinuteBorderColor() As OLE_COLOR
Attribute HandsHourMinuteBorderColor.VB_Description = "Returns/sets the border color for the hour/minute hands of the clock control."

   HandsHourMinuteBorderColor = m_HandsHourMinuteBorderColor

End Property

Public Property Let HandsHourMinuteBorderColor(ByVal NewHandsHourMinuteBorderColor As OLE_COLOR)

   m_HandsHourMinuteBorderColor = NewHandsHourMinuteBorderColor
   PropertyChanged ("HandsHourMinuteBorderColor")
   
   Call InitialiseClock

End Property

Public Property Get HandsHourMinuteColor() As OLE_COLOR
Attribute HandsHourMinuteColor.VB_Description = "Returns/sets the inner color for the hour/minute hands of the clock control."

   HandsHourMinuteColor = m_HandsHourMinuteColor

End Property

Public Property Let HandsHourMinuteColor(ByVal NewHandsHourMinuteColor As OLE_COLOR)

   m_HandsHourMinuteColor = NewHandsHourMinuteColor
   PropertyChanged ("HandsHourMinuteColor")
   
   Call InitialiseClock

End Property

Public Property Get HandsHourMinuteScroll() As ClockHandsMoves
Attribute HandsHourMinuteScroll.VB_Description = "Returns/sets the scrollstyle for the hour/minute hands of the clock control."

   HandsHourMinuteScroll = m_HandsHourMinuteScroll

End Property

Public Property Let HandsHourMinuteScroll(ByVal NewHandsHourMinuteScrollScroll As ClockHandsMoves)

   m_HandsHourMinuteScroll = NewHandsHourMinuteScrollScroll
   PropertyChanged ("HandsHourMinuteScroll")
   
   Call InitialiseClock

End Property

Public Property Get HandsHourMinuteStyle() As ClockHandsHourMinuteStyles
Attribute HandsHourMinuteStyle.VB_Description = "Returns/sets the style for the hour/minute hands of the clock control."

   HandsHourMinuteStyle = m_HandsHourMinuteStyle

End Property

Public Property Let HandsHourMinuteStyle(ByVal NewHandsHourMinuteStyle As ClockHandsHourMinuteStyles)

   m_HandsHourMinuteStyle = NewHandsHourMinuteStyle
   PropertyChanged ("HandsHourMinuteStyle")
   
   Call InitialiseClock

End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."

   hWnd = UserControl.hWnd

End Property

Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Returns/sets to lock/unlock of the clock control to fasten the property changes."

   Locked = m_Locked

End Property

Public Property Let Locked(ByVal NewLocked As Boolean)

   m_Locked = NewLocked
   PropertyChanged ("Locked")
   
   Call InitialiseClock

End Property

Public Property Get MarksHourColor() As OLE_COLOR
Attribute MarksHourColor.VB_Description = "Returns/sets the color for the hour marks of the clock control."

   MarksHourColor = m_MarksHourColor

End Property

Public Property Let MarksHourColor(ByVal NewMarksHourColor As OLE_COLOR)

   m_MarksHourColor = NewMarksHourColor
   PropertyChanged ("MarksHourColor")
   
   Call InitialiseClock

End Property

Public Property Get MarksMinuteColor() As OLE_COLOR
Attribute MarksMinuteColor.VB_Description = "Returns/sets the color for the minute marks of the clock control."

   MarksMinuteColor = m_MarksMinuteColor

End Property

Public Property Let MarksMinuteColor(ByVal NewMarksMinuteColor As OLE_COLOR)

   m_MarksMinuteColor = NewMarksMinuteColor
   PropertyChanged ("MarksMinuteColor")
   
   Call InitialiseClock

End Property

Public Property Get MarksShows() As ClockMarksShows
Attribute MarksShows.VB_Description = "Returns/sets the type that will be shows by the marks of the clock control."

   MarksShows = m_MarksShows

End Property

Public Property Let MarksShows(ByVal NewMarksShows As ClockMarksShows)

   m_MarksShows = NewMarksShows
   PropertyChanged ("MarksShows")
   
   Call InitialiseClock

End Property

Public Property Get MarksStyle() As ClockMarksStyles
Attribute MarksStyle.VB_Description = "Returns/sets the marks style of the clock control."

   MarksStyle = m_MarksStyle

End Property

Public Property Let MarksStyle(ByVal NewMarksStyle As ClockMarksStyles)

   m_MarksStyle = NewMarksStyle
   PropertyChanged ("MarksStyle")
   
   Call InitialiseClock

End Property

Public Property Get NameBackColor() As OLE_COLOR
Attribute NameBackColor.VB_Description = "Returns/sets the background color for the name of the clock control."

   NameBackColor = m_NameBackColor

End Property

Public Property Let NameBackColor(ByVal NewNameBackColor As OLE_COLOR)

   m_NameBackColor = NewNameBackColor
   PropertyChanged ("NameBackColor")
   
   Call InitialiseClock

End Property

Public Property Get NameBackStyle() As BackStyles
Attribute NameBackStyle.VB_Description = "Indicates the name backstyle of the clock control is transparent or opaque."

   NameBackStyle = m_NameBackStyle

End Property

Public Property Let NameBackStyle(ByVal NewNameBackStyle As BackStyles)

   m_NameBackStyle = NewNameBackStyle
   PropertyChanged ("NameBackStyle")
   
   Call InitialiseClock

End Property

Public Property Get NameBold() As Boolean
Attribute NameBold.VB_Description = "Returns/sets to switch the name fontbold of the clock date on/off."

   NameBold = m_NameBold

End Property

Public Property Let NameBold(ByVal NewNameBold As Boolean)

   m_NameBold = NewNameBold
   PropertyChanged ("NameBold")
   
   Call InitialiseClock

End Property

Public Property Get NameBorderStyle() As ClockBorderStyles
Attribute NameBorderStyle.VB_Description = "Returns/sets the name border style of the clock control."

   NameBorderStyle = m_NameBorderStyle

End Property

Public Property Let NameBorderStyle(ByVal NewNameBorderStyle As ClockBorderStyles)

   m_NameBorderStyle = NewNameBorderStyle
   PropertyChanged ("NameBorderStyle")
   
   Call InitialiseClock

End Property

Public Property Get NameClock() As String
Attribute NameClock.VB_Description = "Returns/sets the name of the clock control that will be shows in the namelabel."

   NameClock = m_NameClock

End Property

Public Property Let NameClock(ByVal NewNameClock As String)

   m_NameClock = NewNameClock
   PropertyChanged ("NameClock")
   
   Call InitialiseClock

End Property

Public Property Get NameForeColor() As OLE_COLOR
Attribute NameForeColor.VB_Description = "Returns/sets the foreground color for the name of the clock control."

   NameForeColor = m_NameForeColor

End Property

Public Property Let NameForeColor(ByVal NewNameForeColor As OLE_COLOR)

   m_NameForeColor = NewNameForeColor
   PropertyChanged ("NameForeColor")
   
   Call InitialiseClock

End Property

Public Property Get NamePosition() As ClockLabelPositions
Attribute NamePosition.VB_Description = "Returns/sets the position for the name of the clock control."

   NamePosition = m_NamePosition

End Property

Public Property Let NamePosition(ByVal NewNamePosition As ClockLabelPositions)

   m_NamePosition = NewNamePosition
   PropertyChanged ("NamePosition")
   
   Call InitialiseClock

End Property

Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."

   Set Picture = m_Picture

End Property

Public Property Let Picture(ByVal NewPicture As Picture)

   Set Picture = NewPicture

End Property

Public Property Set Picture(ByVal NewPicture As Picture)

   If Not NewPicture Is Nothing Then
      If NewPicture = 0 Then Set NewPicture = Nothing
      
      m_ClockPlateGradientStyle = NoGradient
   End If
   
   Set m_Picture = NewPicture
   Set NewPicture = Nothing
   PropertyChanged "Picture"
   
   Call GeneratePicture
   Call InitialiseClock

End Property

Public Property Get ScaleSize() As Long
Attribute ScaleSize.VB_Description = "Returns the size of the clock control scale."

   ScaleSize = UserControl.ScaleWidth

End Property

Public Property Get SetDate() As Date
Attribute SetDate.VB_Description = "Returns/sets the starting date of the clock control."

   SetDate = m_SetDate

End Property

Public Property Let SetDate(ByVal NewStartDate As Date)

   m_SetDate = NewStartDate
   PropertyChanged ("SetDate")

End Property

Public Property Get SetTime() As String
Attribute SetTime.VB_Description = "Returns/sets the starting time of the clock control."

   SetTime = m_SetTime

End Property

Public Property Let SetTime(ByVal NewStartTime As String)

   m_SetTime = NewStartTime
   PropertyChanged ("SetTime")
   
   Call GenerateTimeValue

End Property

Public Property Get TimeZoneValue() As Long
Attribute TimeZoneValue.VB_Description = "Returns/sets the value for the timezone of the clock control."

   TimeZoneValue = m_TimeZoneValue

End Property

Public Property Let TimeZoneValue(ByVal NewTimeZoneValue As Long)

   m_TimeZoneValue = NewTimeZoneValue
   PropertyChanged ("TimeZoneValue")
   
   Call GenerateTimeValue
   Call InitialiseClock

End Property

Private Function CheckTime(ByVal IsTime As String) As String

Dim intCount As Integer
Dim intColon As Integer

   If Len(IsTime) Then
      On Local Error GoTo ErrorTime
      
      If Right(IsTime, 1) = ":" Then IsTime = IsTime & "0"
      
      For intCount = 1 To Len(IsTime)
         If Mid(IsTime, intCount, 1) = ":" Then intColon = intColon + 1
      Next 'intCount
      
      If intColon < 1 Then IsTime = IsTime & ":0"
      
      IsTime = Format(IsTime, "hh:mm:ss")
      intCount = Hour(IsTime) + Minute(IsTime) + Second(IsTime)
      CheckTime = IsTime
   End If
   
ErrorTime:
   On Local Error GoTo 0

End Function

Private Function GetDigitsFormat() As String

   GetDigitsFormat = "hh:mm:ss"
   
   If m_DigitsAmPm Then GetDigitsFormat = GetDigitsFormat & " AM/PM"

End Function

Private Function NextAction(ByVal Seconds As Single) As String

   NextAction = Format(DateAdd("s", Seconds, Time), "hh:mm:ss")

End Function

Private Sub CreateTimeNameBox(ByVal Index As Integer, ByVal Text As String, ByVal Position As Integer, ByVal TimePosition As Integer, ByVal ShowTime As Boolean, ByVal FontBold As Boolean)

Dim intSwitch As Integer

   If m_DigitsPosition + ShowTime Then intSwitch = TimePosition
   
   With picClock
      .Tag = .FontBold
      .FontBold = FontBold
      TimeNameBox(Index).BoxX = .Width / 2
      
      If Position = High Then
         TimeNameBox(Index).BoxY = .Height - ClockSize * 0.5 + ((.TextHeight("X") * 1.3) And ((Index = 1) And (intSwitch = High)))
         
      ElseIf Position = Low Then
         TimeNameBox(Index).BoxY = ClockSize * 0.5 - ((.TextHeight("X") * 1.3) And ((Index = 1) And (intSwitch = Low)))
      End If
      
      TimeNameBox(Index).TextX = (.Width - .TextWidth(Text)) / 2
      TimeNameBox(Index).TextY = TimeNameBox(Index).BoxY - .TextHeight("X") / 2
      .FontBold = Val(.Tag)
   End With

End Sub

Private Sub DrawClock()

   With picClock
      .Cls
      DoEvents
      ClockTimer = Timer + TimeValue
      ClockTimer = ClockTimer - (86400 And (ClockTimer > 86400))
      ClockHour = Int(ClockTimer) / 3600
      ClockMinute = Int(ClockTimer - Int(ClockHour) * 3600) / 60
      
      If m_HandSecondScroll Then
         ClockSecond = ClockTimer - Int(ClockHour) * 3600 - Int(ClockMinute) * 60
         
      Else
         ClockSecond = Int(ClockTimer - Int(ClockHour) * 3600 - Int(ClockMinute) * 60)
      End If
      
      m_DateTime = DateAdd("s", TimeValue, CDate(m_SetDate) + Time)
      TimeHour = Hour(m_DateTime)
      TimeMinute = Minute(m_DateTime)
      TimeSecond = Second(m_DateTime)
      
      If m_DateStyle Then
         Buffer(0) = FormatDateTime(m_DateTime, vbLongDate)
         Buffer(0) = CapsText(Buffer(0)) 'UCase(Left(Buffer(0), 1)) & Mid(Buffer(0), 2)
         
         If (TimeSecond > 14) And (TimeSecond < 45) Then
            DatePosition = False
            
         Else
            DatePosition = True
         End If
         
         If m_DateStyle = Down Then DatePosition = True
         If m_DateStyle = Up Then DatePosition = False
         
         If DatePosition Then
            Parts(0) = 270 + Len(Buffer(0)) / 2 * -6 + TextWidth("A") / -3
            
         Else
            Parts(0) = 90 + Len(Buffer(0)) / 2 * 6 + TextWidth("A") / 3
         End If
         
         For Counter = 1 To Len(Buffer(0))
            Parts(1) = -6 + (12 And DatePosition)
            Buffer(1) = Mid(Buffer(0), Counter, 1)
            Parts(1) = PI_PART * (Parts(0) + Counter * Parts(1))
            ClockX(0) = ClockSize + Cos(Parts(1)) * Lengths(3)
            ClockY(0) = ClockSize + Sin(Parts(1)) * Lengths(3)
            .FontBold = m_DateBold
            .ForeColor = m_DateColor
            .CurrentX = ClockX(0) - .TextWidth(Buffer(1)) / 2
            .CurrentY = ClockY(0) - .TextHeight("X") / 2
            picClock.Print Buffer(1)
         Next 'Counter
      End If
      
      If Not m_Active Then
         ClockHour = 0
         ClockMinute = 0
         ClockSecond = 0
         TimeHour = 0
         TimeMinute = 0
         TimeSecond = 0
         m_DateTime = ""
         
      ElseIf m_AlarmTime & ":00" = Format(m_DateTime, "hh:mm:ss") Then
         If Not GiveAlarm Then
            RaiseEvent Alarm
            GiveAlarm = True
         End If
         
      Else
         GiveAlarm = False
      End If
      
      If m_DigitsPosition Then Call PrintDigitNameBox(0, Format(TimeSerial(TimeHour, TimeMinute, TimeSecond), DigitsFormat), m_DigitsBackStyle, m_DigitsBold, m_DigitsForeColor, m_DigitsBackColor, m_DigitsBorderStyle)
      If m_NamePosition Then Call PrintDigitNameBox(1, m_NameClock, m_NameBackStyle, m_NameBold, m_NameForeColor, m_NameBackColor, m_NameBorderStyle)
      
      If m_HandsHourMinuteScroll = Smooth Then
         MoveHour = ClockMinute \ 2
         MoveMinute = ClockSecond \ 10
         
      Else
         MoveHour = 0
         MoveMinute = 0
      End If
       
      Call MakeClockHands(LengthHour, Lengths(0), Int(ClockHour) * 30 + MoveHour, m_HandsHourMinuteColor, True)
      Call MakeClockHands(LengthMinute, Lengths(0), Int(ClockMinute) * 6 + MoveMinute, m_HandsHourMinuteColor, True)
      Call MakeClockHands(LengthSecond, 0, ClockSecond * 6, m_HandSecondColor, False)
      
      .FillStyle = vbFSSolid
      .FillColor = m_HandSecondColor
      picClock.Circle (ClockSize, ClockSize), Lengths(2) - (7 And (m_HandsHourMinuteStyle > 1)) / 10, m_HandSecondColor
      ExtFloodFill .hDC, ClockSize, ClockSize, m_HandSecondColor, FLOODFILLBORDER
      
      If m_ClockHourBell + m_Active = -2 Then
         If (CountBells = -1) And (TimeSecond < 2) And Not GiveBell Then
            If TimeMinute = 0 Then
               CountBells = TimeHour - 1 - (12 And (TimeHour > 12)) + (11 And (TimeHour = 0))
               NextBell = NextAction(0)
               
            ElseIf TimeMinute = 30 Then
               CountBells = 0
               NextBell = NextAction(0)
            End If
         End If
         
         If CountBells > -1 Then
            If Format(Time, "hh:mm:ss") = NextBell Then
               Beep
               GiveBell = True
               CountBells = CountBells - 1
               NextBell = NextAction(1.25)
            End If
            
         ElseIf TimeSecond > 2 Then
            GiveBell = False
         End If
      End If
      
      If m_ClockMirrored = NoMirror Then
         BitBlt hDC, 2, 2, .ScaleWidth, .ScaleWidth, .hDC, 0, 0, vbSrcCopy
         
      Else
         If m_ClockMirrored = HorizontalMirror Then
            StretchBlt hDC, .ScaleWidth + 2, 2, -.ScaleWidth, .ScaleWidth, .hDC, 0, 0, .ScaleWidth, .ScaleWidth, vbSrcCopy
            
         ElseIf m_ClockMirrored = VerticalMirror Then
            StretchBlt hDC, 2, .ScaleWidth + 2, .ScaleWidth, -.ScaleWidth, .hDC, 0, 0, .ScaleWidth, .ScaleWidth, vbSrcCopy
            
         Else
            StretchBlt hDC, .ScaleWidth + 2, .ScaleWidth + 2, -.ScaleWidth, -.ScaleWidth, .hDC, 0, 0, .ScaleWidth, .ScaleWidth, vbSrcCopy
         End If
      End If
   End With

End Sub

Private Sub DrawClockGradient(ByVal ColorBegin As Long, ByVal ColorEnd As Long)

Dim intCircle   As Integer
Dim intStart    As Integer
Dim intX        As Integer
Dim intY        As Integer
Dim sngBlue(1)  As Single
Dim sngGreen(1) As Single
Dim sngRed(1)   As Single

   With picBackGround
      ' calculate start position
      intX = .ScaleWidth / 3.5
      intY = .ScaleHeight / 3.5
      intStart = Sqr(((.ScaleWidth - intX) ^ 2) + ((.ScaleHeight - intY) ^ 2))
      
      ' set the gradient light color position
      If m_ClockPlateGradientPosition = TopCenter Or m_ClockPlateGradientPosition = Center Or m_ClockPlateGradientPosition = BottomCenter Then intX = .ScaleWidth / 2
      If m_ClockPlateGradientPosition = TopRight Or m_ClockPlateGradientPosition = CenterRight Or m_ClockPlateGradientPosition = BottomRight Then intX = .ScaleWidth - intX
      If m_ClockPlateGradientPosition > TopRight And m_ClockPlateGradientPosition < BottomLeft Then intY = .ScaleHeight / 2
      If m_ClockPlateGradientPosition > CenterRight Then intY = .ScaleHeight - intY
      
      If m_ClockPlateGradientStyle = OutToIn Then
         ' swap the colors
         ColorBegin = ColorBegin Xor ColorEnd
         ColorEnd = ColorBegin Xor ColorEnd
         ColorBegin = ColorBegin Xor ColorEnd
      End If
      
      ' fill the RGB color values
      Call LongToRGB(ColorEnd, sngRed(0), sngGreen(0), sngBlue(0))
      Call LongToRGB(ColorBegin, sngRed(1), sngGreen(1), sngBlue(1))
      
      sngRed(1) = (sngRed(1) - sngRed(0)) / intStart
      sngGreen(1) = (sngGreen(1) - sngGreen(0)) / intStart
      sngBlue(1) = (sngBlue(1) - sngBlue(0)) / intStart
      .FillStyle = vbFSSolid
      .DrawStyle = vbInvisible
      
      ' draw gradient circles
      For intCircle = intStart - 1 To 0 Step -1
         .FillColor = RGB(sngRed(0) + (sngRed(1) * intCircle), sngGreen(0) + (sngGreen(1) * intCircle), sngBlue(0) + (sngBlue(1) * intCircle))
         picBackGround.Circle (intX, intY), intCircle
      Next 'intCircle
      
      ' fill the background settings
      .FillStyle = vbFSTransparent
      .DrawStyle = vbSolid
      
      If m_BackStyle = Opaque Then
         .DrawWidth = 4
         .ForeColor = .BackColor
         
         ' make the clock background
         For intCircle = .ScaleWidth / 2 + 1 To .ScaleWidth - .ScaleWidth / 4 Step 2
            picBackGround.Circle (.ScaleWidth / 2, .ScaleHeight / 2), intCircle - 1
         Next 'intCircle
      End If
      
      .DrawWidth = 1
   End With
   
   Erase sngBlue, sngGreen, sngRed

End Sub

Private Sub DrawLines(ByVal HandColor As Long)

   picClock.Line (ClockX(3), ClockY(3))-(ClockX(1), ClockY(1)), HandColor
   picClock.Line (ClockX(1), ClockY(1))-(ClockX(0), ClockY(0)), HandColor
   picClock.Line (ClockX(0), ClockY(0))-(ClockX(2), ClockY(2)), HandColor
   picClock.Line (ClockX(2), ClockY(2))-(ClockX(3), ClockY(3)), HandColor

End Sub

Private Sub GeneratePicture()

   With picBackGround
      If m_Picture Is Nothing Then
         .Picture = Nothing
         .Width = 161
         .Height = 161
         
      Else
         .Picture = m_Picture
      End If
   End With

End Sub

Private Sub GenerateTimeValue()

Dim lngValue As Long

   m_SetTime = Format(CheckTime(m_SetTime), "hh:mm:ss")
   
   If Len(m_SetTime) Then lngValue = Hour(m_SetTime) * 3600 + Minute(m_SetTime) * 60 + Second(m_SetTime) - Int(Timer)
   
   TimeValue = m_TimeZoneValue * 60 + lngValue

End Sub

Private Sub InitialiseClock()

'Const BDR_EDGED  As Long = &H16
'Const BDR_RAISED As Long = &H5
'Const BDR_SUNKEN As Long = &HA
'Const BF_RIGHT   As Long = &H4
'Const BF_TOP     As Long = &H2
'Const BF_LEFT    As Long = &H1
'Const BF_BOTTOM  As Long = &H8
'Const BF_RECT    As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Dim lngColor     As Long

   If m_Locked Then Exit Sub
   If m_SetDate = "" Then m_SetDate = Date
   If m_ClockPlateGradientStyle Then Call DrawClockGradient(m_ClockPlateGradientColor, m_ClockPlateColor)
   
   With picClock
     .Picture = Nothing
      .BackColor = m_BackColor
      UserControl.BackColor = .BackColor
      picClock.Scale (0, 0)-(.Width, .Height)
      AutoRedraw = True
      Cls
      DrawEdge hDC, FrameRect, Choose(m_BorderStyle + 1, NoBorder, BDR_RAISED, BDR_SUNKEN, BDR_EDGED), BF_RECT
      AutoRedraw = False
      
      If m_ClockPlateBackStyle = Opaque Then
         If m_ClockPlateStyle = CircleClock Then
            .FillStyle = vbFSSolid
            .FillColor = m_ClockPlateColor
            .DrawWidth = 2
            picClock.Circle (ClockSize, ClockSize), ClockSize - 2, m_ClockPlateColor
            .DrawWidth = 1
            
         Else
            .BackColor = m_ClockPlateColor
         End If
      End If
      
      If picBackGround.Image Then
         .PaintPicture picBackGround.Image, 0, 0, .ScaleWidth * 2.031, .ScaleWidth * 2.031, 0, 0, picBackGround.ScaleWidth * 2.031, picBackGround.ScaleHeight * 2.031, vbSrcCopy
         
         If m_ClockPlateStyle = CircleClock Then
            If m_ClockPlateBackStyle = Opaque Then
               lngColor = m_BackColor
               
            Else
               lngColor = m_ClockPlateColor
            End If
            
            .FillStyle = vbFSTransparent
            .DrawWidth = ClockSize * 2
            picClock.Circle (ClockSize, ClockSize), ClockSize * 2 - 3, lngColor
            .DrawWidth = 1
         End If
      End If
      
      If m_ClockPlateBorder Then
         .DrawWidth = 2 + (2 And m_ClockPlateBorderBold)
         .ForeColor = m_ClockPlateBorderColor
         
         If m_ClockPlateStyle = CircleClock Then
            .FillStyle = vbFSTransparent
            picClock.Circle (ClockSize, ClockSize), ClockSize - 2
            
         Else
            TempValue = .DrawWidth / 2
            picClock.Line (TempValue, TempValue)-(.Width - TempValue, TempValue)
            picClock.Line -(.Width - TempValue, .Height - TempValue)
            picClock.Line -(TempValue, .Height - TempValue)
            picClock.Line -(TempValue, TempValue)
         End If
         
         .DrawWidth = 1
      End If
      
      .FillStyle = vbFSSolid
      ClockWidth = 0 + (75 And (Not m_ClockPlateBorder And (m_DateStyle = NoDate))) / 1000
      ClockWidth = ClockWidth + 0.84 - (2 And m_ClockBold) / 100
      
      For Counter = 0 To 359
         Parts(0) = PI_PART * (90 - Counter)
         ClockX(0) = ClockSize + Cos(Parts(0)) * ClockWidth * ClockSize
         ClockY(0) = ClockSize + Sin(Parts(0)) * ClockWidth * ClockSize
         ClockX(1) = ClockX(0) + 13
         ClockY(1) = ClockY(0) + 13
         Parts(1) = ClockSize * (ClockWidth - (0.11 + (1 And m_ClockBold) / 100))
         TempValue = 6 - Counter / 30 + (12 And (Counter / 30 > 5))
         
         If Counter Mod 30 = 0 Then
            If m_ClockHours Then
               Parts(2) = 2.25 - (9 And (TempValue = 12)) / 10
               .FontBold = m_ClockBold
               .CurrentX = ClockSize + Cos(Parts(0)) * Parts(1) - .TextHeight("X") / Parts(2) + Abs(Not m_ClockBold)
               .CurrentY = ClockSize + Sin(Parts(0)) * Parts(1) - .TextHeight("X") / 2
               .ForeColor = m_ClockHoursColor
               
               If (m_MarksShows = Quarters) And Counter Mod 90 Then TempValue = 0
               If TempValue Then picClock.Print TempValue
            End If
            
            If (m_MarksShows = Quarters) And Counter Mod 90 Then ClockX(0) = -1
            
            If ClockX(0) > -1 Then
               Select Case m_MarksStyle
                  Case CircleMarks
                     .FillColor = m_MarksHourColor
                     picClock.Circle (ClockX(0), ClockY(0)), 1 + Abs(m_ClockBold) * 2.1, m_MarksHourColor
                     
                  Case SquareMarks
                     Call MakeRectangle(ClockX(0), ClockY(0), 0.017 * ClockSize, 0.017 * ClockSize, m_MarksHourColor, Sunken)
                     
                  Case LineMarks
                     If Counter Mod 90 = 0 Then
                        If (TempValue = 3) Or (TempValue = 9) Then
                           Parts(2) = 0.017
                           Parts(3) = 0.042
                           
                        Else
                           Parts(2) = 0.042
                           Parts(3) = 0.017
                        End If
                        
                        Call MakeRectangle(ClockX(0), ClockY(0), Parts(2) * ClockSize, Parts(3) * ClockSize, m_MarksHourColor, NoBorder)
                     End If
                     
                     If Counter Mod 90 = 0 Then
                        lngColor = m_MarksMinuteColor
                        
                     Else
                        lngColor = m_MarksHourColor
                     End If
                     
                     picClock.Line (ClockX(0) + Cos(Parts(0)) * 4, ClockY(0) + Sin(Parts(0)) * 4)-(ClockX(0) + Cos(Parts(0)) * -5, ClockY(0) + Sin(Parts(0)) * -5), lngColor
               End Select
            End If
            
         ElseIf (Counter Mod 6 = 0) And (m_MarksShows = HoursMinutes) Then
            If m_MarksStyle = CircleMarks Then
               .FillColor = m_MarksMinuteColor
               picClock.Circle (ClockX(0), ClockY(0)), 1 + Abs(m_ClockBold) * 1.1, m_MarksMinuteColor
               
            ElseIf (m_MarksStyle = SquareMarks) Or (m_MarksStyle = LineMarks) Then
               picClock.Line (ClockX(0) + m_ClockBold, ClockY(0) + m_ClockBold)-(ClockX(0) - m_ClockBold, ClockY(0) - m_ClockBold), m_MarksMinuteColor, BF
            End If
         End If
      Next 'Counter
      
      With imgToolTipText
         Scale (0, 0)-(ScaleWidth, Abs(ScaleHeight))
         
         If (m_DigitsPosition + m_NamePosition = 4) And m_DigitsPosition Then
            .Top = .Height + ClockSize * 0.25
            
         Else
            .Top = ClockSize + ClockSize * 0.15
         End If
         
         .Left = ClockSize - picAlarm.Width / 2
         .Height = picAlarm.Height + 4
         .Width = picAlarm.Width + 4
         .Visible = CBool(Len(.ToolTipText))
         
         If Len(m_AlarmTime) Then picClock.PaintPicture picAlarm.Picture, .Left, .Top
      End With
      
      Lengths(3) = 5 - (5 And m_ClockPlateBorder)
      LengthHour = (0.52 - (9 And m_ClockBold) / 100) * (ClockSize + Lengths(3))
      LengthMinute = (0.7 - (9 And m_ClockBold) / 100) * (ClockSize + Lengths(3))
      LengthSecond = (0.76 - (9 And m_ClockBold) / 100) * (ClockSize + Lengths(3))
      Lengths(3) = Lengths(3) * 3
      Lengths(0) = 0.04 * (ClockSize + Lengths(3))
      Lengths(1) = 0.11 * (ClockSize + Lengths(3))
      Lengths(2) = 0.017 * (ClockSize + Lengths(3)) * 1.7
      Lengths(3) = (0.92 - (2 And m_ClockBold) / 100) * ClockSize
      .Picture = .Image
      .FillStyle = vbFSSolid
      picClock.Scale (0, .Height)-(.Width, 0)
      Scale (0, Abs(ScaleHeight))-(ScaleWidth, 0)
      
      Call CreateTimeNameBox(0, Format("00:00:00", DigitsFormat), m_DigitsPosition, False, True, m_DigitsBold)
      Call CreateTimeNameBox(1, m_NameClock, m_NamePosition, m_DigitsPosition, False, m_NameBold)
   End With
   
   If m_Active Then
      Call DrawClock
      
   Else
      AutoRedraw = True
      
      Call DrawClock
      
      Refresh
      AutoRedraw = False
   End If

End Sub

Private Sub LongToRGB(ByVal ColorRGB As OLE_COLOR, ByRef Red As Single, ByRef Green As Single, ByRef Blue As Single)

   Red = ColorRGB And &HFF&
   Green = (ColorRGB \ &H100) And &HFF&
   Blue = (ColorRGB \ &H10000) And &HFF&

End Sub

Private Sub MakeClockHands(ByVal Length As Integer, ByVal HandBold As Integer, ByVal HandTime As Single, ByVal HandColor As Long, ByVal IsHourMinute As Boolean)

   With picClock
      Parts(0) = PI_PART * (90 - HandTime)
      Parts(1) = PI_PART * (180 - HandTime)
      Parts(2) = PI_PART * -HandTime
      
      If m_HandsHourMinuteStyle = Bold Then
         ClockX(0) = ClockSize + Cos(Parts(0)) * Length
         ClockY(0) = ClockSize + Sin(Parts(0)) * Length
         ClockX(1) = ClockSize + Cos(Parts(1)) * HandBold
         ClockY(1) = ClockSize + Sin(Parts(1)) * HandBold
         ClockX(2) = ClockSize + Cos(Parts(2)) * HandBold
         ClockY(2) = ClockSize + Sin(Parts(2)) * HandBold
         ClockX(3) = ClockSize - Cos(Parts(0)) * Lengths(1)
         ClockY(3) = ClockSize - Sin(Parts(0)) * Lengths(1)
         
         Call DrawLines(1)
         
         With picClock
            .FillColor = m_HandsHourMinuteColor
            ExtFloodFill .hDC, ClockSize, ClockSize, 1, FLOODFILLBORDER
         End With
         
         If IsHourMinute Then
            Parts(3) = m_HandsHourMinuteBorderColor
            
         Else
            Parts(3) = m_HandSecondColor
         End If
         
         Call DrawLines(Parts(3))
         
      Else
         If IsHourMinute Then
            .DrawWidth = Choose(m_HandsHourMinuteStyle, 5, 3, 1)
            
         Else
            .DrawWidth = 1
         End If
         
         picClock.Line (ClockSize, ClockSize)-(ClockSize + Cos(Parts(0)) * Length, ClockSize + Sin(Parts(0)) * Length), HandColor
         picClock.Line (ClockSize, ClockSize)-(ClockSize - Cos(Parts(0)) * Lengths(1), ClockSize - Sin(Parts(0)) * Lengths(1)), HandColor
      End If
   End With

End Sub

Private Sub MakeRectangle(ByVal X As Single, ByVal Y As Single, ByVal TextHeight As Single, ByVal TextWidth As Single, ByVal HandColor As Long, ByVal BoxBorderStyle As Integer)

   picClock.Line (X - TextWidth, Y - TextHeight)-(X + TextWidth, Y + TextHeight), HandColor, BF
   
   If BoxBorderStyle Then
      If BoxBorderStyle = Raised Then
         Parts(3) = vb3DHighlight
         
      Else
         Parts(3) = vb3DShadow
      End If
      
      picClock.Line (X - TextWidth, Y - TextHeight)-(X - TextWidth, Y + TextHeight), Parts(3)
      picClock.Line -(X + TextWidth, Y + TextHeight), Choose(BoxBorderStyle, vb3DShadow, vb3DHighlight, vb3DShadow)
      picClock.Line -(X + TextWidth, Y - TextHeight), Choose(BoxBorderStyle, vb3DShadow, vb3DHighlight, vb3DShadow)
      picClock.Line -(X - TextWidth, Y - TextHeight), Parts(3)
   End If

End Sub

Private Sub MakeTransparent()

   SetWindowRgn UserControl.hWnd, CreateRectRgn(0, 0, UserControl.Width, UserControl.Height), True
   
   If m_BackStyle = Transparent Then SetWindowRgn UserControl.hWnd, CreateEllipticRgn(2, 2, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), True

End Sub

Private Sub PrintDigitNameBox(ByVal Index As Integer, ByVal Text As String, ByVal BackStyle As Integer, ByVal ShowBold As Boolean, ByVal LabelForeColor As Long, ByVal LabelBackColor As Long, ByVal BoxBorderStyle As Integer)

   With picClock
      .DrawWidth = 1
      .ForeColor = LabelForeColor
      .FontBold = ShowBold
      
      If BackStyle = Opaque Then Call MakeRectangle(TimeNameBox(Index).BoxX, TimeNameBox(Index).BoxY, 1.1 * .TextHeight("X") / 2, 1.05 * .TextWidth(Text) / 2, LabelBackColor, BoxBorderStyle)
      
      .CurrentX = TimeNameBox(Index).TextX
      .CurrentY = TimeNameBox(Index).TextY
      picClock.Print Text
   End With

End Sub

Private Sub Resize()

   Scale (0, 0)-(ScaleWidth, Abs(ScaleHeight))
   
   With FrameRect
      .Top = 0
      .Left = 0
      .Right = ScaleWidth
      .Bottom = ScaleHeight
   End With
   
   With picClock
      CountBells = -1
      Height = Width
      ClockSize = (Width / Screen.TwipsPerPixelX - 5) / 2
      .Width = ClockSize * 2
      .Height = .Width
   End With
   
   Call InitialiseClock

End Sub

Private Sub SetToolTipText()

   If Len(m_AlarmTime & m_AlarmToolTipText) Then
      If InStr(m_AlarmToolTipText, "#") Then
         imgToolTipText.ToolTipText = Split(m_AlarmToolTipText, "#")(0) & m_AlarmTime & Split(m_AlarmToolTipText, "#", 2)(1)
         
      Else
         imgToolTipText.ToolTipText = m_AlarmToolTipText & " " & m_AlarmTime
      End If
      
   Else
      imgToolTipText.ToolTipText = ""
   End If

End Sub

Private Sub imgToolTipText_Click()

   RaiseEvent Click

End Sub

Private Sub imgToolTipText_DblClick()

   RaiseEvent DblClick

End Sub

Private Sub imgToolTipText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Private Sub imgToolTipText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

Private Sub imgToolTipText_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

   RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub tmrClock_Timer()

   Call DrawClock

End Sub

Private Sub UserControl_Click()

   RaiseEvent Click

End Sub

Private Sub UserControl_DblClick()

   RaiseEvent DblClick

End Sub

Private Sub UserControl_Initialize()

   tmrClock.Enabled = False

End Sub

Private Sub UserControl_InitProperties()

   m_BackColor = Ambient.BackColor
   m_BackStyle = Opaque
   m_BorderStyle = Raised
   m_ClockHours = True
   m_ClockHoursColor = CLOCK_HOURSCOLOR
   m_ClockPlateBackStyle = Opaque
   m_ClockPlateBorder = True
   m_ClockPlateBorderColor = CLOCKPLATE_BORDERCOLOR
   m_ClockPlateColor = CLOCKPLATE_COLOR
   m_ClockPlateGradientColor = GRADIENT_COLOR
   m_ClockPlateGradientPosition = Center
   m_DateColor = DATE_COLOR
   m_DigitsAmPm = True
   m_DigitsBackColor = DIGITS_BACKCOLOR
   m_DigitsBackStyle = Opaque
   m_DigitsBorderStyle = Sunken
   m_DigitsForeColor = DIGITS_FORECOLOR
   m_DigitsPosition = High
   Set m_Font = Ambient.Font
   picClock.FontSize = m_Font.Size
   m_HandSecondColor = HAND_SECONDCOLOR
   m_HandSecondScroll = Smooth
   m_HandsHourMinuteBorderColor = HANDS_HOURMINUTEBORDERCOLOR
   m_HandsHourMinuteColor = HANDS_HOURMINUTECOLOR
   m_HandsHourMinuteScroll = Smooth
   m_MarksHourColor = MARKS_HOURCOLOR
   m_MarksMinuteColor = MARKS_MINUTECOLOR
   m_MarksStyle = CircleMarks
   m_NameBackColor = NAME_BACKCOLOR
   m_NameBackStyle = Opaque
   m_NameBorderStyle = Sunken
   m_NameForeColor = NAME_FORECOLOR
   Set m_Picture = Nothing
   m_SetDate = Date
   m_SetTime = ""
   CountBells = -1
   m_DateTime = ""
   DigitsFormat = GetDigitsFormat
   picClock.Font = m_Font
   
   Call GenerateTimeValue

End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)

   If (KeyAscii = vbKeySpace) Or (KeyAscii = vbKeyReturn) Then RaiseEvent DblClick

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
      m_Active = .ReadProperty("Active", False)
      m_AlarmTime = .ReadProperty("AlarmTime", "")
      m_AlarmToolTipText = .ReadProperty("AlarmToolTipText", "")
      m_BackColor = .ReadProperty("BackColor", Ambient.BackColor)
      m_BackStyle = .ReadProperty("BackStyle", Opaque)
      m_BorderStyle = .ReadProperty("ClockBorderStyles", Raised)
      m_ClockBold = .ReadProperty("ClockBold", False)
      m_ClockHourBell = .ReadProperty("ClockHourBell", False)
      m_ClockHours = .ReadProperty("ClockHours", True)
      m_ClockHoursColor = .ReadProperty("ClockHoursColor", CLOCK_HOURSCOLOR)
      m_ClockMirrored = .ReadProperty("ClockMirrored", NoMirror)
      m_ClockPlateBackStyle = .ReadProperty("ClockPlateBackStyle", Opaque)
      m_ClockPlateBorder = .ReadProperty("ClockPlateBorder", True)
      m_ClockPlateBorderBold = .ReadProperty("ClockPlateBorderBold", False)
      m_ClockPlateBorderColor = .ReadProperty("ClockPlateBorderColor", CLOCKPLATE_BORDERCOLOR)
      m_ClockPlateColor = .ReadProperty("ClockPlateColor", CLOCKPLATE_COLOR)
      m_ClockPlateGradientColor = .ReadProperty("ClockPlateGradientColor", GRADIENT_COLOR)
      m_ClockPlateGradientPosition = .ReadProperty("ClockPlateGradientPosition", Center)
      m_ClockPlateGradientStyle = .ReadProperty("ClockPlateGradientStyle", NoGradient)
      m_ClockPlateStyle = .ReadProperty("ClockPlateStyle", CircleClock)
      m_DateBold = .ReadProperty("DateBold", False)
      m_DateColor = .ReadProperty("DateColor", DATE_COLOR)
      m_DateStyle = .ReadProperty("DateStyle", NoDate)
      m_DigitsAmPm = .ReadProperty("DigitsAmPm", True)
      m_DigitsBackColor = .ReadProperty("DigitsBackColor", DIGITS_BACKCOLOR)
      m_DigitsBackStyle = .ReadProperty("DigitsBackStyle", Opaque)
      m_DigitsBold = .ReadProperty("DigitsBold", False)
      m_DigitsBorderStyle = .ReadProperty("DigitsBorderStyle", Sunken)
      m_DigitsForeColor = .ReadProperty("DigitsForeColor", DIGITS_FORECOLOR)
      m_DigitsPosition = .ReadProperty("DigitsPosition", High)
      Set m_Font = .ReadProperty("Font", Ambient.Font)
      picClock.FontSize = m_Font.Size
      m_Locked = .ReadProperty("Locked", False)
      m_HandSecondColor = .ReadProperty("HandSecondColor", HAND_SECONDCOLOR)
      m_HandSecondScroll = .ReadProperty("HandSecondScroll", Smooth)
      m_HandsHourMinuteBorderColor = .ReadProperty("HandsHourMinuteBorderColor", HANDS_HOURMINUTEBORDERCOLOR)
      m_HandsHourMinuteColor = .ReadProperty("HandsHourMinuteColor", HANDS_HOURMINUTECOLOR)
      m_HandsHourMinuteScroll = .ReadProperty("HandsHourMinuteScroll", Smooth)
      m_HandsHourMinuteStyle = .ReadProperty("HandsHourMinuteStyle", Bold)
      m_MarksHourColor = .ReadProperty("MarksHourColor", MARKS_HOURCOLOR)
      m_MarksMinuteColor = .ReadProperty("MarksMinuteColor", MARKS_MINUTECOLOR)
      m_MarksShows = .ReadProperty("MarksShows", HoursMinutes)
      m_MarksStyle = .ReadProperty("MarksStyle", CircleMarks)
      m_NameBackColor = .ReadProperty("NameBackColor", NAME_BACKCOLOR)
      m_NameBackStyle = .ReadProperty("NameBackStyle", Opaque)
      m_NameBold = .ReadProperty("NameBold", False)
      m_NameBorderStyle = .ReadProperty("NameBorderStyle", Sunken)
      m_NameClock = .ReadProperty("NameClock", "")
      m_NameForeColor = .ReadProperty("NameForeColor", NAME_FORECOLOR)
      m_NamePosition = .ReadProperty("NamePosition", Off)
      Set m_Picture = .ReadProperty("Picture", picBackGround.Picture)
      m_SetDate = .ReadProperty("SetDate", Date)
      m_SetTime = .ReadProperty("SetTime", "")
      m_TimeZoneValue = .ReadProperty("TimeZoneValue", 0)
      DigitsFormat = GetDigitsFormat
      picClock.Font = m_Font
   End With
   
   Call SetToolTipText
   Call GenerateTimeValue
   Call GeneratePicture
   Call InitialiseClock
   
   If m_BackStyle = Opaque Then Call MakeTransparent

End Sub

Private Sub UserControl_Resize()

   If Width < 1812 Then Width = 1812
   
   Call MakeTransparent
   Call Resize

End Sub

Private Sub UserControl_Terminate()

   Set m_Font = Nothing
   Set m_Picture = Nothing
   Erase Lengths, ClockX, ClockY, Parts, Buffer, TimeNameBox

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   With PropBag
      .WriteProperty "Active", m_Active, False
      .WriteProperty "AlarmTime", m_AlarmTime, ""
      .WriteProperty "AlarmToolTipText", m_AlarmToolTipText, ""
      .WriteProperty "BackColor", m_BackColor, Ambient.BackColor
      .WriteProperty "BackStyle", m_BackStyle, Opaque
      .WriteProperty "ClockBorderStyles", m_BorderStyle, Raised
      .WriteProperty "ClockBold", m_ClockBold, False
      .WriteProperty "ClockHourBell", m_ClockHourBell, False
      .WriteProperty "ClockHours", m_ClockHours, True
      .WriteProperty "ClockHoursColor", m_ClockHoursColor, CLOCK_HOURSCOLOR
      .WriteProperty "ClockMirrored", m_ClockMirrored, NoMirror
      .WriteProperty "ClockPlateBackStyle", m_ClockPlateBackStyle, Opaque
      .WriteProperty "ClockPlateBorder", m_ClockPlateBorder, True
      .WriteProperty "ClockPlateBorderBold", m_ClockPlateBorderBold, False
      .WriteProperty "ClockPlateBorderColor", m_ClockPlateBorderColor, CLOCKPLATE_BORDERCOLOR
      .WriteProperty "ClockPlateColor", m_ClockPlateColor, CLOCKPLATE_COLOR
      .WriteProperty "ClockPlateGradientColor", m_ClockPlateGradientColor, GRADIENT_COLOR
      .WriteProperty "ClockPlateGradientPosition", m_ClockPlateGradientPosition, Center
      .WriteProperty "ClockPlateGradientStyle", m_ClockPlateGradientStyle, NoGradient
      .WriteProperty "ClockPlateStyle", m_ClockPlateStyle, CircleClock
      .WriteProperty "DateBold", m_DateBold, False
      .WriteProperty "DateColor", m_DateColor, DATE_COLOR
      .WriteProperty "DateStyle", m_DateStyle, NoDate
      .WriteProperty "DigitsAmPm", m_DigitsAmPm, True
      .WriteProperty "DigitsBackColor", m_DigitsBackColor, DIGITS_BACKCOLOR
      .WriteProperty "DigitsBackStyle", m_DigitsBackStyle, Opaque
      .WriteProperty "DigitsBold", m_DigitsBold, False
      .WriteProperty "DigitsBorderStyle", m_DigitsBorderStyle, Sunken
      .WriteProperty "DigitsForeColor", m_DigitsForeColor, DIGITS_FORECOLOR
      .WriteProperty "DigitsPosition", m_DigitsPosition, High
      .WriteProperty "Font", m_Font, Ambient.Font
      .WriteProperty "Locked", m_Locked, False
      .WriteProperty "HandSecondColor", m_HandSecondColor, HAND_SECONDCOLOR
      .WriteProperty "HandSecondScroll", m_HandSecondScroll, Smooth
      .WriteProperty "HandsHourMinuteBorderColor", m_HandsHourMinuteBorderColor, HANDS_HOURMINUTEBORDERCOLOR
      .WriteProperty "HandsHourMinuteColor", m_HandsHourMinuteColor, HANDS_HOURMINUTECOLOR
      .WriteProperty "HandsHourMinuteScroll", m_HandsHourMinuteScroll, Smooth
      .WriteProperty "HandsHourMinuteStyle", m_HandsHourMinuteStyle, Bold
      .WriteProperty "MarksHourColor", m_MarksHourColor, MARKS_HOURCOLOR
      .WriteProperty "MarksMinuteColor", m_MarksMinuteColor, MARKS_MINUTECOLOR
      .WriteProperty "MarksShows", m_MarksShows, HoursMinutes
      .WriteProperty "MarksStyle", m_MarksStyle, CircleMarks
      .WriteProperty "NameBackColor", m_NameBackColor, NAME_BACKCOLOR
      .WriteProperty "NameBackStyle", m_NameBackStyle, Opaque
      .WriteProperty "NameBold", m_NameBold, False
      .WriteProperty "NameBorderStyle", m_NameBorderStyle, Sunken
      .WriteProperty "NameClock", m_NameClock, ""
      .WriteProperty "NameForeColor", m_NameForeColor, NAME_FORECOLOR
      .WriteProperty "NamePosition", m_NamePosition, Off
      .WriteProperty "Picture", m_Picture, Nothing
      .WriteProperty "SetDate", m_SetDate, Date
      .WriteProperty "SetTime", m_SetTime, ""
      .WriteProperty "TimeZoneValue", m_TimeZoneValue, 0
   End With

End Sub
