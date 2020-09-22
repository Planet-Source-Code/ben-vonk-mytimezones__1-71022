VERSION 5.00
Begin VB.UserControl Calendar 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   ClientHeight    =   3036
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3996
   ScaleHeight     =   253
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   333
   ToolboxBitmap   =   "Calendar.ctx":0000
   Begin VB.PictureBox picCalCell 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   492
      Left            =   600
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   51
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Timer tmrIsOtherDay 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1560
      Top             =   2520
   End
   Begin VB.PictureBox picToDay 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00EBEBEB&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   492
      Left            =   120
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Image imgZodiac 
      Height          =   384
      Index           =   1
      Left            =   2040
      Picture         =   "Calendar.ctx":0312
      Stretch         =   -1  'True
      Tag             =   "Aries"
      ToolTipText     =   "21-03 20-04"
      Top             =   1560
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgZodiac 
      Height          =   384
      Index           =   2
      Left            =   2520
      Picture         =   "Calendar.ctx":0BDC
      Stretch         =   -1  'True
      Tag             =   "Taurus"
      ToolTipText     =   "21-04 20-05"
      Top             =   1560
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgZodiac 
      Height          =   384
      Index           =   3
      Left            =   3000
      Picture         =   "Calendar.ctx":14A6
      Stretch         =   -1  'True
      Tag             =   "Gemini"
      ToolTipText     =   "21-05 20-06"
      Top             =   1560
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgZodiac 
      Height          =   384
      Index           =   4
      Left            =   3480
      Picture         =   "Calendar.ctx":1D70
      Stretch         =   -1  'True
      Tag             =   "Cancer"
      ToolTipText     =   "21-06 22-07"
      Top             =   1560
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgZodiac 
      Height          =   384
      Index           =   5
      Left            =   2040
      Picture         =   "Calendar.ctx":263A
      Stretch         =   -1  'True
      Tag             =   "Leo"
      ToolTipText     =   "23-07 22-08"
      Top             =   2040
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgZodiac 
      Height          =   384
      Index           =   6
      Left            =   2520
      Picture         =   "Calendar.ctx":2F04
      Stretch         =   -1  'True
      Tag             =   "Virgo"
      ToolTipText     =   "23-08 22-09"
      Top             =   2040
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgZodiac 
      Height          =   384
      Index           =   7
      Left            =   3000
      Picture         =   "Calendar.ctx":37CE
      Stretch         =   -1  'True
      Tag             =   "Libra"
      ToolTipText     =   "23-09 22-10"
      Top             =   2040
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgZodiac 
      Height          =   384
      Index           =   8
      Left            =   3480
      Picture         =   "Calendar.ctx":4098
      Stretch         =   -1  'True
      Tag             =   "Scorpio"
      ToolTipText     =   "23-10 22-11"
      Top             =   2040
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgZodiac 
      Height          =   384
      Index           =   9
      Left            =   2040
      Picture         =   "Calendar.ctx":4962
      Stretch         =   -1  'True
      Tag             =   "Sagittarius"
      ToolTipText     =   "23-11 21-12"
      Top             =   2520
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgZodiac 
      Height          =   384
      Index           =   10
      Left            =   2520
      Picture         =   "Calendar.ctx":522C
      Stretch         =   -1  'True
      Tag             =   "Capricorn"
      ToolTipText     =   "22-12 20-01"
      Top             =   2520
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgZodiac 
      Height          =   384
      Index           =   11
      Left            =   3000
      Picture         =   "Calendar.ctx":5AF6
      Stretch         =   -1  'True
      Tag             =   "Aguarius"
      ToolTipText     =   "21-01 20-02"
      Top             =   2520
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgZodiac 
      Height          =   384
      Index           =   12
      Left            =   3480
      Picture         =   "Calendar.ctx":63C0
      Stretch         =   -1  'True
      Tag             =   "Pisces"
      ToolTipText     =   "21-02 20-03"
      Top             =   2520
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgSeason 
      Height          =   384
      Index           =   1
      Left            =   2040
      Picture         =   "Calendar.ctx":6C8A
      ToolTipText     =   "21-03 20-06"
      Top             =   600
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgSeason 
      Height          =   384
      Index           =   2
      Left            =   2520
      Picture         =   "Calendar.ctx":7554
      ToolTipText     =   "21-06 21-09"
      Top             =   600
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgSeason 
      Height          =   384
      Index           =   3
      Left            =   3000
      Picture         =   "Calendar.ctx":7E1E
      ToolTipText     =   "22-09 20-12"
      Top             =   600
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgSeason 
      Height          =   384
      Index           =   4
      Left            =   3480
      Picture         =   "Calendar.ctx":86E8
      ToolTipText     =   "21-12 20-03"
      Top             =   600
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgToDay 
      Height          =   384
      Left            =   120
      Picture         =   "Calendar.ctx":8FB2
      Top             =   120
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgMoonPhase 
      Height          =   384
      Index           =   4
      Left            =   3480
      Picture         =   "Calendar.ctx":92BC
      Top             =   1080
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgMoonPhase 
      Height          =   384
      Index           =   3
      Left            =   3000
      Picture         =   "Calendar.ctx":9F86
      Top             =   1080
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgMoonPhase 
      Height          =   384
      Index           =   2
      Left            =   2520
      Picture         =   "Calendar.ctx":AC50
      Top             =   1080
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgMoonPhase 
      Height          =   384
      Index           =   1
      Left            =   2040
      Picture         =   "Calendar.ctx":B91A
      Top             =   1080
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgQuarter 
      Height          =   384
      Index           =   4
      Left            =   3480
      Picture         =   "Calendar.ctx":C5E4
      ToolTipText     =   "01-10 31-12"
      Top             =   120
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgQuarter 
      Height          =   384
      Index           =   3
      Left            =   3000
      Picture         =   "Calendar.ctx":CEAE
      ToolTipText     =   "01-07 30-09"
      Top             =   120
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgQuarter 
      Height          =   384
      Index           =   2
      Left            =   2520
      Picture         =   "Calendar.ctx":D778
      ToolTipText     =   "01-04 30-06"
      Top             =   120
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgQuarter 
      Height          =   384
      Index           =   1
      Left            =   2040
      Picture         =   "Calendar.ctx":E042
      ToolTipText     =   "01-01 31-03"
      Top             =   120
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgButton 
      Height          =   384
      Index           =   8
      Left            =   1560
      Top             =   600
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgButton 
      Height          =   384
      Index           =   7
      Left            =   1080
      Top             =   600
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgButton 
      Height          =   384
      Index           =   6
      Left            =   600
      Top             =   600
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgButton 
      Height          =   384
      Index           =   5
      Left            =   120
      Top             =   600
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgButton 
      Height          =   384
      Index           =   0
      Left            =   600
      Picture         =   "Calendar.ctx":E90C
      Top             =   120
      Visible         =   0   'False
      Width           =   384
   End
End
Attribute VB_Name = "Calendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Calendar Control
'
'Author Ben Vonk
'20-08-2004 First version
'29-10-2005 Second version (based on Stefaan Casier's 'Owner Drawn Calendar Control' at http://www.codeguru.com/vb/controls/vb_othctrl/ocxcontrols/article.php/c1521/)
'09-12-2005 Some bug fixes and add Event DateChanged

'Notes:
' The moonphase calculations are based on Scott Seligman's 'Moon Cycles' at http://www.scottandmichelle.net/scott/code/index2.mv?codenum=031
' And he used adopted routines from 'pcal', which bares this copyright notice:
'
' Routines to accurately calculate the phase of the moon
'
' Originally adapted from 'moontool.c' by John Walker, Release 2.0.
'
' This routine (calc_phase) and its support routines were adapted
' from phase.c (v 1.2 88/08/26 22:29:42 jef) in the program
' 'xphoon' (v 1.9 88/08/26 22:29:47 jef) by Jef Poskanzer and
' Craig Leres. The necessary notice follows...
'
' Copyright (C) 1988 by Jef Poskanzer and Craig Leres.
'
' Permission to use, copy, modify, and distribute this software
' and its documentation for any purpose and without fee is hereby
' granted, provided that the above copyright notice appear in all
' copies and that both that copyright notice and this permission
' notice appear in supporting documentation. This software is
' provided "as is" without express or implied warranty.
'
' These were added to 'pcal' by RLD on 19-MAR-1991
'
' The GetJulianDay function is adopted from Peter Duffett-Smith's book
' 'Astronomy With Your Personal Computer' by Rick Dyson 18-MAR-1991

Option Explicit

' Public Events
Public Event ButtonClick(ButtonID As ButtonTypes)
Public Event DateChanged(ButtonID As ButtonTypes)
Public Event DayClick(Button As Integer, Shift As Integer, IsDay As Integer, IsMonth As Integer, Cancel As Boolean)
Public Event DayDblClick(IsDay As Integer)
Public Event SelChanged()

' Private Constants
Private Const PI                    As Double = 3.14159265358979
'Private Const BDR_RAISED            As Long = &H5
'Private Const BDR_RAISEDINNER       As Long = &H4
'Private Const BDR_SUNKEN            As Long = &HA
Private Const BDR_SUNKENINNER       As Long = &H8
'Private Const BF_RIGHT              As Long = &H4
'Private Const BF_TOP                As Long = &H2
'Private Const BF_LEFT               As Long = &H1
'Private Const BF_BOTTOM             As Long = &H8
'Private Const BF_RECT               As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
'Private Const DI_NORMAL             As Long = &H3
Private Const CELL_SPACE            As String = "   "
Private Const SEPARATOR_DEFAULT     As String = ", "

' Public Enumerations
Public Enum AppearanceStyles
   [3D]
   [Flat]
End Enum

'Public Enum BackStyles
'   Transparent
'   Opaque
'End Enum

Public Enum BarFrameStyles
   Off
   Small
   Large
End Enum

'Public Enum BorderStyles
'   [None]
'   [Fixed Single]
'End Enum

Public Enum ButtonTypes
   ToDayButton
   LeftButton
   RightButton
   UpButton
   DownButton
   QuarterButton
   SeasonButton
   MoonPhaseButton
   ZodiacButton
End Enum

Public Enum CellStyles
   EmptyCell
   BlueSelect
   DarkGraySelect
   LightGraySelect
   BevelIn
   BevelOut
   UserColor
End Enum

Public Enum CellTypes
   Header
   OtherMonths
   Normal
   Selected
   DayOfTheYear
   WeekNumber
End Enum

Public Enum DateFormats
   [dd-mm-yyyy]
   [mm-dd-yyyy]
   [yyyy-mm-dd]
End Enum

Public Enum DayWidthFormats
   D
   Dd
   Ddd
End Enum

Public Enum FirstWeekDays
   Sunday
   Monday
   Tuesday
   Wednesday
   Thursday
   Friday
   Saturday
End Enum

Public Enum GradientStyles
   NoGradient
   LeftToRight
   RightToLeft
   TopToBottom
   BottomToTop
End Enum

Public Enum GridStyles
   NoGrid
   FullGrid
   MonthOnlyGrid
   HorizontalGrid
   VerticalGrid
End Enum

Public Enum LabelBorderStyles
   Sunken
   Raised
   Edged
End Enum

Public Enum Languages
   System
   English
   Spanish
   Dutch
   French
   Italian
End Enum

Public Enum Hemispheres
   North
   South
End Enum

Public Enum MoonTypes
   NoMoon
   NewMoon
   FirstQuarter
   FullMoon
   LastQuarter
End Enum

Public Enum SelectionTypes
   SingleCell
   MultiCell
End Enum

' Private Enumeration
Private Enum ColorsRGB
   IsRed
   IsGreen
   IsBlue
End Enum

' Private Types
'Private Type Rect
'   Left                             As Long
'   Top                              As Long
'   Right                            As Long
'   Bottom                           As Long
'End Type

Private Type CalButtons
   TipText                          As String     ' button tooltiptext
   Rect                             As Rect       ' button rectangle
End Type

Private Type CalCells
   Type                             As CellTypes  ' cell type
   Mark                             As Integer    ' cell marking state
   MarkTipText(4)                   As String     ' cell marking tooltiptext
   Text                             As String * 3 ' cell text
   TipText                          As String     ' cell tooltiptext
   X                                As Single     ' cell X location
   Y                                As Single     ' cell Y location
End Type

Private Type CalLabels
   Text                             As String     ' label text
   Rect                             As Rect       ' label rectangle
End Type

Private Type GradientRect
   UpperLeft                        As Long
   LowerRight                       As Long
End Type

Private Type LanguagesText
   DayNames                         As String
   Miscellaneous                    As String
   MonthNames                       As String
   QuarterNames                     As String
   SeasonNames                      As String
   MoonPhaseNames                   As String
   MoonPhaseText                    As String
   ZodiacNames                      As String
End Type

Private Type MoonPhaseResult
   Phase                            As Integer
   Days                             As Integer
End Type

'Private Type PointAPI
'   X                                As Long
'   Y                                As Long
'End Type

Private Type TriVertex
   X                                As Long
   Y                                As Long
   Red                              As Integer
   Green                            As Integer
   Blue                             As Integer
   Alpha                            As Integer
End Type

' Private Variables
Private m_Appearance                As AppearanceStyles  ' for calendar appearance
Private m_FrameStyle                As BarFrameStyles    ' set calendar frameline style
Private m_ShowInfoBar               As BarFrameStyles    ' set calendar information style
Private m_ShowNavigationBar         As BarFrameStyles    ' set calendar navigation style
Private IsChanged                   As Boolean           ' checked is day in month is changed
Private IsClicked                   As Boolean           ' checked if button is clicked
Private m_CellOtherMonthView        As Boolean           ' set the viewing of the other month parts
Private m_LabelFontBold             As Boolean           ' set fontbold for labels on/off
Private m_Locked                    As Boolean           ' set calender Locked on/off
Private m_LockInfoBar               As Boolean           ' set infobar as buttons on/off
Private m_SelectedDayMark           As Boolean           ' for showing calendar today mark on/off
Private m_ShowDayOfYear             As Boolean           ' set day of the year on/off
Private m_ShowToolTipText           As Boolean           ' set showing tooltiptext on/off
Private MouseIn                     As Boolean           ' checked if mouse is in button
Private MouseOut                    As Boolean           ' checked if mouse left button
Private OtherMonthSelected          As Boolean           ' checked if the selected day is in a other month
Private SetLastDay                  As Boolean           ' holds lastday of current month if selected
Private CalButtonID                 As ButtonTypes       ' for witch button is clicked
Private CalButton()                 As CalButtons        ' calendar button properties
Private CalCell()                   As CalCells          ' calendar cell properties
Private CalLabel                    As CalLabels         ' calendr label properties
Private m_CellDayOfYearStyle        As CellStyles        ' style of the day of the year cell
Private m_CellDaysStyle             As CellStyles        ' style of the day cells
Private m_CellHeaderStyle           As CellStyles        ' style of header cells
Private m_CellOtherMonthStyle       As CellStyles        ' style of other monht cells
Private m_CellSelectStyle           As CellStyles        ' style of selected cells
Private ToDay                       As Date              ' holds calendar start date
Private m_WeekDayViewChar           As DayWidthFormats   ' total characters of daynames
Private m_DateFormat                As DateFormats       ' set the date format type
Private m_FirstWeekDay              As FirstWeekDays     ' set calendars firstday of the week
Private m_ButtonGradientStyle       As GradientStyles    ' set GradientStyle for buttons
Private m_GradientStyle             As GradientStyles    ' set calendar GradientStyle
Private m_GridStyle                 As GridStyles        ' set the type of calender gridlines
Private CurrentCell                 As Integer           ' always one is selected
Private CalLanguage                 As Integer           ' set calendar language
Private m_CalYear                   As Integer           ' set calendar year
Private m_CalMonth                  As Integer           ' set calendar month
Private m_CalDay                    As Integer           ' set calendar day
Private MonthDays                   As Integer           ' total days of current month
Private MouseCell                   As Integer           ' holds the cell for DayDblClick
Private MouseButton                 As Integer           ' holds the pressed mousebutton
Private SelectedDay                 As Integer           ' holds selected daycell
Private OffsetCell                  As Integer           ' starting from cell
Private SizeX                       As Integer           ' width of a cell
Private SizeY                       As Integer           ' height of a cell
Private WeekDayOfFirstDay           As Integer           ' first day week day
Private m_LabelBackStyle            As BackStyles        ' set backstyle for calendar labels
Private m_LabelBorderStyle          As LabelBorderStyles ' set the label borderstyle
Private LanguageText(5)             As LanguagesText     ' hold the calendar language text
Private m_Language                  As Languages         ' hold selected calendar language
Private m_Hemisphere                As Hemispheres       ' set calender in North or South
Private m_ArrowColor                As OLE_COLOR         ' set forecolor for arrows
Private m_ButtonColor               As OLE_COLOR         ' set backcolor for buttons
Private m_ButtonGradientColor       As OLE_COLOR         ' set gradientcolor for buttons
Private m_CellDayOfYearBackColor    As OLE_COLOR         ' set backcolor for day of the year cell
Private m_CellDayOfYearForeColor    As OLE_COLOR         ' set forecolor for day of the year cell
Private m_CellDaysBackColor         As OLE_COLOR         ' set backcolor for day cells
Private m_CellHeaderBackColor       As OLE_COLOR         ' set backcolor for header cells
Private m_WeekNumberForeColor       As OLE_COLOR         ' set forecolor for header cells
Private m_CellForeColorSunday       As OLE_COLOR         ' set forecolor for sunday cells
Private m_CellForeColorMonday       As OLE_COLOR         ' set forecolor for monday cells
Private m_CellForeColorTuesday      As OLE_COLOR         ' set forecolor for tuesday cells
Private m_CellForeColorWednesday    As OLE_COLOR         ' set forecolor for wednesday cells
Private m_CellForeColorThursday     As OLE_COLOR         ' set forecolor for thursday cells
Private m_CellForeColorFriday       As OLE_COLOR         ' set forecolor for friday cells
Private m_CellForeColorSaturday     As OLE_COLOR         ' set forecolor for saturday cells
Private m_CellOtherMonthBackColor   As OLE_COLOR         ' set backcolor for other month cells
Private m_CellOtherMonthForeColor   As OLE_COLOR         ' set forecolor for other month cells
Private m_CellSelectBackColor       As OLE_COLOR         ' set backcolor for selected cells
Private m_CellSelectForeColor       As OLE_COLOR         ' set forecolor for selected cells
Private m_CellSelectHeaderForeColor As OLE_COLOR         ' set forecolor for selected headercells
Private m_FrameColor                As OLE_COLOR         ' set color for calendar framelines
Private m_GradientColor             As OLE_COLOR         ' set color for calendar gradientcolor
Private m_GridColor                 As OLE_COLOR         ' set color for calendar gridlines
Private m_LabelBackColor            As OLE_COLOR         ' set color for calendar labels backgroundcolor
Private m_LabelForeColor            As OLE_COLOR         ' set color for calendar labels foregroundcolor
Private MarkColor(4)                As OLE_COLOR         ' set mark colors
Private CalendarRect                As Rect              ' holds the calendar rectangle
Private m_SelectionType             As SelectionTypes    ' set type of day selection
Private MouseX                      As Single            ' holds the X position of the mouse
Private MouseY                      As Single            ' holds the Y position of the mouse
Private SizeFont                    As Single            ' holds the fontsize
Private m_Picture                   As StdPicture        ' holds the calendar picture
Private Separator(2)                As String            ' holds the Separator for giving the results of the InfoBar,
                                                         ' for Internal 0 = " ",  1 = "  (" and 2 = ")"
                                                         ' and External 0 = ", ", 1 = ", "  and 2 = ""

' Private API's
'Private Declare Function BitBlt Lib "GDI32" (ByVal hDCDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreatePolygonRgn Lib "GDI32" (lpPoint As Any, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
'Private Declare Function CreateSolidBrush Lib "GDI32" (ByVal crColor As Long) As Long
'Private Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Private Declare Function FillRgn Lib "GDI32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
'Private Declare Function StretchBlt Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
'Private Declare Function GetLocaleInfo Lib "Kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function GradientFill Lib "MSIMG32" (ByVal hDC As Long, ByRef pVertex As TriVertex, ByVal dwNumVertex As Long, pMesh As Any, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Integer
'Private Declare Function OleTranslateColor Lib "OLEPro32" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
'Private Declare Function DrawEdge Lib "User32" (ByVal hDC As Long, qrc As Rect, ByVal edge As Long, ByVal grfFlags As Long) As Long
'Private Declare Function DrawIconEx Lib "User32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyHeight As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
'Private Declare Function PtInRect Lib "User32" (Rect As Rect, ByVal lPtX As Long, ByVal lPtY As Long) As Integer

Public Property Get Appearance() As AppearanceStyles
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."

   Appearance = m_Appearance

End Property

Public Property Let Appearance(ByVal NewAppearance As AppearanceStyles)

   m_Appearance = NewAppearance
   PropertyChanged "Appearance"
   
   Call CalcCalendar

End Property

Public Property Get ArrowColor() As OLE_COLOR
Attribute ArrowColor.VB_Description = "Returns/sets the color used to display the arrows of an calendar control."

   ArrowColor = m_ArrowColor

End Property

Public Property Let ArrowColor(ByVal NewArrowColor As OLE_COLOR)

   m_ArrowColor = NewArrowColor
   PropertyChanged "ArrowColor"
   
   Call CalcCalendar

End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."

   BackColor = UserControl.BackColor

End Property

Public Property Let BackColor(ByVal NewBackColor As OLE_COLOR)

   UserControl.BackColor = NewBackColor
   PropertyChanged "BackColor"
   
   Call CalcCalendar

End Property

Public Property Get BorderStyle() As BorderStyles
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."

   BorderStyle = UserControl.BorderStyle

End Property

Public Property Let BorderStyle(ByVal NewBorderStyle As BorderStyles)

   UserControl.BorderStyle = NewBorderStyle
   PropertyChanged "BorderStyle"

End Property

Public Property Get ButtonColor() As OLE_COLOR
Attribute ButtonColor.VB_Description = "Returns/sets the background color used to display the buttons of an calendar control."

   ButtonColor = m_ButtonColor

End Property

Public Property Let ButtonColor(ByVal NewButtonColor As OLE_COLOR)

   m_ButtonColor = NewButtonColor
   PropertyChanged "ButtonColor"
   
   Call CalcCalendar

End Property

Public Property Get ButtonGradientColor() As OLE_COLOR
Attribute ButtonGradientColor.VB_Description = "Returns/sets the color used to display the button gradient of an calendar control."

   ButtonGradientColor = m_ButtonGradientColor

End Property

Public Property Let ButtonGradientColor(ByVal NewButtonGradientColor As OLE_COLOR)

   m_ButtonGradientColor = NewButtonGradientColor
   PropertyChanged ("ButtonGradientColor")
   
   If m_ButtonGradientStyle Then Call CalcCalendar

End Property

Public Property Get ButtonGradientStyle() As GradientStyles
Attribute ButtonGradientStyle.VB_Description = "Returns/sets the view type used to display the button gradient of an calendar control."

   ButtonGradientStyle = m_ButtonGradientStyle

End Property

Public Property Let ButtonGradientStyle(ByVal NewButtonGradientStyle As GradientStyles)

   m_ButtonGradientStyle = NewButtonGradientStyle
   PropertyChanged ("ButtonGradientStyle")
   
   Call CalcCalendar

End Property

Public Property Get CalDay() As Integer
Attribute CalDay.VB_Description = "Returns/sets the value of the calendar day."

   CalDay = m_CalDay

End Property

Public Property Let CalDay(ByVal NewDay As Integer)

   If NewDay < 1 Then NewDay = 1
   If NewDay > MonthDays Then NewDay = MonthDays
   
   If m_SelectionType = SingleCell Then
      Call DaySelect(m_CalDay, False)
      Call DaySelect(NewDay, True)
      
   ElseIf m_SelectionType = MultiCell Then
      Call DaySelect(NewDay, Not IsDaySel(NewDay))
   End If
   
   m_CalDay = NewDay
   SelectedDay = NewDay
   SetLastDay = False
   PropertyChanged "CalDay"
   
   Call CalcCalendar

End Property

Public Property Get CalMonth() As Integer
Attribute CalMonth.VB_Description = "Returns/sets the value of the calendar month."

   CalMonth = m_CalMonth

End Property

Public Property Let CalMonth(ByVal NewMonth As Integer)

   If NewMonth < 1 Then NewMonth = 1
   If NewMonth > 12 Then NewMonth = 12
   
   m_CalMonth = NewMonth
   PropertyChanged "CalMonth"
   
   Call CalcCalendar

End Property

Public Property Get CalYear() As Integer
Attribute CalYear.VB_Description = "Returns/sets the value of the calendar year."

   CalYear = m_CalYear

End Property

Public Property Let CalYear(ByVal NewYear As Integer)

   If NewYear < 1583 Then NewYear = 1583
   If NewYear > 9999 Then NewYear = 9999
   
   m_CalYear = NewYear
   PropertyChanged "CalYear"
   
   Call CalcCalendar

End Property

Public Property Get CellDayOfYearBackColor() As OLE_COLOR
Attribute CellDayOfYearBackColor.VB_Description = "Returns/sets the background color used to display day of year-cell."

   CellDayOfYearBackColor = m_CellDayOfYearBackColor

End Property

Public Property Let CellDayOfYearBackColor(ByVal NewDayOfYearBackColor As OLE_COLOR)

   m_CellDayOfYearBackColor = NewDayOfYearBackColor
   PropertyChanged "CellDayOfYearBackColor"
   
   Call DrawCalendar

End Property

Public Property Get CellDayOfYearForeColor() As OLE_COLOR
Attribute CellDayOfYearForeColor.VB_Description = "Returns/sets the foreground color used to display day of year-cell."

   CellDayOfYearForeColor = m_CellDayOfYearForeColor

End Property

Public Property Let CellDayOfYearForeColor(ByVal NewDayOfYearForeColor As OLE_COLOR)

   m_CellDayOfYearForeColor = NewDayOfYearForeColor
   PropertyChanged "CellDayOfYearForeColor"
   
   Call DrawCalendar

End Property

Public Property Get CellDayOfYearStyle() As CellStyles
Attribute CellDayOfYearStyle.VB_Description = "Returns/sets the view style of day of year-cell."

   CellDayOfYearStyle = m_CellDayOfYearStyle

End Property

Public Property Let CellDayOfYearStyle(ByVal NewViewDayOfYearCell As CellStyles)

   m_CellDayOfYearStyle = NewViewDayOfYearCell
   PropertyChanged "CellDayOfYearStyle"
   
   Call DrawCalendar

End Property

Public Property Get CellDaysBackColor() As OLE_COLOR
Attribute CellDaysBackColor.VB_Description = "Returns/sets the background color used to display day-cells."

   CellDaysBackColor = m_CellDaysBackColor

End Property

Public Property Let CellDaysBackColor(ByVal NewDaysBackColor As OLE_COLOR)

   m_CellDaysBackColor = NewDaysBackColor
   PropertyChanged "CellDaysBackColor"
   
   Call DrawCalendar

End Property

Public Property Get CellDaysStyle() As CellStyles
Attribute CellDaysStyle.VB_Description = "Returns/sets the view style of non-selected calender day-cells."

   CellDaysStyle = m_CellDaysStyle

End Property

Public Property Let CellDaysStyle(ByVal NewCellDaysStyle As CellStyles)

   m_CellDaysStyle = NewCellDaysStyle
   PropertyChanged "CellDaysStyle"
   
   Call DrawCalendar

End Property

Public Property Get CellForeColorFriday() As OLE_COLOR
Attribute CellForeColorFriday.VB_Description = "Returns/sets the foreground color of the friday-cells."

   CellForeColorFriday = m_CellForeColorFriday

End Property

Public Property Let CellForeColorFriday(ByVal NewCellForeColorFriday As OLE_COLOR)

   m_CellForeColorFriday = NewCellForeColorFriday
   PropertyChanged "CellForeColorFriday"
   
   Call DrawCalendar

End Property

Public Property Get CellForeColorMonday() As OLE_COLOR
Attribute CellForeColorMonday.VB_Description = "Returns/sets the foreground color of the monday-cells."

   CellForeColorMonday = m_CellForeColorMonday

End Property

Public Property Let CellForeColorMonday(ByVal NewCellForeColorMonday As OLE_COLOR)

   m_CellForeColorMonday = NewCellForeColorMonday
   PropertyChanged "CellForeColorMonday"
   
   Call DrawCalendar

End Property

Public Property Get CellForeColorSaturday() As OLE_COLOR
Attribute CellForeColorSaturday.VB_Description = "Returns/sets the foreground color of the saturday-cells."

   CellForeColorSaturday = m_CellForeColorSaturday

End Property

Public Property Let CellForeColorSaturday(ByVal NewCellForeColorSaturday As OLE_COLOR)

   m_CellForeColorSaturday = NewCellForeColorSaturday
   PropertyChanged "CellForeColorSaturday"
   
   Call DrawCalendar

End Property

Public Property Get CellForeColorSunday() As OLE_COLOR
Attribute CellForeColorSunday.VB_Description = "Returns/sets the foreground color of the sunday-cells."

   CellForeColorSunday = m_CellForeColorSunday

End Property

Public Property Let CellForeColorSunday(ByVal NewCellForeColorSunday As OLE_COLOR)

   m_CellForeColorSunday = NewCellForeColorSunday
   PropertyChanged "CellForeColorSunday"
   
   Call DrawCalendar

End Property

Public Property Get CellForeColorThursday() As OLE_COLOR
Attribute CellForeColorThursday.VB_Description = "Returns/sets the foreground color of the thursday-cells."

   CellForeColorThursday = m_CellForeColorThursday

End Property

Public Property Let CellForeColorThursday(ByVal NewCellForeColorThursday As OLE_COLOR)

   m_CellForeColorThursday = NewCellForeColorThursday
   PropertyChanged "CellForeColorThursday"
   
   Call DrawCalendar

End Property

Public Property Get CellForeColorTuesday() As OLE_COLOR
Attribute CellForeColorTuesday.VB_Description = "Returns/sets the foreground color of the tuesday-cells."

   CellForeColorTuesday = m_CellForeColorTuesday

End Property

Public Property Let CellForeColorTuesday(ByVal NewCellForeColorTuesday As OLE_COLOR)

   m_CellForeColorTuesday = NewCellForeColorTuesday
   PropertyChanged "CellForeColorTuesday"
   
   Call DrawCalendar

End Property

Public Property Get CellForeColorWednesday() As OLE_COLOR
Attribute CellForeColorWednesday.VB_Description = "Returns/sets the foreground color of the wednesday-cells."

   CellForeColorWednesday = m_CellForeColorWednesday

End Property

Public Property Let CellForeColorWednesday(ByVal NewCellForeColorWednesday As OLE_COLOR)

   m_CellForeColorWednesday = NewCellForeColorWednesday
   PropertyChanged "CellForeColorWednesday"
   
   Call DrawCalendar

End Property

Public Property Get CellHeaderBackColor() As OLE_COLOR
Attribute CellHeaderBackColor.VB_Description = "Returns/sets the background color used to display headers-cells."

   CellHeaderBackColor = m_CellHeaderBackColor

End Property

Public Property Let CellHeaderBackColor(ByVal NewCellHeaderBackColor As OLE_COLOR)

   m_CellHeaderBackColor = NewCellHeaderBackColor
   PropertyChanged "CellHeaderBackColor"
   
   Call DrawCalendar

End Property

Public Property Get CellHeaderStyle() As CellStyles
Attribute CellHeaderStyle.VB_Description = "Returns/sets the view style of header-cells or weekday-names."

   CellHeaderStyle = m_CellHeaderStyle

End Property

Public Property Let CellHeaderStyle(ByVal NewCellHeaderStyle As CellStyles)

   m_CellHeaderStyle = NewCellHeaderStyle
   PropertyChanged "CellHeaderStyle"
   
   Call DrawCalendar

End Property

Public Property Get CellOtherMonthBackColor() As OLE_COLOR
Attribute CellOtherMonthBackColor.VB_Description = "Returns/sets the background color used to display other month-cells."

   CellOtherMonthBackColor = m_CellOtherMonthBackColor

End Property

Public Property Let CellOtherMonthBackColor(ByVal NewCellOtherMonthBackColor As OLE_COLOR)

   m_CellOtherMonthBackColor = NewCellOtherMonthBackColor
   PropertyChanged "CellOtherMonthBackColor"
   
   Call DrawCalendar

End Property

Public Property Get CellOtherMonthForeColor() As OLE_COLOR
Attribute CellOtherMonthForeColor.VB_Description = "Returns/sets the foreground color used to display other month-cells."

   CellOtherMonthForeColor = m_CellOtherMonthForeColor

End Property

Public Property Let CellOtherMonthForeColor(ByVal NewCellOtherMonthForeColor As OLE_COLOR)

   m_CellOtherMonthForeColor = NewCellOtherMonthForeColor
   PropertyChanged "CellOtherMonthForeColor"
   
   Call DrawCalendar

End Property

Public Property Get CellOtherMonthStyle() As CellStyles
Attribute CellOtherMonthStyle.VB_Description = "Returns/sets the view style of other month-cells."

   CellOtherMonthStyle = m_CellOtherMonthStyle

End Property

Public Property Let CellOtherMonthStyle(ByVal NewCellOtherMonthStyle As CellStyles)

   m_CellOtherMonthStyle = NewCellOtherMonthStyle
   PropertyChanged "CellOtherMonthStyle"
   
   Call DrawCalendar

End Property

Public Property Get CellOtherMonthView() As Boolean
Attribute CellOtherMonthView.VB_Description = "Returns/sets the viewing the parts of other months."

   CellOtherMonthView = m_CellOtherMonthView

End Property

Public Property Let CellOtherMonthView(ByVal NewCellOtherMonthView As Boolean)

   m_CellOtherMonthView = NewCellOtherMonthView
   PropertyChanged "CellOtherMonthView"
   
   Call CalcCalendar

End Property

Public Property Get CellSelectBackColor() As OLE_COLOR
Attribute CellSelectBackColor.VB_Description = "Returns/sets the background color used to display selected-cells."

   CellSelectBackColor = m_CellSelectBackColor

End Property

Public Property Let CellSelectBackColor(ByVal NewCellSelectedBackColor As OLE_COLOR)

   m_CellSelectBackColor = NewCellSelectedBackColor
   PropertyChanged "CellSelectBackColor"
   
   Call DrawCalendar

End Property

Public Property Get CellSelectForeColor() As OLE_COLOR
Attribute CellSelectForeColor.VB_Description = "Returns/sets the foreground color used to display selected-cells."

   CellSelectForeColor = m_CellSelectForeColor

End Property

Public Property Let CellSelectForeColor(ByVal NewCellSelectedForeColor As OLE_COLOR)

   m_CellSelectForeColor = NewCellSelectedForeColor
   PropertyChanged "CellSelectForeColor"
   
   Call DrawCalendar

End Property

Public Property Get CellSelectHeaderForeColor() As OLE_COLOR
Attribute CellSelectHeaderForeColor.VB_Description = "Returns/sets the foreground color used to display selected header-cells."

   CellSelectHeaderForeColor = m_CellSelectHeaderForeColor

End Property

Public Property Let CellSelectHeaderForeColor(ByVal NewCellSelectHeaderForeColor As OLE_COLOR)

   m_CellSelectHeaderForeColor = NewCellSelectHeaderForeColor
   PropertyChanged "CellSelectHeaderForeColor"
   
   Call DrawCalendar

End Property

Public Property Get CellSelectStyle() As CellStyles
Attribute CellSelectStyle.VB_Description = "Returns/sets the view style of selected days-cells."

   CellSelectStyle = m_CellSelectStyle

End Property

Public Property Let CellSelectStyle(ByVal NewCellSelectStyle As CellStyles)

   m_CellSelectStyle = NewCellSelectStyle
   PropertyChanged "CellSelectStyle"
   
   Call DrawCalendar

End Property

Public Property Get DateFormat() As DateFormats
Attribute DateFormat.VB_Description = "Returns/sets the short dateformat of the calendar."

   DateFormat = m_DateFormat

End Property

Public Property Let DateFormat(ByVal NewDateFormat As DateFormats)

   m_DateFormat = NewDateFormat
   PropertyChanged ("DateFormat")
   
   If m_Locked Then Exit Property
   
   Call DrawCalendar
   Call SetToDayImage

End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."

   Enabled = UserControl.Enabled

End Property

Public Property Let Enabled(ByVal NewEnabled As Boolean)

   UserControl.Enabled = NewEnabled
   PropertyChanged "Enabled"

End Property

Public Property Get FirstWeekDay() As FirstWeekDays
Attribute FirstWeekDay.VB_Description = "Returns/sets the calendars first day of the week."

   FirstWeekDay = m_FirstWeekDay - 1

End Property

Public Property Let FirstWeekDay(ByVal NewFirstWeekDay As FirstWeekDays)

   m_FirstWeekDay = NewFirstWeekDay + 1
   PropertyChanged "FirstWeekDay"
   
   Call CalcCalendar

End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."

   Set Font = UserControl.Font

End Property

Public Property Let Font(ByVal NewFont As StdFont)

   Set Font = NewFont

End Property

Public Property Set Font(ByVal NewFont As StdFont)

   Set UserControl.Font = NewFont
   PropertyChanged "Font"
   
   Call CalcCalendar

End Property

Public Property Get FrameStyle() As BarFrameStyles
Attribute FrameStyle.VB_Description = "Returns/sets the view style of the framelines of an calendar control."

   FrameStyle = m_FrameStyle

End Property

Public Property Let FrameStyle(ByVal NewFrameStyle As BarFrameStyles)

   m_FrameStyle = NewFrameStyle
   PropertyChanged "FrameStyle"
   
   Call CalcCalendar

End Property

Public Property Get FrameColor() As OLE_COLOR
Attribute FrameColor.VB_Description = "Returns/sets the color used to display the framelines of an calendar control."

   FrameColor = m_FrameColor

End Property

Public Property Let FrameColor(ByVal NewFrameColor As OLE_COLOR)

   m_FrameColor = NewFrameColor
   PropertyChanged "FrameColor"
   
   Call DrawCalendar

End Property

Public Property Get GradientColor() As OLE_COLOR
Attribute GradientColor.VB_Description = "Returns/sets the color used to display the background gradient of an calendar control."

   GradientColor = m_GradientColor

End Property

Public Property Let GradientColor(ByVal NewGradientColor As OLE_COLOR)

   m_GradientColor = NewGradientColor
   PropertyChanged ("GradientColor")
   
   If m_GradientStyle Then Call CalcCalendar

End Property

Public Property Get GradientStyle() As GradientStyles
Attribute GradientStyle.VB_Description = "Returns/sets the view type used to display the background gradient of an calendar control."

   GradientStyle = m_GradientStyle

End Property

Public Property Let GradientStyle(ByVal NewGradientStyle As GradientStyles)

   m_GradientStyle = NewGradientStyle
   PropertyChanged ("GradientStyle")
   
   Call CalcCalendar

End Property

Public Property Get GridColor() As OLE_COLOR
Attribute GridColor.VB_Description = "Returns/sets the color used to display the gridlines of an calendar control."

   GridColor = m_GridColor

End Property

Public Property Let GridColor(ByVal NewGridColor As OLE_COLOR)

   m_GridColor = NewGridColor
   PropertyChanged "GridColor"
   
   Call DrawGrid

End Property

Public Property Get GridStyle() As GridStyles
Attribute GridStyle.VB_Description = "Returns/sets the view type to display the gridlines of an calendar control."

   GridStyle = m_GridStyle

End Property

Public Property Let GridStyle(ByVal NewGridStyle As GridStyles)

   m_GridStyle = NewGridStyle
   PropertyChanged ("GridStyle")
   
   Call DrawCalendar

End Property

Public Property Get Hemisphere() As Hemispheres
Attribute Hemisphere.VB_Description = "Returns/sets the hemisphere to north or south of an calendar control."

   Hemisphere = m_Hemisphere

End Property

Public Property Let Hemisphere(ByVal NewHemisphere As Hemispheres)

   m_Hemisphere = NewHemisphere
   PropertyChanged "Hemisphere"
   
   Call CalcCalendar

End Property

Public Property Get LabelBackColor() As OLE_COLOR
Attribute LabelBackColor.VB_Description = "Returns/sets the color used to display the label of an calendar control."

   LabelBackColor = m_LabelBackColor

End Property

Public Property Let LabelBackColor(ByVal NewLabelBackColor As OLE_COLOR)

   m_LabelBackColor = NewLabelBackColor
   PropertyChanged "LabelBackColor"
   
   Call CalcCalendar

End Property

Public Property Get LabelBackStyle() As BackStyles
Attribute LabelBackStyle.VB_Description = "Indicates the calendar label is transparent or opaque."

   LabelBackStyle = m_LabelBackStyle

End Property

Public Property Let LabelBackStyle(ByVal NewLabelBackStyle As BackStyles)

   m_LabelBackStyle = NewLabelBackStyle
   PropertyChanged "LabelBackStyle"
   
   Call CalcCalendar

End Property

Public Property Get LabelBorderStyle() As LabelBorderStyles
Attribute LabelBorderStyle.VB_Description = "Returns/sets the border style for the label of an calendar control."

   LabelBorderStyle = m_LabelBorderStyle

End Property

Public Property Let LabelBorderStyle(ByVal NewLabelBorderStyle As LabelBorderStyles)

   m_LabelBorderStyle = NewLabelBorderStyle
   PropertyChanged "LabelBorderStyle"
   
   Call CalcCalendar

End Property

Public Property Get LabelFontBold() As Boolean
Attribute LabelFontBold.VB_Description = "Returns/sets the font boldstyle in the label of an calendar control."

   LabelFontBold = m_LabelFontBold

End Property

Public Property Let LabelFontBold(ByVal NewLabelFontBold As Boolean)

   m_LabelFontBold = NewLabelFontBold
   PropertyChanged "LabelFontBold"
   
   Call CalcCalendar

End Property

Public Property Get LabelForeColor() As OLE_COLOR
Attribute LabelForeColor.VB_Description = "Returns/sets the color used to display text and graphics in the label of an calendar control."

   LabelForeColor = m_LabelForeColor

End Property

Public Property Let LabelForeColor(ByVal NewLabelForeColor As OLE_COLOR)

   m_LabelForeColor = NewLabelForeColor
   PropertyChanged "LabelForeColor"
   
   Call CalcCalendar

End Property

Public Property Get Language() As Languages
Attribute Language.VB_Description = "Returns or sets the Calendar control language."

   Language = m_Language

End Property

Public Property Let Language(ByVal NewLanguage As Languages)

   m_Language = NewLanguage

   If m_Language = System Then
      CalLanguage = GetSystemLanguage
   Else
      CalLanguage = m_Language
   End If

   PropertyChanged "Language"
   
   If m_Locked Then Exit Property
   
   Call SetWeekDayHeaderText
   Call DrawCalendar

End Property

Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Returns/sets to lock/unlock of an calendar control to fasten the property changes."

   Locked = m_Locked

End Property

Public Property Let Locked(ByVal NewLocked As Boolean)

   m_Locked = NewLocked
   PropertyChanged "Locked"
   
   Call Resize
   Call CalcCalendar
   
   If m_Locked Then Call DrawCalendar

End Property

Public Property Get LockInfoBar() As Boolean
Attribute LockInfoBar.VB_Description = "Returns/sets a value that determines whether the infobar can be used as buttons or bar only."

   LockInfoBar = m_LockInfoBar

End Property

Public Property Let LockInfoBar(ByVal NewLockInfoBar As Boolean)

   m_LockInfoBar = NewLockInfoBar
   PropertyChanged "LockInfoBar"

End Property

Public Property Get Picture() As StdPicture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."

   Set Picture = m_Picture

End Property

Public Property Let Picture(ByRef NewPicture As StdPicture)

   Set Picture = NewPicture

End Property

Public Property Set Picture(ByRef NewPicture As StdPicture)

   Set m_Picture = NewPicture
   PropertyChanged "Picture"
   Set NewPicture = Nothing
   
   Call CalcCalendar

End Property

Public Property Get SelectionType() As SelectionTypes
Attribute SelectionType.VB_Description = "Returns/sets the select type of selected-cells. Can be Single or Multi."

   SelectionType = m_SelectionType

End Property

Public Property Let SelectionType(ByVal NewSelectionType As SelectionTypes)

   m_SelectionType = NewSelectionType
   PropertyChanged "SelectionType"
   
   If Not m_Locked And (m_SelectionType = SingleCell) Then Call SetSingleSelect

End Property

Public Property Get SelectedDayMark() As Boolean
Attribute SelectedDayMark.VB_Description = "Returns/sets a value that determines to show or hide the selectedday-marker."

   SelectedDayMark = m_SelectedDayMark

End Property

Public Property Let SelectedDayMark(ByVal NewSelectedDayMark As Boolean)

   m_SelectedDayMark = NewSelectedDayMark
   PropertyChanged "SelectedDayMark"
   
   Call DrawCalendar

End Property

Public Property Get ShowDayOfYear() As Boolean
Attribute ShowDayOfYear.VB_Description = "Returns/sets to show or hide the day of year."

   ShowDayOfYear = m_ShowDayOfYear

End Property

Public Property Let ShowDayOfYear(ByVal NewShowDayOfYear As Boolean)

   m_ShowDayOfYear = NewShowDayOfYear
   PropertyChanged "ShowDayOfYear"
   
   Call DrawCalendar

End Property

Public Property Get ShowInfoBar() As BarFrameStyles
Attribute ShowInfoBar.VB_Description = "Returns/sets the show style to display the infobar of an calendar control."

   ShowInfoBar = m_ShowInfoBar

End Property

Public Property Let ShowInfoBar(ByVal NewShowInfoBar As BarFrameStyles)

   m_ShowInfoBar = NewShowInfoBar
   PropertyChanged "ShowInfoBar"
   
   Call UserControl_Resize

End Property

Public Property Get ShowNavigationBar() As BarFrameStyles
Attribute ShowNavigationBar.VB_Description = "Returns/sets the show style to display the navigationbar of an calendar control."

   ShowNavigationBar = m_ShowNavigationBar

End Property

Public Property Let ShowNavigationBar(ByVal NewShowNavigationBar As BarFrameStyles)

   m_ShowNavigationBar = NewShowNavigationBar
   PropertyChanged "ShowNavigationBar"
   
   Call UserControl_Resize

End Property

Public Property Get ShowToolTipText() As Boolean
Attribute ShowToolTipText.VB_Description = "Returns/sets to show or hide the calendar tooltiptexts."

   ShowToolTipText = m_ShowToolTipText

End Property

Public Property Let ShowToolTipText(ByVal NewShowToolTipText As Boolean)

   m_ShowToolTipText = NewShowToolTipText
   PropertyChanged "ShowToolTipText"

End Property

Public Property Get WeekDayViewChar() As DayWidthFormats
Attribute WeekDayViewChar.VB_Description = "Returns/sets the number of character used in the header-cells  (1-3)."

   WeekDayViewChar = m_WeekDayViewChar - 1

End Property

Public Property Let WeekDayViewChar(ByVal NewWeekDayViewChar As DayWidthFormats)

   If NewWeekDayViewChar < D Then NewWeekDayViewChar = D
   If NewWeekDayViewChar > Ddd Then NewWeekDayViewChar = Ddd
   
   m_WeekDayViewChar = NewWeekDayViewChar + 1
   PropertyChanged "WeekDayViewChar"
   
   If m_Locked Then Exit Property
   
   Call SetWeekDayHeaderText
   Call DrawCalendar

End Property

Public Property Get WeekNumberForeColor() As OLE_COLOR
Attribute WeekNumberForeColor.VB_Description = "Returns/sets the foreground color used to display header-cells."

   WeekNumberForeColor = m_WeekNumberForeColor

End Property

Public Property Let WeekNumberForeColor(ByVal NewWeekNumberForeColor As OLE_COLOR)

   m_WeekNumberForeColor = NewWeekNumberForeColor
   PropertyChanged "WeekNumberForeColor"
   
   Call DrawCalendar

End Property

' returns the days between the calendar date and specified date
Public Function DaysBetween(ByVal IsDate As Date) As Long

   If Year(IsDate) < 1583 Then Exit Function
   
   DaysBetween = DateDiff("d", DateSerial(m_CalYear, m_CalMonth, m_CalDay), IsDate)

End Function

' returns the day of the current year
Public Function DayOfYear() As Integer

   DayOfYear = DateDiff("d", "01-01-" & CStr(m_CalYear), DateSerial(m_CalYear, m_CalMonth, m_CalDay)) + 1

End Function

' fills the language variables with external language data
Public Function FillExternalLanguage(ByRef NewLanguageText() As String, Optional ByRef ErrorReturn As String) As Boolean

Const VAR_NAME As String = "NewLanguageText("

Dim intCount   As Integer
Dim intItems   As Integer
Dim strItems() As String

   ErrorReturn = ""
   
   If UBound(NewLanguageText) < 7 Then
      ErrorReturn = VAR_NAME & ") array must have 8 lines! (0 to 7)" & vbCrLf & vbCrLf
      
      For intCount = 0 To 7
         ErrorReturn = ErrorReturn & VAR_NAME & GetLanguageVarInfo(intCount) & vbCrLf
      Next 'intCount
      
      Exit Function
   End If
   
   For intCount = 0 To 7
      If NewLanguageText(intCount) = "" Then
         If (intCount = 0) Or (intCount = 2) Then
            If intCount Then
               For intItems = 1 To 12
                  NewLanguageText(2) = NewLanguageText(2) & StrConv(MonthName(intCount), vbProperCase) & IIf(intItems < 12, ",", "")
               Next 'intItems
               
            Else
               For intItems = 1 To 7
                  NewLanguageText(0) = NewLanguageText(0) & StrConv(WeekdayName(intItems, , vbSunday), vbProperCase) & IIf(intItems < 7, ",", "")
               Next 'intItems
            End If
         End If
      End If
      
      strItems = Split(NewLanguageText(intCount), ",")
      intItems = GetLanguageVarInfo(intCount, True)
      
      If UBound(strItems) <> intItems Then ErrorReturn = ErrorReturn & VAR_NAME & GetLanguageVarInfo(intCount) & vbCrLf
      
      Erase strItems
   Next 'intCount
   
   If Len(ErrorReturn) Then Exit Function
   
   With LanguageText(CalLanguage)
      .DayNames = NewLanguageText(0)
      .Miscellaneous = NewLanguageText(1)
      .MonthNames = NewLanguageText(2)
      .QuarterNames = NewLanguageText(3)
      .SeasonNames = NewLanguageText(4)
      .MoonPhaseNames = NewLanguageText(5)
      .MoonPhaseText = NewLanguageText(6)
      .ZodiacNames = NewLanguageText(7)
      FillExternalLanguage = True
   End With

End Function

' returns the total days of specified month
Public Function GetMonthDays(Optional ByVal IsMonth As Integer) As Integer

   If (IsMonth < 1) Or (IsMonth > 12) Then IsMonth = m_CalMonth
   
   If (IsMonth = 12) And (m_CalYear = 9999) Then
      GetMonthDays = 31
      
   Else
      GetMonthDays = Day(DateAdd("d", -1, DateSerial(m_CalYear, IsMonth + 1, 1)))
   End If

End Function

' returns the name of specified month
Public Function GetMonthName(ByVal IsMonth As Integer) As String

   If (IsMonth < 1) Or (IsMonth > 12) Then IsMonth = m_CalMonth
   
   GetMonthName = GetTextPart(LanguageText(CalLanguage).MonthNames, IsMonth)

End Function

' returns the deatail number of the moonphase from 0.00 to 0.99
Public Function GetMoonPhaseDetail(ByVal IsDate As Date) As Double

   GetMoonPhaseDetail = CalcMoonPhase(DatePart("m", IsDate), DatePart("d", IsDate), DatePart("yyyy", IsDate))

End Function

' returns the moonphase of given date
Public Function GetMoonPhaseExact(ByVal IsDate As Date) As MoonTypes

Dim dteDate As Date
Dim dblPrev As Double
Dim dblCurr As Double
Dim dblNext As Double

   dteDate = Int(IsDate) - 1
   dblPrev = CalcMoonPhase(DatePart("m", dteDate), DatePart("d", dteDate), DatePart("yyyy", dteDate))
   dteDate = Int(IsDate)
   dblCurr = CalcMoonPhase(DatePart("m", dteDate), DatePart("d", dteDate), DatePart("yyyy", dteDate))
   dteDate = Int(IsDate) + 1
   dblNext = CalcMoonPhase(DatePart("m", dteDate), DatePart("d", dteDate), DatePart("yyyy", dteDate))
   GetMoonPhaseExact = GetMoonQuarter(dblPrev, dblCurr, dblNext)

End Function

' returns the moonphase info for the specified date
Public Function GetMoonPhaseInfo(ByVal IsDate As Date, Optional ByRef IsIcon As StdPicture) As String

Dim intCount   As Integer
Dim intPart    As Integer
Dim sngValue   As Single
Dim strText(2) As String

   intPart = GetMoonPhaseExact(IsDate)
   
   If intPart = NoMoon Then
      sngValue = GetMoonPhaseDetail(IsDate)
      
      If sngValue < 0.125 Then
         intPart = NewMoon
         intCount = -1
         
      ElseIf sngValue < 0.25 Then
         intPart = FirstQuarter
         intCount = 1
         
      ElseIf sngValue < 0.375 Then
         intPart = FirstQuarter
         intCount = -1
         
      ElseIf sngValue < 0.5 Then
         intPart = FullMoon
         intCount = 1
         
      ElseIf sngValue < 0.626 Then
         intPart = FullMoon
         intCount = -1
         
      ElseIf sngValue < 0.75 Then
         intPart = LastQuarter
         intCount = 1
         
      ElseIf sngValue < 0.875 Then
         intPart = LastQuarter
         intCount = -1
         
      Else
         intPart = NewMoon
         intCount = 1
      End If
      
      intCount = DateDiff("d", IsDate, NearestQuarterDate(IsDate, intCount))
   End If
   
   strText(0) = GetTextPart(LanguageText(CalLanguage).MoonPhaseNames, intPart)
   
   If intCount Then
      strText(0) = strText(0) & Separator(1) & Format(DateAdd("d", intCount, IsDate), GetDateFormat(False)) & Separator(2)
      strText(1) = GetTextPart(LanguageText(CalLanguage).MoonPhaseText, ((1 And (intCount = 1)) + (2 And (intCount = -1)) + (3 And (intCount > 1)) + (4 And (intCount < -1))))
      strText(1) = Split(strText(1), "#")(0) & Abs(intCount) & Split(strText(1), "#", 2)(1) & " "
      
      If intPart / 2 = intPart \ 2 Then strText(2) = GetTextPart(LanguageText(CalLanguage).MoonPhaseText, 5) & " "
      
      strText(0) = strText(1) & strText(2) & LCase(strText(0))
   End If
   
   If m_Hemisphere Then
      If intPart = FirstQuarter Then
         intPart = LastQuarter
         
      ElseIf intPart = LastQuarter Then
         intPart = FirstQuarter
      End If
   End If
   
   Set IsIcon = imgMoonPhase.Item(intPart).Picture
   GetMoonPhaseInfo = strText(0)
   Erase strText

End Function

' returns the quarter info for the specified quarter
Public Function GetQuarterInfo(ByVal IsQuarter As Integer, Optional ByRef IsIcon As StdPicture) As String

   If IsQuarter < 1 Then IsQuarter = 1
   If IsQuarter > 4 Then IsQuarter = 4
   
   Set IsIcon = imgQuarter.Item(IsQuarter).Picture
   GetQuarterInfo = GetTextPart(LanguageText(CalLanguage).QuarterNames, IsQuarter) & Separator(1) & GetDate(imgQuarter.Item(IsQuarter).ToolTipText) & Separator(2)

End Function

' returns the season info for the specified season
Public Function GetSeasonInfo(ByVal IsSeason As Integer, Optional ByRef IsIcon As StdPicture) As String

Dim intPeriod As Integer

   If IsSeason < 1 Then IsSeason = 1
   If IsSeason > 4 Then IsSeason = 4
   
   Set IsIcon = imgSeason.Item(IsSeason).Picture
   intPeriod = IsSeason
   
   If m_Hemisphere Then intPeriod = IsSeason + (2 And (IsSeason < 3)) - (2 And (IsSeason > 2))
   
   GetSeasonInfo = GetTextPart(LanguageText(CalLanguage).SeasonNames, IsSeason) & Separator(1) & GetDate(imgSeason.Item(intPeriod).ToolTipText) & Separator(2)

End Function

' returns the dayname of the week specified by IsDate
Public Function GetWeekdayName(ByRef IsDayDate As Variant) As String

Dim intDayNumber As Integer

   If InStr(IsDayDate, "-") = 0 Or InStr(IsDayDate, "/") = 0 Or InStr(IsDayDate, "\") = 0 Then GoTo GetDay
   
   On Local Error GoTo GetDay
   intDayNumber = WeekDay(Format(IsDayDate, "dd-mm-yyyy"))
   
   GoTo GetDayName
   
GetDay:
   intDayNumber = Val(IsDayDate)
   
GetDayName:
   If (intDayNumber > 0) And (intDayNumber < 8) Then GetWeekdayName = GetTextPart(LanguageText(CalLanguage).DayNames, intDayNumber)
   
   On Local Error GoTo 0

End Function

' returns the zodiacsign info for the specified zodiacsign
Public Function GetZodiacInfo(ByVal IsZodiacSign As Integer, Optional ByRef IsIcon As StdPicture) As String

   If IsZodiacSign < 1 Then IsZodiacSign = 1
   If IsZodiacSign > 12 Then IsZodiacSign = 12
   
   Set IsIcon = imgZodiac.Item(IsZodiacSign).Picture
   GetZodiacInfo = GetTextPart(LanguageText(CalLanguage).ZodiacNames, IsZodiacSign) & Separator(0) & GetDate(imgZodiac.Item(IsZodiacSign).ToolTipText) & Separator(1) & imgZodiac.Item(IsZodiacSign).Tag & Separator(2)

End Function

' returns True if that day is selected indeed
Public Function IsDaySel(ByVal IsDay As Integer) As Boolean

   IsDaySel = (CalCell(CalcDay(IsDay)).Type = Selected)

End Function

' marks or demarks a calendar day for the specified day and sets the tooltiptext of it
Public Sub DayMarking(ByVal IsDay As Integer, ByVal MarkIndex As Integer, ByVal OnOff As Boolean, Optional ByVal TipText As String)

   If MarkIndex < 1 Then MarkIndex = 1
   If MarkIndex > 5 Then MarkIndex = 5
   
   With CalCell(CalcDay(IsDay))
      MarkIndex = MarkIndex - 1
      
      If OnOff Then
         .Mark = .Mark Or (2 ^ MarkIndex)
         
      Else
         .Mark = .Mark And Not (2 ^ MarkIndex)
      End If
      
      .MarkTipText(MarkIndex) = TipText
   End With

End Sub

' selects or deselects a day for the specified day
Public Sub DaySelect(ByVal IsDay As Integer, ByVal OnOff As Boolean)

   If OnOff Then
      CalCell(CalcDay(IsDay)).Type = Selected
      
   Else
      CalCell(CalcDay(IsDay)).Type = Normal
   End If

End Sub

' to refresh the whole calendar
Public Sub Refresh()

   Call DrawCalendar

End Sub

' sets the color of the markers
Public Sub SetMarkColors(Optional ByVal Color1 As Long = vbRed, Optional ByVal Color2 As Long = vbGreen, Optional ByVal Color3 As Long = vbMagenta, Optional ByVal Color4 As Long = vbYellow, Optional ByVal Color5 As Long = vbCyan)

   MarkColor(0) = Color1
   MarkColor(1) = Color2
   MarkColor(2) = Color3
   MarkColor(3) = Color4
   MarkColor(4) = Color5

End Sub

' calculate the day position in the month-grid
Private Function CalcDay(ByVal IsDay As Integer) As Integer

   CalcDay = OffsetCell + IsDay + (OffsetCell + IsDay - ((OffsetCell \ 8) * 8 + 1)) \ 7

End Function

' returns the phase of the moon for given date
Private Function CalcMoonPhase(ByVal IsMonth As Long, ByVal IsDay As Long, ByVal IsYear As Long) As Double

Const ECCENT       As Double = 0.016718         'eccentricity of Earth's orbit Elements of the Moon's orbit, epoch 1980.0

Dim dblAnnEquation As Double
Dim dblJulianDay   As Double
Dim dblEpoch       As Double
Dim dblEvection    As Double
Dim dblLambdaSun   As Double
Dim dblLongitude   As Double
Dim dblMoonLongE   As Double
Dim dblMoonLongP   As Double

   ' convert month, day, year into a Julian daynumber
   dblJulianDay = GetJulianDay(IsYear, IsMonth, IsDay) - 2444238.5               '2444238.5 = 1980 January 0.0
   ' calculation of the Sun's position
   dblEpoch = FixAngle(FixAngle((360 / 365.2422) * dblJulianDay) + -3.76286299999998)
   dblLambdaSun = FixAngle((2 * ToDeg(Atn(Sqr((1 + ECCENT) / (1 - ECCENT)) * Tan(Kepler(dblEpoch, ECCENT) / 2)))) + 282.596403) '282.596403 = ecliptic longitude of the Sun at perigee
   ' calculation of the Moon's position
   dblMoonLongE = FixAngle(13.1763966 * dblJulianDay + 64.975464)                '64.975464 = moon's mean lonigitude at the epoch
   dblMoonLongP = FixAngle(dblMoonLongE - 0.1114041 * dblJulianDay - 349.383063) '349.383063 = mean longitude of the perigee at the epoch
   dblEvection = 1.2739 * Sin(ToRad(2 * (dblMoonLongE - dblLambdaSun) - dblMoonLongP))
   dblAnnEquation = 0.1858 * Sin(ToRad(dblEpoch))
   dblMoonLongP = dblMoonLongP + dblEvection - dblAnnEquation - (0.37 * Sin(ToRad(dblEpoch)))
   dblLongitude = dblMoonLongE + dblEvection + (6.2886 * Sin(ToRad(dblMoonLongP))) - dblAnnEquation + (0.214 * Sin(ToRad(2 * dblMoonLongP)))
   ' calculation of the phase of the Moon
   CalcMoonPhase = (FixAngle((dblLongitude + (0.6583 * Sin(ToRad(2 * (dblLongitude - dblLambdaSun))))) - dblLambdaSun) / 360#)

End Function

' check if mouse is in button
Private Function CheckButton(ByRef Button() As CalButtons, ByVal X As Long, ByVal Y As Long, ByVal Max As Integer, Optional ByVal Index As Integer = -1) As Integer

Dim intCount As Integer

   CheckButton = -1
   
   If Index > -1 Then
      If PtInRect(Button(Index).Rect, X, Y) Then CheckButton = intCount
      
   Else
      For intCount = 0 To Max
         If PtInRect(Button(intCount).Rect, X, Y) Then
            CheckButton = intCount
            Exit For
         End If
      Next 'intCount
   End If

End Function

' returns fix angle number
Private Function FixAngle(ByVal Age As Double) As Double

   FixAngle = (Age - 360 * Int(Age / 360))

End Function

Private Function GetColor(ByVal IsColor As Integer) As Integer

   GetColor = Val("&H" & Hex((IsColor / &HFF&) * &HFFFF&))

End Function

' returns the date month format
Private Function GetDate(ByVal IsDate As String) As String

Dim strBuffer() As String

   If InStr(IsDate, " ") = 0 Then Exit Function
   
   strBuffer = Split(IsDate, " ", 2)
   GetDate = SetDateFormat(Day(CDate(strBuffer(0))), LCase(GetTextPart(LanguageText(CalLanguage).MonthNames, Month(CDate(strBuffer(0)))))) & " - " & SetDateFormat(Day(CDate(strBuffer(1))), LCase(GetTextPart(LanguageText(CalLanguage).MonthNames, Month(CDate(strBuffer(1))))))
   Erase strBuffer

End Function

' returns the selected date format
Private Function GetDateFormat(ByVal LongFormat As Boolean)

   If LongFormat Then
      GetDateFormat = "dddd, " & Choose(m_DateFormat + 1, "d mmmm yyyy", "mmmm d yyyy", "yyyy mmmm d")
      
   Else
      GetDateFormat = Choose(m_DateFormat + 1, "dd-mm-yyyy", "mm-dd-yyyy", "yyyy-mm-dd")
   End If

End Function

' returns the color of the specified day
Private Function GetDayColor(ByVal IsDay As Integer)

   If IsDay = vbSunday Then
      GetDayColor = m_CellForeColorSunday
      
   ElseIf IsDay = vbMonday Then
      GetDayColor = m_CellForeColorMonday
      
   ElseIf IsDay = vbTuesday Then
      GetDayColor = m_CellForeColorTuesday
      
   ElseIf IsDay = vbWednesday Then
      GetDayColor = m_CellForeColorWednesday
      
   ElseIf IsDay = vbThursday Then
      GetDayColor = m_CellForeColorThursday
      
   ElseIf IsDay = vbFriday Then
      GetDayColor = m_CellForeColorFriday
      
   ElseIf IsDay = vbSaturday Then
      GetDayColor = m_CellForeColorSaturday
   End If

End Function

' returns the julianday number of the given date
' the Julian date is computed for noon UT
Private Function GetJulianDay(ByVal IsYear As Long, ByVal IsMonth As Long, ByVal IsDay As Long) As Double

Dim dblTemp(1) As Double
Dim lngMonth   As Long
Dim lngYear    As Long

   lngMonth = IsMonth
   lngYear = IsYear
   
   If lngYear < 0 Then lngYear = lngYear + 1
   
   If IsMonth < 3 Then
       lngMonth = IsMonth + 12
       lngYear = lngYear - 1
   End If
   
   If (IsYear > 1582) Or ((IsYear = 1582) And (IsMonth = 10) And (IsDay > 14)) Then
      dblTemp(0) = Int(lngYear / 100)
      dblTemp(0) = 2 - dblTemp(0) + Int(dblTemp(0) / 4)
   End If
   
   dblTemp(1) = (365.25 * lngYear) - 0.75
   
   If lngYear >= 0 Then
      dblTemp(1) = Int(365.25 * lngYear) - 694025
      
   Else
      dblTemp(1) = Sgn(dblTemp(1)) * Int(Abs(dblTemp(1))) - 694025
   End If
   
   GetJulianDay = dblTemp(0) + dblTemp(1) + Int(30.6001 * (lngMonth + 1)) + IsDay + 2415020
   Erase dblTemp

End Function

' returns name and/or total items of the language variable, specified by index
Private Function GetLanguageVarInfo(ByVal Index As Integer, Optional ByVal GetItems As Boolean) As Variant

   GetLanguageVarInfo = Choose(Index + 1, 6, 4, 11, 3, 3, 3, 4, 11)
   
   If Not GetItems Then GetLanguageVarInfo = Index & " = " & Choose(Index + 1, "DayNames", "Miscellaneous", "MonthNames", "QuarterNames", "SeasonNames", "MoonPhaseNames", "MoonPhaseText", "ZodiacNames") & ") needs: " & GetLanguageVarInfo + 1 & " items!"

End Function

' returns the most nearest phase of the moon
Private Function GetMoonQuarter(ByVal PrevQuarter As Double, ByVal CurrQuarter As Double, ByVal NextQuarter As Double) As Long

Dim dblPhase      As Double
Dim dblDifference As Double
Dim lngQuarter    As Long

   If CurrQuarter < PrevQuarter Then CurrQuarter = CurrQuarter + 1
   If NextQuarter < PrevQuarter Then NextQuarter = NextQuarter + 1
   
   GetMoonQuarter = NoMoon
   
   For lngQuarter = 1 To 4
      dblPhase = lngQuarter / 4
      
      If PrevQuarter < dblPhase Then
         If NextQuarter > dblPhase Then
            dblDifference = Abs(CurrQuarter - dblPhase)
            
            If dblDifference < dblPhase - PrevQuarter Then
               If dblDifference < NextQuarter - dblPhase Then
                  GetMoonQuarter = (lngQuarter Mod 4) + 1
                  Exit Function
               End If
            End If
         End If
      End If
   Next 'lngQuarter

End Function

' returns the Locked season
Private Function GetSeason() As Integer

   ' Spring / Autumn
   If m_CalMonth < 4 Then
      GetSeason = 4 - (3 And ((m_CalMonth = 3) And (m_CalDay > 20)))
      
   ' Summer / Winter
   ElseIf m_CalMonth < 7 Then
      GetSeason = 1 + (1 And ((m_CalMonth = 6) And (m_CalDay > 20)))
      
   ' Autumn / Spring
   ElseIf m_CalMonth < 10 Then
      GetSeason = 2 + (1 And ((m_CalMonth = 9) And (m_CalDay > 21)))
      
   ' Winter / Summer
   Else
      GetSeason = 3 + (1 And ((m_CalMonth = 12) And (m_CalDay > 20)))
   End If
   
   If m_Hemisphere Then GetSeason = GetSeason + (2 And (GetSeason < 3)) - (2 And (GetSeason > 2))

End Function

' returns the system language
Private Function GetSystemLanguage() As Integer

'Const LOCALE_SENGLANGUAGE As Long = &H1001
'Const LOCALE_USER_DEFAULT As Long = &H400

Dim strBuffer             As String * 100

   Select Case Left(strBuffer, GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SENGLANGUAGE, strBuffer, Len(strBuffer)) - 1)
      Case "Spanish"
         GetSystemLanguage = Spanish
         
      Case "Dutch"
         GetSystemLanguage = Dutch
         
      Case "French"
         GetSystemLanguage = French
         
      Case "Italian"
         GetSystemLanguage = Italian
         
      Case Else
         GetSystemLanguage = English
   End Select

End Function

' returns the part of text form the specified languages string
Private Function GetTextPart(ByVal Text As String, ByVal Index As Integer) As String

   If Len(Text) And (Index > 0) Then GetTextPart = Split(Text, ",")(Index - 1)

End Function

' returns the Locked zodiacsign
Private Function GetZodiacSign() As Integer

   If m_CalMonth = 1 Then
      GetZodiacSign = 11 - (1 And (m_CalDay < 21))
      
   ElseIf m_CalMonth = 2 Then
      GetZodiacSign = 12 - (1 And (m_CalDay < 21))
      
   ElseIf m_CalMonth = 3 Then
      GetZodiacSign = 1 + (11 And (m_CalDay < 21))
      
   ElseIf m_CalMonth = 4 Then
      GetZodiacSign = 2 - (1 And (m_CalDay < 21))
      
   ElseIf m_CalMonth = 5 Then
      GetZodiacSign = 3 - (1 And (m_CalDay < 21))
      
   ElseIf m_CalMonth = 6 Then
      GetZodiacSign = 4 - (1 And (m_CalDay < 21))
      
   ElseIf m_CalMonth = 7 Then
      GetZodiacSign = 5 - (1 And (m_CalDay < 23))
      
   ElseIf m_CalMonth = 8 Then
      GetZodiacSign = 6 - (1 And (m_CalDay < 23))
      
   ElseIf m_CalMonth = 9 Then
      GetZodiacSign = 7 - (1 And (m_CalDay < 23))
      
   ElseIf m_CalMonth = 10 Then
      GetZodiacSign = 8 - (1 And (m_CalDay < 23))
      
   ElseIf m_CalMonth = 11 Then
      GetZodiacSign = 9 - (1 And (m_CalDay < 23))
      
   ElseIf m_CalMonth = 12 Then
      GetZodiacSign = 10 - (1 And (m_CalDay < 22))
   End If

End Function

' returns the solved equation of Kepler
Private Function Kepler(ByVal Epc As Double, ByVal Ecc As Double) As Double

Dim dblDelta    As Double
Dim dblEquation As Double

   Epc = ToRad(Epc)
   dblEquation = Epc
   
   Do
      dblDelta = dblEquation - Ecc * Sin(dblEquation) - Epc
      dblEquation = dblEquation - (dblDelta / (1 - Ecc * Cos(dblEquation)))
   Loop While (Abs(dblDelta) > 0.000001)
   
   Kepler = dblEquation

End Function

' returns the date of moonquarter for the specified date
Private Function NearestQuarterDate(ByVal IsDate As Date, ByVal AddDay As Integer) As Date

   NearestQuarterDate = IsDate
   
   Do
      NearestQuarterDate = DateAdd("d", AddDay, NearestQuarterDate)
   Loop Until GetMoonPhaseExact(NearestQuarterDate) > NoMoon

End Function

' returns the specified date in the selected dateformat
Private Function SetDateFormat(ByVal IsDay As String, ByVal IsMonth As String) As String

   If m_DateFormat = [yyyy-mm-dd] Then
      SetDateFormat = IsMonth & " " & IsDay
      
   Else
      SetDateFormat = IsDay & " " & IsMonth
   End If

End Function

' returns the Deg from Rad value
Private Function ToDeg(ByVal Rad As Double) As Double

   ToDeg = ((Rad) * (180# / PI))

End Function

' returns the Rad from Deg value
Private Function ToRad(ByVal Deg As Double) As Double

   ToRad = ((Deg) * (PI / 180#))

End Function

' returns -1 if ole_colors cannot be translated
Private Function TranslateColor(ByVal Colors As OLE_COLOR, Optional ByVal Palette As Long) As Long

   If OleTranslateColor(Colors, Palette, TranslateColor) Then TranslateColor = -1

End Function

' returns the RGB value
Private Function TranslateRGB(ByVal ColorVal As Long, ByVal Part As Long) As Long

   TranslateRGB = CLng("&H" + UCase(Mid(Format(Trim(Hex(ColorVal)), "#" & String(6, vbKey0)), 5 - Part * 2, 2)))

End Function

' month or year changed, so reset calendar content
Private Sub CalcCalendar()

Dim blnCalBegin   As Boolean
Dim blnCalEnd     As Boolean
Dim dteDayCounter As Date
Dim dteFirstDay   As Date
Dim dteLastDay    As Date
Dim intCell       As Integer
Dim intCount      As Integer
Dim intOtherDay   As Integer
Dim intStart      As Integer

   If m_Locked Then Exit Sub
   
   ReDim CalCell(55) As CalCells
   
   Cls
   intCell = 9
   
   If m_Picture Is Nothing Then
      UserControl.Picture = Nothing
      
   Else
      PaintPicture m_Picture, 0, 0, ScaleWidth, ScaleHeight, 0, 0, , , vbSrcCopy
   End If
   
   CalCell(0).Text = DayOfYear
   CalCell(0).Type = DayOfTheYear
   dteFirstDay = DateSerial(m_CalYear, m_CalMonth, 1)
   WeekDayOfFirstDay = WeekDay(dteFirstDay, m_FirstWeekDay)
   WeekDayOfFirstDay = WeekDayOfFirstDay + (7 And (WeekDayOfFirstDay < 3))
   
   If (m_CalYear = 1583) And (m_CalMonth = 1) Then blnCalBegin = True
   If (m_CalYear = 9999) And (m_CalMonth = 12) Then blnCalEnd = True
   If Not blnCalBegin Then intOtherDay = GetMonthDays(m_CalMonth - 1) - (WeekDayOfFirstDay - 2)
   
   If blnCalEnd Then
      dteLastDay = DateSerial(9999, 12, 31)
      
   Else
      dteLastDay = DateSerial(m_CalYear, m_CalMonth + 1, 0)
   End If
   
   If m_GradientStyle And (m_Picture Is Nothing) Then Call DrawGradient(CalendarRect, UserControl.BackColor, m_GradientColor, m_GradientStyle)
   
   Call SetWeekDayHeaderText
   
   For intCount = 1 To 42
      With CalCell(intCell)
         If intStart = -1 Then
            .Text = CELL_SPACE
            .Type = OtherMonths
         End If
         
         If intStart = 0 Then                    ' first empty cells part
            If WeekDayOfFirstDay = intCount Then ' cell is first day?
               intStart = 1                      ' start filling days
               dteDayCounter = dteFirstDay
               ' store offset
               OffsetCell = intCount + 7 + (intCount Mod 8 And (intCount > 8))
               
            Else
               If m_CellOtherMonthView And Not blnCalBegin Then
                  .Text = intOtherDay
                  
               Else
                  .Text = CELL_SPACE
               End If
               
               .Type = OtherMonths
               intOtherDay = intOtherDay + 1
            End If
         End If
         
         If intStart = 1 Then
            If dteDayCounter > dteLastDay Then   ' stop at last day
               intStart = 2
               intOtherDay = 1
               
            Else
               .Text = Day(dteDayCounter)
               
               If blnCalEnd Then
                  If Day(dteDayCounter) < 31 Then
                     dteDayCounter = DateSerial(9999, 12, Day(dteDayCounter) + 1)
                     
                  Else
                     intStart = -1
                  End If
                  
               Else
                  dteDayCounter = DateAdd("d", 1, dteDayCounter)
               End If
               
               If CInt(.Text) = m_CalDay Then
                  .Type = Selected               ' if current day Type is Selected
                  CurrentCell = intCell
                  
               Else
                  .Type = Normal
               End If
            End If
         End If
         
         If intStart = 2 Then
            If m_CellOtherMonthView Then
               .Text = intOtherDay
               
            Else
               .Text = CELL_SPACE
            End If
            
            .Type = OtherMonths
            intOtherDay = intOtherDay + 1
         End If
      End With
      
      intCell = intCell + 1 + (1 And ((intCell + 1) Mod 8 = 0))
   Next 'intCount
   
   If blnCalEnd Then
      MonthDays = 31
      
   Else
      MonthDays = GetMonthDays(m_CalMonth)
   End If
   
   Call CalcWeeks
   Call DrawNavigation
   Call DrawInfoBar
   
   UserControl.Picture = UserControl.Image
   BitBlt picCalCell.hDC, 0, 0, picCalCell.ScaleWidth, picCalCell.ScaleHeight, hDC, 0, 0, vbSrcCopy
   
   Call SetButtonTipText
   Call DrawCalendar

End Sub

' calculate the weekdays of the current month
Private Sub CalcWeeks()

Dim intCell  As Integer
Dim intDay   As Integer
Dim intMonth As Integer
Dim intTemp  As Integer
Dim intWeek  As Integer
Dim intYear  As Integer
Dim strText  As String

   For intCell = 8 To 48 Step 8
      strText = CELL_SPACE
      intMonth = m_CalMonth
      intYear = m_CalYear
      intDay = Val(CalCell(intCell + 1).Text)
      
      If intDay = 0 Then intDay = Val(CalCell(intCell + 7).Text)
      
      If intDay Then
         If m_CellOtherMonthView Then
            If (intDay > 15) And (intCell < 20) Then
               intMonth = intMonth - 1
               
               If intMonth < 1 Then
                  intMonth = 12
                  intYear = intYear - 1
               End If
            End If
            
            If (intDay < 16) And (intCell = 48) Then
               intMonth = intMonth + 1
               
               If intMonth > 12 Then
                  intMonth = 1
                  intYear = intYear + 1
               End If
            End If
         End If
         
         intWeek = DatePart("ww", DateSerial(intYear, intMonth, intDay), m_FirstWeekDay, vbFirstFourDays)
         
         If intWeek = 53 Then
            ' be shure it's week 53! (there are some bugs in the weekcalculation of VB5!)
            If intMonth = 12 Then
               If intYear < 9999 Then intTemp = DatePart("ww", DateSerial(intYear + 1, 1, 1), m_FirstWeekDay, vbFirstFourDays)
               
            Else
               If intYear > 1582 Then intTemp = DatePart("ww", DateSerial(intYear - 1, 12, 31), m_FirstWeekDay, vbFirstFourDays)
            End If
            
            If intWeek <> intTemp Then intWeek = intTemp
         End If
         
         strText = intWeek
      End If
      
      CalCell(intCell).Text = strText
      CalCell(intCell).Type = WeekNumber
      
      If strText <> CELL_SPACE Then CalCell(intCell).TipText = GetTextPart(LanguageText(CalLanguage).Miscellaneous, 2) & " " & strText
   Next 'intCell

End Sub

Private Sub CheckForMonth(ByRef IsMonth As Integer, Optional ByRef IsYear As Integer, Optional ByVal Add As Boolean)

   If Add Then
      If IsMonth = 12 Then
         IsMonth = 1
         IsYear = m_CalYear + 1
         
         If IsYear = 9999 Then IsMonth = 12
         
      Else
         IsMonth = m_CalMonth + 1
      End If
      
   Else
      If IsMonth = 1 Then
         IsMonth = 12
         IsYear = m_CalYear - 1
         
         If IsYear = 1583 Then IsMonth = 1
         
      Else
         IsMonth = m_CalMonth - 1
      End If
   End If

End Sub

' when a header (weekday) is clicked in Multi selection mode
Private Sub DaySelectAll(ByVal WeekDay As Integer)

Dim intCell As Integer

   If (CalCell(WeekDay).Type = Header) Or (CalCell(WeekDay).Type = WeekDay) Then
      For intCell = OffsetCell + 1 To CalcDay(MonthDays)
         If intCell Mod 8 Then
            If intCell Mod 8 = WeekDay Then
               CalCell(intCell).Type = Selected
               
            Else
               CalCell(intCell).Type = Normal
            End If
         End If
      Next 'intCell
      
   Else
      For intCell = OffsetCell + 1 To CalcDay(MonthDays)
         If intCell Mod 8 Then
            If intCell \ 8 = WeekDay \ 8 Then
               CalCell(intCell).Type = Selected
               
            Else
               CalCell(intCell).Type = Normal
            End If
         End If
      Next 'intCell
   End If
   
   Call DrawCalendar

End Sub

' routine to draw arrows on navigationbuttons
Private Sub DrawArrow(ByVal Index As Integer, ByVal X As Long, ByVal Y As Long)

Const ALTERNATE       As Long = 1

Dim lngBrush          As Long
Dim lngRegion         As Long
Dim ptaRegion(1 To 3) As PointAPI

   ptaRegion(1).X = X + Choose(Index, 0, 12, 7, 7)
   ptaRegion(1).Y = Y + Choose(Index, 6, 6, 0, 12)
   ptaRegion(2).X = X + Choose(Index, 12, 0, 13, 13)
   ptaRegion(2).Y = Y + Choose(Index, 0, 0, 12, 1)
   ptaRegion(3).X = X + Choose(Index, 12, 0, 0, 0)
   ptaRegion(3).Y = Y + Choose(Index, 12, 12, 12, 1)
   lngBrush = CreateSolidBrush(m_ArrowColor)
   lngRegion = CreatePolygonRgn(ptaRegion(1), 3, ALTERNATE)
   
   If lngRegion Then FillRgn hDC, lngRegion, lngBrush
   
   DeleteObject lngRegion
   DeleteObject lngBrush
   Erase ptaRegion

End Sub

' routine to set Raised or Sunken day-border
Private Sub DrawBevel(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal BevelType As Integer)

'Const BDR_RAISEDOUTER As Long = &H1
Const BDR_SUNKENOUTER As Long = &H2

Dim lngBorder         As Long
Dim rctBorder         As Rect

   With rctBorder
      .Left = X1
      .Top = Y1
      .Right = X2
      .Bottom = Y2
   End With
   
   Line (X1 + 2, Y1 + 2)-Step(X2 - X1 - 4, Y2 - Y1 - 4), m_ButtonColor, BF
   
   If m_Appearance = [3D] Then
      If BevelType - BevelIn Then
         lngBorder = BDR_RAISED
         
      Else
         lngBorder = BDR_SUNKEN
      End If
      
   Else
      If BevelType - BevelIn Then
         lngBorder = BDR_RAISEDOUTER
         
      Else
         lngBorder = BDR_SUNKENOUTER
      End If
   End If
   
   DrawEdge hDC, rctBorder, lngBorder, BF_RECT

End Sub

' routine to draw buttons Raised or Sunken
Private Sub DrawButton(ByRef Button() As CalButtons, ByVal Index As Integer, ByVal State As Boolean, ByVal MouseMoves As Boolean)

Dim lngBorder As Long

   If Not MouseMoves Then
      With Button(Index).Rect
         BitBlt hDC, .Left + 3, .Top + 3, .Right - .Left - 6, .Bottom - .Top - 6, hDC, .Left + 2, .Top + 2, vbSrcCopy
      End With
   End If
   
   If m_Appearance = [3D] Then
      If State Then
         lngBorder = BDR_SUNKEN
         
      Else
         lngBorder = BDR_RAISED
      End If
      
   ElseIf State Then
      lngBorder = BDR_SUNKENINNER
      
   Else
      lngBorder = BDR_RAISEDINNER
   End If
   
   DrawEdge hDC, Button(Index).Rect, lngBorder, BF_RECT

End Sub

' routine to draw cells in different modes
Private Sub DrawCalendar()

Dim intCount As Integer

   If m_Locked Then Exit Sub
   
   Cls
   
   For intCount = 0 To 55
      With CalCell(intCount)
         .X = (intCount Mod 8) * SizeX
         .Y = (intCount \ 8) * SizeY + (SizeY * 2 And CBool(m_ShowNavigationBar))
         
         Call DrawCell(intCount)
      End With
      
      If CalCell(intCount).Mark Then Call DrawMarkers(intCount)
   Next 'intCount
   
   Call SetCurrentDay
   Call DrawLabel
   Call DrawGrid
   Call DrawFrame

End Sub

' routine to draw cells in different modes
Private Sub DrawCell(ByVal Index As Integer, Optional ByVal Zoom As Boolean, Optional ByVal Selection As Boolean)

Dim cstMode      As CellStyles
Dim intChoose    As Integer
Dim intTemp      As Integer
Dim intX1        As Integer
Dim intX2        As Integer
Dim intY1        As Integer
Dim intY2        As Integer
Dim lngBackColor As Long
Dim lngForeColor As Long

   If Not m_ShowDayOfYear And (Index = 0) Then Exit Sub
   
   With CalCell(Index)
      intX1 = .X + (1 And (.Type = Selected))
      intY1 = .Y + (1 And (.Type = Selected))
      intX2 = SizeX - (2 And (.Type = Selected))
      intY2 = SizeY - (2 And (.Type = Selected))
      
      If .Type = DayOfTheYear Then
         cstMode = m_CellDayOfYearStyle
         lngBackColor = m_CellDayOfYearBackColor
         lngForeColor = m_CellDayOfYearForeColor
         
      ElseIf .Type = Header Or .Type = WeekNumber Then
         cstMode = m_CellHeaderStyle
         lngBackColor = m_CellHeaderBackColor
         
         If Selection Then
            lngForeColor = m_CellSelectHeaderForeColor
            
         ElseIf .Type = Header Then
            lngForeColor = GetDayColor(DatePart("w", DateSerial(m_CalYear, m_CalMonth, Val(CalCell(Index + 24).Text))))
            
         Else
            lngForeColor = m_WeekNumberForeColor
         End If
         
      ElseIf .Type = Normal Then
         cstMode = m_CellDaysStyle
         lngBackColor = m_CellDaysBackColor
         lngForeColor = GetDayColor(DatePart("w", DateSerial(m_CalYear, m_CalMonth, Val(CalCell(Index).Text))))
         
      ElseIf .Type = OtherMonths Then
         cstMode = m_CellOtherMonthStyle
         lngBackColor = m_CellOtherMonthBackColor
         intTemp = Val(CalCell(Index).Text)
         intChoose = -1 + (2 And (intTemp < 16))
         
         If intTemp Then intTemp = DatePart("w", DateSerial(m_CalYear, m_CalMonth + intChoose, intTemp))
         
         If intTemp = vbSunday Then
            lngForeColor = m_CellOtherMonthForeColor And &H808080
            
         Else
            lngForeColor = m_CellOtherMonthForeColor
         End If
         
      ElseIf .Type = Selected Then
         cstMode = m_CellSelectStyle
         lngBackColor = m_CellSelectBackColor
         lngForeColor = m_CellSelectForeColor
      End If
      
      If cstMode = EmptyCell Then
         BitBlt hDC, .X + 1, .Y + 1, SizeX - 2, SizeY - 2, picCalCell.hDC, .X + 1, .Y + 1, vbSrcCopy
         
      ElseIf cstMode = BevelIn Then
         Line (intX1, intY1)-Step(intX2, intY2), m_ButtonColor, BF
         
         Call DrawBevel(intX1, intY1, intX1 + intX2, intY1 + intY2, BevelIn)
         
      ElseIf cstMode = BevelOut Then
         Line (intX1, intY1)-Step(intX2, intY2), m_ButtonColor, BF
         
         Call DrawBevel(intX1, intY1, intX1 + intX2, intY1 + intY2, BevelOut)
         
      ElseIf cstMode = UserColor Then
         Line (intX1 + 1, intY1 + 1)-Step(intX2 - 2, intY2 - 2), lngBackColor, BF
         
      ' BlueSelect, DarkGraySelect or LightGraySelect
      Else
         Line (intX1 + 1, intY1 + 1)-Step(intX2 - 2, intY2 - 2), QBColor(9 - (cstMode - BlueSelect)), BF
         Line (intX1, intY1)-Step(intX2 - 1, intY2 - 1), QBColor(15), B
         
         If cstMode = LightGraySelect Then
            lngForeColor = QBColor(0)
            
         Else
            lngForeColor = QBColor(15)
         End If
      End If
      
      FontSize = SizeFont + (3 And Zoom)
      ForeColor = lngForeColor
      CurrentX = intX1 + (intX2 - TextWidth(Trim(.Text))) \ 2
      CurrentY = intY1 + (intY2 - TextHeight("X")) \ 2
      Print Trim(.Text)
      ForeColor = vbButtonText
   End With

End Sub

' routine to draw the calendar FrameStyle lines
Private Sub DrawFrame()

Dim intSize As Integer

   If m_FrameStyle = Off Then Exit Sub
   If m_FrameStyle = Small Then intSize = 8
   
   Line (0 + intSize, CalCell(0).Y + SizeY)-(CalCell(7).X + SizeX - intSize \ 2, CalCell(0).Y + SizeY), m_FrameColor
   Line (CalCell(0).X + SizeX, CalCell(0).Y + intSize)-(CalCell(0).X + SizeX, CalCell(49).Y + SizeY), m_FrameColor
   Line (0 + intSize, CalCell(49).Y + SizeY)-(CalCell(7).X + SizeX - intSize \ 2, CalCell(49).Y + SizeY), m_FrameColor

End Sub

' routine to draw the calendar or buttons gradient
Private Sub DrawGradient(ByRef ObjectRect As Rect, ByVal ObjectBackColor As Long, ByVal GradientColor As Long, ByVal GradientStyle As GradientStyles)

Dim lngRGB         As Long
Dim tvxGradient(1) As TriVertex
Dim rctGradient    As GradientRect

   If GradientStyle = NoGradient Then Exit Sub
   
   If (GradientStyle = RightToLeft) Or (GradientStyle = BottomToTop) Then
      lngRGB = TranslateColor(GradientColor)
      
   Else
      lngRGB = TranslateColor(ObjectBackColor)
   End If
   
   With tvxGradient(0)
      .X = ObjectRect.Left
      .Y = ObjectRect.Top
      .Red = GetColor(TranslateRGB(lngRGB, IsRed))
      .Green = GetColor(TranslateRGB(lngRGB, IsGreen))
      .Blue = GetColor(TranslateRGB(lngRGB, IsBlue))
   End With
   
   If (GradientStyle = RightToLeft) Or (GradientStyle = BottomToTop) Then
      lngRGB = TranslateColor(ObjectBackColor)
      
   Else
      lngRGB = TranslateColor(GradientColor)
   End If
   
   With tvxGradient(1)
      .X = ScaleX(ObjectRect.Right, ScaleMode, vbPixels)
      .Y = ScaleY(ObjectRect.Bottom, ScaleMode, vbPixels)
      .Red = GetColor(TranslateRGB(lngRGB, IsRed))
      .Green = GetColor(TranslateRGB(lngRGB, IsGreen))
      .Blue = GetColor(TranslateRGB(lngRGB, IsBlue))
   End With
   
   rctGradient.UpperLeft = 1
   rctGradient.LowerRight = 0
   GradientFill hDC, tvxGradient(0), 4, rctGradient, 1, 1 - (1 And (GradientStyle < TopToBottom))
   Erase tvxGradient

End Sub

' routine to draw the calendar grid
Private Sub DrawGrid()

Dim intChoose   As Integer
Dim intCount    As Integer
Dim intLastCell As Integer
Dim intSize(1)  As Integer

   If m_GridStyle = NoGrid Then Exit Sub
   If m_FrameStyle = Small Then intSize(0) = 6
   
   intSize(1) = intSize(0) \ 3
   
   If m_GridStyle = FullGrid Then
      For intCount = 17 To 49 Step 8
         Line (CalCell(intCount).X + 1 + intSize(1), CalCell(intCount).Y)-(CalCell(intCount + 6).X + SizeX - intSize(0), CalCell(intCount + 6).Y), m_GridColor
      Next 'intCount
      
      For intCount = 9 To 14
         Line (CalCell(intCount).X + SizeX, CalCell(intCount).Y + 1 + intSize(1))-(CalCell(intCount + 40).X + SizeX, CalCell(intCount + 40).Y + SizeY - intSize(1)), m_GridColor
      Next 'intCount
      
      If m_FrameStyle = Large Then Line (CalCell(7).X + SizeX, CalCell(7).Y)-(CalCell(55).X + SizeX, CalCell(55).Y + SizeY + 1), m_GridColor
      
   ElseIf m_GridStyle = MonthOnlyGrid Then
      For intCount = 17 To 41 Step 8
         Line (CalCell(intCount).X + 1 + intSize(1) + ((SizeX - 1 - intSize(1)) And (CalCell(intCount).Type = OtherMonths)), CalCell(intCount).Y)-(CalCell(intCount + 6).X + SizeX - intSize(0), CalCell(intCount + 6).Y), m_GridColor
      Next 'intCount
      
      intLastCell = CalcDay(MonthDays)
      
      If intLastCell < 48 Then
         intChoose = intLastCell
         
      Else
         intChoose = 55
      End If
      
      Line (CalCell(49).X + 1 + intSize(1), CalCell(49).Y)-(CalCell(intChoose).X + SizeX + 1 - ((intSize(0) + 1) And (intLastCell > 46)), CalCell(55).Y), m_GridColor
      
      For intCount = 9 To 14
         Line (CalCell(intCount).X + SizeX, CalCell(intCount).Y + 1 + intSize(1) + (SizeY - intSize(1) And (CalCell(intCount + 1).Type = OtherMonths)))-(CalCell(intCount + 40).X + SizeX, CalCell(intCount + 40 - (8 And (intCount + 32 > intLastCell))).Y + ((SizeY - intSize(1)) And (intCount + 40 <= intLastCell))), m_GridColor
      Next 'intCount
      
   ElseIf m_GridStyle = HorizontalGrid Then
      For intCount = 17 To 49 Step 8
         Line (CalCell(intCount).X + 1 + intSize(1), CalCell(intCount).Y)-(CalCell(intCount + 6).X + SizeX - intSize(0), CalCell(intCount + 6).Y), m_GridColor
      Next 'intCount
      
   ElseIf m_GridStyle = VerticalGrid Then
      For intCount = 9 To 14
         Line (CalCell(intCount).X + SizeX, CalCell(intCount).Y + 1 + intSize(1))-(CalCell(intCount + 40).X + SizeX, CalCell(intCount + 40).Y + SizeY - intSize(1)), m_GridColor
      Next 'intCount
   End If
   
   Erase intSize

End Sub

' routine to draw the information of the calendar
Private Sub DrawInfoBar()

Dim dteDate   As Date
Dim intCount  As Integer
Dim intLength As Integer
Dim lngBorder As Long
Dim stdIcon   As StdPicture

   If m_ShowInfoBar = Off Then Exit Sub
   
   With CalButton(9).Rect
      .Left = CalButton(5).Rect.Left
      .Top = CalButton(5).Rect.Top
      .Right = CalButton(8).Rect.Right
      .Bottom = CalButton(5).Rect.Bottom
      
      If m_ButtonGradientStyle Then
         Call DrawGradient(CalButton(9).Rect, m_ButtonColor, m_ButtonGradientColor, m_ButtonGradientStyle)
         
      Else
         Line (.Left + 2, .Top + 2)-Step(.Right - .Left - 4, .Bottom - .Top - 4), m_ButtonColor, BF
      End If
      
      If m_Appearance = [3D] Then
         lngBorder = BDR_RAISED
         
      Else
         lngBorder = BDR_RAISEDINNER
      End If
      
      DrawEdge hDC, CalButton(9).Rect, lngBorder, BF_RECT
   End With
   
   Separator(0) = " "
   Separator(1) = "  ("
   Separator(2) = ")"
   dteDate = DateSerial(m_CalYear, m_CalMonth, m_CalDay)
   CalButton(5).TipText = GetQuarterInfo(DatePart("q", dteDate), stdIcon)
   imgButton.Item(5).Picture = stdIcon
   CalButton(6).TipText = GetSeasonInfo(GetSeason, stdIcon)
   imgButton.Item(6).Picture = stdIcon
   CalButton(7).TipText = GetMoonPhaseInfo(dteDate, stdIcon)
   imgButton.Item(7).Picture = stdIcon
   CalButton(8).TipText = GetZodiacInfo(GetZodiacSign, stdIcon)
   imgButton.Item(8).Picture = stdIcon
   Set stdIcon = Nothing
   Separator(0) = SEPARATOR_DEFAULT
   Separator(1) = SEPARATOR_DEFAULT
   Separator(2) = ""
   intLength = 1 + (3 And m_LockInfoBar)
   
   For intCount = 5 To 8
      With CalButton(intCount).Rect
         If intCount < 8 Then
            Line (.Right - 1, .Top + intLength)-(.Right - 1, .Bottom - intLength), vb3DShadow
            Line (.Right, .Top + intLength)-(.Right, .Bottom - intLength), vb3DHighlight
         End If
         
         DrawIconEx hDC, .Left + (.Right - .Left - 32) \ 2, .Top + 5, imgButton.Item(intCount).Picture.Handle, 32, 32, 0, 0, DI_NORMAL
      End With
   Next 'intCount

End Sub

' routine to draw the calendar month year label
Private Sub DrawLabel()

Dim blnBold  As Boolean
Dim intSize  As Integer
Dim lngColor As Long
Dim sngCheck As Single
Dim strDate  As String

   If m_ShowNavigationBar = Off Then Exit Sub
   
   With CalLabel.Rect
      intSize = (.Right - .Left)
      lngColor = ForeColor
      FontSize = SizeFont + 2
      blnBold = FontBold
      FontBold = m_LabelFontBold
      ForeColor = m_LabelForeColor
      strDate = GetTextPart(LanguageText(CalLanguage).MonthNames, 9) & " " & m_CalYear
      
      If intSize And (TextWidth(strDate) > intSize - 4) Then
         Do While TextWidth(strDate) > intSize - 4
            If sngCheck = FontSize Then Exit Do
            
            sngCheck = FontSize
            FontSize = FontSize - 1
         Loop
      End If
      
      strDate = SetDateFormat(GetTextPart(LanguageText(CalLanguage).MonthNames, m_CalMonth), CStr(m_CalYear))
      CurrentX = .Left + (intSize - TextWidth(strDate)) \ 2
      CurrentY = .Top + (.Bottom - .Top - TextHeight("X")) \ 2 - 1
      Print strDate
      FontBold = blnBold
      FontSize = SizeFont
      ForeColor = lngColor
   End With

End Sub

' routine to draw one of five types of markers
Private Sub DrawMarkers(ByVal Index As Integer)

Dim intLeft   As Integer
Dim intTop    As Integer
Dim intWidth  As Integer
Dim intHeight As Integer

   intWidth = (SizeX - 9) / 5
   intHeight = SizeY / 10
   
   With CalCell(Index)
      intLeft = .X + 3
      intTop = .Y + SizeY - intHeight - 4
      
      If (.Mark And 1) = 1 Then Line (intLeft + intWidth * 4, intTop)-Step(intWidth, intHeight), MarkColor(0), BF
      If (.Mark And 2) = 2 Then Line (intLeft + intWidth * 3, intTop)-Step(intWidth, intHeight), MarkColor(1), BF
      If (.Mark And 4) = 4 Then Line (intLeft + intWidth * 2, intTop)-Step(intWidth, intHeight), MarkColor(2), BF
      If (.Mark And 8) = 8 Then Line (intLeft + intWidth * 1, intTop)-Step(intWidth, intHeight), MarkColor(3), BF
      If (.Mark And 16) = 16 Then Line (intLeft, intTop)-Step(intWidth, intHeight), MarkColor(4), BF
   End With

End Sub

' routine to draw the navigation of the calendar
Private Sub DrawNavigation()

'Const BDR_EDGED As Long = &H16

Dim intCount    As Integer
Dim intSize     As Integer
Dim lngBorder   As Long

   If m_ShowNavigationBar = Off Then Exit Sub
   
   For intCount = 0 To 4
      With CalButton(intCount).Rect
         If m_ButtonGradientStyle Then
            Call DrawGradient(CalButton(intCount).Rect, m_ButtonColor, m_ButtonGradientColor, m_ButtonGradientStyle)
            
         Else
            Line (.Left, .Top)-Step(.Right - .Left - 2, .Right - .Left - 2), m_ButtonColor, BF
         End If
         
         If m_Appearance = [3D] Then
            lngBorder = BDR_RAISED
            
         Else
            lngBorder = BDR_RAISEDINNER
         End If
         
         DrawEdge hDC, CalButton(intCount).Rect, lngBorder, BF_RECT
         
         If intCount = 0 Then
            DrawIconEx hDC, .Left + 5, .Top + 5, imgButton.Item(intCount - (m_Hemisphere And (intCount = 5))).Picture.Handle, 32, 32, 0, 0, DI_NORMAL
            
         Else
            With CalButton(intCount).Rect
               intSize = (.Right - .Left - 12) \ 2
               
               Call DrawArrow(intCount, .Left + intSize, .Top + intSize)
            End With
         End If
      End With
   Next 'intCount
   
   If m_Appearance = [3D] Then
      lngBorder = Choose(m_LabelBorderStyle + 1, BDR_SUNKEN, BDR_RAISED, BDR_EDGED)
      
   Else
      lngBorder = Choose(m_LabelBorderStyle + 1, BDR_SUNKENINNER, BDR_RAISEDINNER, BDR_EDGED)
   End If
   
   With CalLabel.Rect
      .Left = CalButton(2).Rect.Right
      .Top = CalButton(1).Rect.Top
      .Right = CalButton(3).Rect.Left
      .Bottom = CalButton(1).Rect.Bottom + (1 And (m_LabelBorderStyle = Edged))
      
      If m_LabelBackStyle = Opaque Then
         Line (.Left + 1, .Top + 1)-(.Right - 2, .Bottom - 2), m_LabelBackColor, BF
         
         Call DrawLabel
      End If
      
      DrawEdge hDC, CalLabel.Rect, lngBorder, BF_RECT
   End With
   
   Call SetToDayImage

End Sub

' for resizing the calendar
Private Sub Resize()

Dim intCount As Integer
Dim intLeft  As Integer
Dim intX     As Integer
Dim intY     As Integer

   If m_Locked Then Exit Sub
   
   picToDay.FontSize = 10 + ((3 And (Screen.TwipsPerPixelY > 12)) / 2)
   SizeX = (ScaleWidth - 2) / 8
   SizeY = (ScaleHeight - 1 - (5 And CBool(m_ShowInfoBar))) / (7 + (2 And CBool(m_ShowNavigationBar)) + (2 And CBool(m_ShowInfoBar)))
   intY = (SizeY * 2 - 42) \ 2
   
   If m_ShowNavigationBar Then
      For intCount = 0 To 4
         With CalButton(intCount).Rect
            If intCount < 3 Then
               .Left = 42 * intCount
               
            Else
               .Left = (ScaleWidth - 42 * Abs(intCount - 5) - 1)
            End If
            
            .Top = intY
            .Right = .Left + 42
            .Bottom = .Top + 42
         End With
      Next 'intCount
      
      If m_ShowNavigationBar = Small Then
         For intCount = 1 To 4
            With CalButton(intCount).Rect
               If intCount < 3 Then
                  .Left = 42 + intCount * 26
                  
               Else
                  .Left = ScaleWidth - ((5 - intCount) * 26) - 13
               End If
               
               .Top = .Top + 8
               .Right = .Left + 26
               .Bottom = .Bottom - 8
            End With
         Next 'intCount
         
         With CalButton(0).Rect
            .Left = .Left + 13
            .Right = .Left + 42
         End With
      End If
   End If
   
   If m_ShowInfoBar Then
      intX = ScaleWidth \ (4 + (1 And (m_ShowInfoBar = Small)))
      intY = ScaleHeight - intY - 44
      intLeft = (ScaleWidth - intX * 4) \ 2
      
      For intCount = 5 To 8
         With CalButton(intCount).Rect
            .Left = intLeft + intX * (intCount - 5)
            .Top = intY
            .Right = .Left + intX
            .Bottom = .Top + 44
         End With
      Next 'intCount
   End If

End Sub

' routine to set the button tooltiptext
Private Sub SetButtonTipText()

Dim intCount As Integer
Dim intMonth As Integer
Dim intYear  As Integer
Dim strText  As String

   For intCount = LeftButton To DownButton
      If intCount = LeftButton Then
         intMonth = m_CalMonth - 1
         intYear = m_CalYear - (1 And (intMonth = 0))
         
         If intMonth = 0 Then intMonth = 12
         
         If (m_CalMonth = 1) And (m_CalYear = 1583) Then
            strText = GetTextPart(LanguageText(CalLanguage).Miscellaneous, 3)
            intYear = 0
         End If
         
      ElseIf intCount = RightButton Then
         intMonth = m_CalMonth + 1
         intYear = m_CalYear + (1 And (intMonth = 13))
         
         If intMonth = 13 Then intMonth = 1
         
         If (m_CalMonth = 12) And (m_CalYear = 9999) Then
            strText = GetTextPart(LanguageText(CalLanguage).Miscellaneous, 4)
            intYear = 0
         End If
         
      ElseIf intCount = UpButton Then
         intYear = m_CalYear + (1 And (m_CalYear < 9999))
         intMonth = m_CalMonth
         
         If m_CalYear = 9999 Then
            strText = GetTextPart(LanguageText(CalLanguage).Miscellaneous, 4)
            intYear = 0
         End If
         
      ElseIf intCount = DownButton Then
         intYear = m_CalYear - (1 And (m_CalYear > 1583))
         intMonth = m_CalMonth
      
         If m_CalYear = 1583 Then
            strText = GetTextPart(LanguageText(CalLanguage).Miscellaneous, 3)
            intYear = 0
         End If
      End If
      
      If intYear Then
         strText = GetTextPart(LanguageText(CalLanguage).MonthNames, intMonth)
         
         If ((intCount = LeftButton) And (m_CalMonth = 1)) Or ((intCount = RightButton) And (m_CalMonth = 12)) Or ((intCount > RightButton) And (intYear > 0)) Then strText = strText & " " & intYear
      End If
      
      CalButton(intCount).TipText = strText
      strText = ""
   Next 'intCount

End Sub

' routine to set the current day marker
Private Sub SetCurrentDay()

Static intPrevDay  As Integer
Static intPrevWeek As Integer

Dim intSize        As Integer

   If Not m_SelectedDayMark Then Exit Sub
   If intPrevDay Then Call DrawCell(intPrevDay)
   If intPrevWeek Then Call DrawCell(intPrevWeek)
   
   CalCell(CurrentCell).Type = (Normal And (CurrentCell <= CalcDay(MonthDays))) + (OtherMonths And (CurrentCell > CalcDay(MonthDays)))
   
   Call DrawCell(CurrentCell)
   Call DrawMarkers(CurrentCell)
   
   If SetLastDay Or (SelectedDay > MonthDays) Then
      CurrentCell = CalcDay(MonthDays)
      
   Else
      CurrentCell = CalcDay(SelectedDay)
   End If
   
   CalCell(CurrentCell).Type = Selected
   
   Call DrawCell(CurrentCell)
   Call DrawMarkers(CurrentCell)
   
   If SizeY < 32 Then
      intSize = SizeY + 2
      
   Else
      intSize = 32
   End If
   
   DrawIconEx hDC, CalCell(CurrentCell).X + (SizeX - intSize) \ 2, CalCell(CurrentCell).Y + (SizeY - intSize) \ 2, imgToDay.Picture.Handle, intSize, intSize, 0, 0, DI_NORMAL
   m_CalDay = Val(CalCell(CurrentCell).Text)
   intPrevDay = CurrentCell Mod 8
   intPrevWeek = (CurrentCell \ 8) * 8
   
   Call DrawCell(intPrevDay, , True)
   Call DrawCell(intPrevWeek, , True)
   
   If m_ShowDayOfYear Then
      CalCell(0).Text = DayOfYear
      CalCell(0).TipText = GetTextPart(LanguageText(CalLanguage).Miscellaneous, 1) & " " & m_CalYear
      CalCell(0).TipText = Split(CalCell(0).TipText, "#")(0) & DayOfYear & Split(CalCell(0).TipText, "#", 2)(1)
      
      Call DrawCell(0)
   End If

End Sub

' switch to Single selection mode
' if more than one day is selected at the present
' deselect all but current calendar day
Private Sub SetSingleSelect()

Dim intCell As Integer
   
   For intCell = OffsetCell + 1 To 55
      If Val(CalCell(intCell).Text) = m_CalDay Then
         CalCell(intCell).Type = Selected
         CurrentCell = intCell
         
      ElseIf intCell Mod 8 Then
         CalCell(intCell).Type = Normal
      End If
   Next 'intCell
   
   Call DrawCalendar

End Sub

' for setting day value in today button
Private Sub SetToDayImage()

Dim intCount As Integer
Dim strDay   As String

   With CalButton(0).Rect
      DrawIconEx hDC, .Left + 5, .Top + 5, imgButton.Item(0).Picture.Handle, 32, 32, 0, 0, DI_NORMAL
   End With
   
   With picToDay
      strDay = Right(" " & Day(Date), 2)
      .Cls
      .CurrentX = (.ScaleWidth - .TextWidth(strDay)) \ (Len(strDay) * (10 - (7 And (Screen.TwipsPerPixelX > 12))))
      
      For intCount = 1 To 2
         picToDay.Print Mid(strDay, intCount, 1);
         .CurrentX = .CurrentX - 1
      Next 'intCount
   End With
   
   With CalButton(0).Rect
      For intCount = 2 To 16
         StretchBlt hDC, .Left + 10 + intCount * 0.53, .Top + 14 + intCount, 17, 1, picToDay.hDC, 0, intCount, 18, 1, vbSrcCopy
      Next 'intCount
   End With
   
   strDay = Format(Now, GetDateFormat(True))
   CalButton(0).TipText = GetTextPart(LanguageText(CalLanguage).Miscellaneous, 5) & " " & UCase(Left(strDay, 1)) & Mid(strDay, 2)

End Sub

' set the tooltiptext on/off
Private Sub SetToolTipText(ByVal TipText As String)

   If m_ShowToolTipText And Len(TipText) Then
      Extender.ToolTipText = " " & TipText & " "
      
   Else
      Extender.ToolTipText = ""
   End If

End Sub

' multi langual weekday names
Private Sub SetWeekDayHeaderText()

Dim intCell As Integer
Dim intDay  As Integer

   intDay = m_FirstWeekDay
   
   For intCell = 1 To 7
      With CalCell(intCell)
         .Type = Header
         .TipText = GetTextPart(LanguageText(CalLanguage).DayNames, intDay)
         .Text = Left(.TipText, m_WeekDayViewChar)
         intDay = intDay + 1 - (7 And (intDay = 7))
      End With
   Next 'intCell

End Sub

' when the daymarker is set on the current day and the systemdate
' will be changed in tomorrow, the daymarker also jumps to tomorrow
Private Sub tmrIsOtherDay_Timer()

   If ToDay = Date Then Exit Sub
   
   ToDay = Date
   
   If ToDay = DateAdd("d", 1, DateSerial(m_CalYear, m_CalMonth, Val(CalCell(CurrentCell).Text))) Then
      m_CalDay = Day(Date)
      m_CalMonth = Month(Date)
      SelectedDay = m_CalDay
      CalYear = Year(Date)
      RaiseEvent ButtonClick(ToDayButton)
      
   Else
      m_CalDay = Day(Date)
      m_CalMonth = Month(Date)
      m_CalYear = Year(Date)
      
      Call DrawCalendar
   End If

End Sub

Private Sub UserControl_DblClick()

   IsClicked = True
   
   Call UserControl_MouseUp(MouseButton, 0, MouseX, MouseY)
   
   If CalCell(MouseCell).Type <> Selected Then Exit Sub
   
   RaiseEvent DayDblClick(Val(CalCell(MouseCell).Text))

End Sub

Private Sub UserControl_Initialize()

   LanguageText(Dutch).DayNames = "Zondag,Maandag,Dinsdag,Woensdag,Donderdag,Vrijdag,Zaterdag"
   LanguageText(Dutch).Miscellaneous = "Dag # in: ,Week: ,Begin van kalender!,Einde van kalender!,Vandaag:"
   LanguageText(Dutch).MonthNames = "Januari,Februari,Maart,April,Mei,Juni,Juli,Augustus,September,Oktober,November,December"
   LanguageText(Dutch).QuarterNames = "Eerste kwartaal,Tweede kwartaal,Derde kwartaal,Vierde kwartaal"
   LanguageText(Dutch).SeasonNames = "Lente,Zomer,Herfst,Winter"
   LanguageText(Dutch).MoonPhaseNames = "Nieuwe maan,Eerste kwartier,Volle maan,Laatste kwartier"
   LanguageText(Dutch).MoonPhaseText = "# dag voor,# dag na,# dagen voor,# dagen na,het"
   LanguageText(Dutch).ZodiacNames = "Ram,Stier,Tweelingen,Kreeft,Leeuw,Maagd,Weegschaal,Schorpioen,Boogschutter,Steenbok,Waterman,Vissen"
   LanguageText(English).DayNames = "Sunday,Monday,Tuesday,Wednesday,Thursday,Friday,Saturday"
   LanguageText(English).Miscellaneous = "Day # in: ,Week: ,Begin of calendar!,End of calendar!,Today:"
   LanguageText(English).MonthNames = "January,February,March,April,May,June,July,August,September,October,November,December"
   LanguageText(English).QuarterNames = "First quarter,Second quarter,Thirth quarter,Fourth quarter"
   LanguageText(English).SeasonNames = "Spring,Summer,Autumn,Winter"
   LanguageText(English).MoonPhaseNames = "New moon,First quarter,Full moon,Last quarter"
   LanguageText(English).MoonPhaseText = "# day before,# day after,# days before,# days after,the"
   LanguageText(English).ZodiacNames = "Ram,Bull,Twins,Crab,Lion,Virgin,Scales,Scorpion,Archer,Goat,Water Carrier,Fishes"
   LanguageText(French).DayNames = "Dimanche,Lundì,Mardì,Mercredì,Jeudì,Vendredì,Samedi"
   LanguageText(French).Miscellaneous = "Jour # dans: ,Semaine: ,4Initialiser le calendrier!,Arrêter le calendrier!,Aujourd'hui:"
   LanguageText(French).MonthNames = "Janvier,Février,Mars,Avril,Mai,Juin,Juillet,Août,Septembre,Octobre,Novembre,Décembre"
   LanguageText(French).QuarterNames = "Premier trimestre,Deuxième trimestre,Troisième trimestre,Quatrième trimestre"
   LanguageText(French).SeasonNames = "Printemps,Eté,Autonne,Hiver"
   LanguageText(French).MoonPhaseNames = "Lune nouveau,Premier trimestre,Pleine lune,Dernier trimestre"
   LanguageText(French).MoonPhaseText = "# jour à,# jour après,# journée à,# journée après,la"
   LanguageText(French).ZodiacNames = "Bélier,Taureau,Gemeaux,Cancer,Lion,Vierge,Balance,Scorpion,Sagittaire,Capricorne,Verseau,Poissons"
   LanguageText(Italian).DayNames = "Domenica,Lunedì,Martedì,Mercoledì,Giovedì,Venerdì,Sabato"
   LanguageText(Italian).Miscellaneous = "Giorno # in: ,Settimana: ,Inizio del calendario!,Fine del calendario!,Oggi:"
   LanguageText(Italian).MonthNames = "Gennaio,Febbraio,Marzo,Aprile,Maggio,Giugno,Luglio,Agosto,Settembre,Ottobre,Novembre,Dicembre"
   LanguageText(Italian).QuarterNames = "Primo quarto,Secondo quarto,Terzo quarto,Quarto quarto"
   LanguageText(Italian).SeasonNames = "Primavera,Estate,Autunno,Inverno"
   LanguageText(Italian).MoonPhaseNames = "Luna nuevo,Primo quarto,Plenilunio,Ultimo quarto"
   LanguageText(Italian).MoonPhaseText = "# dia a,# dia dietro,# dias a,# dias dietro,el"
   LanguageText(Italian).ZodiacNames = "Ariete,Toro,Gemelli,Cancro,Leone,Vergine,Bilancia,Scorpione,Sagitario,Capricorno,Acquario,Pesci"
   LanguageText(Spanish).DayNames = "Domingo,Lunes,Martes,Miércoles,Jueves,Viernes,Sábado"
   LanguageText(Spanish).Miscellaneous = "Día # de: ,Semana: ,¡Inicio del calendario!,¡Fin del calendario!,Hoy:"
   LanguageText(Spanish).MonthNames = "Enero,Febrero,Marzo,Abril,Mayo,Junio,Julio,Agosto,Septiembre,Octubre,Noviembre,Diciembre"
   LanguageText(Spanish).QuarterNames = "Primero trimestre,Segundo trimestre,Tercero trimestre,Cuarto trimestre"
   LanguageText(Spanish).SeasonNames = "Primavera,Verano,Otoño,Invierno"
   LanguageText(Spanish).MoonPhaseNames = "Luna nueva,Primero cuarto,Luna lena,Último cuarto"
   LanguageText(Spanish).MoonPhaseText = "# dia antes,# dia después,# dias antes,# dias después,de"
   LanguageText(Spanish).ZodiacNames = "Aries,Tauro,Géminis,Cáncer,Leo,Virgo,Libra,Escorpio,Sagitario,Capricornio,Acuario,Piscis"
   SizeFont = FontSize
   ToDay = Date
   SelectedDay = Day(Date)
   Separator(0) = SEPARATOR_DEFAULT
   Separator(1) = SEPARATOR_DEFAULT
   
   ReDim CalButton(9) As CalButtons
   
   Call SetMarkColors
   Call Resize
   Call CalcCalendar

End Sub

' initialize Properties for User Control
Private Sub UserControl_InitProperties()

   m_ArrowColor = vbButtonText
   UserControl.BackColor = Ambient.BackColor
   m_ButtonColor = vbButtonFace
   m_ButtonGradientColor = vbButtonFace
   m_CalYear = Year(Now)
   m_CalMonth = Month(Now)
   m_CalDay = Day(Now)
   m_CellDayOfYearBackColor = vbButtonFace
   m_CellDayOfYearForeColor = vbButtonText
   m_CellDayOfYearStyle = BlueSelect
   m_CellDaysBackColor = vbButtonFace
   m_CellForeColorSunday = vbButtonText
   m_CellForeColorMonday = vbButtonText
   m_CellForeColorTuesday = vbButtonText
   m_CellForeColorWednesday = vbButtonText
   m_CellForeColorThursday = vbButtonText
   m_CellForeColorFriday = vbButtonText
   m_CellForeColorSaturday = vbButtonText
   m_CellHeaderBackColor = vbButtonFace
   m_WeekNumberForeColor = vbButtonText
   m_CellHeaderStyle = DarkGraySelect
   m_WeekDayViewChar = 2
   m_CellOtherMonthBackColor = vbButtonFace
   m_CellOtherMonthForeColor = vbButtonText
   m_CellSelectBackColor = vbButtonFace
   m_CellSelectForeColor = vbButtonText
   m_FirstWeekDay = vbMonday
   Set UserControl.Font = Ambient.Font
   SizeFont = FontSize
   m_FrameColor = vbBlue
   m_GradientColor = Ambient.BackColor
   m_GridColor = vbBlue
   m_LabelBackColor = vbButtonFace
   m_LabelBackStyle = Opaque
   m_LabelBorderStyle = Raised
   m_LabelForeColor = vbButtonText
   CalLanguage = GetSystemLanguage
   m_SelectedDayMark = True
   m_ShowDayOfYear = True
   m_ShowInfoBar = Large
   m_ShowNavigationBar = Large
   m_ShowToolTipText = True
   
   Call CalcCalendar

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim blnCancel As Boolean
Dim intCell   As Integer
Dim intChoose As Integer
Dim intMonth  As Integer

   intCell = (Y \ SizeY) * 8 + X \ SizeX - (16 And CBool(m_ShowNavigationBar))
   MouseButton = Button
   IsChanged = False
   MouseX = X
   MouseY = Y
   
   If (intCell < 0) Or (intCell > UBound(CalCell)) Then
      If Button = vbLeftButton Then
         intChoose = 8 - (4 And m_LockInfoBar)
         CalButtonID = CheckButton(CalButton, X, Y, intChoose)
         
         If CalButtonID > -1 Then
            Call DrawButton(CalButton, CalButtonID, True, False)
            
            UserControl.Refresh
            MouseIn = True
            MouseOut = False
            IsClicked = True
         End If
      End If
      
      Exit Sub
      
   Else
      CalButtonID = -1
   End If
   
   MouseCell = intCell
   
   If CalCell(intCell).Type = DayOfTheYear Then Exit Sub
   
   If CalCell(intCell).Type = OtherMonths Then
      If m_CellOtherMonthView Then
         intChoose = Val(CalCell(intCell).Text)
         intMonth = m_CalMonth
         
         If intChoose > 15 Then
            Call CheckForMonth(intMonth)
            
         Else
            Call CheckForMonth(intMonth, , True)
         End If
         
         RaiseEvent DayClick(Button, Shift, intChoose, intMonth, blnCancel)
         
         If blnCancel Then Exit Sub  ' user code canceled
         
         m_CalDay = intChoose
         
         If m_CalDay > 15 Then
            CalButtonID = LeftButton
            
         Else
            CalButtonID = RightButton
         End If
         
         OtherMonthSelected = True
         SelectedDay = m_CalDay
         MouseIn = True
         IsClicked = True
         SetLastDay = False
      End If
      
      Exit Sub
   End If
   
   If (CalCell(intCell).Type <> Header) And (CalCell(intCell).Type <> WeekNumber) Then
      RaiseEvent DayClick(Button, Shift, Val(CalCell(intCell).Text), m_CalMonth, blnCancel)
      
      If blnCancel Then Exit Sub  ' user code canceled
   End If
   
   If m_SelectionType = SingleCell Then
      If (CalCell(intCell).Type = Header) Or (CalCell(intCell).Type = WeekNumber) Or (intCell = CurrentCell) Then
         If intCell = CurrentCell Then RaiseEvent DayClick(Button, Shift, Val(CalCell(intCell).Text), m_CalMonth, blnCancel)
         
         Exit Sub
      End If
      
      If Button <> vbLeftButton Then Exit Sub
      
      CalCell(CurrentCell).Type = Normal
      
      Call DrawCell(CurrentCell)
      Call DrawMarkers(CurrentCell)
      
      CalCell(intCell).Type = Selected
      m_CalDay = Val(CalCell(intCell).Text)
      CurrentCell = intCell
      SelectedDay = m_CalDay
      SetLastDay = (m_CalDay = MonthDays)
      IsChanged = True
      
   ' MultiCell
   Else
      If Button <> vbLeftButton Then Exit Sub
      
      Select Case CalCell(intCell).Type
         Case Header, WeekNumber
            Call DaySelectAll(intCell)
            
         Case Selected
            CalCell(intCell).Type = Normal
            m_CalDay = Val(CalCell(intCell).Text)
            
         Case Normal
            CalCell(intCell).Type = Selected
            m_CalDay = Val(CalCell(intCell).Text)
      End Select
   End If
   
   Call DrawCell(intCell)
   Call DrawMarkers(intCell)
   
   RaiseEvent SelChanged

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Static intPrevCell As Integer

Dim intCell        As Integer
Dim intTemp        As Integer
Dim strTipText     As String

   If intPrevCell = 0 Then intPrevCell = CurrentCell
   
   Call DrawCell(intPrevCell)
   Call DrawMarkers(intPrevCell)
   
   intCell = (Y \ SizeY) * 8 + X \ SizeX - (16 And CBool(m_ShowNavigationBar))
   
   If (intCell < 0) Or (intCell > UBound(CalCell)) Then
      intTemp = 8 - (4 And m_LockInfoBar)
      MouseIn = (CheckButton(CalButton, X, Y, intTemp, CalButtonID) > -1)
      intCell = CheckButton(CalButton, X, Y, 8)
      
      If IsClicked Then Call DrawButton(CalButton, CalButtonID, MouseIn, Not MouseOut)
      If intCell > -1 Then strTipText = CalButton(intCell).TipText
      
      MouseOut = False
      
      If Not MouseIn Then
         MouseOut = True
         
         Call DrawNavigation
         Call DrawInfoBar
      End If
      
   Else
      strTipText = CalCell(intCell).TipText
      
      If (intCell Mod 8) And ((intCell > 8) And (intCell <= UBound(CalCell))) Then
         Call DrawCell(intCell, True)
         Call DrawMarkers(intCell)
         
         intPrevCell = intCell
         
         If CalCell(intCell).Type = OtherMonths Then
            intTemp = m_CalMonth + (1 And (intCell > 19)) - (1 And (intCell < 20))
            intTemp = intTemp + (12 And (intTemp < 1)) - (12 And (intTemp > 12))
            
            If m_CellOtherMonthView Then
               strTipText = GetMonthName(intTemp)
               
            Else
               strTipText = ""
            End If
            
            If (intTemp = m_CalMonth - 11) Or (intTemp = m_CalMonth + 11) Then strTipText = strTipText & " " & m_CalYear + (1 And (intCell > 19)) - (1 And (intCell < 20))
            
         ElseIf CalCell(intCell).Mark Then
            For intTemp = 0 To 4
               If Point(X, Y) = MarkColor(intTemp) Then
                  strTipText = CalCell(intCell).MarkTipText(intTemp)
                  Exit For
               End If
            Next 'intTemp
         End If
      End If
   End If
   
   If Not OtherMonthSelected Then Call SetCurrentDay
   
   Call SetToolTipText(strTipText)
   Call DrawGrid
   Call DrawFrame

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If IsClicked Then
      IsClicked = False
      OtherMonthSelected = False
      
      If MouseIn And (Button = vbLeftButton) Then
         If CalButtonID = ToDayButton Then
            m_CalDay = Day(Now)
            m_CalMonth = Month(Now)
            SelectedDay = m_CalDay
            SetLastDay = False
            CalYear = Year(Now)
            
         ElseIf CalButtonID = LeftButton Then
            Call CheckForMonth(m_CalMonth, m_CalYear)
            
            CalMonth = m_CalMonth
            CalYear = m_CalYear
            
         ElseIf CalButtonID = RightButton Then
            Call CheckForMonth(m_CalMonth, m_CalYear, True)
            
            CalMonth = m_CalMonth
            CalYear = m_CalYear
            
         ElseIf CalButtonID = UpButton Then
            CalYear = m_CalYear + 1
            
         ElseIf CalButtonID = DownButton Then
            CalYear = m_CalYear - 1
         End If
         
         Call SetButtonTipText
         
         If CalButtonID > -1 Then
            If CalButtonID > DownButton Then
               Call CalcCalendar
               
               DoEvents
            End If
            
            RaiseEvent ButtonClick(CalButtonID)
         End If
      End If
      
      RaiseEvent DateChanged(CalButtonID)
      
   ElseIf IsChanged Then
      Call DrawInfoBar
      
      RaiseEvent DateChanged(CalButtonID)
   End If

End Sub

' load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

   With PropBag
      m_Locked = .ReadProperty("Locked", False)
      m_Appearance = .ReadProperty("Appearance", [3D])
      m_ArrowColor = .ReadProperty("ArrowColor", vbButtonText)
      UserControl.BackColor = .ReadProperty("BackColor", Ambient.BackColor)
      UserControl.BorderStyle = .ReadProperty("BorderStyle", vbBSNone)
      m_ButtonColor = .ReadProperty("ButtonColor", vbButtonFace)
      m_ButtonGradientColor = .ReadProperty("ButtonGradientColor", vbButtonFace)
      m_ButtonGradientStyle = .ReadProperty("ButtonGradientStyle", NoGradient)
      m_CalDay = .ReadProperty("CalDay", Day(Now))
      m_CalMonth = .ReadProperty("CalMonth", Month(Now))
      m_CalYear = .ReadProperty("CalYear", Year(Now))
      m_CellDaysBackColor = .ReadProperty("CellDaysBackColor", vbButtonFace)
      m_CellDaysStyle = .ReadProperty("CellDaysStyle", EmptyCell)
      m_CellDayOfYearBackColor = .ReadProperty("CellDayOfYearBackColor", vbButtonFace)
      m_CellDayOfYearForeColor = .ReadProperty("CellDayOfYearForeColor", vbButtonText)
      m_CellDayOfYearStyle = .ReadProperty("CellDayOfYearStyle", BlueSelect)
      m_CellForeColorSunday = .ReadProperty("CellForeColorSunday", vbButtonText)
      m_CellForeColorMonday = .ReadProperty("CellForeColorMonday", vbButtonText)
      m_CellForeColorTuesday = .ReadProperty("CellForeColorTuesday", vbButtonText)
      m_CellForeColorWednesday = .ReadProperty("CellForeColorWednesday", vbButtonText)
      m_CellForeColorThursday = .ReadProperty("CellForeColorThursday", vbButtonText)
      m_CellForeColorFriday = .ReadProperty("CellForeColorFriday", vbButtonText)
      m_CellForeColorSaturday = .ReadProperty("CellForeColorSaturday", vbButtonText)
      m_CellHeaderBackColor = .ReadProperty("CellHeaderBackColor", vbButtonFace)
      m_CellHeaderStyle = .ReadProperty("CellHeaderStyle", DarkGraySelect)
      m_CellOtherMonthBackColor = .ReadProperty("CellOtherMonthBackColor", vbButtonFace)
      m_CellOtherMonthForeColor = .ReadProperty("CellOtherMonthForeColor", vbButtonText)
      m_CellOtherMonthStyle = .ReadProperty("CellOtherMonthStyle", EmptyCell)
      m_CellOtherMonthView = .ReadProperty("CellOtherMonthView", False)
      m_CellSelectBackColor = .ReadProperty("CellSelectBackColor", vbButtonFace)
      m_CellSelectForeColor = .ReadProperty("CellSelectForeColor", vbButtonText)
      m_CellSelectHeaderForeColor = .ReadProperty("CellSelectHeaderForeColor", vbBlue)
      m_CellSelectStyle = .ReadProperty("CellSelectStyle", EmptyCell)
      m_DateFormat = .ReadProperty("DateFormat", [dd-mm-yyyy])
      UserControl.Enabled = .ReadProperty("Enabled", True)
      m_FirstWeekDay = .ReadProperty("FirstWeekDay", vbMonday)
      Set UserControl.Font = .ReadProperty("Font", Ambient.Font)
      SizeFont = FontSize
      m_FrameStyle = .ReadProperty("FrameStyle", Off)
      m_FrameColor = .ReadProperty("FrameColor", vbBlue)
      m_GradientColor = .ReadProperty("GradientColor", Ambient.BackColor)
      m_GradientStyle = .ReadProperty("GradientStyle", NoGradient)
      m_GridColor = .ReadProperty("GridColor", vbBlue)
      m_GridStyle = .ReadProperty("GridStyle", NoGrid)
      m_LabelBackColor = .ReadProperty("LabelBackColor", vbButtonFace)
      m_LabelBackStyle = .ReadProperty("LabelBackStyle", Opaque)
      m_LabelBorderStyle = .ReadProperty("LabelBorderStyle", Raised)
      m_LabelFontBold = .ReadProperty("LabelFontBold", False)
      m_LabelForeColor = .ReadProperty("LabelForeColor", vbButtonText)
      m_Language = .ReadProperty("Language", System)
      
      If m_Language = System Then
         CalLanguage = GetSystemLanguage
         
      Else
         CalLanguage = m_Language
      End If
      
      m_Hemisphere = .ReadProperty("Hemisphere", North)
      m_LockInfoBar = .ReadProperty("LockInfoBar", False)
      Set m_Picture = .ReadProperty("Picture", Nothing)
      m_SelectedDayMark = .ReadProperty("SelectedDayMark", True)
      m_SelectionType = .ReadProperty("SelectionType", SingleCell)
      m_ShowDayOfYear = .ReadProperty("ShowDayOfYear", True)
      m_ShowInfoBar = .ReadProperty("ShowInfoBar", Large)
      m_ShowNavigationBar = .ReadProperty("ShowNavigationBar", Large)
      m_ShowToolTipText = .ReadProperty("ShowToolTipText", True)
      m_WeekDayViewChar = .ReadProperty("WeekDayViewChar", Dd)
      m_WeekNumberForeColor = .ReadProperty("WeekNumberForeColor", vbButtonText)
      tmrIsOtherDay.Enabled = Ambient.UserMode
      
      Call Resize
      Call CalcCalendar
   End With

End Sub

Private Sub UserControl_Resize()

Static blnBusy As Boolean

   If blnBusy Then Exit Sub
   
   blnBusy = True
   
   If m_ShowNavigationBar + m_ShowInfoBar = Off Then
      If Height < 2052 Then Height = 2052
      If Width < 3012 Then Width = 3012
      
   ElseIf (m_ShowNavigationBar = Off) Or (m_ShowInfoBar = Off) Then
      If Height < 2652 Then Height = 2652
      If Width < 3804 Then Width = 3804
      
   Else
      If Height < 3252 Then Height = 3252
      If Width < 3804 Then Width = 3804
   End If
   
   CalendarRect.Bottom = ScaleHeight
   CalendarRect.Right = ScaleWidth
   picCalCell.Height = ScaleHeight
   picCalCell.Width = ScaleWidth
   blnBusy = False
   
   Call Resize
   Call CalcCalendar

End Sub

Private Sub UserControl_Terminate()

   tmrIsOtherDay.Enabled = False
   Set m_Picture = Nothing
   Erase MarkColor, Separator, CalButton

End Sub

' write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   With PropBag
      .WriteProperty "Locked", m_Locked, False
      .WriteProperty "Appearance", m_Appearance, [3D]
      .WriteProperty "ArrowColor", m_ArrowColor, vbButtonText
      .WriteProperty "BackColor", UserControl.BackColor, Ambient.BackColor
      .WriteProperty "BorderStyle", UserControl.BorderStyle, 0
      .WriteProperty "ButtonColor", m_ButtonColor, vbButtonFace
      .WriteProperty "ButtonGradientColor", m_ButtonGradientColor, vbButtonFace
      .WriteProperty "ButtonGradientStyle", m_ButtonGradientStyle, NoGradient
      .WriteProperty "CalDay", m_CalDay, Day(Now)
      .WriteProperty "CalMonth", m_CalMonth, Month(Now)
      .WriteProperty "CalYear", m_CalYear, Year(Now)
      .WriteProperty "CellDayOfYearBackColor", m_CellDayOfYearBackColor, vbButtonFace
      .WriteProperty "CellDayOfYearForeColor", m_CellDayOfYearForeColor, vbButtonText
      .WriteProperty "CellDayOfYearStyle", m_CellDayOfYearStyle, BlueSelect
      .WriteProperty "CellDaysBackColor", m_CellDaysBackColor, vbButtonFace
      .WriteProperty "CellDaysStyle", m_CellDaysStyle, EmptyCell
      .WriteProperty "CellForeColorSunday", m_CellForeColorSunday, vbButtonText
      .WriteProperty "CellForeColorMonday", m_CellForeColorMonday, vbButtonText
      .WriteProperty "CellForeColorTuesday", m_CellForeColorTuesday, vbButtonText
      .WriteProperty "CellForeColorWednesday", m_CellForeColorWednesday, vbButtonText
      .WriteProperty "CellForeColorThursday", m_CellForeColorThursday, vbButtonText
      .WriteProperty "CellForeColorFriday", m_CellForeColorFriday, vbButtonText
      .WriteProperty "CellForeColorSaturday", m_CellForeColorSaturday, vbButtonText
      .WriteProperty "CellHeaderBackColor", m_CellHeaderBackColor, vbButtonFace
      .WriteProperty "CellHeaderStyle", m_CellHeaderStyle, DarkGraySelect
      .WriteProperty "CellOtherMonthBackColor", m_CellOtherMonthBackColor, vbButtonFace
      .WriteProperty "CellOtherMonthForeColor", m_CellOtherMonthForeColor, vbButtonText
      .WriteProperty "CellOtherMonthStyle", m_CellOtherMonthStyle, EmptyCell
      .WriteProperty "CellOtherMonthView", m_CellOtherMonthView, False
      .WriteProperty "CellSelectBackColor", m_CellSelectBackColor, vbButtonFace
      .WriteProperty "CellSelectForeColor", m_CellSelectForeColor, vbButtonText
      .WriteProperty "CellSelectHeaderForeColor", m_CellSelectHeaderForeColor, vbBlue
      .WriteProperty "CellSelectStyle", m_CellSelectStyle, EmptyCell
      .WriteProperty "DateFormat", m_DateFormat, [dd-mm-yyyy]
      .WriteProperty "Enabled", UserControl.Enabled, True
      .WriteProperty "FirstWeekDay", m_FirstWeekDay, vbMonday
      .WriteProperty "Font", UserControl.Font, Ambient.Font
      .WriteProperty "FrameStyle", m_FrameStyle, Off
      .WriteProperty "FrameColor", m_FrameColor, vbBlue
      .WriteProperty "GradientColor", m_GradientColor, Ambient.BackColor
      .WriteProperty "GradientStyle", m_GradientStyle, NoGradient
      .WriteProperty "GridColor", m_GridColor, Ambient.BackColor
      .WriteProperty "GridStyle", m_GridStyle, NoGrid
      .WriteProperty "LabelBackColor", m_LabelBackColor, vbButtonFace
      .WriteProperty "LabelBackStyle", m_LabelBackStyle, Opaque
      .WriteProperty "LabelBorderStyle", m_LabelBorderStyle, Raised
      .WriteProperty "LabelFontBold", m_LabelFontBold, False
      .WriteProperty "LabelForeColor", m_LabelForeColor, vbButtonText
      .WriteProperty "Language", m_Language, System
      .WriteProperty "Hemisphere", m_Hemisphere, North
      .WriteProperty "LockInfoBar", m_LockInfoBar, False
      .WriteProperty "Picture", m_Picture, Nothing
      .WriteProperty "SelectedDayMark", m_SelectedDayMark, True
      .WriteProperty "SelectionType", m_SelectionType, SingleCell
      .WriteProperty "ShowDayOfYear", m_ShowDayOfYear, True
      .WriteProperty "ShowInfoBar", m_ShowInfoBar, Large
      .WriteProperty "ShowNavigationBar", m_ShowNavigationBar, Large
      .WriteProperty "ShowToolTipText", m_ShowToolTipText, True
      .WriteProperty "WeekDayViewChar", m_WeekDayViewChar, Dd
      .WriteProperty "WeekNumberForeColor", m_WeekNumberForeColor, vbButtonText
   End With

End Sub
