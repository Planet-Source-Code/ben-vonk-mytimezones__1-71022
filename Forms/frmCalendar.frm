VERSION 5.00
Begin VB.Form frmCalendar 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   6060
   ClientLeft      =   0
   ClientTop       =   96
   ClientWidth     =   9288
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   505
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   774
   ShowInTaskbar   =   0   'False
   Begin MyTimeZones.ThemedComboBox tcbSkinner 
      Left            =   8280
      Top             =   2040
      _ExtentX        =   445
      _ExtentY        =   423
      BorderColorStyle=   1
      ComboBoxBorderColor=   11904913
      DriveListBoxBorderColor=   0
   End
   Begin VB.PictureBox picHide 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   492
      Left            =   6240
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   91
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5400
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.ListBox lstWeekDays 
      Appearance      =   0  'Flat
      BackColor       =   &H00F7EFE2&
      ForeColor       =   &H00D94600&
      Height          =   240
      ItemData        =   "frmCalendar.frx":0000
      Left            =   4200
      List            =   "frmCalendar.frx":0002
      TabIndex        =   14
      Top             =   5520
      Width           =   1452
   End
   Begin MyTimeZones.FlatButton flbDate 
      Height          =   240
      Left            =   6360
      TabIndex        =   5
      Top             =   5040
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
      BackStyle       =   0
      IconX           =   2
      IconY           =   2
      OnlyIconClick   =   -1  'True
   End
   Begin VB.Timer tmrClock 
      Enabled         =   0   'False
      Interval        =   950
      Left            =   8280
      Top             =   2400
   End
   Begin VB.PictureBox picClock 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D94600&
      Height          =   324
      Index           =   0
      Left            =   240
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   276
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox picClock 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D94600&
      Height          =   324
      Index           =   1
      Left            =   6780
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   276
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox picCalendarImage 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1092
      Left            =   8280
      ScaleHeight     =   91
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.Timer tmrTimeToGo 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   8760
      Top             =   2400
   End
   Begin VB.PictureBox picTimeToGo 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4020
      Left            =   504
      ScaleHeight     =   335
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   601
      TabIndex        =   6
      Top             =   1020
      Visible         =   0   'False
      Width           =   7212
      Begin VB.TextBox txtDaysToDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00F7EFE2&
         BorderStyle     =   0  'None
         ForeColor       =   &H00D94600&
         Height          =   228
         Left            =   204
         MaxLength       =   8
         TabIndex        =   9
         Top             =   2976
         Width           =   2340
      End
      Begin VB.ComboBox cmbTimeToGoShowType 
         BackColor       =   &H00F7EFE2&
         ForeColor       =   &H00D94600&
         Height          =   312
         ItemData        =   "frmCalendar.frx":0004
         Left            =   1200
         List            =   "frmCalendar.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   3540
         Width           =   1932
      End
      Begin VB.ComboBox cmbTimeToGo 
         BackColor       =   &H00F7EFE2&
         ForeColor       =   &H00D94600&
         Height          =   312
         ItemData        =   "frmCalendar.frx":0008
         Left            =   192
         List            =   "frmCalendar.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   516
         Width           =   2364
      End
      Begin MyTimeZones.Calendar calTimeToGo 
         Height          =   3372
         Left            =   2760
         TabIndex        =   11
         TabStop         =   0   'False
         Tag             =   "&H00B93B00&"
         Top             =   0
         Width           =   4392
         _ExtentX        =   7747
         _ExtentY        =   5948
         ArrowColor      =   14239232
         ButtonColor     =   13747380
         ButtonGradientColor=   16248802
         ButtonGradientStyle=   4
         CellDayOfYearForeColor=   8396672
         CellDayOfYearStyle=   0
         CellForeColorSunday=   255
         CellForeColorMonday=   8396672
         CellForeColorTuesday=   8396672
         CellForeColorWednesday=   8396672
         CellForeColorThursday=   8396672
         CellForeColorFriday=   8396672
         CellForeColorSaturday=   16711680
         CellHeaderStyle =   0
         CellOtherMonthForeColor=   13089191
         CellOtherMonthView=   -1  'True
         CellSelectForeColor=   12591040
         CellSelectHeaderForeColor=   12591040
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FrameStyle      =   2
         FrameColor      =   11904913
         GridColor       =   11904913
         GridStyle       =   2
         LabelBackColor  =   16248802
         LabelBorderStyle=   2
         LabelFontBold   =   -1  'True
         LabelForeColor  =   14239232
         LockInfoBar     =   -1  'True
         ShowInfoBar     =   0
         ShowNavigationBar=   1
         WeekDayViewChar =   2
         WeekNumberForeColor=   14239232
      End
      Begin MyTimeZones.LEDDisplay ledDisplay 
         Height          =   264
         Left            =   3288
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   3564
         Width           =   3720
         _ExtentX        =   6562
         _ExtentY        =   466
         BackColor       =   14866892
         BorderStyle     =   0
         DisplayColor    =   16248802
         ForeColor       =   14239232
         NoTextScrolling =   -1  'True
         Size            =   1
      End
      Begin VB.Label lblTimeToGo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00D94600&
         Height          =   276
         Index           =   0
         Left            =   180
         TabIndex        =   8
         Top             =   1500
         Width           =   2388
      End
   End
   Begin VB.ListBox lstSort 
      Height          =   264
      Left            =   8280
      Sorted          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3240
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.ListBox lstSpecialDays 
      Height          =   264
      Left            =   8280
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2880
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.PictureBox picList 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   4020
      Left            =   1080
      ScaleHeight     =   335
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   505
      TabIndex        =   3
      Top             =   1020
      Visible         =   0   'False
      Width           =   6060
      Begin MyTimeZones.ThemedScrollBar tsbVertical 
         Height          =   2892
         Left            =   5760
         TabIndex        =   19
         Top             =   600
         Width           =   252
         _ExtentX        =   445
         _ExtentY        =   5101
         MouseWheel      =   0   'False
      End
   End
   Begin MyTimeZones.Calendar calCalendar 
      Height          =   3972
      Left            =   1104
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1032
      Width           =   6012
      _ExtentX        =   10605
      _ExtentY        =   7006
      ArrowColor      =   14239232
      BackColor       =   16577767
      ButtonColor     =   13747380
      ButtonGradientColor=   16248802
      ButtonGradientStyle=   4
      CellDayOfYearForeColor=   8396672
      CellDayOfYearStyle=   0
      CellForeColorSunday=   255
      CellForeColorMonday=   8396672
      CellForeColorTuesday=   8396672
      CellForeColorWednesday=   8396672
      CellForeColorThursday=   8396672
      CellForeColorFriday=   8396672
      CellForeColorSaturday=   16711680
      CellHeaderStyle =   0
      CellOtherMonthForeColor=   13089191
      CellOtherMonthView=   -1  'True
      CellSelectForeColor=   12591040
      CellSelectHeaderForeColor=   12591040
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FrameStyle      =   1
      FrameColor      =   11904913
      GradientColor   =   14998478
      GradientStyle   =   3
      GridColor       =   11904913
      GridStyle       =   2
      LabelBackColor  =   16248802
      LabelBorderStyle=   2
      LabelFontBold   =   -1  'True
      LabelForeColor  =   14239232
      ShowInfoBar     =   1
      ShowNavigationBar=   1
      WeekDayViewChar =   2
      WeekNumberForeColor=   14239232
   End
   Begin MyTimeZones.CheckBox chkTimeToGo 
      Height          =   384
      Left            =   360
      TabIndex        =   13
      Top             =   5472
      Visible         =   0   'False
      Width           =   3612
      _ExtentX        =   6371
      _ExtentY        =   677
      BackStyle       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   8396672
      IconCheckedGrayed=   "frmCalendar.frx":000C
   End
   Begin MyTimeZones.CrystalButton cbtChoose 
      Height          =   504
      Index           =   0
      Left            =   7380
      TabIndex        =   0
      Top             =   5376
      Width           =   672
      _ExtentX        =   1185
      _ExtentY        =   889
      BackColor       =   13542759
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   0
      Shape           =   1
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00B5A791&
      Height          =   372
      Index           =   0
      Left            =   120
      Top             =   240
      Width           =   1332
   End
   Begin VB.Image imgImages 
      Height          =   192
      Index           =   7
      Left            =   8760
      Picture         =   "frmCalendar.frx":08E6
      Top             =   1560
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   2
      Left            =   8280
      Picture         =   "frmCalendar.frx":0E70
      Top             =   600
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   5
      Left            =   8760
      Picture         =   "frmCalendar.frx":173A
      Top             =   1080
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   6
      Left            =   8280
      Top             =   1560
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   3
      Left            =   8760
      Picture         =   "frmCalendar.frx":2404
      Top             =   600
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   4
      Left            =   8280
      Picture         =   "frmCalendar.frx":2CCE
      Top             =   1080
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   1
      Left            =   8760
      Picture         =   "frmCalendar.frx":3598
      Top             =   120
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   0
      Left            =   8280
      Top             =   120
      Visible         =   0   'False
      Width           =   384
   End
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Private Constants
Private Const HDR_HEIGHT     As Integer = 25
Private Const HDR_WIDTH_DATE As Integer = 135
Private Const HDR_WIDTH_DAY  As Integer = 115
Private Const HDR_WIDTH_YEAR As Integer = 42

' Public Variable
Public TimeToGoOnly          As Boolean

' Private Variables
Private HasTipText           As Boolean
Private IsCheckDaysToGo      As Boolean
Private IsLeftButton         As Boolean
Private LastMonthDay         As Boolean
Private ShowInfo             As Boolean
Private WithIcons            As Boolean
Private CalButtonID          As Integer
Private ClockY               As Integer
Private ExtraLineSpace       As Integer
Private ListCount            As Integer
Private ListIndex            As Integer
Private PageItems            As Integer
Private RowIndex             As Integer
Private Tabs(3)              As Integer
Private TimeToGoEdit         As Integer
Private MaxDaysUp            As Long
Private MaxDaysDown          As Long
Private Header(3)            As Rect
Private ItemIcon()           As StdPicture
Private DateItems()          As String
Private ListItems()          As String
Private SelectedDateItem     As String
Private Separator            As String
Private ZodiacTipText(11)    As String

Public Function CalcDateValue(ByRef IsValue As Integer, ByVal IsDate As Integer, ByVal IsPart As Integer) As Integer

Dim intValue As Integer

   intValue = IsDate \ IsPart
   IsValue = IsDate - (IsPart * intValue)
   CalcDateValue = intValue

End Function

Public Function GetBeginValue(ByVal IsYear As Integer) As Integer

Dim intValue(6) As Integer

   intValue(0) = CalcDateValue(intValue(3), IsYear, 100)
   intValue(6) = intValue(0) - CalcDateValue(intValue(4), intValue(0), 4)
   intValue(0) = intValue(0) + 8
   intValue(0) = intValue(0) - CalcDateValue(intValue(1), intValue(0), 25) - 7
   CalcDateValue intValue(2), IsYear, 19
   intValue(0) = 19 * intValue(2) + intValue(6) - CalcDateValue(intValue(1), intValue(0), 3) + 15
   CalcDateValue intValue(5), intValue(0), 30
   intValue(0) = intValue(3)
   intValue(0) = 2 * (intValue(4) + CalcDateValue(intValue(1), intValue(0), 4)) - intValue(5) - intValue(1) + 32
   CalcDateValue intValue(1), intValue(0), 7
   intValue(0) = intValue(2) + 11 * intValue(5) + 22 * intValue(1)
   intValue(6) = intValue(5) + intValue(1) + 3
   GetBeginValue = 47 + intValue(6) - 7 * CalcDateValue(intValue(1), intValue(0), 451)
   Erase intValue

End Function

Public Sub CreateTimeToGo()

   With picTimeToGo
      MousePointer = vbHourglass
      ShowInfo = True
      .PaintPicture Picture, 0, 0, .ScaleWidth, .ScaleHeight, .Left, .Top, .ScaleWidth, .ScaleHeight, vbSrcCopy
      picTimeToGo.Line (calTimeToGo.Left - 1, calTimeToGo.Top + 5)-(calTimeToGo.Left - 1, 279), &HD1C4B4
      picTimeToGo.Line (5, 93)-(calTimeToGo.Left, 93), &HB5A791
      picTimeToGo.Line (5, 155)-(calTimeToGo.Left, 155), &HB5A791
      picTimeToGo.Line (5, 217)-(calTimeToGo.Left, 217), &HB5A791
      picTimeToGo.Line (5, 279)-(calTimeToGo.Left, 279), &HB5A791
      
      Call SetCalendarImage(.hDC, calTimeToGo, picCalendarImage)
      Call SetBackColor(picTimeToGo, cmbTimeToGo)
      Call SetBackColor(picTimeToGo, lblTimeToGo.Item(0))
      Call SetBackColor(picTimeToGo, lblTimeToGo.Item(1))
      Call SetBackColor(picTimeToGo, txtDaysToDate)
      Call SetBackColor(picTimeToGo, cmbTimeToGoShowType)
      Call DrawText(picTimeToGo, AppText(211), 9, &H801F80, 0, cmbTimeToGo.Top - 22, cmbTimeToGo.Left - 1)
      Call DrawText(picTimeToGo, AppText(212), 9, &H801F80, 0, lblTimeToGo.Item(0).Top - 23, lblTimeToGo.Item(0).Left)
      Call DrawText(picTimeToGo, AppText(213), 9, &H801F80, 0, lblTimeToGo.Item(1).Top - 23, lblTimeToGo.Item(0).Left)
      Call DrawText(picTimeToGo, AppText(214), 9, &H801F80, 0, txtDaysToDate.Top - 22, lblTimeToGo.Item(0).Left)
      Call DrawText(picTimeToGo, AppText(215), 9, &H801F80, 0, cmbTimeToGoShowType.Top + 3, lblTimeToGo.Item(0).Left)
      
      If Len(TimeToGo(0)) Then lblTimeToGo.Item(0).Caption = Format(TimeToGo(0), "d mmmm yyyy")
      
      If Len(TimeToGo(1)) Then
         lblTimeToGo.Item(1).Caption = Format(TimeToGo(1), "d mmmm yyyy")
         cmbTimeToGo.ListIndex = TimeToGoShow
         
         Call SetDisplay
         Call lblTimeToGo_Click(1)
         
      Else
         cmbTimeToGo.ListIndex = 0
         
         Call ResetDateInput
      End If
      
      DoEvents
      calTimeToGo.LabelBackColor = .Point(calTimeToGo.Left, calTimeToGo.Top + 35)
      Cls
      
      With cbtChoose.Item(1)
         .Shape = ShapeRight
         .ToolTipText = GetToolTipText(AppText(5))
         .Visible = True
      End With
      
      With chkTimeToGo
         .Visible = True
         .Value = Abs(AppSettings(SET_AUTODELETETIMETOGO))
         .Caption = AppText(216)
         .AutoSize = True
         .BackStyle = Transparent
      End With
      
      Call ToggleControls(True)
      
      cbtChoose.Item(1).Picture = imgImages.Item(6).Picture
      .Visible = True
   End With

End Sub

Public Sub ResetSpecialDays(Optional ByVal Forced As Boolean)

Dim dteDate     As Date
Dim intCount    As Integer
Dim intDay      As Integer
Dim intValue(2) As Integer
Dim strSign     As String

   If NoSpecialDays Then Exit Sub
   If (SpecialDays(0) = "") Or (Not Forced And (Year(CalendarDate) = calCalendar.CalYear)) Then Exit Sub
   
   With calCalendar
      CalendarDate = DateSerial(.CalYear, .CalMonth, .CalDay)
   End With
   
   intValue(0) = GetBeginValue(Year(CalendarDate))
   lstSpecialDays.Clear
   lstSort.Clear
   
   For intCount = 0 To UBound(SpecialDays)
      strSign = Left(SpecialDays(intCount), 1)
      
      If strSign Like "[+--]" Then
         intValue(2) = Val(Split(SpecialDays(intCount), ",")(0))
         intValue(1) = intValue(0) + (intValue(2) And (strSign = "+"))
         
         If strSign = "+" Then intValue(2) = 1
         
         dteDate = DateSerial(Year(CalendarDate), CalcDateValue(intDay, intValue(1) + (1 And (intValue(1) > 91)), 31) + 2, intDay + intValue(2))
         
      ElseIf strSign Like "[0-1]" Then
         dteDate = DateSerial(Year(CalendarDate), Val(Left(SpecialDays(intCount), 2)), Val(Mid(SpecialDays(intCount), 3, 2)))
         
      ElseIf InStr("MTWFS", UCase(strSign)) Then
         dteDate = GetGivenMonthDay(Year(CalendarDate), RTrim(Split(SpecialDays(intCount), ",")(0)))
      End If
      
      With lstSort
         .AddItem Format(dteDate, "yyyymmdd") & "," & Split(SpecialDays(intCount), ",", 2)(1)
         .ItemData(.NewIndex) = intCount
      End With
   Next 'intCount
   
   Erase intValue
   
   With lstSort
      For intCount = 0 To .ListCount - 1
         strSign = Mid(.List(intCount), 7, 2) & "-" & Mid(.List(intCount), 5, 2) & "-" & Left(.List(intCount), 4)
         lstSpecialDays.AddItem calCalendar.GetWeekdayName(WeekDay(strSign)) & ", " & Format(strSign, "d mmmm") & Mid(.List(intCount), 9)
      Next 'intCount
   End With

End Sub

Public Sub SetSpecialDays(ByRef Calendar As Calendar)

Dim intCount     As Integer
Dim intListIndex As Integer
Dim intMarker    As Integer
Dim strBuffer    As String
Dim strDate      As String
Dim strDay       As String
Dim strTipText   As String

   With Calendar
      If .Locked Then Exit Sub
      
      strDate = .CalYear & Format(.CalMonth, "#00")
      intListIndex = GetListIndex(lstSort.hWnd, LB_FINDSTRING, strDate)
      
      If intListIndex = -1 Then Exit Sub
      
      For intListIndex = intListIndex To lstSort.ListCount - 1
         If Left(lstSort.List(intListIndex), 6) <> strDate Then Exit For
         
         strBuffer = lstSort.List(intListIndex)
         
         If strDay = Mid(strBuffer, 7, 2) Then
            intMarker = intMarker + (1 And (intMarker < 5))
            
         Else
            intMarker = 1
         End If
         
         strDay = Mid(strBuffer, 7, 2)
         strTipText = Trim(Split(strBuffer, ",", 2)(1))
         
         For intCount = 1 To 2
            strBuffer = lstSort.List(intListIndex + intCount * 5)
            
            If Left(strBuffer, 6) <> strDate & strDay Then Exit For
            
            strTipText = strTipText & ", " & Trim(Split(strBuffer, ",", 2)(1))
         Next 'intCount
         
         Call .DayMarking(strDay, intMarker, True, strTipText)
      Next 'intListIndex
      
      .Refresh
   End With

End Sub

Private Function CheckDaysToGo() As Boolean

Dim lngTotalDays As Long
Dim strPrompt    As String

   If IsCheckDaysToGo Then Exit Function
   
   lngTotalDays = Val(txtDaysToDate.Text)
   
   If (lngTotalDays < MaxDaysDown) Or (lngTotalDays > MaxDaysUp) Then
      If AppSettings(SET_ASKCONFIRM) Then
         IsCheckDaysToGo = True
         
         If lngTotalDays < MaxDaysDown Then
            strPrompt = Split(AppError(41), ",")(0) & " " & Format(DateSerial(1583, 1, 1), DefaultDateFormat) & "."
            lngTotalDays = MaxDaysDown
            
         Else
            strPrompt = Split(AppError(41), ",")(1) & " " & Format(DateSerial(9999, 12, 31), DefaultDateFormat) & "."
            lngTotalDays = MaxDaysUp
         End If
         
         ShowMessage AppText(15) & " " & strPrompt & vbCrLf & Replace(AppError(42), "#", lngTotalDays), vbInformation, AppText(15), AppError(40), TimeToWait
         
         With txtDaysToDate
            .Text = lngTotalDays
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
         End With
         
         IsCheckDaysToGo = False
         Exit Function
      End If
   End If
   
   CheckDaysToGo = True

End Function

Private Function DrawListHeader(ByVal Index As Integer, ByVal Text As String, Optional ByVal Width As Integer) As Integer

Dim intLeft As Integer

   With picList
      If Index Then intLeft = Header(Index - 1).Right
      If Width = 0 Then Width = .ScaleWidth - intLeft
      
      Header(Index).Left = intLeft
      Header(Index).Right = intLeft + Width
      Header(Index).Bottom = HDR_HEIGHT
      DrawEdge .hDC, Header(Index), BDR_RAISEDINNER, BF_RECT
      intLeft = Header(Index).Left + 5
      
      If (Index = 3) Or ((Index = 0) And WithIcons) Then intLeft = Header(Index).Left + (Header(Index).Right - Header(Index).Left - .TextWidth(Text)) \ 2
      
      .ForeColor = &HC01FC0
      .CurrentY = (Header(Index).Bottom - .TextHeight("X")) \ 2
      .CurrentX = intLeft
      picList.Print Text
      DrawListHeader = intLeft
   End With

End Function

Private Function GetSelectedItem(ByVal WithKeyboard As Boolean) As String

   If (RowIndex = -1) Or ((RowIndex > ListIndex) And (WithKeyboard And (ListIndex < 0))) Then Exit Function
   
   GetSelectedItem = DateItems(ListIndex)

End Function

Private Function SetActiveListIndex(ByVal Index As Integer, ByVal IsDate As String, ByVal LowDate As String, ByVal HighDate As String) As Boolean

Dim strDate(2) As String

   With calCalendar
      strDate(0) = Format(CDate(DateSerial(.CalYear, .CalMonth, .CalDay)), "mmdd")
      strDate(2) = Format(CDate(Trim(Split(IsDate, "-", 2)(1)) & " " & .CalYear), "mmdd")
      strDate(1) = Format(CDate(Trim(Split(IsDate, "-")(0)) & " " & .CalYear), "mmdd")
   End With
   
   If (strDate(0) < LowDate) Or (strDate(0) > HighDate) Then
      If LowDate = "0320" Then
         ListIndex = 1 + (2 And (calCalendar.Hemisphere = North))
         
      Else
         ListIndex = 9
      End If
      
      SetActiveListIndex = True
      
   ElseIf (strDate(0) >= strDate(1)) And (strDate(0) <= strDate(2)) Then
      ListIndex = Index
      SetActiveListIndex = True
   End If
   
   Erase strDate

End Function

Private Sub ChangeHemisphere()

   calCalendar.Hemisphere = Abs(AppSettings(SET_HEMISPHERE))
   cbtChoose.Item(3).ToolTipText = GetToolTipText(AppText(207 + AppSettings(SET_HEMISPHERE)))
   cbtChoose.Item(3).Picture = imgImages.Item(4 + AppSettings(SET_HEMISPHERE)).Picture

End Sub

Private Sub ClearNextDate()

   flbDate.Visible = False
   SelectedDateItem = ""

End Sub

Private Sub CreateSpecialDays()

   MousePointer = vbHourglass
   Load frmSpecialDays
   picList.Visible = False
   frmSpecialDays.Show vbModal, Me
   picList.Visible = True
   MousePointer = vbDefault

End Sub

Private Sub CreateWindow(Optional ByVal State As Boolean = True)

   Call DrawHeader(Me, AppText(56 + (154 And (chkTimeToGo.Visible Or TimeToGoOnly))), True)
   
   If State Then Call DrawText(Me, AppText(13), 9, &H801F80, 0, lstWeekDays.Top, 30)

End Sub

Private Sub DrawListItems(ByVal Start As Integer)

Dim intCount      As Integer
Dim intHeader     As Integer
Dim intSize       As Integer
Dim intTextHeight As Integer
Dim intValue      As Integer
Dim lngX          As Long
Dim lngY          As Long
Dim strText()     As String

   With picList
      .Cls
      .CurrentY = HDR_HEIGHT - ExtraLineSpace
      intTextHeight = .TextHeight("X")
      
      If WithIcons Then intSize = (32 - intTextHeight) \ 2
      
      For intCount = Start To Start + PageItems
         strText = Split(ListItems(intCount), Separator)
         lngY = .CurrentY + intSize + ((intSize + 1) And WithIcons) + (2 And (ScreenResize > 1))
         intHeader = Abs(WithIcons)
         
         If intCount = ListIndex Then
            lngX = 0 + (lngX = Header(1).Left And WithIcons)
            picList.Line (lngX, lngY)-(.ScaleWidth, lngY + intTextHeight), &HB93B00, BF
            .ForeColor = &HFFFFD6
            
         Else
            .ForeColor = &HD94600
         End If
         
         For intValue = 0 To UBound(strText)
            If intHeader > 2 Then Exit For
            
            .CurrentY = lngY
            .CurrentX = Tabs(intHeader)
            picList.Print Trim(strText(intValue))
            intHeader = intHeader + 1
         Next 'intValue
         
         If WithIcons Then
            DrawIconEx .hDC, Tabs(0), lngY - intSize, ItemIcon(intCount).Handle, 32, 32, 0, 0, DI_NORMAL
            picList.Print
         End If
      Next 'intCount
   End With
   
   Erase strText

End Sub

Private Sub EndCalendar()

   MousePointer = vbHourglass
   
   Call frmMyTimeZones.ToggleControls(True)
   
   With calCalendar
      CalendarDate = CDate(DateSerial(.CalYear, .CalMonth, .CalDay))
   End With
   
   Hide
   DoEvents
   Unload Me
   Set frmCalendar = Nothing

End Sub

Private Sub FillList(ByVal Index As Integer, ByRef Box As Object)

Dim intCount  As Integer
Dim strList() As String

   strList = Split(AppText(Index), ",")
   
   For intCount = 0 To UBound(strList)
      Box.AddItem strList(intCount)
   Next 'intCount
   
   Erase strList

End Sub

Private Sub FillTimeToGo()

Dim intIndex As Integer

   With calTimeToGo
      TimeToGo(TimeToGoEdit) = CDate(DateSerial(.CalYear, .CalMonth, .CalDay))
      
      Call CheckDatesTimeToGo(True, lblTimeToGo)
      
      lblTimeToGo.Item(TimeToGoEdit).Caption = Format(TimeToGo(TimeToGoEdit), "d mmmm yyyy")
   End With
   
   DoEvents
   intIndex = Abs(Not CBool(TimeToGoEdit))
   
   If Len(TimeToGo(intIndex)) Then intIndex = TimeToGoEdit
   
   Call lblTimeToGo_Click(intIndex)
   Call SetDisplay

End Sub

Private Sub GetListItems(ByVal Index As Integer, Optional ByVal CreateDisplay As Boolean = True)

Dim blnListIndex As Boolean
Dim intCount     As Integer
Dim intHeader(3) As Integer
Dim strDate(1)   As String
Dim strHeader(3) As String
Dim strText      As String

   MousePointer = vbHourglass
   ListIndex = -1
   
   With tsbVertical
      .Visible = False
      .Value = 0
      .Max = 0
   End With
   
   If CreateDisplay Then
      With picList
         BitBlt .hDC, 0, 0, .ScaleWidth, .ScaleHeight, hDC, .Left, .Top, vbSrcCopy
      End With
      
      ShowInfo = Not ShowInfo
      picList.Visible = ShowInfo
      calCalendar.Visible = Not ShowInfo
      
      If Not ShowInfo Then
         Call RestoreButtons
         Call ChangeHemisphere
         Call SetSpecialDays(calCalendar)
         Call ToggleControls
         Call SetToolTipText
         
         Erase ItemIcon
         Exit Sub
      End If
      
      Call ToggleControls
      Call DrawFooter(Me, AppText(23) & " " & LCase(AppText(18 + Index)), 12)
      
      If Index Then
         With cbtChoose.Item(0)
            .Shape = ShapeNone
            .Left = .Left - (.Width * 1.3 - .Width)
            .Width = .Width * 1.3
         End With
         
         cbtChoose.Item(1).Visible = False
         WithIcons = True
         strHeader(0) = calCalendar.CalYear
         strHeader(1) = AppText(16)
         strHeader(2) = AppText(17)
         intHeader(0) = HDR_WIDTH_YEAR
         intHeader(1) = picList.ScaleWidth \ 2 - HDR_WIDTH_YEAR
         
         ' Zodiac
         If Index = 4 Then
            HasTipText = True
            ListCount = 11
            intHeader(1) = intHeader(1) - 42
            Separator = ","
            
         ' MoonPhases
         ElseIf Index = 3 Then
            HasTipText = False
            strHeader(2) = AppText(15)
            Separator = vbNullChar
            
         ' 2 or 1 Quarters & Seasons
         Else
            HasTipText = False
            ListCount = 3
            Separator = ","
         End If
         
      ' 0 = SpecialDays
      Else
         HasTipText = False
         WithIcons = False
         strHeader(0) = AppText(14)
         strHeader(1) = AppText(15)
         strHeader(2) = AppText(16)
         strHeader(3) = calCalendar.CalYear
         intHeader(0) = HDR_WIDTH_DAY
         intHeader(1) = HDR_WIDTH_DATE
         intHeader(2) = DescriptionWidth
         Separator = ","
         
         With cbtChoose.Item(1)
            .ToolTipText = GetToolTipText(AppText(208))
            .Picture = imgImages.Item(5).Picture
            .Shape = ShapeRight
         End With
      End If
      
      With tsbVertical
         .Top = HDR_HEIGHT
         .Left = picList.ScaleWidth - .Width
         .Height = picList.ScaleHeight - HDR_HEIGHT
      End With
      
      For intCount = 0 To 3
         Tabs(intCount) = DrawListHeader(intCount, strHeader(intCount), intHeader(intCount))
      Next 'intCount
      
      DoEvents
      picList.Picture = picList.Image
   End If
   
   With calCalendar
      ' MoonPhases
      If Index = 3 Then
         ReDim ItemIcon(0) As StdPicture
         
         ListCount = 0
         strDate(0) = CDate("1-1-" & .CalYear)
         
         Do Until Year(strDate(0)) <> .CalYear
            strText = .GetMoonPhaseInfo(strDate(0), ItemIcon(ListCount))
            
            If InStr(strText, ",") Then
               strDate(1) = CDate(RTrim(Split(strText, ",", 2)(1)))
               
               If Format(strDate(0), "yyyymmdd") < Format(strDate(1), "yyyymmdd") Then
                  strDate(0) = strDate(1)
                  intCount = 0
                  
               Else
                  intCount = 3
               End If
               
            Else
               ReDim Preserve ListItems(ListCount) As String
               ReDim Preserve DateItems(ListCount) As String
               
               strDate(1) = FormatDateTime(strDate(0), vbLongDate)
               strDate(1) = Split(strDate(1), .CalYear)(0) & Split(strDate(1), .CalYear, 2)(1)
               
               If Right(strDate(1), 4) = " de " Then strDate(1) = Left(strDate(1), Len(strDate(1)) - 4) ' delete ' de ' in spanish windows
               
               ListItems(ListCount) = strText & vbNullChar & CapsText(strDate(1))
               DateItems(ListCount) = strDate(0)
               ListCount = ListCount + 1
               intCount = 5
               
               ReDim Preserve ItemIcon(ListCount) As StdPicture
               
               If Not blnListIndex Then
                  If Month(strDate(0)) = .CalMonth Then
                     If Day(strDate(0)) >= .CalDay Then
                        blnListIndex = True
                        ListIndex = ListCount - 1
                     End If
                     
                  ElseIf Month(strDate(0)) > .CalMonth Then
                     blnListIndex = True
                     ListIndex = ListCount - 2
                  End If
               End If
            End If
            
            strDate(0) = DateAdd("d", intCount, strDate(0))
         Loop
         
      Else
         If Index Then
            ReDim ItemIcon(ListCount) As StdPicture
            
         Else
            ListCount = lstSpecialDays.ListCount - 1
         End If
         
         ReDim ListItems(ListCount) As String
         ReDim DateItems(ListCount) As String
         
         strDate(0) = Format(DateSerial(.CalYear, .CalMonth, .CalDay), DefaultDateFormat)
         
         For ListCount = 0 To ListCount
            ' SpecialDays
            If Index = 0 Then
               strText = lstSpecialDays.List(ListCount)
               strDate(1) = LTrim(Split(strText, ",", 2)(1))
               strDate(1) = Format(CDate(RTrim(Split(strDate(1), ",")(0)) & " " & .CalYear), DefaultDateFormat)
               
               If Not blnListIndex Then
                  If ((Month(strDate(1)) = Month(strDate(0))) And (Day(strDate(1)) >= Day(strDate(0)))) Or (Month(strDate(1)) > Month(strDate(0))) Then
                     blnListIndex = True
                     ListIndex = ListCount
                  End If
               End If
               
            ' Quarters
            ElseIf Index = 1 Then
               strText = .GetQuarterInfo(ListCount + 1, ItemIcon(ListCount))
               strDate(1) = Trim(Split(strText, ",")(1))
               ListIndex = DatePart("q", strDate(0)) - 1
               blnListIndex = True
               
            ' Seasons
            ElseIf Index = 2 Then
               strText = .GetSeasonInfo(ListCount + 1, ItemIcon(ListCount))
               strDate(1) = Trim(Split(strText, ",")(1))
               
               If Not blnListIndex Then blnListIndex = SetActiveListIndex(ListCount, LTrim(Split(strText, ",", 2)(1)), "0320", "1220")
               
            ' 4 =  ZodiacSigns
            Else
               strText = .GetZodiacInfo(ListCount + 1, ItemIcon(ListCount))
               strDate(1) = Trim(Split(strText, ",")(1))
               ZodiacTipText(ListCount) = GetToolTipText(Trim(Split(strText, ",")(2)))
               
               If Not blnListIndex Then blnListIndex = SetActiveListIndex(ListCount, LTrim(Split(strText, ",")(1)), "0121", "1221")
            End If
            
            ListItems(ListCount) = strText
            DateItems(ListCount) = strDate(1)
         Next 'ListCount
      End If
   End With
   
   If Not blnListIndex Then
      blnListIndex = True
      ListIndex = ListCount - 1
   End If
   
   With tsbVertical
      MousePointer = vbDefault
      intCount = 1 + (1 And WithIcons)
      ExtraLineSpace = (2 And (ScreenResize > 1) And Not WithIcons)
      PageItems = (picList.ScaleHeight - HDR_HEIGHT) \ (picList.TextHeight("X") + ExtraLineSpace) - 1
      PageItems = PageItems \ intCount + WithIcons * 3 - (1 And (ScreenResize > 1) And WithIcons)
      
      If PageItems >= ListCount Then PageItems = ListCount - 1
      
      .Max = ListCount - PageItems - 1
      .LargeChange = PageItems - ((PageItems - .Max) And (.Max < PageItems))
      .Visible = (.Max > 0)
      
      If blnListIndex And (.Value <> .Max) Then
         intCount = -(ListCount - ListIndex) + PageItems + 1
         
         If .Max + intCount = .Value Then
            Call DrawListItems(.Value)
            
         Else
            .Value = .Max + (intCount - (intCount And (intCount > 0)))
         End If
         
      Else
         Call DrawListItems(0)
      End If
   End With
   
   picList.SetFocus
   Erase intHeader, strDate, strHeader

End Sub

Private Sub ResetDateInput()

   ledDisplay.Text = ""
   DoEvents
   ledDisplay.Active = False
   lblTimeToGo.Item(0).Caption = ""
   lblTimeToGo.Item(1).Caption = ""
   TimeToGo(0) = ""
   TimeToGo(1) = ""
   TimeToGoEdit = 1
   
   Call SetCalendarDate(calTimeToGo, CalendarDate)
   Call lblTimeToGo_Click(0)

End Sub

Private Sub RestoreButtons()

   With cbtChoose.Item(0)
      .Width = cbtChoose.Item(1).Width
      .Left = cbtChoose.Item(1).Left + .Width - 1
      .Shape = ShapeLeft
   End With
   
   With cbtChoose.Item(1)
      .Picture = imgImages.Item(1).Picture
      .Shape = ShapeSides
      .Visible = True
   End With

End Sub

Private Sub SetDisplay()

   With ledDisplay
      .Text = "  0000000000000  "
      .BackColor = &HCFC1AF
      .Active = True
   End With
   
   If CheckTimeToGo Then
      tmrTimeToGo.Enabled = True
      ShowTimeToGo ledDisplay
   End If

End Sub

Private Sub SetFirstWeekDay()

   With lstWeekDays
      If FirstWeekDay <> .TopIndex Then Call ClearNextDate
      
      FirstWeekDay = .TopIndex
      .ListIndex = FirstWeekDay
      calCalendar.FirstWeekDay = FirstWeekDay
      calTimeToGo.FirstWeekDay = FirstWeekDay
   End With
   
   Call SetSpecialDays(calCalendar)

End Sub

Private Sub SetItemDate(ByVal Reset As Boolean)

Static intDatePart As Integer

Dim intYear        As Integer
Dim strDateBegin   As String
Dim strDateEnd     As String
Dim strYear        As String

   If SelectedDateItem = "" Then Exit Sub
   If Reset Then intDatePart = 1
   
   intDatePart = Abs(Not CBool(intDatePart))
   
   If InStr(SelectedDateItem, " - ") Then
      strDateBegin = DateValue(Split(SelectedDateItem, " - ")(intDatePart) & " " & Year(Date))
      strDateEnd = DateValue(Split(SelectedDateItem, " - ")(Abs(Not CBool(intDatePart))) & " " & Year(Date))
      
      If Month(strDateBegin) > Month(Date) Then strDateBegin = DateAdd("yyyy", -1, strDateBegin)
      
   Else
      strDateBegin = SelectedDateItem
      strDateEnd = strDateBegin
      
      If Not Reset Then Exit Sub
   End If
   
   With calCalendar
      If (Month(strDateBegin) = 12) And (Month(strDateEnd) < 4) Then
         .CalYear = Year(strDateBegin)
         intYear = .CalYear + 1
         
      ElseIf (Month(strDateBegin) < 4) And (Month(strDateEnd) = 12) Then
         .CalYear = Year(strDateEnd)
         intYear = .CalYear - 1
         
      Else
         intYear = .CalYear
      End If
      
      .CalMonth = Month(strDateBegin)
      .CalDay = Day(strDateBegin)
      IsLeftButton = False
      
      If (CalButtonID > DownButton) And (CalButtonID <> MoonPhaseButton) Then
         With flbDate
            If Reset Then
               .Visible = True
               .Left = shpBorder.Item(2).Left + (CalButtonID - 4) * 99 - 10 + CalButtonID
            End If
            
            .ToolTipText = GetToolTipText(AppText(201) & " " & CapsText(Format(DateValue(Split(SelectedDateItem, " - ")(Abs(Not CBool(intDatePart))) & " " & intYear), LongDateFormat)))
         End With
      End If
      
      Call ResetSpecialDays
      
      If Reset Then
         Call GetListItems(0)
         
      Else
         Call SetSpecialDays(calCalendar)
      End If
      
      DoEvents
      .SetFocus
   End With

End Sub

Private Sub SetScrollBarWheelScan(ByVal State As Boolean)

   tsbVertical.MouseWheel = State
   tsbVertical.MouseWheelInContainer = State

End Sub

Private Sub SetTimeToGoDate()

   If txtDaysToDate.Text = "" Then Exit Sub
   
   If TimeToGoEdit = 0 Then
      If Len(lblTimeToGo.Item(1).Caption) Then
         TimeToGo(0) = DateAdd("d", Val(txtDaysToDate.Text), TimeToGo(1))
         
      Else
         TimeToGo(0) = DateAdd("d", Val(txtDaysToDate.Text), Date)
      End If
      
   ElseIf Len(lblTimeToGo.Item(0).Caption) Then
      TimeToGo(1) = DateAdd("d", Val(txtDaysToDate.Text), TimeToGo(0))
      
   Else
      TimeToGo(1) = DateAdd("d", Val(txtDaysToDate.Text), Date)
   End If
   
   lblTimeToGo.Item(TimeToGoEdit).Caption = Format(TimeToGo(TimeToGoEdit), "d mmmm yyyy")
   txtDaysToDate.Text = ""
   
   With calTimeToGo
      .CalDay = Day(TimeToGo(TimeToGoEdit))
      .CalMonth = Month(TimeToGo(TimeToGoEdit))
      .CalYear = Year(TimeToGo(TimeToGoEdit))
   End With
   
   Call CheckDatesTimeToGo(True, lblTimeToGo)
   Call SetDisplay

End Sub

Private Sub SetToolTipText()

   With cbtChoose
      .Item(0).ToolTipText = GetToolTipText(AppText(12))
      .Item(1).ToolTipText = GetToolTipText(AppText(205 + ShowInfo))
      .Item(2).ToolTipText = GetToolTipText(AppText(210))
   End With

End Sub

Private Sub ToggleControls(Optional IsTimeToGo As Boolean)

Dim objBorder As Object

   Call CreateWindow(Not ShowInfo)
   
   picHide.Visible = IsTimeToGo
   lstWeekDays.Visible = Not ShowInfo
   DoEvents
   
   If IsTimeToGo And ShowInfo Then
      Set objBorder = picTimeToGo
      
   Else
      calCalendar.Visible = Not ShowInfo
      
      If ShowInfo Then
         Set objBorder = picList
      
      Else
         Set objBorder = calCalendar
      End If
   End If
   
   With shpBorder.Item(2)
      .Top = objBorder.Top - 1
      .Left = objBorder.Left - 1
      .Width = objBorder.Width + 2
      .Height = objBorder.Height + 2
   End With
   
   Set objBorder = Nothing
   shpBorder.Item(3).Visible = Not ShowInfo
   cbtChoose.Item(2).Visible = Not ShowInfo
   cbtChoose.Item(3).Visible = Not ShowInfo
   picHide.Visible = False
   MousePointer = vbDefault
   DoEvents

End Sub

Private Sub calCalendar_ButtonClick(ButtonID As ButtonTypes)

   Call ClearNextDate
   
   If ButtonID > DownButton Then
      ' 5 = quarterbutton,   6 = seasonbutton
      ' 7 = moonphasebutton, 8 = zodiacbutton
      Call GetListItems(ButtonID - DownButton)
      
      CalButtonID = ButtonID
      
   Else
      Call ResetSpecialDays
      Call SetSpecialDays(calCalendar)
   End If

End Sub

Private Sub calCalendar_DayClick(Button As Integer, Shift As Integer, IsDay As Integer, IsMonth As Integer, Cancel As Boolean)

   Call ClearNextDate

End Sub

Private Sub calTimeToGo_ButtonClick(ButtonID As ButtonTypes)

   If ButtonID < QuarterButton Then
      If ButtonID = ToDayButton Then LastMonthDay = False
      If LastMonthDay Then calTimeToGo.CalDay = calTimeToGo.GetMonthDays
   End If

End Sub

Private Sub calTimeToGo_DateChanged(ButtonID As ButtonTypes)

   LastMonthDay = (calTimeToGo.CalDay = calTimeToGo.GetMonthDays)
   
   If ButtonID < ToDayButton Then Call FillTimeToGo

End Sub

Private Sub calTimeToGo_DayClick(Button As Integer, Shift As Integer, IsDay As Integer, IsMonth As Integer, Cancel As Boolean)

Static blnBusy As Boolean

   With calTimeToGo
      LastMonthDay = (IsDay = .GetMonthDays(IsMonth))
      
      If (IsMonth = .CalMonth) And (IsDay = .CalDay) Then
         If blnBusy Then
            blnBusy = False
            
         Else
            blnBusy = True
            
            Call FillTimeToGo
         End If
      End If
   End With

End Sub

Private Sub cbtChoose_Click(Index As Integer)

Dim strPrompt As String

   If IsCheckDaysToGo Then Exit Sub
   
   Call ClearNextDate
   
   Select Case Index
      Case 0
         If picList.Visible Then
            ' exit listings
            Call GetListItems(0)
            
         ElseIf picTimeToGo.Visible Then
            ' go back to calandar
            MousePointer = vbHourglass
            ShowInfo = False
            
            If Len(lblTimeToGo.Item(1).Caption) Then
               TimeToGoShow = cmbTimeToGo.ListIndex
               AppSettings(SET_AUTODELETETIMETOGO) = CBool(chkTimeToGo.Value)
               
            Else
               TimeToGo(0) = ""
               TimeToGo(1) = ""
               TimeToGoShow = -1
            End If
            
            If TimeToGoOnly Then
               Call EndCalendar
               
            Else
               picTimeToGo.Visible = False
               chkTimeToGo.Visible = False
               
               Call RestoreButtons
               Call SetToolTipText
               Call ToggleControls(True)
            End If
            
         Else
            ' exit
            Call EndCalendar
         End If
         
      Case 1
         If picList.Visible Then
            ' show window to edit special days
            Call CreateSpecialDays
            
            If lstSort.ListCount Then
               Call SetSpecialDays(calCalendar)
               Call GetListItems(0, False)
               
            Else
               Call GetListItems(0)
            End If
            
         ElseIf picTimeToGo.Visible Then
            If ledDisplay.Active Then If AppSettings(SET_ASKCONFIRM) Then If ShowMessage(AppError(48), vbQuestion, AppText(210), AppError(3), TimeToWait) = vbNo Then Exit Sub
            
            Call ResetDateInput
            
         Else
            ' show special days
            If NoSpecialDays Then
               strPrompt = Replace(AppError(26), "$", SPECIAL_DAYS) & vbCrLf & AppError(27)
               
               If ShowMessage(strPrompt, vbQuestion, AppText(18), AppError(3), TimeToWait) = vbNo Then
                  Exit Sub
                  
               Else
                  Call CreateSpecialDays
                  
                  If NoSpecialDays Then
                     picList.Visible = False
                     Exit Sub
                  End If
               End If
            End If
            
            CalButtonID = -1
            
            Call GetListItems(0)
         End If
         
      Case 2
         Call CreateTimeToGo
         
      Case 3
         AppSettings(SET_HEMISPHERE) = Not AppSettings(SET_HEMISPHERE)
         
         Call ChangeHemisphere
         Call SetSpecialDays(calCalendar)
   End Select

End Sub

Private Sub chkTimeToGo_Click()

   AppSettings(SET_AUTODELETETIMETOGO) = Not AppSettings(SET_AUTODELETETIMETOGO)

End Sub

Private Sub cmbTimeToGo_Click()

   TimeToGoShow = cmbTimeToGo.ListIndex
   
   If CheckTimeToGo Then ShowTimeToGo ledDisplay

End Sub

Private Sub cmbTimeToGoShowType_Click()

Dim intTimeToGoShowType As Integer

   intTimeToGoShowType = TimeToGoShowType
   TimeToGoShowType = cmbTimeToGoShowType.ListIndex
   
   If (TimeToGoShowType = 1) And (Format(TimeToGo(0), "yyyymmdd") > Format(Date, "yyyymmdd")) Then
      TimeToGoShowType = intTimeToGoShowType
      cmbTimeToGoShowType.ListIndex = intTimeToGoShowType
      
   ElseIf (TimeToGoShowType = 2) And (Format(TimeToGo(1), "yyyymmdd") < Format(Date, "yyyymmdd")) Then
      TimeToGoShowType = intTimeToGoShowType
      cmbTimeToGoShowType.ListIndex = intTimeToGoShowType
   End If
   
   If CheckTimeToGo Then ShowTimeToGo ledDisplay

End Sub

Private Sub flbDate_Click()

   Call SetItemDate(False)

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

   If IsExit(KeyCode, Shift) Then
      Call cbtChoose_Click(0)
      
   ElseIf calCalendar.Visible Then
      If KeyCode = vbKeyReturn Then Call SetItemDate(False)
   End If

End Sub

Private Sub Form_Load()

Dim intCount As Integer

   Call SetIcon(imgImages.Item(6), 30)
   Call InitForm(Me, 0)
   Call RoundForm(hWnd, ScaleWidth, ScaleHeight, 62)
   Call CenterForm(Me)
   Call ResizeAllControls(Me)
   Call CreateWindow
   Call FillList(217, cmbTimeToGo)
   Call FillList(218, cmbTimeToGoShowType)
   Call RemoveListBoxBorder(lstWeekDays.hWnd)
   
   With tsbVertical
      .Left = picList.Width - .Width
      .Height = .Height * ScreenResize
   End With
   
   cbtChoose.Item(0).Picture = imgImages.Item(0).Picture
   DescriptionWidth = picList.ScaleWidth - (HDR_WIDTH_DAY + HDR_WIDTH_DATE + HDR_WIDTH_YEAR)
   lstWeekDays.ToolTipText = GetToolTipText(AppText(200))
   lstWeekDays.Left = 40 + TextWidth(AppText(13))
   SetClockBackground hDC, picClock.Item(0)
   ClockY = SetClockBackground(hDC, picClock.Item(1))
   
   For intCount = 1 To 3
      Load cbtChoose.Item(intCount)
      
      With cbtChoose.Item(intCount)
         If intCount < 4 Then
            .Left = cbtChoose.Item(intCount - 1).Left - .Width + 1
            .Shape = Choose(intCount, ShapeSides, ShapeSides, ShapeRight)
            .Visible = True
            .Picture = imgImages.Item(intCount).Picture
         End If
      End With
   Next 'intCount
   
   With lblTimeToGo
      Load .Item(1)
      .Item(1).Top = .Item(0).Top + 62
      .Item(1).Visible = True
   End With
   
   For intCount = 1 To 7
      Load shpBorder.Item(intCount)
      
      With shpBorder.Item(intCount)
         .Visible = True
         
         If intCount > 3 Then
            Set .Container = picTimeToGo
            
            If intCount = 4 Then .BorderColor = &HC01FC0
         End If
      End With
   Next 'intCount
   
   For intCount = 0 To 1
      With shpBorder.Item(intCount)
         .Top = picClock.Item(intCount).Top - 1
         .Left = picClock.Item(intCount).Left - 1
         .Height = picClock.Item(intCount).Height + 2
         .Width = picClock.Item(intCount).Width + 2
      End With
      
      With shpBorder.Item(intCount + 4)
         .Top = lblTimeToGo.Item(intCount).Top - 2
         .Left = lblTimeToGo.Item(intCount).Left + 1
         .Height = lblTimeToGo.Item(intCount).Height - 2
         .Width = lblTimeToGo.Item(intCount).Width - 2
      End With
   Next 'intCount
   
   With shpBorder.Item(3)
      .Top = lstWeekDays.Top - 1
      .Left = lstWeekDays.Left - 1
      .Height = lstWeekDays.Height + 2
      .Width = lstWeekDays.Width + 2
   End With
   
   With shpBorder.Item(6)
      .Top = txtDaysToDate.Top - 1
      .Left = txtDaysToDate.Left - 1
      .Width = txtDaysToDate.Width + 2
      .Height = txtDaysToDate.Height + 2
   End With
   
   With shpBorder.Item(7)
      .Top = ledDisplay.Top - 1
      .Left = ledDisplay.Left - 1
      .Width = ledDisplay.Width + 2
      .Height = ledDisplay.Height + 2
   End With
   
   With calCalendar
      shpBorder.Item(2).Top = .Top - 1
      shpBorder.Item(2).Left = .Left - 1
      shpBorder.Item(2).Width = .Width + 2
      shpBorder.Item(2).Height = .Height + 3
      
      Call ChangeHemisphere
      
      For intCount = 1 To 7
         lstWeekDays.AddItem .GetWeekdayName(intCount)
      Next 'intCount
      
      If Left(LCase(DefaultDateFormat), 2) = "mm" Then
         .DateFormat = [mm-dd-yyyy]
         
      ElseIf Left(LCase(DefaultDateFormat), 2) = "dd" Then
         .DateFormat = [dd-mm-yyyy]
            
      Else
         .DateFormat = [yyyy-mm-dd]
      End If
      
      .FillExternalLanguage LanguageText
      .ShowToolTipText = AppSettings(SET_SHOWTIPTEXT)
      .Top = (ScaleHeight - .Height) \ 2
      shpBorder.Item(2).Top = .Top - 2
      calTimeToGo.FillExternalLanguage LanguageText
      
      Call SetCalendarDate(calCalendar, CalendarDate)
      Call .SetMarkColors(Color2:=QBColor(2))
      Call SetToolTipText
      Call ResetSpecialDays(True)
   End With
   
   With flbDate
      .Top = shpBorder.Item(2).Top + shpBorder.Item(2).Height + 2
      .IconX = 0
      .IconY = 0
      .Icon = imgImages.Item(7).Picture
   End With
   
   Call tmrClock_Timer
   
   With cbtChoose.Item(2)
      picHide.Width = .Width * 2
      picHide.Top = .Top
      picHide.Left = .Left - .Width
      picHide.Height = .Height
   End With
   
   With picHide
      BitBlt .hDC, 0, 0, .ScaleWidth, .ScaleHeight, hDC, .Left, .Top, vbSrcCopy
   End With
   
   tmrClock.Enabled = True
   cmbTimeToGoShowType.ListIndex = TimeToGoShowType
   lstWeekDays.ListIndex = FirstWeekDay

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   Call SetScrollBarWheelScan(False)

End Sub

Private Sub lblTimeToGo_Click(Index As Integer)

   Call SetTimeToGoDate
   
   If Index Then
      If Len(lblTimeToGo.Item(0).Caption) Then
         MaxDaysDown = DateDiff("d", CDate(lblTimeToGo.Item(0).Caption), DateSerial(1583, 1, 1))
         MaxDaysUp = DateDiff("d", CDate(lblTimeToGo.Item(0).Caption), DateSerial(9999, 12, 31))
         
      Else
         MaxDaysDown = DateDiff("d", Date, DateSerial(1583, 1, 1))
         MaxDaysUp = DateDiff("d", Date, DateSerial(9999, 12, 31))
      End If
      
   Else
      MaxDaysDown = DateDiff("d", Date, DateSerial(1583, 1, 1))
      MaxDaysUp = DateDiff("d", Date, DateSerial(9999, 12, 31))
   End If
   
   shpBorder.Item(TimeToGoEdit + 4).BorderColor = &HB5A791
   shpBorder.Item(Index + 4).BorderColor = &HC01FC0
   TimeToGoEdit = Index
   txtDaysToDate.Text = ""
   
   If Len(TimeToGo(Index)) Then Call SetCalendarDate(calTimeToGo, TimeToGo(Index))
   
   LastMonthDay = (calTimeToGo.CalDay = calTimeToGo.GetMonthDays)

End Sub

Private Sub lstWeekDays_Click()

   Call SetFirstWeekDay

End Sub

Private Sub lstWeekDays_Scroll()

   Call SetFirstWeekDay

End Sub

Private Sub picList_Click()

   If Not IsLeftButton Then Exit Sub
   
   If RowIndex > -1 Then
      If (RowIndex < ListCount) And (ListIndex <> tsbVertical.Value + RowIndex) Then
         ListIndex = tsbVertical.Value + RowIndex
         
         Call DrawListItems(tsbVertical.Value)
      End If
   End If

End Sub

Private Sub picList_DblClick()

   If Not IsLeftButton Then Exit Sub
   
   SelectedDateItem = GetSelectedItem(False)
   
   Call SetItemDate(True)

End Sub

Private Sub picList_GotFocus()

   tsbVertical.ContainerArrowKeys = True

End Sub

Private Sub picList_KeyPress(KeyAscii As Integer)

   If KeyAscii = vbKeyReturn Then
      SelectedDateItem = GetSelectedItem(True)
      
      Call SetItemDate(True)
   End If

End Sub

Private Sub picList_LostFocus()

   tsbVertical.ContainerArrowKeys = False

End Sub

Private Sub picList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   IsLeftButton = (Button = vbLeftButton)

End Sub

Private Sub picList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   Call SetScrollBarWheelScan(True)
   
   If Y > HDR_HEIGHT Then
      RowIndex = (Y - HDR_HEIGHT - ExtraLineSpace + ((32 - picList.TextHeight("X")) And WithIcons)) \ (picList.TextHeight("X") + (2 And (ScreenResize > 1))) \ (1 + (2 And WithIcons))
      
   Else
      RowIndex = -1
   End If
   
   If HasTipText Then
      If (Y > HDR_HEIGHT) And (X < Header(0).Right) Then
         picList.ToolTipText = ZodiacTipText(tsbVertical.Value + RowIndex)
         
      ElseIf Len(picList.ToolTipText) Then
         picList.ToolTipText = ""
      End If
   End If

End Sub

Private Sub tmrClock_Timer()

   Call DrawDateTime(picClock.Item(0), ClockY, Now, DefaultDateFormat)
   Call DrawDateTime(picClock.Item(1), ClockY, Now, "hh:mm:ss")

End Sub

Private Sub tmrTimeToGo_Timer()

   If CheckTimeToGo Then ShowTimeToGo ledDisplay

End Sub

Private Sub tsbVertical_Change()

   If (UBound(ListItems) = 0) Or Not tsbVertical.Visible Then Exit Sub
   
   Call DrawListItems(tsbVertical.Value)

End Sub

Private Sub tsbVertical_KeyDown(KeyCode As Integer, Shift As Integer)

   With tsbVertical
      If KeyCode = vbKeyUp Then
         ListIndex = ListIndex - (1 And (ListIndex > 0))
         
         If .Value + PageItems < ListIndex Then
            .Value = ListIndex - PageItems
            
         ElseIf ListIndex >= .Value Then
            Call DrawListItems(.Value)
            
         Else
            .Value = .Value - (1 And (.Value > 0))
         End If
         
      ElseIf KeyCode = vbKeyDown Then
         ListIndex = ListIndex + (1 And (ListIndex < ListCount - 1))
         
         If .Value > ListIndex Then
            .Value = ListIndex
            
         ElseIf ListIndex <= PageItems + .Value Then
            Call DrawListItems(.Value)
            
         Else
            .Value = .Value + (1 And (.Value < .Max))
         End If
         
      ElseIf KeyCode = vbKeyPageUp Then
         ListIndex = ListIndex - PageItems
         
         If ListIndex < 0 Then ListIndex = 0
         
         .Value = ListIndex
         
         If (ListIndex < PageItems) Or (.Value = .Max) Then Call DrawListItems(.Value)
         
      ElseIf KeyCode = vbKeyPageDown Then
         ListIndex = ListIndex + PageItems
         
         If ListIndex > ListCount - 1 Then ListIndex = ListCount - 1
         
         .Value = ListIndex - PageItems
         
         If (ListIndex > PageItems) Or (.Value = 0) Then Call DrawListItems(.Value)
         
      ElseIf KeyCode = vbKeyHome Then
         If ListIndex <= PageItems Then
            ListIndex = 0
            
            Call DrawListItems(0)
            
         Else
            ListIndex = 0
         End If
         
         .Value = 0
         
      ElseIf KeyCode = vbKeyEnd Then
         If ListIndex >= ListCount - PageItems - 1 Then
            ListIndex = ListCount - 1
            
            Call DrawListItems(.Max)
            
         Else
            ListIndex = ListCount - 1
         End If
         
         .Value = .Max
      End If
   End With

End Sub

Private Sub tsbVertical_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   Call SetScrollBarWheelScan(True)

End Sub

Private Sub tsbVertical_MouseWheel(ScrollLines As Integer)

   tsbVertical.Value = tsbVertical.Value + ScrollLines

End Sub

Private Sub tsbVertical_Scroll()

   Call DrawListItems(tsbVertical.Value)

End Sub

Private Sub txtDaysToDate_Change()

   With txtDaysToDate
      If Left(.Text, 1) = "-" And Mid(.Text, 2, 1) = "0" Then
         .Text = "-" & Mid(.Text, 3)
         .SelStart = 1
      End If
   End With

End Sub

Private Sub txtDaysToDate_KeyPress(KeyAscii As Integer)

Dim ptaTextCursor As PointAPI

   GetCaretPos ptaTextCursor
   
   If KeyAscii = vbKeyReturn Then
      If CheckDaysToGo Then
         Call SetTimeToGoDate
         Call lblTimeToGo_Click(1)
      End If
   End If
   
   If Not Chr(KeyAscii) Like "[0-9]" And (KeyAscii <> vbKeyBack) And (KeyAscii <> 45) Then KeyAscii = 0
   If KeyAscii = vbKey0 Then If (txtDaysToDate.Text = "") Or (ptaTextCursor.X < 2) Then KeyAscii = 0
   If KeyAscii = 45 Then If (Left(txtDaysToDate.Text, 1) = "-") Or (ptaTextCursor.X > 1) Then KeyAscii = 0

End Sub

Private Sub txtDaysToDate_LostFocus()

   If CheckDaysToGo Then Call SetTimeToGoDate

End Sub
