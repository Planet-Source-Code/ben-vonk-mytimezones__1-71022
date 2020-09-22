VERSION 5.00
Begin VB.Form frmInfo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   6060
   ClientLeft      =   0
   ClientTop       =   0
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
   Begin VB.PictureBox picBorder 
      BackColor       =   &H00EAC183&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   4632
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   52
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5520
      Visible         =   0   'False
      Width           =   624
      Begin VB.TextBox txtTime 
         Alignment       =   2  'Center
         BackColor       =   &H00F7EFE2&
         BorderStyle     =   0  'None
         ForeColor       =   &H00D94600&
         Height          =   216
         Left            =   12
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   5
         Top             =   12
         Width           =   600
      End
   End
   Begin MyTimeZones.ThemedComboBox tcbSkinner 
      Left            =   8760
      Top             =   2040
      _ExtentX        =   445
      _ExtentY        =   423
      BorderColorStyle=   1
      ComboBoxBorderColor=   15384963
      DriveListBoxBorderColor=   0
   End
   Begin MyTimeZones.ThumbWheel twhTime 
      Height          =   336
      Left            =   5280
      TabIndex        =   4
      Top             =   5472
      Visible         =   0   'False
      Width           =   708
      _ExtentX        =   1249
      _ExtentY        =   593
      Max             =   100
      ShadeControl    =   15523788
      ShadeWheel      =   16248802
      Value           =   50
   End
   Begin VB.ComboBox cmbClockName 
      BackColor       =   &H00F7EFE2&
      ForeColor       =   &H00D94600&
      Height          =   312
      Left            =   924
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   5472
      Visible         =   0   'False
      Width           =   2208
   End
   Begin VB.Timer tmrClock 
      Enabled         =   0   'False
      Interval        =   980
      Left            =   8280
      Top             =   2040
   End
   Begin VB.PictureBox picClock 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1212
      Left            =   4320
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   306
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3876
      Visible         =   0   'False
      Width           =   3672
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
      Height          =   1332
      Left            =   4248
      Top             =   3840
      Width           =   3792
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   7
      Left            =   8760
      Picture         =   "frmInfo.frx":0000
      Top             =   1560
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   4
      Left            =   8280
      Picture         =   "frmInfo.frx":0CCA
      Top             =   1080
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   6
      Left            =   8280
      Picture         =   "frmInfo.frx":1994
      Top             =   1560
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   5
      Left            =   8760
      Top             =   1080
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
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   1
      Left            =   8760
      Top             =   120
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   2
      Left            =   8280
      Picture         =   "frmInfo.frx":225E
      Top             =   600
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   3
      Left            =   8760
      Picture         =   "frmInfo.frx":2B28
      Top             =   600
      Visible         =   0   'False
      Width           =   384
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Private Constants
Private Const MAX_CLOCK_NAMES    As Integer = 10
Private Const CB_FINDSTRINGEXACT As Long = &H158

' Private Class with Events
Private WithEvents clsMouse      As clsMouseWheel
Attribute clsMouse.VB_VarHelpID = -1

' Private Variables
Private ButtonPressed            As Boolean
Private ChangeTime               As Boolean
Private OnlyDefaultNames         As Boolean
Private SetAlarm                 As Boolean
Private SubclassedTextBox        As Boolean
Private ChooseIndex              As Integer
Private FirstLine                As Integer
Private LineSpace                As Integer
Private RightTextX               As Integer
Private PrevSelectedName         As Integer
Private SelectedTimePart         As Integer
Private SelectedName             As Integer
Private DayLightTime             As Long
Private PrevName                 As String

Public Sub CreateWindow()

   InZoneInfo = True
   
   Call DrawHeader(Me)
   Call DrawTimeZoneInfo
   Call tmrClock_Timer
   
   If ShowSettings And (SelectedClock < 2) Then Call DrawFooter(Me, Left(AppText(2), Len(AppText(2)) - 1), 14)

End Sub

Private Function InfoData(ByRef TimeZone As SystemTime, ByVal Text As String) As String

Dim dteDate      As Date
Dim strTextLeft  As String
Dim strTextRight As String

   With TimeZone
      dteDate = Format(WeekDayToDate(Year(Now), .wMonth, .wDayOfWeek, .wDay) & " " & TimeSerial(.wHour, .wMinute, .wSecond), DefaultDateFormat & " hh:mm:ss")
      strTextLeft = Replace(AppText(110), "$", Text)
      
      If .wYear Then
         strTextRight = DateSerial(.wYear, .wMonth, .wDay) & " " & AppText(122)
         
      Else
         strTextRight = TranslateDay(.wDayOfWeek, .wDay) & " " & MonthName(.wMonth) & " " & AppText(122)
      End If
      
      strTextRight = strTextRight & Format(dteDate, " hh:mm:ss ") & AppText(109)
      
      Call DrawText(strTextLeft, strTextRight)
      
      If .wYear = 0 Then
         strTextRight = "(" & AppText(120 + (1 And (DateDiff("d", Now, dteDate) < 0))) & Format(dteDate, " d mmmm yyyy)")
         
         Call DrawText("", strTextRight)
      End If
   End With
   
   InfoData = Format(dteDate, "yyyymmddhhmmss")

End Function

Private Function TranslateDay(ByVal WeekDay As Long, ByVal Days As Long) As String

   TranslateDay = AppText(123) & " " & LCase(GetNamePart(AppText(8), Days)) & " " & WeekdayName(WeekDay + (7 And (WeekDay = 0)), False, vbMonday) & " " & AppText(119)

End Function

Private Function CheckForDoubleAlarmTime() As Boolean

Dim dteTime    As Date
Dim intClock   As Integer
Dim intCount   As Integer
Dim strMessage As String
Dim strNames   As String

   If Not SetAlarm Then Exit Function
   
   dteTime = Format(LocalDateToUTC(Date + CDate(txtTime.Text), AllZones(FavoritsInfo(SelectedFavorit).ZoneID)), "hh:mm")
   
   For intCount = 0 To UBound(FavoritsInfo)
      If Len(FavoritsInfo(intCount).AlarmTime) And (intCount <> SelectedFavorit) Then
         If dteTime = Format(LocalDateToUTC(Date + CDate(FavoritsInfo(intCount).AlarmTime), AllZones(FavoritsInfo(intCount).ZoneID)), "hh:mm") Then
            strNames = strNames & frmMyTimeZones.clkFavorits.Item(intCount).NameClock & " (" & AppText(130) & " " & FavoritsInfo(intCount).AlarmTime & ")" & vbCrLf
            intClock = intClock + 1
         End If
      End If
   Next 'intCount
   
   If intClock Then
      strMessage = Split(AppText(27), ",")(1 And intClock > 1)
      strMessage = Replace(Replace(AppError(39), "$", strMessage), "#", Format(dteTime, "hh:mm")) & vbCrLf & Left(strNames, Len(strNames) - 2)
      strMessage = Replace(strMessage, "$", Split(AppText(28), ",")(1 And intClock > 1))
      
      If ShowMessage(strMessage, vbQuestion, AppError(5), AppText(150), TimeToWait) = vbNo Then
         txtTime.SetFocus
         CheckForDoubleAlarmTime = True
      End If
   End If

End Function

Private Sub DrawClock(ByVal DateTime As String, Optional ByVal ForeColor As Long = &HD94600)

Dim strDateTime As String

   With picClock
      strDateTime = CapsText(DateTime)
      .CurrentX = (.Width - .TextWidth(strDateTime)) \ 2
      .ForeColor = ForeColor
      picClock.Print strDateTime
      .CurrentY = .CurrentY + LineSpace
   End With

End Sub

Private Sub DrawText(ByVal TextLeft As String, ByVal TextRight As String, Optional ByVal LeftForeColor As Long = &H801F80, Optional ByVal RightForeColor As Long = &HB93B00)

   ForeColor = LeftForeColor
   CurrentX = 15
   Print TextLeft;
   ForeColor = RightForeColor
   CurrentX = RightTextX
   Print TextRight

End Sub

Private Sub DrawTimeZoneClock(ByVal DateTimeInfo As String)

Dim intCount   As Integer
Dim lngColor   As Long
Dim strText(3) As String

   intCount = 3 + (1 And (DayLightTime <> 0))
   LineSpace = picClock.ScaleHeight \ intCount
   CurrentY = picClock.Top + picClock.ScaleHeight \ 2 - LineSpace * (intCount / 2)
   LineSpace = LineSpace - TextHeight("X")
   CurrentY = CurrentY + LineSpace \ 2
   FirstLine = CurrentY
   strText(0) = AppText(115)
   strText(1) = AppText(116)
   strText(2) = AppText(117)
   strText(3) = DateTimeInfo
   
   For intCount = 0 To 3
      If Len(strText(intCount)) Then
         If intCount = 2 Then
            lngColor = &HC01FC0
            
         Else
            lngColor = &H801F80
         End If
         
         Call DrawText(strText(intCount), "", lngColor)
         
         CurrentY = CurrentY + LineSpace
      End If
   Next 'intCount
   
   Erase strText

End Sub

Private Sub DrawTimeZoneInfo()

Dim intCount As Integer
Dim strBegin As String
Dim strDate  As String
Dim strEnd   As String

   ReDim strText(3) As String
   
   With AllZones(SelectedTimeZoneID)
      Line (10, 180)-(ScaleWidth - 10, 180), &HB5A791
      Line (10, 309)-(ScaleWidth - 10, 309), &HB5A791
      CurrentY = 82
      RightTextX = ScaleWidth \ 2 - 50
      DayLightTime = False
      intCount = DateDiff("n", Now, UTCToLocalDate(GetSystemDate, AllZones(SelectedTimeZoneID)))
      
      If intCount < 0 Then
         strText(0) = "-"
         
      ElseIf intCount Then
         strText(0) = "+"
      End If
      
      If .StandardBias < 0 Then
         strText(1) = "-"
         
      ElseIf .StandardBias Then
         strText(1) = "+"
      End If
      
      If .Bias < 0 Then
         strText(2) = "-"
         
      ElseIf .Bias Then
         strText(2) = "+"
      End If
      
      Call DrawText(AppText(105), strText(0) & Format(DateAdd("n", Abs(intCount), "00:00:00"), "hh:mm ") & AppText(109), &HC01FC0, &HC01FC0)
      Call DrawText(AppText(124), TimeZoneRegKeyName(SelectedTimeZoneID))
      Call DrawText(AppText(102), .StandardName)
      Call DrawText(AppText(103), strText(1) & Format(DateAdd("n", Abs(.StandardBias), "00:00:00"), "hh:mm ") & AppText(104))
      Call DrawText(AppText(100), strText(2) & Format(DateAdd("n", Abs(.Bias), "00:00:00"), "hh:mm ") & AppText(101))
      
      If .DaylightDate.wMonth = 0 Then
         CurrentY = 235
         RightTextX = (ScaleWidth - TextWidth(AppText(106))) \ 2
         
         Call DrawText("", AppText(106), , &HC01FC0)
         Call DrawTimeZoneClock("")
         
         Exit Sub
      End If
      
      CurrentY = 190
      
      If .DaylightBias < 0 Then
         strText(3) = "-"
         
      ElseIf .DaylightBias Then
         strText(3) = "+"
      End If
      
      Call DrawText(AppText(107), .DaylightName)
      Call DrawText(AppText(108), strText(3) & Format(DateAdd("n", Abs(.DaylightBias), "00:00:00"), "hh:mm ") & AppText(104))
      
      strBegin = InfoData(.DaylightDate, AppText(111))
      strEnd = InfoData(.StandardDate, AppText(112))
      strDate = Format(Now, "yyyymmddhhmmss")
      
      If (strDate > strBegin) And (strDate < strEnd) Then
         strText(0) = AppText(114)
         DayLightTime = .DaylightBias
         
      Else
         strText(0) = AppText(113)
         DayLightTime = Abs(.DaylightBias)
      End If
   End With
   
   Call DrawTimeZoneClock(Replace(AppText(118), "$", strText(0)))
   
   Erase strText

End Sub

Private Sub EndInfo()

   If picBorder.Visible And AppSettings(SET_CHECKDOUBLEALARMS) Then If CheckForDoubleAlarmTime Then Exit Sub
   
   MousePointer = vbHourglass
   
   Call SetClockSettings
   Call ResetComboBox
   
   If SubclassedTextBox Then Call SubclassTextBox(txtTime.hWnd)
   
   Call MouseUnhook
   Call frmMyTimeZones.ToggleControls(True)
   
   Hide
   DoEvents
   Unload Me
   InZoneInfo = False
   Set frmInfo = Nothing

End Sub

Private Sub GetNames()

Dim blnSelected   As Boolean
Dim intCount      As Integer
Dim strName()     As String
Dim strNameSelect As String

   If FavoritsInfo(SelectedFavorit).DisplayName = "" Then Exit Sub
   
   OnlyDefaultNames = True
   strNameSelect = FavoritsInfo(SelectedFavorit).DisplayName
   
   If InStr(strNameSelect, ") ") Then strNameSelect = Replace(strNameSelect, ")", "),", , 1)
   
   With cmbClockName
      .Clear
      strName = Split(strNameSelect, ",")
      
      For intCount = 0 To UBound(strName)
         strName(intCount) = LTrim(strName(intCount))
         
         If Left(strName(intCount), 1) = "*" Then
            strName(intCount) = Mid(strName(intCount), 2)
            strNameSelect = strName(intCount)
            blnSelected = True
         End If
         
         .AddItem strName(intCount)
         
         If Right(strName(intCount), 1) = "]" Then OnlyDefaultNames = False
      Next 'intCount
      
      If Not blnSelected Then strNameSelect = strName(0)
      
      Erase strName
      PrevSelectedName = GetListIndex(.hWnd, CB_FINDSTRINGEXACT, strNameSelect)
      
      If PrevSelectedName > -1 Then .ListIndex = 0 + (PrevSelectedName And ((PrevSelectedName < .ListCount) And (PrevSelectedName > -1)))
   End With

End Sub

Private Sub MouseHook()

   ChangeTime = True
   Set clsMouse = New clsMouseWheel
   
   Call SelectTimePart
   Call clsMouse.Hook(hWnd)

End Sub

Private Sub MouseUnhook()

   If clsMouse Is Nothing Then Exit Sub
   
   Call clsMouse.Unhook
   
   Set clsMouse = Nothing
   ChangeTime = False
   DoEvents

End Sub

Private Sub SelectTimePart()

   If Not ChangeTime Then Exit Sub
   
   With txtTime
      .SelStart = 3 - (3 And (SelectedTimePart = 1))
      .SelLength = 2
      .SetFocus
      DoEvents
   End With

End Sub

Private Sub SetClockSettings()

Dim intPointer   As Integer
Dim strAlarmTime As String
Dim strName      As String

   If Not picBorder.Visible Or (SelectedClock < 2) Then Exit Sub
   
   With cmbClockName
      strName = .Text
      
      For intPointer = 0 To .ListCount - 1
         If .List(intPointer) = strName Then Exit For
      Next 'intPointer
      
      If intPointer >= .ListCount Then strName = "[" & strName & "]"
      
      strName = "*" & strName
   End With
   
   With FavoritsInfo(SelectedFavorit)
      .DisplayName = Replace(.DisplayName, "*", "")
      intPointer = InStr(.DisplayName, Mid(strName, 2))
      
      If intPointer Then
         .DisplayName = Left(.DisplayName, intPointer - 1) & strName & Mid(.DisplayName, intPointer + Len(strName) - 1)
         
      Else
         .DisplayName = .DisplayName & ", " & strName
      End If
      
      If SetAlarm Then
         .AlarmTime = txtTime.Text
         .AlarmTipText = AppText(0)
         
      Else
         .AlarmTime = ""
         .AlarmTipText = ""
      End If
   End With
   
   With frmMyTimeZones.clkFavorits.Item(SelectedFavorit)
      .Locked = True
      .NameClock = TrimClockName(Mid(strName, 2))
      strAlarmTime = .AlarmTime
      .AlarmTime = FavoritsInfo(SelectedFavorit).AlarmTime
      .AlarmToolTipText = Trim(GetToolTipText(FavoritsInfo(SelectedFavorit).AlarmTipText))
      .Locked = False
   End With
   
   With frmMyTimeZones.ledDisplay
      If (.ToolTipText = GetDisplayToolTipText(strName)) Then
         .Text = CreateAlarmMessage(SelectedFavorit)
         .NoTextScrolling = False
         
         If (SelectedFavorit = frmMyTimeZones.AlarmIndex) And (strAlarmTime = Format(Time, "hh:mm")) And (strAlarmTime <> FavoritsInfo(SelectedFavorit).AlarmTime) Then Call frmMyTimeZones.SetAlarmOff
      End If
   End With

End Sub

Private Sub SetToolTipText()

Dim intCount As Integer

   For intCount = 0 To 5
      cbtChoose.Item(intCount).ToolTipText = GetToolTipText(AppText(Choose(intCount + 1, 12, 55, 142, 144 + SetAlarm, 145, 147 + OnlyDefaultNames)))
   Next 'intCount
   
   twhTime.ToolTipText = GetToolTipText(AppText(150))
   txtTime.ToolTipText = GetToolTipText(AppText(24))
   cmbClockName.ToolTipText = GetToolTipText(AppText(151))

End Sub

Private Sub clsMouse_Wheel(ScrollLines As Integer)

   twhTime.ScrollValue = twhTime.ScrollValue - ScrollLines
   ButtonPressed = True

End Sub

Private Sub cmbClockName_Change()

Dim strText As String

   With cmbClockName
      If .Text = "" Then Exit Sub
      
      strText = .Text
      Mid(strText, 1, 1) = UCase(Mid(strText, 1, 1))
      .Text = strText
   End With

End Sub

Private Sub cmbClockName_Click()

   SelectedName = cmbClockName.ListIndex
   PrevName = cmbClockName.Text

End Sub

Private Sub cmbClockName_KeyPress(KeyAscii As Integer)

   If (KeyAscii = 40) Or (KeyAscii = 41) Then KeyAscii = vbEmpty
   
   With cmbClockName
      If .SelStart = 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
      
      If (.ListCount = MAX_CLOCK_NAMES) And (KeyAscii = vbKeyReturn) Then
         .Text = PrevName
         KeyAscii = vbEmpty
         Exit Sub
         
      ElseIf (.ListCount < MAX_CLOCK_NAMES) And (KeyAscii = vbKeyReturn) Then
         .ListIndex = GetListIndex(.hWnd, CB_FINDSTRING, .Text)
         
         If .Text = "" Then .ListIndex = 0
         If .ListIndex > -1 Then Exit Sub
         
         .Text = "[" & .Text & "]"
         .AddItem .Text
         .SelLength = Len(.Text)
         OnlyDefaultNames = False
         cbtChoose.Item(5).ToolTipText = GetToolTipText(AppText(147))
         cbtChoose.Item(5).Picture = imgImages.Item(5).Picture
         
      ElseIf (KeyAscii <> vbKeyBack) And (KeyAscii <> vbKeyDelete) Then
         If Len(.Text) > 14 Then
            .Text = Left(.Text, 15)
            .SelStart = 15
            KeyAscii = vbEmpty
         End If
      End If
   End With

End Sub

Private Sub cmbClockName_LostFocus()

   With cmbClockName
      If (.ListCount = MAX_CLOCK_NAMES) And (.Text <> PrevName) Then .Text = PrevName
   End With

End Sub

Private Sub cbtChoose_Click(Index As Integer)

   Select Case Index
      Case 0
         Call EndInfo
         
      Case 1
         MousePointer = vbHourglass
         Load frmMap
         picClock.Visible = False
         frmMap.Show vbModal, Me
         MousePointer = vbDefault
         
         Call tmrClock_Timer
         Call CreateWindow
         
         DoEvents
         
      Case 2
         Call OpenClockImage(Me)
         
         ImageName = ""
         
      Case 3
         SetAlarm = Not SetAlarm
         cbtChoose.Item(3).Picture = imgImages.Item(3 + (3 And SetAlarm)).Picture
         
         Call SetToolTipText
         
         If Not SetAlarm Then
            Call tmrClock_Timer
            
         Else
            txtTime.SetFocus
         End If
         
      Case 4
         Call OpenAlarmMessage(Me)
         
      Case 5
         If OnlyDefaultNames Then Exit Sub
         If AppSettings(SET_ASKCONFIRM) Then If ShowMessage(AppError(21), vbQuestion, AppError(20), AppError(3), TimeToWait) = vbNo Then Exit Sub
         
         FavoritsInfo(SelectedFavorit).DisplayName = AllZones(SelectedTimeZoneID).DisplayName
         cbtChoose.Item(5).Picture = imgImages.Item(7).Picture
         
         Call GetNames
         Call SetToolTipText
   End Select

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

   If IsExit(KeyCode, Shift) Then Call EndInfo

End Sub

Private Sub Form_Load()

Dim intCount As Integer
Dim intIndex As Integer

   SelectedTimePart = 2
   ChooseIndex = -1
   
   Call SetIcon(imgImages.Item(1), 5)
   Call SetIcon(imgImages.Item(5), 30)
   Call InitForm(Me)
   Call RoundForm(hWnd, ScaleWidth, ScaleHeight, 62)
   Call CenterForm(Me)
   Call ResizeAllControls(Me)
   Call GetNames
   
   For intCount = 0 To 7
      If intCount Then Load cbtChoose.Item(intCount)
      
      With cbtChoose.Item(intCount)
         If intCount = 5 Then
            .Shape = ShapeRight
            .Left = cmbClockName.Left - .Width - 7
            
         ElseIf intCount Then
            .Shape = ShapeSides
            
            If intCount = 3 Then
               .Left = picBorder.Left - .Width - 7
               
            ElseIf intCount > 5 Then
               intIndex = 2 * (intCount - 6)
               .Blocked = True
               .Left = cbtChoose.Item(3 + intIndex).Left + cbtChoose.Item(3 + intIndex).Width - 1
               .Width = cbtChoose.Item(2 + intIndex).Left - .Left + 2
               .Picture = Nothing
               
            Else
               .Left = cbtChoose.Item(intCount - 1).Left - .Width + 1
            End If
         End If
         
         If intCount < 6 Then .Picture = imgImages.Item(intCount).Picture
      End With
   Next 'intCount
   
   With picClock
      BitBlt .hDC, 0, 0, .ScaleWidth, .ScaleHeight, hDC, .Left, .Top, vbSrcCopy
      .Picture = .Image
      shpBorder.Top = .Top - 1
      shpBorder.Left = .Left - 1
      shpBorder.Width = .Width + 2
      shpBorder.Height = .Height + 2
   End With
   
   With cbtChoose
      twhTime.MouseTrap = AppSettings(SET_MOUSEINTHUMBWHEEL)
      .Item(1).Visible = ShowSettings
      .Item(2).Visible = ShowSettings
      
      If (SelectedClock > 1) And ShowSettings Then
         .Item(3).Visible = True
         .Item(4).Visible = True
         .Item(5).Visible = True
         .Item(6).Visible = True
         .Item(7).Visible = True
         cmbClockName.Visible = True
         picBorder.Visible = True
         twhTime.Visible = True
         
      Else
         cbtChoose.Item(2).Shape = ShapeRight
      End If
      
      If SelectedClock > 1 Then
         SetAlarm = (FavoritsInfo(SelectedFavorit).AlarmTime <> "")
         .Item(3).Picture = imgImages.Item(3 + (3 And SetAlarm)).Picture
         .Item(4).Picture = imgImages.Item(4).Picture
         .Item(5).Picture = imgImages.Item(5 + (2 And OnlyDefaultNames)).Picture
      End If
   End With
   
   If Not ShowSettings Then
      With cbtChoose.Item(0)
         .Width = .Width * 1.3
         .Left = ScaleWidth - .Width - 14
         .Shape = ShapeNone
      End With
   End If
   
   tmrClock.Enabled = True
   picClock.Visible = True
   PrevSelectedName = 0
   
   Call SetComboBox(hWnd, cbtChoose.Item(0))
   Call CreateWindow
   Call SetToolTipText
   
   If FavoritsInfo(SelectedFavorit).AlarmTime = "" Then
      txtTime.Text = Format(frmMyTimeZones.clkFavorits.Item(SelectedFavorit).DateTime, "hh:mm")
      
   Else
      txtTime.Text = Format(FavoritsInfo(SelectedFavorit).AlarmTime, "hh:mm")
   End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If SubclassedTextBox Then
      Call SubclassTextBox(txtTime.hWnd)
      
      SubclassedTextBox = False
   End If

End Sub

Private Sub tmrClock_Timer()

Dim dteClock As Date

   If (Second(Time) = 0) And (Minute(Time) \ 15 = Minute(Time) / 15) Then Call CreateWindow
   
   With picClock
      .Cls
      .CurrentY = FirstLine - .Top
   End With
   
   dteClock = UTCToLocalDate(GetSystemDate, AllZones(SelectedTimeZoneID))
   
   Call DrawClock(Format(Now, LongDateFormat & " - hh:mm:ss"))
   Call DrawClock(Format(GetSystemDate, LongDateFormat & " - hh:mm:ss"))
   Call DrawClock(Format(dteClock, LongDateFormat & " - hh:mm:ss"), &HC01FC0)
   
   If DayLightTime Then Call DrawClock(Format(DateAdd("n", DayLightTime, dteClock), LongDateFormat & " - hh:mm:ss"))
   If Not SetAlarm Then txtTime.Text = Format(dteClock, "hh:mm")

End Sub

Private Sub twhTime_Change()

Dim intCount As Integer

   If Not SetAlarm Then Call cbtChoose_Click(3)
   
   With twhTime
      intCount = Sgn(.Value - Val(.Tag))
      .Tag = .Value
   End With
   
   With txtTime
      If ButtonPressed Then
         .Text = Format(TimeSerial(Hour(.Text) - (intCount And (SelectedTimePart = 1)) + (24 And (Hour(.Text) = 0) And (SelectedTimePart = 1)), Minute(.Text) - (intCount And (SelectedTimePart = 2)), 0), "hh:mm")
         ButtonPressed = False
         
      Else
         .Text = Format(TimeSerial(Hour(.Text) + (intCount And (SelectedTimePart = 1)) + (24 And (Hour(.Text) = 0) And (SelectedTimePart = 1)), Minute(.Text) + (intCount And (SelectedTimePart = 2)), 0), "hh:mm")
      End If
   End With
   
   Call SelectTimePart

End Sub

Private Sub twhTime_Click()

   Call MouseUnhook
   Call MouseHook

End Sub

Private Sub twhTime_GotFocus()

   Call MouseHook

End Sub

Private Sub twhTime_LostFocus()

   Call MouseUnhook

End Sub

Private Sub txtTime_DblClick()

   Call SelectTimePart

End Sub

Private Sub txtTime_GotFocus()

   Call MouseHook

End Sub

Private Sub txtTime_KeyDown(KeyCode As Integer, Shift As Integer)

Dim intValue As Integer

   If (KeyCode = vbKeyUp) Or (KeyCode = vbKeyDown) Then
      twhTime.ScrollValue = twhTime.ScrollValue - (1 And (KeyCode = vbKeyUp)) + (1 And (KeyCode = vbKeyDown))
      KeyCode = vbEmpty
      ButtonPressed = True
      
   Else
      intValue = GetSelectedDatePart(KeyCode, SelectedTimePart, 2, 1)
      KeyCode = vbEmpty
      
      If intValue Then SelectedTimePart = intValue
      
      Call SelectTimePart
   End If

End Sub

Private Sub txtTime_KeyUp(KeyCode As Integer, Shift As Integer)

   Call SelectTimePart

End Sub

Private Sub txtTime_LostFocus()

   If SubclassedTextBox Then
      Call SubclassTextBox(txtTime.hWnd)
      
      SubclassedTextBox = False
   End If
   
   Call MouseUnhook

End Sub

Private Sub txtTime_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   SelectedTimePart = 2 - (1 And (txtTime.SelStart < 4))
   
   Call SelectTimePart

End Sub

Private Sub txtTime_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If Not SubclassedTextBox Then
      Call SubclassTextBox(txtTime.hWnd)
      
      SubclassedTextBox = True
   End If

End Sub
