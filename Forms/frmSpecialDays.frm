VERSION 5.00
Begin VB.Form frmSpecialDays 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   6060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9276
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   170
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   505
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   773
   ShowInTaskbar   =   0   'False
   Begin MyTimeZones.ThemedComboBox tcbSkinner 
      Left            =   8400
      Top             =   2640
      _ExtentX        =   445
      _ExtentY        =   423
      BorderColorStyle=   1
      ComboBoxBorderColor=   11904913
      DriveListBoxBorderColor=   0
   End
   Begin VB.PictureBox picDescription 
      BackColor       =   &H00F7EFE2&
      BorderStyle     =   0  'None
      Height          =   264
      Left            =   732
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   220
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1644
      Width           =   2640
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H00F7EFE2&
         BorderStyle     =   0  'None
         ForeColor       =   &H00D94600&
         Height          =   312
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   2640
      End
   End
   Begin VB.Timer tmrClock 
      Enabled         =   0   'False
      Interval        =   950
      Left            =   8760
      Top             =   2040
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
      TabIndex        =   14
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
      Index           =   0
      Left            =   240
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   276
      Visible         =   0   'False
      Width           =   1200
   End
   Begin MyTimeZones.FlatButton flbChoose 
      Height          =   504
      Index           =   0
      Left            =   3960
      TabIndex        =   11
      Top             =   1284
      Width           =   504
      _ExtentX        =   889
      _ExtentY        =   889
      BackStyle       =   0
      OnlyIconClick   =   -1  'True
   End
   Begin MyTimeZones.CheckBox chkDay 
      Height          =   384
      HelpContextID   =   2930
      Index           =   0
      Left            =   672
      TabIndex        =   5
      Top             =   2508
      WhatsThisHelpID =   2930
      Width           =   3012
      _ExtentX        =   5313
      _ExtentY        =   677
      BackStyle       =   0
      Enabled         =   0   'False
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
      IconCheckedGrayed=   "frmSpecialDays.frx":0000
   End
   Begin VB.ComboBox cmbDescription 
      BackColor       =   &H00F7EFE2&
      Enabled         =   0   'False
      ForeColor       =   &H00D94600&
      Height          =   312
      HelpContextID   =   2940
      Left            =   708
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1620
      WhatsThisHelpID =   2940
      Width           =   2940
   End
   Begin VB.ListBox lstSelection 
      Appearance      =   0  'Flat
      BackColor       =   &H00F7EFE2&
      Enabled         =   0   'False
      ForeColor       =   &H00D94600&
      Height          =   240
      HelpContextID   =   2950
      Index           =   0
      ItemData        =   "frmSpecialDays.frx":08DA
      Left            =   2904
      List            =   "frmSpecialDays.frx":08E1
      TabIndex        =   6
      Top             =   3132
      WhatsThisHelpID =   2950
      Width           =   732
   End
   Begin VB.ListBox lstSelection 
      Appearance      =   0  'Flat
      BackColor       =   &H00F7EFE2&
      Enabled         =   0   'False
      ForeColor       =   &H00D94600&
      Height          =   240
      HelpContextID   =   2960
      Index           =   2
      Left            =   2184
      TabIndex        =   9
      Top             =   4428
      WhatsThisHelpID =   2960
      Width           =   1452
   End
   Begin VB.ListBox lstSelection 
      Appearance      =   0  'Flat
      BackColor       =   &H00F7EFE2&
      Enabled         =   0   'False
      ForeColor       =   &H00D94600&
      Height          =   240
      HelpContextID   =   2970
      Index           =   1
      Left            =   720
      TabIndex        =   8
      Top             =   4428
      WhatsThisHelpID =   2970
      Width           =   1212
   End
   Begin MyTimeZones.Calendar calSpecialDays 
      Height          =   2772
      HelpContextID   =   2980
      Left            =   3912
      TabIndex        =   12
      Top             =   2100
      WhatsThisHelpID =   2980
      Width           =   3792
      _ExtentX        =   6689
      _ExtentY        =   4890
      ArrowColor      =   14239232
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
      GridColor       =   16711680
      ShowDayOfYear   =   0   'False
      ShowInfoBar     =   0
      ShowNavigationBar=   0
      WeekDayViewChar =   2
      WeekNumberForeColor=   14239232
   End
   Begin MyTimeZones.CheckBox chkDay 
      Height          =   384
      HelpContextID   =   3000
      Index           =   1
      Left            =   672
      TabIndex        =   7
      Top             =   3828
      WhatsThisHelpID =   3000
      Width           =   3012
      _ExtentX        =   5313
      _ExtentY        =   677
      BackStyle       =   0
      Enabled         =   0   'False
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
      IconCheckedGrayed=   "frmSpecialDays.frx":08F3
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
      Height          =   4020
      Index           =   0
      Left            =   504
      Top             =   1020
      Width           =   7212
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   6
      Left            =   8280
      Picture         =   "frmSpecialDays.frx":11CD
      Top             =   1560
      Visible         =   0   'False
      WhatsThisHelpID =   3010
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   8
      Left            =   8280
      Picture         =   "frmSpecialDays.frx":1A97
      Top             =   2040
      Visible         =   0   'False
      WhatsThisHelpID =   3020
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   7
      Left            =   8760
      Top             =   1560
      Visible         =   0   'False
      WhatsThisHelpID =   3030
      Width           =   384
   End
   Begin VB.Label lblText 
      AutoSize        =   -1  'True
      Height          =   216
      Left            =   840
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      WhatsThisHelpID =   3050
      Width           =   48
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   5
      Left            =   8760
      Picture         =   "frmSpecialDays.frx":2361
      Top             =   1080
      Visible         =   0   'False
      WhatsThisHelpID =   3060
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   2
      Left            =   8280
      Picture         =   "frmSpecialDays.frx":2C2B
      Top             =   600
      Visible         =   0   'False
      WhatsThisHelpID =   3070
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   1
      Left            =   8760
      Top             =   120
      Visible         =   0   'False
      WhatsThisHelpID =   3080
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   0
      Left            =   8280
      Top             =   120
      Visible         =   0   'False
      WhatsThisHelpID =   3090
      Width           =   384
   End
   Begin VB.Label lblMonth 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D94600&
      Height          =   300
      Left            =   4560
      TabIndex        =   10
      Top             =   1332
      WhatsThisHelpID =   3100
      Width           =   2388
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   3
      Left            =   8760
      Picture         =   "frmSpecialDays.frx":34F5
      Top             =   600
      Visible         =   0   'False
      WhatsThisHelpID =   3110
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   4
      Left            =   8280
      Picture         =   "frmSpecialDays.frx":3DBF
      Top             =   1080
      Visible         =   0   'False
      WhatsThisHelpID =   3120
      Width           =   384
   End
End
Attribute VB_Name = "frmSpecialDays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Private Constant
Private Const WEEK_DAYS As String = "SuMoTuWeThFrSa"

' Private Variables
Private IsAdd           As Boolean
Private IsChanged       As Boolean
Private Easter          As Date
Private PrevDate        As Date
Private ClockY          As Integer
Private PrevListIndex   As Integer
Private PrevValue       As Integer
Private IsMonthOfDay    As String

Private Sub CalculateDays()

Dim dteDate  As Date
Dim intDays  As Integer
Dim intIndex As Integer

   With frmCalendar
      intIndex = .GetBeginValue(Year(CalendarDate))
      dteDate = DateSerial(Year(CalendarDate), .CalcDateValue(intDays, intIndex, 31) + 2, intDays + 3)
      intDays = DateDiff("d", dteDate, CDate("31-12-" & .calCalendar.CalYear))
      
      For intDays = intDays To 0 Step -1
         lstSelection.Item(0).AddItem intDays
      Next 'intDays
      
      intIndex = lstSelection.Item(0).ListCount - 1
      intDays = DateDiff("d", dteDate, CDate("01-01-" & .calCalendar.CalYear))
      
      For intDays = -1 To intDays Step -1
         lstSelection.Item(0).AddItem intDays
      Next 'intDays
      
      lstSelection.Item(0).ListIndex = intIndex
   End With

End Sub

Private Sub CreateWindow()

   Call DrawHeader(Me, AppText(18), True)
   Call DrawText(Me, AppText(16) & ":", 9, &H801F80, 0, cmbDescription.Top - 21, cmbDescription.Left - 1)
   Call DrawText(Me, AppText(232), 9, &H801F80, 0, lstSelection.Item(0).Top, lstSelection.Item(1).Left)
   
   Line (320, 91)-(320, 414), &HB5A791
   Line (320, 160)-(637, 160), &HB5A791
   Line (48, 191)-(320, 191), &HB5A791
   Line (48, 301)-(320, 301), &HB5A791

End Sub

Private Sub EndSpecialDays()

   MousePointer = vbHourglass
   
   Call SetCalendarDate(frmCalendar.calCalendar, CalendarDate)
   
   Hide
   DoEvents
   Unload Me
   Set frmSpecialDays = Nothing

End Sub

Private Sub GetDescription()

Dim intListIndex As Integer
Dim strDate      As String

   If IsAdd + IsChanged Then Exit Sub
   
   cmbDescription.Clear
   cmbDescription.Enabled = False
   txtDescription.ToolTipText = ""
   txtDescription.Text = ""
   chkDay.Item(0).Value = Unchecked
   chkDay.Item(1).Value = Unchecked
   
   With calSpecialDays
      strDate = .CalYear & Format(.CalMonth, "#00") & Format(.CalDay, "#00")
   End With
   
   With frmCalendar.lstSort
      intListIndex = GetListIndex(.hWnd, LB_FINDSTRING, strDate)
      
      If intListIndex = -1 Then Exit Sub
      
      .ListIndex = intListIndex
      
      Do While intListIndex < .ListCount
         If Left(.List(intListIndex), 8) <> strDate Then Exit Do
         
         cmbDescription.AddItem Trim(Split(.List(intListIndex), ",", 2)(1))
         cmbDescription.ItemData(cmbDescription.NewIndex) = intListIndex
         intListIndex = intListIndex + 1
      Loop
   End With
   
   Call SetDescriptionInfo(cmbDescription.List(0))

End Sub

Private Sub ItemDelete()

   ' procedure for deleting the item
   With frmCalendar.lstSort
      SpecialDays(.ItemData(cmbDescription.ItemData(cmbDescription.ListIndex))) = ""
      .RemoveItem cmbDescription.ItemData(cmbDescription.ListIndex)
      txtDescription.Text = ""
   End With
   
   Call SaveSpecialDays
   Call GetDescription
   Call SetDescriptionInfo(cmbDescription.List(0))

End Sub

Private Sub ItemSave(ByVal ChangeItem As Boolean)

Dim intCount As Integer
Dim strItem  As String

   ' procedure for saving the item
   MousePointer = vbHourglass
   
   Call DrawFooter(Me, AppText(243), 14)
   
   IsAdd = False
   IsChanged = False
   
   Call ToggleButtons
   
   If chkDay.Item(0).Value = Checked Then
      intCount = CInt(lstSelection.Item(0).Text) + 2
      
      If intCount > -1 Then
         strItem = "+" & intCount
         
      Else
         strItem = intCount
      End If
      
      strItem = Left(strItem & Space(4), 4)
      
   ElseIf chkDay.Item(1).Value = Checked Then
      strItem = IsMonthOfDay
      
   Else
      strItem = Format(calSpecialDays.CalMonth, "#00") & Format(calSpecialDays.CalDay, "#00")
   End If
   
   strItem = strItem & " , " & txtDescription.Text
   
   If ChangeItem Then
      ' change item
      SpecialDays(frmCalendar.lstSort.ItemData(cmbDescription.ItemData(PrevListIndex))) = strItem
      
   Else
      ' add item
      ReDim Preserve SpecialDays(UBound(SpecialDays) + 1) As String
      
      SpecialDays(UBound(SpecialDays)) = strItem
   End If
   
   strItem = txtDescription.Text
   
   Call SaveSpecialDays
   Call GetDescription
   Call SetDescriptionInfo(strItem)

End Sub

Private Sub PreviousInfo(ByVal Store As Boolean)

   If Store Then
      PrevListIndex = cmbDescription.ListIndex
      
      With calSpecialDays
         PrevDate = DateSerial(.CalYear, .CalMonth, .CalDay)
      End With
      
   Else
      With calSpecialDays
         .Locked = True
         .CalMonth = Month(PrevDate)
         .CalDay = Day(PrevDate)
         .Locked = False
      End With
      
      Call frmCalendar.SetSpecialDays(calSpecialDays)
      Call GetDescription
      
      If PrevListIndex > -1 Then cmbDescription.ListIndex = PrevListIndex
   End If

End Sub

Private Sub SaveSpecialDays()

Dim blnSaveItems As Boolean
Dim intCount     As Integer
Dim intFileWrite As Integer

   On Local Error GoTo ExitSub
   
   If Dir(DataPath, vbDirectory) = "" Then MkDir DataPath
   
   intFileWrite = FreeFile
   
   Open DataPath & SPECIAL_DAYS For Output As #intFileWrite
      Print #intFileWrite, "[SpecialDays]"
      
      For intCount = 0 To UBound(SpecialDays)
         If Len(SpecialDays(intCount)) Then
            blnSaveItems = True
            Print #intFileWrite, "Day = "; SpecialDays(intCount)
         End If
      Next 'intCount
   Close #intFileWrite
   
   If Not blnSaveItems Then
      Kill DataPath & SPECIAL_DAYS
      
      ReDim SpecialDays(0) As String
   End If
   
   Call LoadSpecialDays
   Call frmCalendar.ResetSpecialDays(True)
   
ExitSub:
   On Local Error GoTo 0
   Close #intFileWrite

End Sub

Private Sub SetDescriptionText()

Dim intPointer As Integer
Dim strDate    As String

   If IsAdd + IsChanged Then Exit Sub
   
   txtDescription.Text = cmbDescription.Text
   chkDay.Item(0).Value = Unchecked
   chkDay.Item(1).Value = Unchecked
   strDate = SpecialDays(frmCalendar.lstSort.ItemData(cmbDescription.ItemData(cmbDescription.ListIndex)))
   intPointer = (InStr(UCase(WEEK_DAYS), UCase(Left(strDate, 2))) + 1) \ 2
   
   If intPointer Then
      intPointer = intPointer - FirstWeekDay - 1 + (7 And (intPointer <= FirstWeekDay))
      lstSelection.Item(2).ListIndex = intPointer
      lstSelection.Item(1).ListIndex = Val(Mid(strDate, 3, 1)) - 1
      chkDay.Item(1).Value = Checked
   End If
   
   If Left(strDate, 1) Like "[+--]" Then
      With lstSelection.Item(0)
         .ListIndex = .ListCount + .List(.ListCount - 1) - Val(strDate) + 1
      End With
      
      chkDay.Item(0).Value = Checked
   End If

End Sub

Private Sub SetDescriptionInfo(ByVal ItemText As String)

Dim intCount As Integer

   With cmbDescription
      If .ListCount > 1 Then .Enabled = True
      
      txtDescription.ToolTipText = GetToolTipText(AppText(235))
      
      For intCount = 0 To .ListCount - 1
         If .List(intCount) = ItemText Then Exit For
      Next 'intCount
      
      If intCount < .ListCount Then .ListIndex = intCount
   End With

End Sub

Private Sub SetDayOfMonth()

Dim intDay As Integer

   intDay = lstSelection.Item(2).ListIndex + calSpecialDays.FirstWeekDay
   
   If intDay > 6 Then intDay = intDay - 7
   
   IsMonthOfDay = Mid(WEEK_DAYS, intDay * 2 + 1, 2) & lstSelection.Item(1).ListIndex + 1 & Format(calSpecialDays.CalMonth, "#00")
   calSpecialDays.CalDay = Day(GetGivenMonthDay(Year(Date), IsMonthOfDay))
   
   Call frmCalendar.SetSpecialDays(calSpecialDays)

End Sub

Private Sub SetDayOfMonthListIndex()

Dim dteDate As Date

   If (IsAdd + IsChanged = False) Or (chkDay.Item(1).Value = Unchecked) Then Exit Sub
   
   With calSpecialDays
      dteDate = DateSerial(.CalYear, .CalMonth, .CalDay)
      lstSelection.Item(2).ListIndex = WeekDay(dteDate, .FirstWeekDay + 1) - 1
      lstSelection.Item(1).ListIndex = (DateDiff("d", DateSerial(.CalYear, .CalMonth, 1), dteDate) \ 7)
   End With

End Sub

Private Sub SetEasterDate()

Dim dteDate As Date

   dteDate = DateAdd("d", CInt(lstSelection.Item(0).Text), Easter)
   
   With calSpecialDays
      .Locked = True
      .CalDay = Day(dteDate)
      .CalMonth = Month(dteDate)
      .Locked = False
      
      Call SetMonth(.CalMonth)
      Call frmCalendar.SetSpecialDays(calSpecialDays)
   End With

End Sub

Private Sub SetEasterListIndex()

Dim intDays As Integer

   If (IsAdd + IsChanged = False) Or (chkDay.Item(0).Value = Unchecked) Then Exit Sub
   
   With calSpecialDays
      intDays = DateDiff("d", Easter, DateSerial(.CalYear, .CalMonth, .CalDay))
   End With
   
   With lstSelection.Item(0)
      .ListIndex = .ListCount + .List(.ListCount - 1) - intDays - 1
   End With

End Sub

Private Sub SetMonth(ByVal IsMonth As Integer)

   With calSpecialDays
      .CalMonth = IsMonth
      lblMonth.Caption = .GetMonthName(IsMonth)
      flbChoose.Item(0).ToolTipText = GetToolTipText(.GetMonthName(IsMonth - 1 + (12 And (IsMonth = 1))))
      flbChoose.Item(1).ToolTipText = GetToolTipText(.GetMonthName(IsMonth + 1 - (12 And (IsMonth = 12))))
   End With

End Sub

Private Sub ToggleButtons()

Dim intCount As Integer

   For intCount = 1 To 2
      chkDay.Item(intCount - 1).Enabled = IsAdd Or IsChanged
      cbtChoose.Item(intCount).ToolTipText = GetToolTipText(AppText(240 + intCount + (2 And (IsAdd Or IsChanged)) - (236 And ((intCount = 1) And (Not IsAdd And Not IsChanged)))))
      cbtChoose.Item(intCount).Picture = imgImages.Item(intCount + (6 And (IsAdd Or IsChanged))).Picture
   Next 'intCount
   
   intCount = 240 + (1 And IsChanged) + (2 And IsAdd)
   
   Call CreateWindow
   Call DrawFooter(Me, AppText(intCount), 14)
   
   txtDescription.ToolTipText = ""
   lstSelection.Item(0).Enabled = False
   lstSelection.Item(1).Enabled = False
   lstSelection.Item(2).Enabled = False
   DoEvents

End Sub

Private Sub calSpecialDays_DateChanged(ButtonID As ButtonTypes)

   Call GetDescription
   Call SetEasterListIndex
   Call SetDayOfMonthListIndex

End Sub

Private Sub chkDay_Click(Index As Integer)

   If IsAdd + IsChanged = False Then
      chkDay.Item(Index).Value = PrevValue
      Exit Sub
   End If
   
   chkDay.Item(Abs(Index = 0)).Value = Unchecked
   lstSelection.Item(0).Enabled = CBool(chkDay.Item(0).Value)
   lstSelection.Item(1).Enabled = chkDay.Item(1).Value
   lstSelection.Item(2).Enabled = chkDay.Item(1).Value
   
   If chkDay.Item(Index).Value = Unchecked Then Exit Sub
   
   If chkDay.Item(0).Value Then
      Call SetEasterDate
      
   ElseIf chkDay.Item(1).Value Then
      Call SetDayOfMonth
   End If

End Sub

Private Sub chkDay_GotFocus(Index As Integer)

   If IsAdd + IsChanged = False Then PrevValue = chkDay.Item(Index).Value

End Sub

Private Sub cmbDescription_Click()

   Call SetDescriptionText

End Sub

Private Sub cmbDescription_Scroll()

   Call SetDescriptionText

End Sub

Private Sub cbtChoose_Click(Index As Integer)

Dim strPrompt As String

   Select Case Index
      Case 0 ' exit
         If (IsAdd + IsChanged) And Len(txtDescription.Text) And AppSettings(SET_ASKCONFIRM) Then
            Call CreateWindow
            Call DrawFooter(Me, AppText(243), 14)
            
            strPrompt = Replace(AppError(29), "$", txtDescription.Text, , 1)
            
            If ShowMessage(strPrompt, vbQuestion, AppError(28), AppError(3), TimeToWait) = vbYes Then Call ItemSave(IsChanged)
            
            DoEvents
         End If
         
         Call EndSpecialDays
         
      Case 1 ' delete or save
         If txtDescription.Text = "" Then Exit Sub
         
         If IsAdd + IsChanged Then
            ' save
            Call ItemSave(IsChanged)
            
         Else
            ' delete
            Call DrawFooter(Me, AppText(5), 14)
            
            If AppSettings(SET_ASKCONFIRM) Then
               strPrompt = Replace(AppError(30), "$", txtDescription.Text, , 1)
               
               If ShowMessage(strPrompt, vbQuestion, AppError(28), AppError(3), TimeToWait) = vbNo Then
                  Call CreateWindow
                  
                  Exit Sub
               End If
            End If
            
            With calSpecialDays
               PrevDate = DateSerial(.CalYear, .CalMonth, .CalDay)
            End With
            
            MousePointer = vbHourglass
            
            Call ToggleButtons
            Call ItemDelete
         End If
         
         Call SetMonth(Month(PrevDate))
         Call frmCalendar.SetSpecialDays(calSpecialDays)
         
         MousePointer = vbDefault
         
      Case 2 ' add or restore
         MousePointer = vbHourglass
         
         If IsAdd + IsChanged Then
            ' restore
            Call CreateWindow
            Call DrawFooter(Me, AppText(244), 14)
            
            DoEvents
            IsAdd = False
            IsChanged = False
            
         Else
            IsAdd = Not IsAdd
            
            If IsAdd Then
               If cmbDescription.ListCount > 14 Then
                  With calSpecialDays
                     strPrompt = FormatDateTime(DateSerial(.CalYear, .CalMonth, .CalDay), vbLongDate)
                     strPrompt = Replace(AppError(32), "#", CapsText(strPrompt))
                  End With
                  
               ElseIf NoSpecialDays Then
                  ReDim SpecialDays(0) As String
                  
               ElseIf UBound(SpecialDays) > 98 Then
                  strPrompt = AppError(33)
               End If
               
               If Len(strPrompt) Then
                  MousePointer = vbDefault
                  ShowMessage strPrompt, vbStop, AppText(18), AppError(31), TimeToWait
                  IsAdd = False
                  Exit Sub
               End If
               
               txtDescription.Text = ""
               txtDescription.SetFocus
               cmbDescription.Enabled = False
               chkDay.Item(0).Value = Unchecked
               chkDay.Item(1).Value = Unchecked
            End If
            
            Call PreviousInfo(True)
         End If
         
         Call ToggleButtons
         Call SetMonth(Month(PrevDate))
         Call frmCalendar.SetSpecialDays(calSpecialDays)
         Call PreviousInfo(False)
         
         MousePointer = vbDefault
   End Select

End Sub

Private Sub flbChoose_Click(Index As Integer)

Dim intMonth As Integer

   intMonth = calSpecialDays.CalMonth
   
   If Index Then
      intMonth = intMonth + 1 - (12 And (intMonth = 12))
      
   Else
      intMonth = intMonth - 1 + (12 And (intMonth = 1))
   End If
   
   Call SetMonth(intMonth)
   
   If (IsAdd + IsChanged) And (chkDay.Item(1).Value = Checked) Then Call SetDayOfMonth
   
   Call frmCalendar.SetSpecialDays(calSpecialDays)
   Call GetDescription
   Call SetEasterListIndex
   Call SetDayOfMonthListIndex

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

   If IsExit(KeyCode, Shift) Then Call EndSpecialDays

End Sub

Private Sub Form_Load()

Dim cbiCombo As ComboBoxInfo
Dim intCount As Integer

   Call SetIcon(imgImages.Item(1), 30)
   Call SetIcon(imgImages.Item(7), 31)
   Call InitForm(Me)
   Call RoundForm(hWnd, ScaleWidth, ScaleHeight, 62)
   Call CenterForm(Me)
   Call ResizeAllControls(Me)
   Call CalculateDays
   
   For intCount = 0 To 2
      If intCount Then
         Load cbtChoose.Item(intCount)
         
         With cbtChoose.Item(intCount)
            .Left = cbtChoose.Item(intCount - 1).Left - .Width + 1
            .Shape = Choose(intCount, ShapeSides, ShapeRight)
            .Visible = True
         End With
      End If
      
      cbtChoose.Item(intCount).ToolTipText = GetToolTipText(AppText(Choose(intCount + 1, 12, 5, 242)))
      cbtChoose.Item(intCount).Picture = imgImages.Item(intCount).Picture
      lstSelection.Item(1).AddItem GetNamePart(AppText(8), intCount + 1)
   Next 'intCount
   
   For intCount = 0 To 1
      If intCount Then Load flbChoose.Item(intCount)
      
      With flbChoose.Item(intCount)
         .Top = lblMonth.Top - 10
         .Visible = True
         .ToolTipText = GetToolTipText(AppText(intCount + 15 + (238 And (intCount > 0))))
      End With
      
      With chkDay.Item(intCount)
         .BackStyle = Transparent
         .Caption = AppText(230 + intCount)
         Set .IconChecked = imgImages.Item(5).Picture
         Set .IconCheckedGrayed = imgImages.Item(6).Picture
      End With
   Next 'intCount
   
   With lblMonth
      .Left = calSpecialDays.Left + (calSpecialDays.Width - .Width) \ 2 - 2
      flbChoose.Item(0).Left = .Left - flbChoose.Item(0).Width - 12
      flbChoose.Item(1).Left = .Left + .Width + 10
   End With
   
   Call SetCalendarImage(hDC, calSpecialDays, frmCalendar.picCalendarImage)
   Call SetCalendarDate(calSpecialDays, CalendarDate)
   
   With picDescription
      cbiCombo.cbSize = Len(cbiCombo)
      GetComboBoxInfo cmbDescription.hWnd, cbiCombo
      .Top = cmbDescription.Top + 1
      .Left = cmbDescription.Left + 1
      .Height = cmbDescription.Height - 2
      .Width = cmbDescription.Width - (cbiCombo.rcButton.Right - cbiCombo.rcButton.Left + 2)
      txtDescription.Height = .ScaleHeight
      txtDescription.Width = .Width
   End With
   
   With calSpecialDays
      intCount = FirstWeekDay + 1
      .CalDay = frmCalendar.calCalendar.CalDay
      .CalMonth = frmCalendar.calCalendar.CalMonth
      .FillExternalLanguage LanguageText
      
      Do
         lstSelection.Item(2).AddItem .GetWeekdayName(intCount)
         intCount = intCount + 1 - (7 And (intCount = 7))
         
         If intCount = FirstWeekDay + 1 Then Exit Do
      Loop
      
      Call SetMonth(.CalMonth)
      
      lblMonth.Caption = .GetMonthName(.CalMonth)
      .ShowToolTipText = AppSettings(SET_SHOWTIPTEXT)
      .FirstWeekDay = FirstWeekDay
      
      Call SetButton(0, 3, Me)
      Call SetButton(1, 4, Me)
      Call .SetMarkColors(Color2:=QBColor(2))
      Call frmCalendar.SetSpecialDays(calSpecialDays)
   End With
   
   For intCount = 1 To 6
      Load shpBorder.Item(intCount)
      
      With shpBorder.Item(intCount)
         If intCount < 3 Then
            .Top = picClock.Item(intCount - 1).Top - 1
            .Left = picClock.Item(intCount - 1).Left - 1
            .Width = picClock.Item(intCount - 1).Width + 2
            .Height = picClock.Item(intCount - 1).Height + 2
            
         ElseIf intCount = 3 Then
            .Top = lblMonth.Top
            .Left = lblMonth.Left
            .Width = lblMonth.Width
            .Height = lblMonth.Height
            
         Else
            Call RemoveListBoxBorder(lstSelection.Item(intCount - 4).hWnd)
            
            .Top = lstSelection.Item(intCount - 4).Top - 1
            .Left = lstSelection.Item(intCount - 4).Left - 1
            .Width = lstSelection.Item(intCount - 4).Width + 2
            .Height = lstSelection.Item(intCount - 4).Height + 2
         End If
         
         .Visible = True
      End With
   Next 'intCount
   
   SetClockBackground hDC, picClock.Item(0)
   ClockY = SetClockBackground(hDC, picClock.Item(1))
   
   Call tmrClock_Timer
   
   tmrClock.Enabled = True
   lstSelection.Item(1).ListIndex = 0
   lstSelection.Item(2).ListIndex = 0
   Easter = DateAdd("d", Abs(lstSelection.Item(0).List(lstSelection.Item(0).ListCount - 1)), "01-01-" & Year(Now))
   
   Call CreateWindow
   Call GetDescription
   Call SetEasterListIndex
   Call SetDayOfMonthListIndex

End Sub

Private Sub lstSelecttion_GotFocus(Index As Integer)

   If Not IsAdd And Not IsChanged Then calSpecialDays.SetFocus

End Sub

Private Sub lstSelecttion_Scroll(Index As Integer)

   Select Case Index
      Case 0
         lstSelection.Item(0).ListIndex = lstSelection.Item(0).TopIndex
         
         Call SetEasterDate
         
      Case 1
         lstSelection.Item(1).ListIndex = lstSelection.Item(1).TopIndex
         
         Call SetDayOfMonth
         
      Case 2
         lstSelection.Item(2).ListIndex = lstSelection.Item(2).TopIndex
         
         Call SetDayOfMonth
   End Select

End Sub

Private Sub tmrClock_Timer()

   Call DrawDateTime(picClock.Item(0), ClockY, Now, DefaultDateFormat)
   Call DrawDateTime(picClock.Item(1), ClockY, Now, "hh:mm:ss")

End Sub

Private Sub txtDescription_Change()

Dim intPointer As Integer

   With lblText
      .Caption = txtDescription.Text
      intPointer = txtDescription.SelStart
      
      Do
         If .Width <= DescriptionWidth Then Exit Do
         
         .Caption = Left(.Caption, Len(.Caption) - 1)
      Loop
      
      txtDescription.Text = .Caption
      txtDescription.SelStart = intPointer
   End With

End Sub

Private Sub txtDescription_Click()

   With txtDescription
      If Not IsAdd Then
         If txtDescription.Text = "" Then Exit Sub
         
         IsChanged = True
         
         Call PreviousInfo(True)
      End If
      
      .SelStart = 0
      .SelLength = Len(.Text)
      cmbDescription.Enabled = False
   End With
   
   Call ToggleButtons
   
   If chkDay.Item(0).Value = Checked Then lstSelection.Item(0).Enabled = True
   
   If chkDay.Item(1).Value = Checked Then
      lstSelection.Item(1).Enabled = True
      lstSelection.Item(2).Enabled = True
   End If

End Sub

Private Sub txtDescription_GotFocus()

   If Not IsAdd And (txtDescription.Text = "") Then calSpecialDays.SetFocus

End Sub
