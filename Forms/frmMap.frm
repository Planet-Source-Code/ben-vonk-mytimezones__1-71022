VERSION 5.00
Begin VB.Form frmMap 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   6972
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9768
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00800000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "frmMap.frx":0000
   ScaleHeight     =   581
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   814
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picCls 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F7EFE2&
      BorderStyle     =   0  'None
      Height          =   492
      Left            =   8760
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1080
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.PictureBox picClock 
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
      Left            =   7140
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   276
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox picClock 
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
   Begin VB.Timer tmrClock 
      Enabled         =   0   'False
      Interval        =   333
      Left            =   9240
      Top             =   600
   End
   Begin VB.PictureBox picMode 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F7EFE2&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Left            =   120
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   696
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5724
      Width           =   8352
   End
   Begin VB.PictureBox picMap 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4860
      Left            =   120
      MouseIcon       =   "frmMap.frx":E1A5
      Picture         =   "frmMap.frx":E4AF
      ScaleHeight     =   405
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   696
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   864
      Width           =   8352
      Begin VB.Shape shpZone 
         BorderColor     =   &H00400000&
         BorderStyle     =   3  'Dot
         DrawMode        =   12  'Nop
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   4860
         Left            =   4008
         Top             =   0
         Visible         =   0   'False
         Width           =   348
      End
      Begin VB.Line linClock 
         BorderColor     =   &H00FFFF00&
         BorderStyle     =   3  'Dot
         X1              =   348
         X2              =   348
         Y1              =   0
         Y2              =   410
      End
   End
   Begin MyTimeZones.CrystalButton cbtChoose 
      Height          =   504
      Index           =   0
      Left            =   7740
      TabIndex        =   5
      Top             =   6288
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
      Left            =   228
      Top             =   252
      Width           =   1224
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   2
      Left            =   8760
      Picture         =   "frmMap.frx":45A66
      Top             =   600
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   1
      Left            =   9240
      Picture         =   "frmMap.frx":46330
      Top             =   120
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   0
      Left            =   8760
      Top             =   120
      Visible         =   0   'False
      Width           =   384
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Private Variables
Private ClockDate       As Date
Private ClockY          As Integer
Private ZonePosition    As Integer
Private HourInterval    As Single
Private TimeInterval    As Single
Private TimeStep        As Single
Private ZoneStep        As Single
Private ModeToolTipText As String

Public Sub CreateWindow()

   InZoneMap = True
   
   Call DrawHeader(Me)
   Call ClockOrZone

End Sub

Private Sub ClockOrZone()

   shpZone.Visible = Not AppSettings(SET_MODEZONEMAP)
   
   If ShowSettings Then
      ModeToolTipText = GetToolTipText(AppText(2) & " " & GetTimeZoneText(SelectedTimeZoneID))
      
   Else
      ModeToolTipText = GetToolTipText(StrConv(AppText(1), vbProperCase))
   End If
   
   With linClock
      If AppSettings(SET_MODEZONEMAP) Then
         .BorderColor = &HC01FC0
         .BorderStyle = vbBSSolid
         .BorderWidth = 2
         
      Else
         .BorderColor = &HFFFF00
         .BorderStyle = vbBSDot
         .BorderWidth = 1
      End If
   End With
   
   With picMap
      ZonePosition = (.ScaleWidth \ 2) - (AllZones(SelectedTimeZoneID).Bias / 60) * ZoneStep - ZoneStep \ 2
      
      If ZonePosition >= .ScaleWidth Then ZonePosition = ZonePosition - .ScaleWidth
   End With
   
   If AppSettings(SET_MODEZONEMAP) Then
      Call ShowClock
      Call tmrClock_Timer
      
   Else
      Call MapPosition(0)
      Call SetMarks(ZonePosition)
      Call ShowZone
   End If

End Sub

Private Sub EndMap()

   MousePointer = vbHourglass
   
   Call frmMyTimeZones.ToggleControls(True)
   Call ResetComboBox
   
   If InZoneInfo Then frmInfo.picClock.Visible = True
   
   Hide
   DoEvents
   Unload Me
   InZoneMap = False
   Set frmMap = Nothing

End Sub

Private Sub MapPosition(ByVal X As Integer)

Dim sngX(1) As Single

   With picMap
      .PaintPicture .Picture, X, 0, .ScaleWidth, .ScaleHeight
      
      If X Then
         If X < 0 Then
            X = X - 1
            sngX(0) = .ScaleWidth + X
            
         Else
            X = X + 1
            sngX(1) = .ScaleWidth - X
         End If
         
         .PaintPicture .Picture, sngX(0), 0, Abs(X), .ScaleHeight, sngX(1), 0, Abs(X), .ScaleHeight
      End If
   End With
   
   Erase sngX

End Sub

Private Sub PrintClock()

Dim intCurrHour  As Integer
Dim intPrevHour  As Integer
Dim lngHourColor As Long

   intCurrHour = Hour(ClockDate) + (1 And (Minute(ClockDate) > 29))
   intCurrHour = intCurrHour - (24 And (intCurrHour = 24))
   intPrevHour = intCurrHour - 1 + (24 And (intCurrHour = 0))
   linClock.X1 = (Hour(ClockDate) * 60 + Minute(ClockDate)) / HourInterval - (1 And (intCurrHour + Minute(ClockDate) = 59))
   linClock.X2 = linClock.X1
   
   Call MapPosition(Int(linClock.X1 - ZonePosition - shpZone.Width / 2))
   
   With picMode
      Call PrintLine(0, 2, TimeInterval * Minute(ClockDate), 2, , &HC01FC0)
      Call PrintLine(0, 4, TimeInterval * Second(ClockDate), 4, , &HFF&)
      
      If Minute(ClockDate) = 0 Then Call PrintLine(0, 2, .ScaleWidth, 2, , &HEAD199)
      If Second(ClockDate) = 0 Then Call PrintLine(0, 4, .ScaleWidth, 4, , &HEAD199)
   End With
   
   If (intCurrHour = 0) Or (intPrevHour = 0) Then
      If intCurrHour Then
         lngHourColor = &HD94600
         
      Else
         lngHourColor = &HFF&
      End If
      
      Call PrintHour(24, lngHourColor)
   End If
   
   Call PrintHour(Format(intPrevHour, "0#"))
   Call PrintHour(Format(intCurrHour, "0#"), &HFF&)

End Sub

Private Sub PrintHour(ByVal Hour As String, Optional ByVal Color As Long = &HD94600)

Dim intX    As Integer
Dim strText As String

   With picMode
      If CInt(Hour) < 24 Then
         strText = Hour
         
      Else
         strText = "00"
      End If
      
      intX = CInt(Hour) * TimeStep - .TextWidth(strText) \ 2
      BitBlt .hDC, intX, 16, .TextWidth(strText), .TextHeight("X"), picCls.hDC, 0, 0, vbSrcCopy
      .ForeColor = Color
      .CurrentY = 16
      .CurrentX = intX
      picMode.Print strText
   End With

End Sub

Private Sub PrintLine(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, Optional ByVal Width As Integer = 2, Optional ByVal Color As Long = &HD94600)

   picMode.DrawWidth = Width
   picMode.Line (X1, Y1)-(X2, Y2), Color, BF

End Sub

Private Sub SetMarks(ByVal X As Integer)

   With linClock
      .X1 = X + (shpZone.Width \ 2 And Not AppSettings(SET_MODEZONEMAP))
      .X2 = .X1
      shpZone.Left = X
   End With

End Sub

Private Sub SetToolTipText()

   cbtChoose.Item(0).ToolTipText = GetToolTipText(AppText(12))
   cbtChoose.Item(1).ToolTipText = GetToolTipText(AppText(171 + Abs(AppSettings(SET_MODEZONEMAP))))

End Sub

Private Sub ShowClock()

Dim intCount As Integer
Dim sngHalf  As Single
Dim sngHour  As Single
Dim sngQuart As Single

   With picMode
      sngHalf = TimeStep \ 2
      sngQuart = TimeStep \ 4
      .Cls
      
      Call PrintLine(0, 3, .ScaleWidth, 3, 8)
      Call PrintLine(0, 3, .ScaleWidth, 3, 4, &HEAD199)
      Call PrintLine(0, 7, 0, 14, 3)
      Call PrintLine(sngQuart + 1, 7, sngQuart + 1, 9)
      Call PrintHour("00")
      Call PrintHour(24)
      
      For intCount = 1 To 23
         sngHour = intCount * TimeStep
         
         Call PrintLine(sngHour, 7, sngHour, 14, 3)
         Call PrintLine(sngHour + sngQuart + 1, 7, sngHour + sngQuart + 1, 9)
         Call PrintLine(sngHour - sngQuart, 7, sngHour - sngQuart, 9)
         Call PrintLine(sngHour - sngHalf, 7, sngHour - sngHalf, 12)
         Call PrintHour(Format(intCount, "0#"))
      Next 'intCount
      
      sngHour = intCount * TimeStep
      
      Call PrintLine(.ScaleWidth, 7, .ScaleWidth, 14, 3)
      Call PrintLine(sngHour - sngQuart, 7, sngHour - sngQuart, 9)
      Call PrintLine(sngHour - sngHalf, 7, sngHour - sngHalf, 12)
   End With

End Sub

Private Sub ShowZone()

Dim intCount    As Integer
Dim intPosition As Integer
Dim sngCurrX    As Single
Dim sngCurrY    As Single
Dim sngHalf     As Single
Dim strZone     As String

   With picMode
      sngHalf = ZoneStep \ 2
      .Cls
      
      Call PrintLine(0, 2, .ScaleWidth, 2, 4)
      
      For intCount = -12 To 13
         If intCount < 13 Then
            intPosition = intCount + 12
            strZone = Format(Abs(intCount), "0#")
            
            If (intCount = -12) And TimeZone13 Then strZone = strZone & "/13"
            
            If intCount Then
               .ForeColor = &HD94600
               
            Else
               .ForeColor = &HFF&
            End If
            
            .CurrentY = 13
            .CurrentX = intPosition * ZoneStep + sngHalf - .TextWidth(strZone) \ 2 + (1 And (intCount = 0))
            picMode.Print strZone
         End If
         
         Call PrintLine((intCount + 12) * ZoneStep, 4, (intCount + 12) * ZoneStep, 8)
         Call PrintLine((intCount + 12) * ZoneStep - sngHalf, 4, (intCount + 12) * ZoneStep - sngHalf, 10, 3)
      Next 'intCount
      
      sngCurrY = 13 + .TextHeight("X") \ 2 - 1
      sngCurrX = .ScaleWidth \ 2 - sngHalf
      
      For intCount = 0 To 2
         .CurrentY = sngCurrY - (3 And (intCount = 2))
         .CurrentX = sngCurrX + (2 * sngHalf And (intCount > 0)) - (4 And (intCount < 2))
         
         Call PrintLine(.CurrentX - (1 And (intCount = 2)), .CurrentY, .CurrentX + (7 And (intCount < 2)), .CurrentY + 1 + (6 And (intCount = 2)), 1, &HC01FC0)
      Next 'intCount
   End With

End Sub

Private Sub cbtChoose_Click(Index As Integer)

   If Index Then
      AppSettings(SET_MODEZONEMAP) = Not AppSettings(SET_MODEZONEMAP)
      cbtChoose.Item(1).Picture = imgImages.Item(1 + (1 And AppSettings(SET_MODEZONEMAP))).Picture
      
      Call SetToolTipText
      Call ClockOrZone
      
   Else
      Call EndMap
   End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

   If IsExit(KeyCode, Shift) Then Call EndMap

End Sub

Private Sub Form_Load()

Dim intChoose As Integer
Dim intCount  As Integer
Dim strText   As String

   Call InitForm(Me, -1, 8592, 6972)
   Call RoundForm(hWnd, ScaleWidth, ScaleHeight, 62)
   Call CenterForm(Me)
   Call ResizeControl(cbtChoose.Item(0))
   Call ResizeControl(picMode)
   
   With picMap
      .Top = .Top * ScreenResize
      .Left = .Left * ScreenResize
      shpZone.Height = .ScaleHeight
      linClock.Y2 = .ScaleHeight
      HourInterval = 1440 / .ScaleWidth
   End With
   
   With picMode
      .Top = picMap.Top + picMap.Height
      TimeInterval = .ScaleWidth / 59
      TimeStep = .ScaleWidth / 24
      ZoneStep = .ScaleWidth / 25
   End With
   
   For intCount = 0 To 1
      Call ResizeControl(picClock.Item(intCount))
      
      If intCount Then
         intChoose = 1 + Abs(AppSettings(SET_MODEZONEMAP))
         Load cbtChoose.Item(intCount)
         Load shpBorder.Item(intCount)
         
         With cbtChoose.Item(intCount)
            .Left = cbtChoose.Item(intCount - 1).Left - .Width + 1
            .Shape = ShapeRight
            .Visible = True
         End With
         
      Else
         intChoose = intCount
      End If
      
      With shpBorder.Item(intCount)
         .Top = picClock.Item(intCount).Top - 1
         .Left = picClock.Item(intCount).Left - 1
         .Width = picClock.Item(intCount).Width + 2
         .Height = picClock.Item(intCount).Height + 2
         .Visible = True
      End With
      
      cbtChoose.Item(intCount).Picture = imgImages.Item(intChoose).Picture
      ClockY = SetClockBackground(hDC, picClock.Item(intCount))
   Next 'intCount
   
   tmrClock.Enabled = True
   
   Call SetToolTipText
   Call tmrClock_Timer
   Call SetComboBox(hWnd, cbtChoose.Item(0))
   Call CreateWindow
   
   If ShowSettings Then
      If SelectedClock > 1 Then
         strText = frmMyTimeZones.clkFavorits.Item(SelectedFavorit).NameClock
         
      Else
         strText = Left(AppText(2), Len(AppText(2)) - 1)
      End If
      
      Call DrawFooter(Me, strText, 14)
   End If

End Sub

Private Sub picMap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Static intPointer  As Integer
Static intPrevBias As Integer
Static sngX        As Single

Dim intNewBias     As Integer

   With shpZone
      If AppSettings(SET_MODEZONEMAP) Then
         picMap.ToolTipText = ""
         picMap.MousePointer = vbDefault
         Exit Sub
         
      ElseIf (X >= .Left) And (X <= .Left + .Width + 2) Then
         picMap.ToolTipText = ModeToolTipText
         
         If Not ShowSettings Then picMap.MousePointer = vbCustom
         
      Else
         picMap.ToolTipText = ""
         picMap.MousePointer = vbDefault
      End If
      
      If ((X < -15) Or (X > picMap.ScaleWidth - 1)) Or (Button <> vbLeftButton) Or ShowSettings Then Exit Sub
      If picMap.MousePointer = vbCustom Then Call SetMarks(X - .Width / 2)
   End With
   
   intNewBias = Int(((X / ZoneStep - 12.5) * -60) / 15) * 15
   
   If sngX <> X Then
      intPointer = intPointer - (1 And (sngX < X)) + (1 And (sngX > X))
      sngX = X
      
      If (intPointer < 0) Or (intPointer > UBound(AllZones)) Then
         intPrevBias = -1
         sngX = -1
      End If
   End If
   
   If intPrevBias <> intNewBias Then
      intPrevBias = intNewBias
      intPointer = 0
   End If
   
   For intPointer = intPointer To UBound(AllZones)
      If AllZones(intPointer).Bias = intPrevBias Then
         With frmMyTimeZones.cmbTimeZones
            For intNewBias = 0 To .ListCount - 1
               If .ItemData(intNewBias) = intPointer Then
                  .ListIndex = intNewBias
                  Exit For
               End If
            Next 'intNewBias
         End With
         
         Exit For
      End If
   Next 'intPointer
   
   Call SetMarks(ZonePosition)

End Sub

Private Sub picMode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim intMid  As Integer
Dim sngGMT  As Single
Dim sngZone As Single
Dim sngRest As Single

   If AppSettings(SET_MODEZONEMAP) Then Exit Sub
   
   With picMode
      intMid = .ScaleWidth / 2
      sngZone = Int((X / ZoneStep - 12 + (5 And (X < intMid)) / 10) * 100) / 100 - (5 And (X > intMid)) / 10
      sngGMT = Int(sngZone)
      sngRest = sngZone - sngGMT
      
      If X > intMid Then
         If (sngRest > 0.12) And (sngRest < 0.37) Then sngGMT = sngGMT + 0.15
         If (sngRest > 0.36) And (sngRest < 0.63) Then sngGMT = sngGMT + 0.3
         If (sngRest > 0.62) And (sngRest < 0.87) Then sngGMT = sngGMT + 0.45
         If (sngRest > 0.86) And (sngRest < 0.99) Then sngGMT = sngGMT + 1
         
      ElseIf X < intMid Then
         If (sngRest > 0.62) And (sngRest < 0.86) Then sngGMT = sngGMT - 0.15
         If (sngRest > 0.36) And (sngRest < 0.63) Then sngGMT = sngGMT - 0.3
         If (sngRest > 0.12) And (sngRest < 0.37) Then sngGMT = sngGMT - 0.45
         If (sngRest > 0.01) And (sngRest < 0.13) Then sngGMT = sngGMT - 1
      End If
      
      .ToolTipText = GetToolTipText("GMT " & Replace(Format(sngGMT, "#00.00"), Mid(Format(0.1, "0.0"), 2, 1), ":"))
   End With

End Sub

Private Sub tmrClock_Timer()

   ClockDate = UTCToLocalDate(GetSystemDate, AllZones(SelectedTimeZoneID))
   
   Call DrawDateTime(picClock.Item(0), ClockY, ClockDate, DefaultDateFormat)
   Call DrawDateTime(picClock.Item(1), ClockY, ClockDate, "hh:mm:ss")
   
   If AppSettings(SET_MODEZONEMAP) Then
      With picMode
         .ToolTipText = Format(ClockDate, LongDateFormat & "  -  hh:mm:ss")
         .ToolTipText = GetToolTipText(CapsText(.ToolTipText))
      End With
      
      Call PrintClock
   End If

End Sub
