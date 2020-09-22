Attribute VB_Name = "mdlLayout"
Option Explicit

' Private API's
Private Declare Function CreateRectRgn Lib "GDI32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetParent Lib "User32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

' returns top of the text
Public Function SetClockBackground(ByVal hDC As Long, ByRef Box As PictureBox) As Long

   With Box
      BitBlt .hDC, 0, 0, .ScaleWidth, .ScaleHeight, hDC, .Left, .Top, vbSrcCopy
      .Picture = .Image
      .Visible = True
      SetClockBackground = (.ScaleHeight - .TextHeight("X")) / 2
   End With

End Function

' function for calling the messageform
Public Function ShowMessage(ByVal Prompt As String, ByVal MessageType As VbMsgBoxStyle, ByVal Title As String, ByVal Message As String, Optional ByVal Wait As Integer, Optional ByVal DefaultKey As VbMsgBoxStyle = vbNo) As VbMsgBoxResult

   Load frmMessage
   ShowMessage = frmMessage.ShowMessage(Prompt, MessageType, Title, Message, Wait, DefaultKey)

End Function

' handle time difference between dates
Public Function ShowTimeToGo(ByRef Display As Object) As Boolean

Dim blnHold    As Boolean
Dim dteTime(1) As Date
Dim dblTime    As Double
Dim intWeekDay As Integer
Dim strBuffer  As String

   strBuffer = Format(Date, "yyyymmdd")
   
   If (TimeToGoShowType = TotalTime) Or ((TimeToGoShowType = PassedTime) And (Format(TimeToGo(0), "yyyymmdd") > strBuffer)) Or (((TimeToGoShowType = 2) And (Format(TimeToGo(1), "yyyymmdd") < strBuffer))) Then
      dteTime(0) = CDate(TimeToGo(0)) + "00:00:00"
      dteTime(1) = CDate(TimeToGo(1)) + "00:00:00"
      ShowTimeToGo = Not ((Format(TimeToGo(0), "yyyymmdd") <= strBuffer) And (Format(TimeToGo(1), "yyyymmdd") <= strBuffer))
      blnHold = True
      
   ElseIf (TimeToGoShowType = PassedTime) Then
      If Format(TimeToGo(1), "yyyymmdd") > strBuffer Then
         dteTime(0) = CDate(TimeToGo(0)) + "00:00:00"
         dteTime(1) = Now
         
      ElseIf Format(TimeToGo(1), "yyyymmdd") < strBuffer Then
         dteTime(0) = CDate(TimeToGo(1)) + "23:59:59"
         dteTime(1) = Now
         dblTime = 1
      End If
      
      ShowTimeToGo = True
      
   ' RemainingTime
   Else
      If Format(Date, "yyyymmdd") = Format(TimeToGo(1), "yyyymmdd") Then
         dteTime(0) = Date
         
      Else
         dteTime(0) = Now
      End If
      
      dteTime(1) = CDate(TimeToGo(1))
      ShowTimeToGo = True
   End If
   
   If TimeToGoShow > -1 Then
      strBuffer = Choose(TimeToGoShow + 1, "s", "n", "h", "d", "ww", "m", "q", "yyyy")
      intWeekDay = WeekDay(dteTime(0))
      dblTime = dblTime + DateDiff(strBuffer, dteTime(0), dteTime(1), intWeekDay, vbFirstFourDays)
      
      If Not blnHold And (dteTime(0) > dteTime(1)) Then dblTime = dblTime - DateDiff(strBuffer, dteTime(0), Now, intWeekDay, vbFirstFourDays)
   End If
   
   Display.Text = Format(Abs(dblTime), "#" & String(13, vbKey0))
   Erase dteTime

End Function

' center form in the screen
Public Sub CenterForm(ByRef Window As Form)

   With frmMyTimeZones
      Window.Top = .Top + (.Height - Window.Height) \ 2
      Window.Left = .Left + (.Width - Window.Width) \ 2
   End With

End Sub

' show globe position
Public Sub DoAnimation(ByRef Window As Form)

Static intAnimation As Integer

   With Window
      BitBlt .hDC, GlobeXY.X, GlobeXY.Y, 62, 62, AnimationhDC, 1221, 0, vbSrcAnd
      BitBlt .hDC, GlobeXY.X, GlobeXY.Y, 62, 62, AnimationhDC, intAnimation * 61, 0, vbSrcPaint
      .Refresh
      intAnimation = intAnimation + 1 - (20 And (intAnimation > 18))
   End With

End Sub

' print date or time in specified picturebox
Public Sub DrawDateTime(ByRef Box As Object, ByVal Top As Long, ByVal ClockDate As Date, ByVal FormatType As String)

   With Box
      .Cls
      .CurrentY = Top
      .CurrentX = (.ScaleWidth - .TextWidth(Format(ClockDate, FormatType))) \ 2
      Box.Print Format(ClockDate, FormatType)
   End With

End Sub

' print copyright notice
Public Sub DrawCopyright(ByRef Window As Form, ByVal Y As Single)

Dim intCount As Integer

   For intCount = 1 To 0 Step -1
      Call DrawText(Window, InfoCopyright, 10, &HBBAC99, intCount, Y)
      Call DrawText(Window, AppText(25), 6, &HBBAC99, intCount, Y + 18)
   Next 'intCount

End Sub

' print window footer
Public Sub DrawFooter(ByRef Window As Form, ByVal Text As String, ByVal Size As Integer)

Dim sngSize As Single
Dim sngX    As Single
Dim sngY    As Single

   With Window
      sngSize = .FontSize
      .FontSize = Size * ScreenResize
      sngX = (.ScaleWidth - .TextWidth(Text)) \ 2
      sngY = .cbtChoose.Item(0).Top + (.cbtChoose.Item(0).Height - .TextHeight("X")) \ 2
      
      Call DrawText(Window, Text, Size, &HBBAC99, 1, sngY, sngX)
      Call DrawText(Window, Text, Size, &HBBAC99, 0, sngY, sngX)
      
      .FontSize = sngSize
   End With

End Sub

' print window header text
Public Sub DrawHeader(ByRef Window As Form, Optional ByVal Header As String, Optional ByVal DrawOnlyHeader As Boolean)

Dim blnIsGMT       As Boolean
Dim intCount       As Integer
Dim intPosition    As Integer
Dim sngFontSize    As Single
Dim strDisplayName As String

   With Window
      .Cls
      sngFontSize = .FontSize
      
      If Not DrawOnlyHeader Then .FontSize = 9 + AddToFontSize
      
      For intPosition = 1 To 0 Step -1
         If intPosition Then
            .ForeColor = vbWhite
            
         Else
            .ForeColor = &HC01FC0
         End If
         
         If DrawOnlyHeader Then
            Call DrawText(Window, Header, 16, &HC01FC0, intPosition, 20)
            
         Else
            For intCount = 0 To 1
               strDisplayName = GetTimeZoneText(SelectedTimeZoneID, CBool(intCount))
               
               If blnIsGMT Then
                  If InStr(strDisplayName, ";") Then
                     strDisplayName = Trim(Split(strDisplayName, ";", 2)(1))
                     
                  ElseIf InStr(strDisplayName, ":") Then
                     strDisplayName = Trim(Split(strDisplayName, ":", 2)(1))
                  End If
               End If
               
               If intCount = 0 Then
                  blnIsGMT = strDisplayName = "GMT"
                  
                  If Len(Header) Then strDisplayName = Header & strDisplayName
               End If
               
               .CurrentY = intPosition + 17 + (21 And (intCount > 0))
               .CurrentX = (.ScaleWidth - .TextWidth(strDisplayName)) \ 2 + intPosition
               Window.Print strDisplayName
               
               If intPosition Then
                  .ForeColor = &HFCF9F2
                  
               Else
                  .ForeColor = &HD94600
               End If
            Next 'intCount
         End If
      Next 'intPosition
      
      .FontSize = sngFontSize
      frmMyTimeZones.cmbTimeZones.ListIndex = SelectedListIndex
   End With

End Sub

' print the given text on the window
Public Sub DrawText(ByRef Window As Object, ByVal Text As String, ByVal Size As Integer, ByVal Color As Long, ByVal Position As Integer, ByVal Y As Single, Optional ByVal X As Single)

Dim sngFontSize As Single

   With Window
      sngFontSize = .FontSize
      .FontSize = Size + AddToFontSize
      
      If Position = 0 Then
         .ForeColor = Color
         
      Else
         .ForeColor = vbWhite
      End If
      
      If X Then
         .CurrentX = X
         
      Else
         .CurrentX = (.ScaleWidth - .TextWidth(Text)) \ 2 + Position
      End If
      
      .CurrentY = Y + Position
      Window.Print Text
      .FontSize = sngFontSize
   End With

End Sub

Public Sub InitForm(ByRef Window As Form, Optional ByVal SetPicture As Integer = 2, Optional ByVal IsWidth As Long = 8220, Optional ByVal IsHeight As Long = 6060, Optional CopyImages As Boolean = True)

   With Window
      .Height = IsHeight * ScreenResize
      .Width = IsWidth * ScreenResize
      .FontSize = .FontSize + AddToFontSize
      
      If CopyImages Then Call SetIcon(.imgImages.Item(0), 29)
      
      If SetPicture > -1 Then
         .Picture = frmMyTimeZones.picChild.Item(SetPicture - (2 And (SetPicture = 2))).Picture
         
         If SetPicture = 2 Then
            Call SetBackgroundImage(.hDC, .ScaleHeight - 20, .ScaleWidth, frmMyTimeZones.picMasker.Item(0))
            
            .Picture = .Image
         End If
      End If
   End With

End Sub

' removes the border from an control
Public Sub RemoveListBoxBorder(ByVal hWnd As Long)

Dim lngParentWnd       As Long
Dim ptaClient          As PointAPI
Dim rctClient          As Rect

   SetWindowLong hWnd, GWL_STYLE, GetWindowLong(hWnd, GWL_STYLE) And Not WS_BORDER
   SetWindowLong hWnd, GWL_EXSTYLE, GetWindowLong(hWnd, GWL_EXSTYLE) And Not WS_EX_CLIENTEDGE
   SetWindowPos hWnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
   lngParentWnd = GetParent(hWnd)
   GetWindowRect hWnd, rctClient
   
   With ptaClient
      .X = rctClient.Left
      .Y = rctClient.Top
      ScreenToClient lngParentWnd, ptaClient
      rctClient.Left = .X
      rctClient.Top = .Y
      .X = rctClient.Right
      .Y = rctClient.Bottom
      ScreenToClient lngParentWnd, ptaClient
      rctClient.Right = .X
      rctClient.Bottom = .Y
      InflateRect rctClient, 1, 1
   End With

End Sub

' open the message edit form
Public Sub OpenAlarmMessage(ByRef Window As Form)

   Window.MousePointer = vbHourglass
   frmEditMessage.Show vbModal, Window
   Window.MousePointer = vbDefault

End Sub

' open the clock image form
Public Sub OpenClockImage(ByRef Window As Form)

Dim strFileName As String

   If SelectedClock > 1 Then
      strFileName = FavoritsInfo(SelectedFavorit).ImageFile
      
   Else
      strFileName = ZonesInfo(SelectedClock).ImageFile
   End If
   
   Window.MousePointer = vbHourglass
   
   Call frmImage.SetImageFile(strFileName)
   
   frmImage.Show vbModal, Window
   Window.MousePointer = vbDefault
   
   If strFileName <> ImageName Then Call SetClockImage(SelectedClock, ImageName)

End Sub

' set the timezones combobox to the original form frmMyTimeZones
Public Sub ResetComboBox(Optional ByVal State As Boolean = True)

Dim intLeft As Integer
Dim intTop  As Integer

   With frmMyTimeZones
      With .clkTimeZone
         intTop = .Item(0).Top + .Item(0).Height
      End With
      
      With .flbChoose
         intTop = intTop + (.Item(0).Top - intTop - frmMyTimeZones.cmbTimeZones.Height) \ 2 - 3
         intLeft = (.Item(0).Left + .Item(10).Left + .Item(10).Width) \ 2
      End With
   
      With .cmbTimeZones
         If State Then SetParent .hWnd, frmMyTimeZones.hWnd
         
         .Top = intTop
         .Left = intLeft - .Width \ 2
         .ListIndex = SelectedListIndex
         .BackColor = &HE8EAED
      End With
      
      .tcbSkinner.ComboBoxBorderColor = &HD1D3D4
   End With

End Sub

Public Sub ResizeAllControls(ByRef Window As Form, Optional ByVal NoPictureBox As Boolean)

Dim ctlControl As Control

   For Each ctlControl In Window.Controls
      If TypeOf ctlControl Is ThemedScrollBar Then
         ' do nothing
         
      ElseIf NoPictureBox Then
         If Not TypeOf ctlControl Is PictureBox Then Call ResizeControl(ctlControl)
         
      Else
         Call ResizeControl(ctlControl)
      End If
   Next 'ctlControl

End Sub

' resize and relocate control
Public Sub ResizeControl(ByRef Control As Object)

   On Local Error Resume Next
   
   With Control
      .Font.Size = .Font.Size + AddToFontSize
      .Top = .Top * ScreenResize
      .Left = .Left * ScreenResize
      .Width = .Width * ScreenResize
      .Height = .Height * ScreenResize
   End With
   
   On Local Error GoTo 0

End Sub

' make rounded form corners
Public Sub RoundForm(ByVal hWnd As Long, ByVal Width As Long, ByVal Height As Long, ByVal Curve As Integer, Optional ByVal OnlyTop As Boolean)

Const RGN_OR     As Long = 2

Dim lngRegion(1) As Long

   lngRegion(0) = CreateRoundRectRgn(0, 0, Width + 1, Height + 1, Curve, Curve)
   
   If OnlyTop Then
      lngRegion(1) = CreateRectRgn(0, Curve \ 2 + 1, Width, Height)
      CombineRgn lngRegion(0), lngRegion(0), lngRegion(1), RGN_OR
      DeleteObject lngRegion(1)
   End If
   
   SetWindowRgn hWnd, lngRegion(0), True
   DeleteObject lngRegion(0)
   Erase lngRegion

End Sub

' create the backcolor for the specified object
Public Sub SetBackColor(ByRef Box As Object, ControlBackColor As Object)

   With ControlBackColor
      .BackColor = Box.Point(.Left, .Top + .Height \ 2)
   End With

End Sub

' paints an image transparent on the background
Public Sub SetBackgroundImage(ByVal hDC, ByVal ScaleHeight As Long, ByVal ScaleWidth As Long, ByRef ImageBox As PictureBox)

Dim lngMaskBmp     As Long
Dim lngMaskDC      As Long
Dim lngMaskPrevBmp As Long
Dim lngLeft        As Long
Dim lngTop         As Long
Dim picMasker      As PictureBox

   With ImageBox
      Set picMasker = frmMyTimeZones.picMasker.Item(1)
      .BackColor = .Point(0, 0)
      picMasker.ForeColor = .BackColor
      picMasker.Width = .Width
      picMasker.Height = .Height
      lngTop = (ScaleHeight - .ScaleHeight) / 2 + 10
      lngLeft = (ScaleWidth - .ScaleWidth) / 2
      
      With picMasker
         lngMaskDC = CreateCompatibleDC(.hDC)
         lngMaskBmp = CreateBitmap(.ScaleWidth, .ScaleHeight, 1, 1, ByVal 0&)
         lngMaskPrevBmp = SelectObject(lngMaskDC, lngMaskBmp)
         BitBlt lngMaskDC, 0, 0, .ScaleWidth, .ScaleHeight, ImageBox.hDC, 0, 0, vbSrcCopy
         BitBlt .hDC, 0, 0, .ScaleWidth, .ScaleHeight, lngMaskDC, 0, 0, vbSrcCopy
         SelectObject lngMaskDC, lngMaskPrevBmp
         DeleteObject lngMaskBmp
         DeleteDC lngMaskDC
      End With
      
      StretchBlt hDC, lngLeft, lngTop, .ScaleWidth, .ScaleHeight, picMasker.hDC, 0, 0, .ScaleWidth, .ScaleHeight, vbSrcPaint
      StretchBlt hDC, lngLeft, lngTop, .ScaleWidth, .ScaleHeight, .hDC, 0, 0, .ScaleWidth, .ScaleHeight, vbSrcAnd
      Set picMasker = Nothing
   End With

End Sub

' create the flbChoose button on the specified window
Public Sub SetButton(ByVal Index As Integer, ByVal Icon As Integer, ByRef Window As Form)

Dim lngX As Long
Dim lngY As Long

   With Window
      lngX = 5 + (4 And (.Name = frmMyTimeZones.Name))
      lngY = 5 + (42 And (.Name = frmMyTimeZones.Name))
   End With
   
   With Window.flbChoose.Item(Index)
      .IconX = lngX
      .IconY = lngY
      .Icon = Window.imgImages.Item(Icon).Picture
   End With

End Sub

' fill the image of the specified calendar
Public Sub SetCalendarImage(ByVal hDC As Long, ByRef Calendar As Object, ByRef Box As PictureBox)

   With Box
      .Height = Calendar.Height
      .Width = Calendar.Width
      BitBlt .hDC, 0, 0, .ScaleWidth, .ScaleHeight, hDC, Calendar.Left, Calendar.Top, vbSrcCopy
      Calendar.Picture = .Image
      .Cls
   End With

End Sub

' fill the image of the specified clock
Public Sub SetClockImage(ByVal Index As Integer, ByVal FileName As String)

   With frmMyTimeZones
      If Index > 1 Then
         With .clkFavorits.Item(Index - 2)
            .Locked = True
            .Picture = LoadPicture(FileName)
            
            If FileName = "" Then .ClockPlateGradientStyle = OutToIn
            
            .Locked = False
         End With
         
      Else
         With .clkTimeZone.Item(Index)
            .Locked = True
            .Picture = LoadPicture(FileName)
            
            If FileName = "" Then .ClockPlateGradientStyle = OutToIn
            
            .Locked = False
         End With
      End If
   End With

End Sub

' relocate the timezones combobox to the specified window
Public Sub SetComboBox(ByVal hWnd As Long, ByRef Button As Object)

   If Not ShowSettings Then
      With frmMyTimeZones
         With .cmbTimeZones
            SetParent .hWnd, hWnd
            .Left = 30
            .Top = Button.Top + (Button.Height - .Height) \ 2
            .BackColor = &HF7EFE2
         End With
         
         .tcbSkinner.ComboBoxBorderColor = &HEAC183
      End With
   End If

End Sub

' set the imagebox icon specified by index
Public Sub SetIcon(ByRef Box As Image, ByVal Index As Integer)

   Set Box.Picture = frmMyTimeZones.imgImages.Item(Index).Picture

End Sub

' set globe background
Public Sub SetLogoAndGlobe(ByVal hDC As Long)

   With frmMyTimeZones.picImages
      BitBlt hDC, 26, 404, .ScaleWidth, .ScaleHeight, .hDC, 0, 0, vbSrcCopy
      GlobeXY.X = 35
      GlobeXY.Y = 405
   End With

End Sub
