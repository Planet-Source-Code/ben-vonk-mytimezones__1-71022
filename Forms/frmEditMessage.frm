VERSION 5.00
Begin VB.Form frmEditMessage 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   4680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7728
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
   ScaleHeight     =   390
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   644
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrText 
      Interval        =   40
      Left            =   7200
      Top             =   1200
   End
   Begin VB.PictureBox picMessage 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   6720
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1800
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.TextBox txtMessage 
      BackColor       =   &H00F7EFE2&
      BorderStyle     =   0  'None
      ForeColor       =   &H00D94600&
      Height          =   252
      Left            =   396
      MaxLength       =   100
      TabIndex        =   2
      Top             =   1680
      Width           =   5256
   End
   Begin VB.PictureBox picText 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F7EFE2&
      BorderStyle     =   0  'None
      ForeColor       =   &H00D94600&
      Height          =   252
      Left            =   6720
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   438
      TabIndex        =   1
      Top             =   2160
      Visible         =   0   'False
      Width           =   5256
   End
   Begin VB.Timer tmrError 
      Left            =   6720
      Top             =   1200
   End
   Begin MyTimeZones.LEDDisplay ledDisplay 
      Height          =   348
      Left            =   1080
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3096
      Width           =   3708
      _ExtentX        =   6541
      _ExtentY        =   614
      BackColor       =   14866892
      BorderStyle     =   0
      DisplayColor    =   16248802
      ForeColor       =   14239232
   End
   Begin MyTimeZones.CrystalButton cbtChoose 
      Height          =   504
      Index           =   0
      Left            =   5160
      TabIndex        =   4
      Top             =   3996
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
      Height          =   276
      Index           =   0
      Left            =   384
      Top             =   1668
      Width           =   5280
   End
   Begin VB.Image imgImages 
      Height          =   390
      Index           =   1
      Left            =   7200
      Top             =   120
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   2
      Left            =   6720
      Picture         =   "frmEditMessage.frx":0000
      Top             =   600
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   3
      Left            =   7200
      Picture         =   "frmEditMessage.frx":08CA
      Top             =   600
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   0
      Left            =   6720
      Top             =   120
      Visible         =   0   'False
      Width           =   384
   End
End
Attribute VB_Name = "frmEditMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Private Variable
Private InEditor As Boolean

Private Function CheckDoubleDateTime() As Boolean

Static blnBusy As Boolean

Dim intCount   As Integer
Dim intPointer As Integer
Dim intStart   As Integer
Dim intTotal   As Integer
Dim strText    As String

   If blnBusy Then Exit Function
   
   With txtMessage
      For intCount = 0 To 1
         intStart = 1
         intTotal = 0
         
         Do
            intStart = InStr(intStart, UCase(.Text), UCase(AppVar(intCount))) + 1
            
            If intStart = 1 Then
               Exit Do
               
            Else
               If intTotal = 1 Then
                  .Text = Left(.Text, intStart - 2) & Mid(.Text, intStart + Len(AppVar(intCount)))
                  intStart = 1
                  .SelStart = InStr(UCase(.Text), UCase(AppVar(intCount))) - 1
                  .SelLength = Len(AppVar(intCount))
                  .SetFocus
                  CheckDoubleDateTime = True
                  Exit Do
                  
               ElseIf Mid(.Text, intStart - 1, Len(AppVar(intCount))) <> AppVar(intCount) Then
                  blnBusy = True
                  intPointer = .SelStart
                  strText = .Text
                  Mid(strText, intStart - 1, Len(AppVar(intCount))) = AppVar(intCount)
                  .Text = strText
                  .SelStart = intPointer
                  blnBusy = False
               End If
            End If
            
            intTotal = intTotal + 1
         Loop
      Next 'intCount
   End With

End Function

Private Sub AddDateTimeVar(ByVal Index As Integer)

Dim intPointer    As Integer
Dim intPosition   As Integer
Dim ptaTextCursor As PointAPI

   InEditor = True
   GetCaretPos ptaTextCursor
   picText.Cls
   
   With txtMessage
      For intPosition = 1 To Len(.Text)
         If ptaTextCursor.X = picText.CurrentX + 1 Then
            intPosition = intPosition - 1
            Exit For
         End If
         
         picText.Print Mid(.Text, intPosition, 1);
      Next 'intPosition
      
      .SetFocus
      .SelStart = intPosition + Len(AppVar(Index))
      intPointer = InStr(UCase(.Text), UCase(AppVar(Index)))
      
      If intPointer Then
         .SelStart = intPointer - 1
         .SelLength = Len(AppVar(Index))
         
      ElseIf Len(.Text) + Len(AppVar(Index)) < .MaxLength Then
         .Text = Left(.Text, intPosition) & AppVar(Index) & Mid(.Text, intPosition + 1)
         .SelStart = intPosition + Len(AppVar(Index))
         
      Else
         Call ShowTextError
      End If
   End With

End Sub

Private Sub EndEditMessage()

   MousePointer = vbHourglass
   Hide
   DoEvents
   Unload Me
   Set frmEditMessage = Nothing

End Sub

Private Sub ResetMainDisplay()

   If frmMyTimeZones.tmrAlarmOff.Enabled Then Call DisableDisplay(frmMyTimeZones.ledDisplay)

End Sub

Private Sub ShowTextError()

Dim ptaTextCursor As PointAPI
Dim strMessage    As String

   With txtMessage
      GetCaretPos ptaTextCursor
      .Text = Left(.Text, .MaxLength)
      strMessage = GetToolTipText(Replace(AppError(38), "#", CStr(.MaxLength)))
   End With
   
   With picMessage
      .Top = txtMessage.Top + txtMessage.Height / 2
      .Height = .TextHeight("X") + 4
      .Width = .TextWidth(strMessage) + 2
      picMessage.Print strMessage
      .Left = txtMessage.Left + ptaTextCursor.X + 5
      
      If .Left + .Width > ScaleWidth Then .Left = txtMessage.Left + txtMessage.Width - .Width - 5
      
      .Visible = True
   End With
   
   tmrError.Enabled = True
   
   Call PlaySound(SOUND_ATTENTION)

End Sub

Private Sub cbtChoose_Click(Index As Integer)

Dim intCount   As Integer
Dim intPointer As Integer

   If CheckDoubleDateTime And (Index <> 1) Then Exit Sub
   
   Select Case Index
      Case 0
         For intCount = 0 To 1
            intPointer = InStr(UCase(txtMessage.Text), UCase(AppVar(intCount)))
            
            If intPointer Then txtMessage.Text = Left(txtMessage.Text, intPointer - 1) & AppVar(intCount) & Mid(txtMessage.Text, intPointer + Len(AppVar(intCount)))
         Next 'intCount
         
         FavoritsInfo(SelectedFavorit).AlarmMessage = txtMessage.Text
         
         Call EndEditMessage
         
      Case 1
         ' clear input
         With txtMessage
            If Len(.Text) Then If AppSettings(SET_ASKCONFIRM) Then If ShowMessage(AppError(46), vbQuestion, AppText(145), AppError(3), TimeToWait) = vbNo Then Exit Sub
            
            If .SelLength Then
               intPointer = .SelStart
               .Text = Left(.Text, .SelStart) & Mid(.Text, .SelStart + .SelLength + 1)
               
               If intPointer <= Len(.Text) Then
                  .SelStart = intPointer
                  
               Else
                  .SelStart = 0
               End If
               
            Else
               .Text = ""
            End If
            
            .SetFocus
         End With
         
      Case 2, 3
         ' add 2 = time or 3 = date variable to string
         Call AddDateTimeVar(3 - Index)
   End Select

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

   If IsExit(KeyCode, Shift) Then Call EndEditMessage

End Sub

Private Sub Form_Load()

Dim intCount As Integer

   Call SetIcon(imgImages.Item(0), 29)
   Call SetIcon(imgImages.Item(1), 30)
   Call InitForm(Me, 1, 6000, 4680)
   Call RoundForm(hWnd, ScaleWidth, ScaleHeight, 62)
   Call CenterForm(Me)
   Call DrawHeader(Me, AppText(145), True)
   Call ResizeAllControls(Me)
   Call ResizeControl(picMessage)
   
   Line (20, 218)-(ScaleWidth - 20, 218), &HB5A791
   picMessage.BackColor = vbInfoBackground
   tmrError.Interval = TimeToWait
   Load shpBorder.Item(1)
   
   With txtMessage
      .Height = .Height / ScreenResize
      .Left = (ScaleWidth - .Width) \ 2
      .Text = FavoritsInfo(SelectedFavorit).AlarmMessage
      
      Call DrawText(Me, AppText(165), 9, &H801F80, 0, .Top - .Height - 1, .Left)
      Call tmrText_Timer
   End With
   
   With shpBorder.Item(0)
      .Top = txtMessage.Top - 1
      .Left = txtMessage.Left - 1
      .Width = txtMessage.Width + 2
      .Height = txtMessage.Height + 2
   End With
   
   With shpBorder.Item(1)
      ledDisplay.Speed = DisplaySpeed
      .Top = ledDisplay.Top - 1
      .Left = ledDisplay.Left - 1
      .Width = ledDisplay.Width + 2
      .Height = ledDisplay.Height + 2
      .Visible = True
   End With
   
   For intCount = 0 To 3
      If intCount Then Load cbtChoose.Item(intCount)
      
      With cbtChoose.Item(intCount)
         If intCount Then
            .Left = cbtChoose.Item(intCount - 1).Left - .Width + 1
            .Shape = Choose(intCount, ShapeSides, ShapeSides, ShapeRight)
         End If
         
         .ToolTipText = GetToolTipText(AppText(12 - (7 And (intCount = 1)) + ((152 + intCount) And (intCount > 1))))
         .Picture = imgImages.Item(intCount).Picture
         .Visible = True
      End With
   Next 'intCount

End Sub

Private Sub picMessage_Click()

   Call tmrError_Timer

End Sub

Private Sub tmrError_Timer()

   picMessage.Visible = False
   tmrError.Enabled = False

End Sub

Private Sub tmrText_Timer()

   With ledDisplay
      If Len(txtMessage.Text) Then
         .Text = CreateAlarmMessage(SelectedFavorit, txtMessage.Text)
         
      Else
         .Text = ""
      End If
      
      .Active = True
      
      If .Active Then
         If .BackColor <> &HCFC1AF Then .BackColor = &HCFC1AF
         
      Else
         If .BackColor <> &HE2D9CC Then .BackColor = &HE2D9CC
      End If
   End With

End Sub

Private Sub txtMessage_Change()

   If txtMessage.Text = "" Then
      Call ResetMainDisplay
      
   Else
      CheckDoubleDateTime
   End If

End Sub

Private Sub txtMessage_GotFocus()

   If InEditor Then Exit Sub
   
   With txtMessage
      .SelStart = 0
      .SelLength = Len(.Text)
      InEditor = True
   End With

End Sub

Private Sub txtMessage_KeyDown(KeyCode As Integer, Shift As Integer)

   If (KeyCode = vbKeyLeft) Or (KeyCode = vbKeyRight) Or (KeyCode = vbKeyUp) Or (KeyCode = vbKeyDown) Or (KeyCode = vbKeyHome) Or (KeyCode = vbKeyEnd) Or (KeyCode = vbKeyDelete) Or (KeyCode = vbKeyBack) Then Call tmrError_Timer

End Sub

Private Sub txtMessage_KeyPress(KeyAscii As Integer)

   If Len(txtMessage.Text) >= txtMessage.MaxLength Then Call ShowTextError

End Sub

Private Sub txtMessage_LostFocus()

   CheckDoubleDateTime

End Sub
