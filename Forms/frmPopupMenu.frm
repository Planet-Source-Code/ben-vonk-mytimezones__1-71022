VERSION 5.00
Begin VB.Form frmPopupMenu 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   3840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3396
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPopupMenu.frx":0000
   ScaleHeight     =   320
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   283
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer tmrCheck 
      Interval        =   10
      Left            =   3000
      Top             =   720
   End
   Begin VB.Image imgImages 
      Height          =   240
      Index           =   1
      Left            =   3000
      Picture         =   "frmPopupMenu.frx":5C34
      Top             =   360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgImages 
      Height          =   240
      Index           =   0
      Left            =   3000
      Picture         =   "frmPopupMenu.frx":625E
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgSelected 
      Height          =   252
      Index           =   0
      Left            =   360
      Top             =   120
      Width           =   2400
   End
   Begin VB.Image imgOption 
      Height          =   240
      Index           =   0
      Left            =   360
      Stretch         =   -1  'True
      Top             =   120
      Width           =   240
   End
   Begin VB.Label lblOption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   216
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   48
   End
End
Attribute VB_Name = "frmPopupMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Private Type
Private Type LogFont
   lfHeight         As Long
   lfWidth          As Long
   lfEscapement     As Long
   lfOrientation    As Long
   lfWeight         As Long
   lfItalic         As Byte
   lfUnderline      As Byte
   lfStrikeOut      As Byte
   lfCharSet        As Byte
   lfOutPrecision   As Byte
   lfClipPrecision  As Byte
   lfQuality        As Byte
   lfPitchAndFamily As Byte
   lfFacename       As String * 33
End Type

' Private Variable
Private MenuItem    As Integer

' Private API's
Private Declare Function CreateFontIndirect Lib "GDI32" Alias "CreateFontIndirectA" (lpLogFont As LogFont) As Long
Private Declare Function GetActiveWindow Lib "User32" () As Long

Public Sub ActivatePopupMenu()

Dim intCount   As Integer
Dim ptaMouseXY As PointAPI

   If Visible Then Call EndPopupMenu
   
   GetCursorPos ptaMouseXY
   Left = ptaMouseXY.X * Screen.TwipsPerPixelX - Width
   Top = ptaMouseXY.Y * Screen.TwipsPerPixelY - Height
   SelectedPopupMenu = -1
   MenuItem = -1
   tmrCheck.Enabled = True
   
   For intCount = 0 To 6
      imgOption.Item(intCount).Picture = imgImages.Item(0 + (1 And (intCount = SelectedTrayClock))).Picture
      
      If intCount < 2 Then
         lblOption.Item(intCount).Caption = frmMyTimeZones.clkTimeZone.Item(intCount).NameClock
         
      Else
         lblOption.Item(intCount).Caption = frmMyTimeZones.clkFavorits.Item(intCount - 2).NameClock
      End If
   Next 'intCount
   
   Call ColorMenuItems
   
   SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
   SetForegroundWindow hWnd

End Sub

Public Sub EndPopupMenu()

   Hide
   DoEvents
   Unload Me
   Set frmPopupMenu = Nothing

End Sub

Private Sub ColorMenuItems()

Dim intCount As Integer

   For intCount = 0 To lblOption.UBound
      With lblOption.Item(intCount)
         If .Caption = AppText(6) Then
            .ForeColor = &HBBAC99
            
         Else
            .ForeColor = &H801F80
         End If
         
         If intCount < 7 Then If intCount = SelectedTrayClock Then .ForeColor = &HC01FC0
      End With
   Next 'intCount

End Sub

Private Sub SetControl(ByVal ControlType As Object, ByVal Top As Single, ByVal Left As Single)

   With ControlType
      .Top = Top
      .Left = Left
      .Visible = True
   End With

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

   If IsExit(KeyCode, Shift) Then
      Call EndPopupMenu
      
   Else
      If KeyCode = vbKeyDown Then
         Call ColorMenuItems
         
         MenuItem = MenuItem + 1
         
         If MenuItem > 8 Then MenuItem = 0
         If lblOption.Item(MenuItem).Caption = AppText(6) Then MenuItem = MenuItem + 1
         If MenuItem > 8 Then MenuItem = 0
         
         lblOption.Item(MenuItem).ForeColor = &HD94600
         
      ElseIf (KeyCode = vbKeyEnd) Or (KeyCode = vbKeyHome) Then
         Call ColorMenuItems
         
         MenuItem = 8 - (8 And KeyCode = vbKeyHome)
         lblOption.Item(MenuItem).ForeColor = &HD94600
         
      ElseIf KeyCode = vbKeyReturn Then
         If MenuItem < 7 Then SelectedTrayClock = MenuItem
         
         SelectedPopupMenu = MenuItem
         
         Call ColorMenuItems
         Call EndPopupMenu
         
      ElseIf KeyCode = vbKeyUp Then
         Call ColorMenuItems
         
         MenuItem = MenuItem - 1
         
         If MenuItem < 0 Then MenuItem = 8
         If lblOption.Item(MenuItem).Caption = AppText(6) Then MenuItem = MenuItem - 1
         If MenuItem < 0 Then MenuItem = 8
         
         lblOption.Item(MenuItem).ForeColor = &HD94600
      End If
   End If

End Sub

Private Sub Form_Load()

Dim intCount    As Integer
Dim intLength   As Integer
Dim lgfFont     As LogFont
Dim lngPrevFont As Long
Dim sngCurrY    As Single
Dim strChar     As String

   Height = 3840 * ScreenResize
   Width = 2880 * ScreenResize
   
   Call ResizeAllControls(Me)
   
   For intCount = 1 To 8
      Load imgSelected.Item(intCount)
      Load lblOption.Item(intCount)
      Load imgOption.Item(intCount)
      sngCurrY = imgSelected.Item(intCount - 1).Top + 30
      
      If (intCount = 2) Or (intCount = 7) Then sngCurrY = sngCurrY + 20
      
      Call SetControl(lblOption.Item(intCount), sngCurrY, lblOption.Item(0).Left)
      Call SetControl(imgOption.Item(intCount), sngCurrY, imgOption.Item(0).Left)
      Call SetControl(imgSelected.Item(intCount), sngCurrY, imgSelected.Item(0).Left)
   Next 'intCount
   
   With frmMyTimeZones
      imgOption.Item(7).Picture = .Icon
      lblOption.Item(7).Caption = AppText(26) & " " & App.Title
      imgOption.Item(8).Picture = .flbChoose.Item(10).Icon
      lblOption.Item(8).Caption = AppText(60)
   End With
   
   With lgfFont
      .lfFacename = FontName & vbNullChar
      .lfEscapement = 900
      .lfHeight = 16
   End With
   
   For intCount = 0 To 1
      If intCount Then
         ForeColor = &HD94600
         
      Else
         ForeColor = &HE0E0E0
      End If
      
      sngCurrY = ScaleHeight - 6 + intCount
      
      For intLength = 1 To Len(APP_PRODUCTNAME)
         lngPrevFont = SelectObject(hDC, CreateFontIndirect(lgfFont))
         strChar = Mid(APP_PRODUCTNAME, intLength, 1)
         CurrentX = 4 - intCount
         CurrentY = sngCurrY
         sngCurrY = sngCurrY - TextWidth(strChar)
         Print strChar
         DeleteObject SelectObject(hDC, lngPrevFont)
      Next 'intLength
   Next 'intCount

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If MenuItem = -1 Then Exit Sub
   
   Call ColorMenuItems
   
   MenuItem = -1

End Sub

Private Sub Form_Paint()

   SetFocus

End Sub

Private Sub imgSelected_Click(Index As Integer)

   If Index < 7 Then
      If lblOption.Item(Index).Caption = AppText(6) Then Exit Sub
      
      imgOption.Item(SelectedTrayClock).Picture = imgImages.Item(0).Picture
      lblOption.Item(SelectedTrayClock).ForeColor = &H801F80
      imgOption.Item(Index).Picture = imgImages.Item(1).Picture
      SelectedTrayClock = Index
      
   Else
      SelectedPopupMenu = Index
   End If
   
   Call ColorMenuItems
   Call EndPopupMenu

End Sub

Private Sub imgSelected_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim strDate As String

   If MenuItem = Index Then Exit Sub
   
   Call ColorMenuItems
   
   imgSelected.Item(Index).ToolTipText = ""
   
   If Index < 2 Then
      strDate = frmMyTimeZones.clkTimeZone.Item(Index).DateTime
      
   ElseIf Index < 7 Then
      strDate = frmMyTimeZones.clkFavorits.Item(Index - 2).DateTime
   End If
   
   If lblOption.Item(Index).Caption = AppText(6) Then Exit Sub
   If Len(strDate) Then imgSelected.Item(Index).ToolTipText = GetToolTipText(CapsText(Format(strDate, LongDateFormat)) & " - " & Format(strDate, "hh:mm"))
   
   lblOption.Item(Index).ForeColor = &HD94600
   MenuItem = Index

End Sub

Private Sub tmrCheck_Timer()

   If hWnd <> GetActiveWindow Then Call EndPopupMenu

End Sub
