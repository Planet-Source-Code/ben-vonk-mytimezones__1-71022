VERSION 5.00
Begin VB.Form frmSettings 
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   505
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   773
   ShowInTaskbar   =   0   'False
   Begin MyTimeZones.ThemedComboBox tcbSkinner 
      Left            =   8760
      Top             =   1560
      _ExtentX        =   445
      _ExtentY        =   423
      BorderColorStyle=   1
      ComboBoxBorderColor=   11904913
      DriveListBoxBorderColor=   0
   End
   Begin MyTimeZones.CheckBox chkSettings 
      Height          =   384
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   1320
      Width           =   3540
      _ExtentX        =   6244
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
      IconCheckedGrayed=   "frmSettings.frx":0000
   End
   Begin VB.ComboBox cmbTimeServers 
      BackColor       =   &H00F7EFE2&
      ForeColor       =   &H00D94600&
      Height          =   312
      ItemData        =   "frmSettings.frx":08DA
      Left            =   2280
      List            =   "frmSettings.frx":0905
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   4440
      Width           =   5280
   End
   Begin VB.Timer tmrAnimation 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   8280
      Top             =   1560
   End
   Begin MyTimeZones.CrystalButton cbtChoose 
      Height          =   504
      Index           =   0
      Left            =   7380
      TabIndex        =   2
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
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   5
      Left            =   8760
      Picture         =   "frmSettings.frx":0B86
      Top             =   1080
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   4
      Left            =   8280
      Picture         =   "frmSettings.frx":1450
      Top             =   1080
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   3
      Left            =   8760
      Picture         =   "frmSettings.frx":1D1A
      Top             =   600
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   2
      Left            =   8280
      Picture         =   "frmSettings.frx":25E4
      Top             =   600
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
      Index           =   0
      Left            =   8280
      Top             =   120
      Visible         =   0   'False
      Width           =   384
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Private Constant
Private Const CB_FINDSTRING As Long = &H14C

' Private Variables
Private Settings(11)        As Boolean
Private TimeServer          As String

Private Function DefaultValues(Optional ByVal OnlyCheck As Boolean) As Boolean

Dim intCount As Integer

   With chkSettings
      For intCount = 0 To .UBound
         With .Item(intCount)
            If (intCount < 5) Or (intCount = 8) Or (intCount = 11) Then
               If .Value = vbUnchecked Then Exit For
               
            ElseIf .Value = vbChecked Then
               Exit For
            End If
         End With
      Next 'intCount
      
      If intCount = .UBound + 1 Then
         If Left(cmbTimeServers.Text, Len(DEFAULT_TIMESERVER)) = DEFAULT_TIMESERVER Then
            DefaultValues = True
            Exit Function
         End If
      End If
   End With
   
   If OnlyCheck Then Exit Function
   If AppSettings(SET_ASKCONFIRM) Then If ShowMessage(AppError(45), vbQuestion, AppError(44), AppError(3), TimeToWait) = vbNo Then Exit Function
   
   For intCount = 0 To 11
      With chkSettings.Item(intCount)
         If (intCount < 5) Or (intCount = 8) Or (intCount = 11) Then
            .Value = vbChecked
            
         Else
            .Value = vbUnchecked
         End If
         
         AppSettings(intCount) = .Value
      End With
   Next 'intCount
   
   cmbTimeServers.ListIndex = GetListIndex(cmbTimeServers.hWnd, CB_FINDSTRING, DEFAULT_TIMESERVER)
   
   Call SetButtonInfo

End Function

Private Sub EndSettings()

   MousePointer = vbHourglass
   
   If AppSettings(SET_AUTOSAVE) Then Call SaveSettings
   
   TimeServerURL = TimeServer
   
   Call SwapSettings(Settings, AppSettings)
   Call frmMyTimeZones.ToggleControls(True)
   Call SetButtonInfo
   
   Hide
   DoEvents
   Unload Me
   Set frmSettings = Nothing

End Sub

Private Sub SaveSettings()

Dim blnChanged As Boolean
Dim intCount   As Integer

   For intCount = 0 To UBound(Settings)
      If AppSettings(intCount) <> Settings(intCount) Then
         blnChanged = True
         Exit For
      End If
   Next 'intCount
   
   If TimeServerURL <> TimeServer Then blnChanged = True
   If Not blnChanged Then Exit Sub
   If AppSettings(SET_ASKCONFIRM) Then If ShowMessage(AppError(19), vbQuestion, AppError(18), AppError(3), TimeToWait) = vbNo Then Exit Sub
   
   ' Complete SaveSettings!
   TimeServer = TimeServerURL
   
   Call SetRegValues
   Call SwapSettings(AppSettings, Settings)

End Sub

Private Sub SetButtonInfo()

   Call SetButton(1, GetIcon(1), frmMyTimeZones)
   Call SetButton(2, GetIcon(2), frmMyTimeZones)
   Call SetButton(7, GetIcon(7), frmMyTimeZones)
   Call SetButton(8, GetIcon(8), frmMyTimeZones)
   Call frmMyTimeZones.CheckFavoritIsSet
   
   cmbTimeServers.ToolTipText = GetToolTipText(AppText(199))
   cbtChoose.Item(0).ToolTipText = GetToolTipText(AppText(12))
   cbtChoose.Item(1).ToolTipText = GetToolTipText(AppText(196 + AppSettings(SET_AUTOSAVE) * 3))
   cbtChoose.Item(1).Picture = imgImages.Item(1 + Abs(AppSettings(SET_AUTOSAVE)) * 3).Picture
   cbtChoose.Item(2).ToolTipText = GetToolTipText(AppText(194 + Abs(DefaultValues(True)) * 3))
   cbtChoose.Item(2).Picture = imgImages.Item(2 + Abs(DefaultValues(True)) * 3).Picture
   cbtChoose.Item(3).ToolTipText = GetToolTipText(AppText(195))
   frmMyTimeZones.shpActiveClock.Visible = AppSettings(SET_ACTIVEBORDER)
   frmMyTimeZones.twvDate.MouseTrap = AppSettings(SET_MOUSEINTHUMBWHEEL)
   DoEvents

End Sub

Private Sub ShowInformation()

Const vbInfo As Integer = &H60

   MousePointer = vbHourglass
   ShowMessage APP_PRODUCTNAME & vbCrLf & InfoVersion, vbInfo, AppText(195), ""
   MousePointer = vbDefault

End Sub

Private Sub SwapSettings(ByRef FromArray() As Boolean, ByRef ToArray() As Boolean)

Dim intIndex As Integer

   For intIndex = 0 To UBound(Settings)
      ToArray(intIndex) = FromArray(intIndex)
   Next 'intIndex

End Sub

Private Sub chkSettings_Click(Index As Integer)

   AppSettings(Index) = Not AppSettings(Index)
   
   Call SetButtonInfo

End Sub

Private Sub cmbTimeServers_Click()

   TimeServerURL = Trim(Split(cmbTimeServers.Text, "(", 2)(0))
   
   Call SetButtonInfo

End Sub

Private Sub cbtChoose_Click(Index As Integer)

   Select Case Index
      Case 0
         Call EndSettings
         
      Case 1
         If Not AppSettings(SET_AUTOSAVE) Then Call SaveSettings
         
      Case 2
         If DefaultValues Then
            If AppSettings(SET_ASKCONFIRM) Then
               ShowMessage AppError(49), vbStop, AppError(44), AppError(16), TimeToWait
               Exit Sub
            End If
         End If
         
      Case 3
         Call ShowInformation
   End Select

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

   If IsExit(KeyCode, Shift) Then Call EndSettings

End Sub

Private Sub Form_Load()

Dim intCount As Integer

   Call SetIcon(imgImages.Item(1), 31)
   Call InitForm(Me)
   Call RoundForm(hWnd, ScaleWidth, ScaleHeight, 62)
   Call CenterForm(Me)
   Call ResizeAllControls(Me)
   Call DrawHeader(Me, AppText(59), True)
   Call SetLogoAndGlobe(hDC)
   Call DrawText(Me, AppText(9), 9, &H801F80, 0, cmbTimeServers.Top + 4, 40)
   Call SwapSettings(AppSettings, Settings)
   
   TimeServer = TimeServerURL
   tmrAnimation.Enabled = True
   
   For intCount = 0 To 3
      If intCount Then
         Load cbtChoose.Item(intCount)
         
         With cbtChoose.Item(intCount)
            .Left = cbtChoose.Item(intCount - 1).Left - .Width + 1
            .Shape = Choose(intCount, ShapeSides, ShapeSides, ShapeRight)
            .Visible = True
         End With
      End If
      
      cbtChoose.Item(intCount).Picture = imgImages.Item(intCount)
   Next 'intCount
   
   For intCount = 0 To UBound(Settings)
      If intCount Then Load chkSettings.Item(intCount)
      
      With chkSettings.Item(intCount)
         .Visible = True
         .Caption = AppText(180 + intCount)
         .Value = Abs(AppSettings(intCount))
         Settings(intCount) = .Value
         .AutoSize = True
         
         If intCount = 6 Then
            .Top = chkSettings.Item(0).Top
            
         ElseIf intCount Then
            .Top = chkSettings.Item(intCount - 1).Top + .Height + 10
         End If
         
         If intCount > 5 Then
            .AlignCaption = [Right Justify]
            .Alignment = [Right Justify]
            .Left = ScaleWidth - .Width - 40
         End If
         
         .BackStyle = Transparent
      End With
   Next 'intCount
   
   With cmbTimeServers
      .Left = ScaleWidth - .Width - 40
      .List(0) = AppText(198)
      .ListIndex = GetListIndex(.hWnd, CB_FINDSTRING, TimeServerURL)
   End With
   
   Call SetButtonInfo

End Sub

Private Sub tmrAnimation_Timer()

   Call DoAnimation(Me)

End Sub
