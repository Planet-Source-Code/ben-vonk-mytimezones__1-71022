VERSION 5.00
Begin VB.Form frmImage 
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
   Begin VB.PictureBox picMonitor 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2232
      Left            =   6120
      Picture         =   "frmImage.frx":0000
      ScaleHeight     =   186
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   202
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2520
      Visible         =   0   'False
      Width           =   2424
   End
   Begin MyTimeZones.CheckBox chkImage 
      Height          =   384
      Left            =   360
      TabIndex        =   1
      Top             =   4080
      Width           =   3492
      _ExtentX        =   6160
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
      IconCheckedGrayed=   "frmImage.frx":0E82
   End
   Begin VB.PictureBox picImage 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00550000&
      BorderStyle     =   0  'None
      Height          =   1716
      Left            =   1872
      ScaleHeight     =   143
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   188
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1344
      Width           =   2256
      Begin VB.Shape shpBorder 
         BackColor       =   &H80000008&
         BorderColor     =   &H00400040&
         DrawMode        =   12  'Nop
         FillColor       =   &H0040A0FF&
         FillStyle       =   0  'Solid
         Height          =   1656
         Left            =   300
         Shape           =   3  'Circle
         Top             =   24
         Width           =   1656
      End
   End
   Begin VB.PictureBox picSource 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1332
      Left            =   6120
      ScaleHeight     =   111
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   121
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1080
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   2
      Left            =   6120
      Picture         =   "frmImage.frx":175C
      Top             =   600
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   1
      Left            =   6600
      Top             =   120
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   0
      Left            =   6120
      Top             =   120
      Visible         =   0   'False
      Width           =   384
   End
End
Attribute VB_Name = "frmImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Private Type
Private Type OpenFileName
   lStructSize       As Long
   hWndOwner         As Long
   hInstance         As Long
   lpstrFilter       As String
   lpstrCustomFilter As String
   nMaxCustFilter    As Long
   nFilterIndex      As Long
   lpstrFile         As String
   nMaxFile          As Long
   lpstrFileTitle    As String
   nMaxFileTitle     As Long
   lpstrInitialDir   As String
   lpstrTitle        As String
   lFlags            As Long
   nFileOffset       As Integer
   nFileExtension    As Integer
   lpstrDefExt       As String
   lCustData         As Long
   lpfnHook          As Long
   lpTemplateName    As String
End Type

' Private Variables
Private ImageFile    As String
Private ImageMap     As String

' Private API's
Private Declare Function GetOpenFileName Lib "ComDlg32" Alias "GetOpenFileNameA" (pOpenfilename As OpenFileName) As Long
Private Declare Function GetCurrentThreadId Lib "Kernel32" () As Long
Private Declare Function SetWindowsHookEx Lib "User32" Alias "SetWindowsHookExA" (ByVal IDHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long

Public Sub SetImageFile(ByVal FileName As String)

Dim intPointer As Integer

   If Len(FileName) Then
      intPointer = InStrRev(FileName, "\")
      ImageFile = Mid(FileName, intPointer + 1)
      ImageMap = Left(FileName, intPointer)
      ImageFile = CheckImageHeader(ImageMap, ImageFile)
      
      Call SetImage(True)
      
   Else
      ImageFile = ""
      ImageMap = ""
   End If

End Sub

Private Function CheckImageHeader(ByVal Map As String, ByVal FileName As String) As String

Dim blnIsOk     As Boolean
Dim intFileRead As Integer
Dim strHeader   As String * 50

   On Local Error GoTo ExitSub
   intFileRead = FreeFile
   
   Open Map & FileName For Binary Access Read As #intFileRead
      Get #intFileRead, 1, strHeader
   Close #intFileRead
   
   Select Case UCase(Right(FileName, 3))
      Case "BMP"
         blnIsOk = (Left(strHeader, 2) = "BM")
         
      Case "JPG"
         blnIsOk = (Left(strHeader, 3) = "ÿØÿ")
         
      Case "EMF"
         blnIsOk = (InStr(UCase(strHeader), "EMF"))
         
      Case "WMF"
         blnIsOk = (Left(strHeader, 3) = "×ÍÆ")
         
      Case "GIF"
         blnIsOk = (Left(strHeader, 3) = "GIF")
         
      Case "ICO"
         blnIsOk = (Left(strHeader, 4) = vbNullChar & vbNullChar & Chr(1) & vbNullChar)
   End Select
   
   If blnIsOk Then CheckImageHeader = FileName
   
ExitSub:
   On Local Error GoTo 0
   Close #intFileRead

End Function

Private Function GetImageFile() As String

Const GWL_HINSTANCE          As Long = -6
Const OFN_CREATEPROMPT       As Long = &H2000
Const OFN_EXPLORER           As Long = &H80000
Const OFN_HIDEREADONLY       As Long = &H4
Const OFN_LONGNAMES          As Long = &H200000
Const OFN_NODEREFERENCELINKS As Long = &H100000
Const WH_CBT                 As Long = 5
Const FILE_EXT               As String = "*.bmp;*.jpg;*.emf;*.wmf;*.gif;*.ico"

Dim ofnImageFile             As OpenFileName

   With ofnImageFile
      On Local Error GoTo ErrorLoad
      hWndParent = frmMyTimeZones.hWnd
      .lStructSize = Len(ofnImageFile)
      .hWndOwner = hWndParent
      .hInstance = App.hInstance
      .lpstrFilter = AppText(3) & " (" & FILE_EXT & ")" & vbNullChar & FILE_EXT & vbNullChar
      .lpstrFile = ImageFile & String(257 - Len(ImageFile), vbNullChar)
      .nMaxFile = Len(.lpstrFile) - 1
      .lpstrFileTitle = .lpstrFile
      .nMaxFileTitle = .nMaxFile
      .lpstrInitialDir = ImageMap
      .lpstrTitle = AppText(4)
      .lFlags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_CREATEPROMPT Or OFN_NODEREFERENCELINKS Or OFN_HIDEREADONLY
      WindowHook = SetWindowsHookEx(WH_CBT, AddressOf CenterCommonDialog, GetWindowLong(hWnd, GWL_HINSTANCE), GetCurrentThreadId)
      
      If GetOpenFileName(ofnImageFile) Then
         .lpstrFile = Trim(.lpstrFile)
         ImageFile = StripNull(.lpstrFileTitle)
         GetImageFile = Left(.lpstrFile, InStrRev(.lpstrFile, "\"))
         
      Else
         GetImageFile = ImageMap
      End If
      
      GoTo ExitFunction
   End With
   
ErrorLoad:
   If AppSettings(SET_ASKCONFIRM) Then ShowMessage AppError(10), vbCritical, AppError(12), AppError(9), TimeToWait
   
ExitFunction:
   On Local Error GoTo 0

End Function

Private Sub EndImage()

   MousePointer = vbHourglass
   ImageName = ""
   
   If Len(ImageFile) Then ImageName = ImageMap & ImageFile
   
   If SelectedClock > 1 Then
      FavoritsInfo(SelectedFavorit).ImageFile = ImageName
      
   Else
      ZonesInfo(SelectedClock).ImageFile = ImageName
   End If
   
   Hide
   DoEvents
   Unload Me
   Set frmImage = Nothing

End Sub

Private Sub SetImage(ByVal State As Boolean)

   With picSource
      If State And Len(ImageMap) Then .Picture = LoadPicture(ImageMap & ImageFile)
      
      If .Picture Then
         picImage.PaintPicture .Image, 0, 0, picImage.ScaleWidth, picImage.ScaleHeight, 0, 0, .ScaleWidth, .ScaleHeight, vbSrcCopy
         
         With picImage
            If AppSettings(SET_SHOWCLOCKIMAGE) Then
               .FillStyle = vbFSTransparent
               .DrawWidth = .ScaleWidth / 2
               picImage.Circle (.ScaleWidth / 2 + 1, .ScaleHeight / 2), (.ScaleWidth / 2) * 1.26, &HFF5F00
               .DrawWidth = 1
            End If
         End With
      End If
   End With

End Sub

Private Sub ShowClockBorder()

Dim lngColor As Long

   If Len(ImageFile) Then
      picImage.ToolTipText = GetToolTipText(AppText(164) & " " & ImageMap & ImageFile)
      shpBorder.Visible = False
      lngColor = &HFFFF00 ' Cyan
      
      Call SetImage(True)
      
   Else
      picImage.Cls
      picImage.ToolTipText = ""
      shpBorder.Visible = AppSettings(SET_SHOWCLOCKIMAGE)
      lngColor = &H8F4B29 ' Dark Blue (= color from monitor image)
   End If
   
   Line (335, 257)-(338, 259), lngColor, BF

End Sub

Private Sub cbtChoose_Click(Index As Integer)

   Select Case Index
      Case 0
         Call EndImage
         
      Case 1
         If Len(ImageFile) Then If AppSettings(SET_ASKCONFIRM) Then If ShowMessage(AppError(47), vbQuestion, AppText(160), AppError(3), TimeToWait) = vbNo Then Exit Sub
         
         ImageFile = ""
         
         Call ShowClockBorder
         
      Case 2
         ImageMap = GetImageFile
         
         If Len(ImageMap) And Len(ImageFile) Then
            ImageFile = CheckImageHeader(ImageMap, ImageFile)
            
            If AppSettings(SET_ASKCONFIRM) And (ImageFile = "") Then ShowMessage AppError(11), vbInformation, AppError(12), AppError(9), TimeToWait
         End If
         
         Call ShowClockBorder
   End Select

End Sub

Private Sub chkImage_Click()

   AppSettings(SET_SHOWCLOCKIMAGE) = Not AppSettings(SET_SHOWCLOCKIMAGE)
   
   Call ShowClockBorder

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

   If IsExit(KeyCode, Shift) Then Call EndImage

End Sub

Private Sub Form_Load()

Dim intCount As Integer

   Call SetIcon(imgImages.Item(1), 30)
   Call InitForm(Me, 1, 6000, 4680)
   Call RoundForm(hWnd, ScaleWidth, ScaleHeight, 62)
   Call CenterForm(Me)
   Call DrawHeader(Me, AppText(160), True)
   Call ResizeAllControls(Me, True)
   Call SetBackgroundImage(hDC, ScaleHeight - 20, ScaleWidth, picMonitor)
   Call ResizeControl(picImage)
   Call ShowClockBorder
   
   For intCount = 0 To 2
      If intCount Then Load cbtChoose.Item(intCount)
      
      With cbtChoose.Item(intCount)
         If intCount Then
            .Left = cbtChoose.Item(intCount - 1).Left - .Width + 1
            .Shape = Choose(intCount, ShapeSides, ShapeRight)
         End If
         
         .ToolTipText = GetToolTipText(AppText(12 + (148 + intCount And (intCount > 0))))
         .Picture = imgImages.Item(intCount).Picture
         .Visible = True
      End With
   Next 'intCount
   
   With chkImage
      .BackStyle = Transparent
      .Caption = AppText(163)
      .AutoSize = True
      .Value = Abs(AppSettings(SET_SHOWCLOCKIMAGE))
   End With
   
   With shpBorder
      .Width = picImage.ScaleWidth - 2
      .Height = picImage.ScaleHeight - 2
      .Top = (picImage.ScaleHeight - .Height) \ 2
      .Left = (picImage.ScaleWidth - .Width) \ 2
   End With

End Sub
