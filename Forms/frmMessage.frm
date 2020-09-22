VERSION 5.00
Begin VB.Form frmMessage 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   3840
   ClientLeft      =   36
   ClientTop       =   12
   ClientWidth     =   7320
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
   Picture         =   "frmMessage.frx":0000
   ScaleHeight     =   320
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   610
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrWait 
      Enabled         =   0   'False
      Left            =   6840
      Top             =   2520
   End
   Begin MyTimeZones.CrystalButton cbtChoose 
      Height          =   504
      Index           =   0
      Left            =   5160
      TabIndex        =   0
      Top             =   3072
      Width           =   672
      _extentx        =   1185
      _extenty        =   889
      backcolor       =   13542759
      caption         =   ""
      picalign        =   0
      shape           =   1
      font            =   "frmMessage.frx":9530
   End
   Begin VB.Image imgMessages 
      Height          =   576
      Index           =   5
      Left            =   6720
      Picture         =   "frmMessage.frx":9554
      Top             =   1320
      Visible         =   0   'False
      Width           =   576
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   2
      Left            =   6240
      Top             =   2520
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgMessages 
      Height          =   576
      Index           =   3
      Left            =   6720
      Picture         =   "frmMessage.frx":B21E
      Top             =   720
      Visible         =   0   'False
      Width           =   576
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   0
      Left            =   6240
      Picture         =   "frmMessage.frx":CEE8
      Top             =   2040
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   1
      Left            =   6840
      Picture         =   "frmMessage.frx":D7B2
      Top             =   2040
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgMessages 
      Height          =   576
      Index           =   2
      Left            =   6120
      Picture         =   "frmMessage.frx":E07C
      Top             =   720
      Visible         =   0   'False
      Width           =   576
   End
   Begin VB.Image imgMessages 
      Height          =   576
      Index           =   4
      Left            =   6120
      Picture         =   "frmMessage.frx":FD46
      Top             =   1320
      Visible         =   0   'False
      Width           =   576
   End
   Begin VB.Image imgMessages 
      Height          =   576
      Index           =   1
      Left            =   6720
      Picture         =   "frmMessage.frx":11A10
      Top             =   120
      Visible         =   0   'False
      Width           =   576
   End
   Begin VB.Image imgMessages 
      Height          =   576
      Index           =   0
      Left            =   6120
      Picture         =   "frmMessage.frx":136DA
      Top             =   120
      Visible         =   0   'False
      Width           =   576
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Private Variables
Private IsDefaultKey  As VbMsgBoxStyle
Private MessageResult As VbMsgBoxResult

Public Function ShowMessage(ByVal Prompt As String, ByVal MessageType As VbMsgBoxStyle, ByVal Title As String, ByVal Message As String, ByVal Wait As Integer, ByVal DefaultKey As VbMsgBoxStyle) As VbMsgBoxResult

Dim intCount     As Integer
Dim intCurrY     As Integer
Dim intIcon      As Integer
Dim intLineCount As Integer
Dim strText()    As String

   intIcon = Choose(MessageType \ 16, 0, 1, 2, 3, 4, 5)
   IsDefaultKey = DefaultKey
   
   If MessageType <> vbQuestion Then
      With cbtChoose.Item(1)
         .Width = .Width * 1.3
         .Shape = ShapeNone
         cbtChoose.Item(0).Visible = False
         
         If intIcon = 5 Then
            .Top = ScaleHeight - .Height - 16
            .Left = ScaleWidth - .Width - 14
            
         Else
            .Left = cbtChoose.Item(0).Left - (.Width - cbtChoose.Item(0).Width)
         End If
      End With
   End If
   
   Call DrawText(Me, Title, 10, &H766B4D, 0, 6, 41)
   Call DrawText(Me, Title, 10, &HFFFFDF, 0, 5, 40)
   Call DrawText(Me, Message, 14 - (2 And intIcon = 5), &HC01FC0, 0, 60 + (5 And (intIcon = 5)), 100)
   
   DrawIconEx hDC, 29, 49, imgMessages.Item(intIcon).Picture.Handle, 48, 48, 0, 0, DI_NORMAL
   
   If intIcon = 5 Then
      With cbtChoose.Item(1)
         intCurrY = .Top + (.Height - TextHeight("X") * 1.7) \ 2
         .Picture = imgImages.Item(2).Picture
         .ToolTipText = GetToolTipText(AppText(12))
      End With
      
      Call DrawCopyright(Me, intCurrY)
      
   Else
      cbtChoose.Item(0).Picture = imgImages.Item(0).Picture
      
      For intCount = 1 To 2
         Line (20, 120 * intCount)-(ScaleWidth - 20, 120 * intCount + 1), &HD1C4B4, BF
      Next 'intCount
      
      ForeColor = &HD94600
   End If
   
   strText = Split(WordWrap(Prompt), vbCrLf)
   intLineCount = UBound(strText)
   FontSize = 9 + (6 And (intIcon = 5))
   CurrentY = 180 - (((intLineCount + 1) * TextHeight("X"))) \ 2 - ((TextHeight("X") * intLineCount + ((intLineCount - 2) And intLineCount = 3)) \ (2 + (2 And (intLineCount = 4)))) \ 2
   
   If intIcon = 5 Then CurrentY = 144
   
   For intCount = 0 To intLineCount
      If intIcon = 5 Then
         If intCount = intLineCount Then
            FontSize = 8 + AddToFontSize
            CurrentY = CurrentY - TextHeight("X") \ 1.5
         End If
         
         intCurrY = CurrentY
         CurrentX = (ScaleWidth - TextWidth(strText(intCount))) / 2 + 1
         CurrentY = CurrentY + 1
         ForeColor = vbWhite
         Print strText(intCount)
         CurrentY = intCurrY
         
         If intCount = intLineCount Then
            ForeColor = &HC01FC0
            
         Else
            ForeColor = &HD94600
         End If
      End If
      
      CurrentX = (ScaleWidth - TextWidth(strText(intCount))) / 2
      Print strText(intCount)
      CurrentY = CurrentY + TextHeight("X") \ (2 + (2 And (intLineCount = 4)))
   Next 'intCount
   
   Erase strText
   tmrWait.Interval = Wait
   tmrWait.Enabled = CBool(Wait)
   
   If intIcon <> 5 Then Call PlaySound(SOUND_ATTENTION)
   
   Show vbModal, frmMyTimeZones
   ShowMessage = MessageResult

End Function

Private Function WordWrap(ByVal Text As String) As String

Dim intNewLine   As Integer
Dim intTextWidth As Integer

   intTextWidth = ScaleWidth - 40
   
   Do While TextWidth(Text) > intTextWidth
      intNewLine = Len(Text)
      
      Do
         intNewLine = InStrRev(Text, " ", intNewLine) - 1
         
         If TextWidth(Left(Text, intNewLine)) <= intTextWidth Then Exit Do
      Loop
      
      WordWrap = WordWrap & Left(Text, intNewLine) & vbCrLf
      Text = Mid(Text, intNewLine + 2)
   Loop
   
   WordWrap = WordWrap & Text

End Function

Private Sub EndMessage()

   Hide
   DoEvents
   Unload Me
   Set frmMessage = Nothing

End Sub

Private Sub FormCenter()

   If frmMyTimeZones.Visible Then
      Call CenterForm(Me)
      
   Else
      Top = (Screen.Height - Height) \ 2
      Left = (Screen.Width - Width) \ 2
   End If

End Sub

Private Sub cbtChoose_Click(Index As Integer)

   Select Case Index
      Case 0
         MessageResult = vbNo
         
      Case 1
         MessageResult = vbYes
   End Select
   
   Call EndMessage

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

   If IsExit(KeyCode, Shift) Then Call EndMessage

End Sub

Private Sub Form_Load()

   Height = 3840 * ScreenResize
   Width = 6060 * ScreenResize
   
   Call SetIcon(imgImages.Item(2), 29)
   Call ResizeAllControls(Me, True)
   Call RoundForm(hWnd, ScaleWidth, ScaleHeight, 23, True)
   Call FormCenter
   
   With cbtChoose
      Load .Item(1)
      .Item(1).Left = .Item(0).Left - .Item(0).Width + 1
      .Item(1).Shape = ShapeRight
      .Item(1).Visible = True
      .Item(0).ToolTipText = GetToolTipText(AppText(260))
      .Item(1).ToolTipText = GetToolTipText(AppText(261))
      .Item(1).Picture = imgImages.Item(1).Picture
   End With

End Sub

Private Sub tmrWait_Timer()

   MessageResult = IsDefaultKey
   
   Call EndMessage

End Sub
