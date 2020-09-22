VERSION 5.00
Begin VB.Form frmIntro 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   6060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8856
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
   Moveable        =   0   'False
   ScaleHeight     =   505
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   738
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrAnimation 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   8400
      Top             =   120
   End
   Begin VB.PictureBox picMap 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2868
      Left            =   1704
      Picture         =   "frmIntro.frx":0000
      ScaleHeight     =   239
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   399
      TabIndex        =   0
      Top             =   1560
      Width           =   4788
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00958761&
      Height          =   2892
      Left            =   1692
      Top             =   1548
      Width           =   4812
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Private Variables
Private LineStart As Single
Private LineX     As Single

Public Sub EndIntro()

   If InIntro Then Exit Sub
   
   Hide
   frmMyTimeZones.Show
   DoEvents
   Unload Me
   Set frmIntro = Nothing

End Sub

Private Sub Form_Click()

   Call EndIntro

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

   Call EndIntro

End Sub

Private Sub Form_Load()

Dim intCount As Integer

   Call InitForm(Me, 0, , , False)
   Call RoundForm(hWnd, ScaleWidth, ScaleHeight, 62)
   Call SetLogoAndGlobe(hDC)
   
   With Screen
      Top = (.Height - Height) \ 2 - 16 * .TwipsPerPixelY
      Left = (.Width - Width) \ 2
   End With
   
   With picMap
      .Top = (ScaleHeight - .Height) \ 2
      .Left = (ScaleWidth - .Width) \ 2
      shpBorder.Top = .Top - 1
      shpBorder.Left = .Left - 1
      shpBorder.Width = .ScaleWidth + 2
      shpBorder.Height = .ScaleHeight + 2
      LineStart = .ScaleWidth \ 2 - 14
      LineX = LineStart
   End With
   
   For intCount = 1 To 0 Step -1
      Call DrawText(Me, APP_PRODUCTNAME, 11, &HC01FC0, intCount, 16)
      Call DrawText(Me, InfoVersion, 9, &HD94600, intCount, 39)
   Next 'intCount
   
   Call DrawCopyright(Me, 453)
   Call DoAnimation(Me)
   
   tmrAnimation.Enabled = True
   SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
   Show
   DoEvents

End Sub

Private Sub picMap_Click()

   Call EndIntro

End Sub

Private Sub tmrAnimation_Timer()

Dim lngColor As Long

   With picMap
      If Visible Then .SetFocus
      
      If InIntro Then
         lngColor = &HD36229
         
      Else
         lngColor = &HE7D1A1
      End If
      
      .Cls
      picMap.Line (LineX, 0)-(LineX + 1, .ScaleHeight), lngColor, B
      LineX = LineX + 2
      
      If LineX > .ScaleWidth Then LineX = -1
      
      If (LineX > LineStart - 3) And (LineX < LineStart) Then
         Call EndIntro
         
      ElseIf LineX Mod 3 = 0 Then
         Call DoAnimation(Me)
      End If
   End With

End Sub
