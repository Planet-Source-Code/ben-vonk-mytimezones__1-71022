VERSION 5.00
Begin VB.UserControl ThumbWheel 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   ClientHeight    =   1128
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1128
   ScaleHeight     =   94
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   94
   ToolboxBitmap   =   "ThumbWheel.ctx":0000
   Begin VB.PictureBox picMaskVertical 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H80000008&
      ForeColor       =   &H80000008&
      Height          =   516
      Left            =   480
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   108
   End
   Begin VB.PictureBox picMaskHorizontal 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H80000008&
      ForeColor       =   &H80000008&
      Height          =   108
      Left            =   600
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   516
   End
   Begin VB.PictureBox picTurnVertical 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   516
      Left            =   360
      Picture         =   "ThumbWheel.ctx":0312
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   108
   End
   Begin VB.PictureBox picWheel 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   360
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox picTurnHorizontal 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   108
      Left            =   600
      Picture         =   "ThumbWheel.ctx":082D
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   516
   End
   Begin VB.Shape shpShadeControl 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      BorderStyle     =   0  'Transparent
      DrawMode        =   9  'Not Mask Pen
      FillColor       =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   1056
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   1092
   End
   Begin VB.Image imgVertical 
      Height          =   708
      Left            =   0
      Picture         =   "ThumbWheel.ctx":0D3F
      Top             =   360
      Visible         =   0   'False
      Width           =   336
   End
   Begin VB.Image imgHorizontal 
      Height          =   336
      Left            =   360
      Picture         =   "ThumbWheel.ctx":144F
      Top             =   0
      Visible         =   0   'False
      Width           =   708
   End
End
Attribute VB_Name = "ThumbWheel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'ThumbWheel Control
'
'Author Ben Vonk
'24-08-2005 First version (based on Nero's 'ThumbWheel Control' at http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=11856&lngWId=1)
'10-11-2007 Second version, add Enabled, MouseTrap and SpinOver properties

Option Explicit

' Public Events
Public Event Change()
Public Event Click()

' Public Enumeration
'Public Enum Orientations
'   Horizontal
'   Vertical
'End Enum

' Private Types
'Private Type Rect
'   Left                As Integer
'   Top                 As Integer
'   Right               As Integer
'   Bottom              As Integer
'End Type

'Private Type PointAPI
'   X                   As Long
'   Y                   As Long
'End Type

' Private Variables
Private Clicked        As Boolean
Private m_MouseTrap    As Boolean
Private m_SpinOver     As Boolean
Private Increment      As Integer
Private m_Max          As Integer
Private m_Min          As Integer
Private m_Orientation  As Integer
Private m_Value        As Integer
Private WheelPosition  As Integer
Private m_ShadeControl As Long
Private m_ShadeWheel   As Long
Private LastX          As Single
Private LastY          As Single

' Private API's
'Private Declare Function BitBlt Lib "GDI32" (ByVal hDCDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal dwRop As Long) As Long
'Private Declare Function StretchBlt Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function ClientToScreen Lib "User32" (ByVal hWnd As Long, lpPoint As PointAPI) As Long
Private Declare Function ClipCursor Lib "User32" (lpRect As Any) As Long
'Private Declare Function GetClientRect Lib "User32" (ByVal hWnd As Long, lpRect As Rect) As Long
Private Declare Function OffsetRect Lib "User32" (lpRect As Rect, ByVal X As Long, ByVal Y As Long) As Long

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines wheter an object can respond to user-generated events."

   Enabled = UserControl.Enabled

End Property

Public Property Let Enabled(NewEnabled As Boolean)

   UserControl.Enabled = NewEnabled
   PropertyChanged "Enabled"

End Property

Public Property Get Max() As Integer
Attribute Max.VB_Description = "Returns/sets a thumb wheel position's maximum Value property setting."

   Max = m_Max

End Property

Public Property Let Max(ByVal NewMax As Integer)

   If NewMax < m_Min Then NewMax = m_Max
   
   m_Max = NewMax
   m_Value = m_Max
   PropertyChanged "Max"

End Property

Public Property Get Min() As Integer
Attribute Min.VB_Description = "Returns/sets a thumb wheel position's minimum Value property setting."

   Min = m_Min

End Property

Public Property Let Min(ByVal NewMin As Integer)

   If NewMin > m_Max Then NewMin = m_Min
   
   m_Min = NewMin
   m_Value = m_Min
   PropertyChanged "Min"

End Property

Public Property Get MouseTrap() As Boolean
Attribute MouseTrap.VB_Description = "Determines whether the mouse can only move in the control."

    MouseTrap = m_MouseTrap

End Property

Public Property Let MouseTrap(ByVal NewMouseTrap As Boolean)

   m_MouseTrap = NewMouseTrap
   PropertyChanged "MouseTrap"

End Property

Public Property Get Orientation() As Orientations
Attribute Orientation.VB_Description = "Returns/sets a thumb wheel orientation, horizontal or vertical."

   Orientation = m_Orientation

End Property

Public Property Let Orientation(ByVal NewOrientation As Orientations)

   If NewOrientation < Horizontal Then NewOrientation = Horizontal
   If NewOrientation > Vertical Then NewOrientation = Vertical
   
   m_Orientation = NewOrientation
   PropertyChanged "Orientation"
   
   Call UserControl_Resize

End Property

Public Property Get ScrollValue() As Integer
Attribute ScrollValue.VB_Description = "Returns/sets the increased or decreased value of a object. Increasing or decreasing is set by 1!"

   ScrollValue = m_Value

End Property

Public Property Let ScrollValue(ByVal NewScrollValue As Integer)

   NewScrollValue = CheckValue(NewScrollValue)
   
   If m_Value <> NewScrollValue Then
      Increment = 1 - (2 And (NewScrollValue > m_Value))
      
      Call WheelMove
   End If
   
   m_Value = NewScrollValue

End Property

Public Property Get ShadeControl() As OLE_COLOR
Attribute ShadeControl.VB_Description = "Returns/sets the control shade color, used to change the object color effect."

   ShadeControl = m_ShadeControl

End Property

Public Property Let ShadeControl(ByVal NewShadeControl As OLE_COLOR)

   m_ShadeControl = NewShadeControl
   PropertyChanged "ShadeControl"
   
   Call UserControl_Resize

End Property

Public Property Get ShadeWheel() As OLE_COLOR
Attribute ShadeWheel.VB_Description = "Returns/sets the wheel shade color, used to change the wheel color effect."

   ShadeWheel = m_ShadeWheel

End Property

Public Property Let ShadeWheel(ByVal NewShadeWheel As OLE_COLOR)

   m_ShadeWheel = NewShadeWheel
   PropertyChanged "ShadeWheel"
   
   Call UserControl_Resize

End Property

Public Property Get SpinOver() As Boolean
Attribute SpinOver.VB_Description = "Determines whether the min/max values are spinover or not."

   SpinOver = m_SpinOver

End Property

Public Property Let SpinOver(NewRollover As Boolean)

   m_SpinOver = NewRollover
   PropertyChanged "SpinOver"

End Property

Public Property Get Value() As Integer
Attribute Value.VB_Description = "Returns/sets the value of a object."

   Value = m_Value

End Property

Public Property Let Value(ByVal NewValue As Integer)

   NewValue = CheckValue(NewValue)
   m_Value = NewValue
   PropertyChanged "Value"
   RaiseEvent Change

End Property

Private Function CheckValue(ByVal IsValue As Integer) As Integer

   If IsValue > m_Max Then
      If m_SpinOver Then
         IsValue = m_Min
         
      Else
         IsValue = m_Max
      End If
      
   ElseIf IsValue < m_Min Then
      If m_SpinOver Then
         IsValue = m_Max
         
      Else
         IsValue = m_Min
      End If
   End If
   
   CheckValue = IsValue

End Function

' draws the thumbwheel
Private Sub DrawWheel()

   With picWheel
      If m_Orientation = Horizontal Then
         StretchBlt .hDC, 0, 0, .ScaleWidth, .ScaleHeight, picTurnHorizontal.hDC, 0, WheelPosition, .ScaleWidth, 1, vbSrcCopy
         BitBlt .hDC, 0, 0, .ScaleWidth, .ScaleHeight, picMaskHorizontal.hDC, 16, 16, vbSrcAnd
         
      Else
         StretchBlt .hDC, 0, 0, .ScaleWidth, .ScaleHeight, picTurnVertical.hDC, WheelPosition, 0, 1, .ScaleHeight, vbSrcCopy
         BitBlt .hDC, 0, 0, .ScaleWidth, .ScaleHeight, picMaskVertical.hDC, 0, 0, vbSrcAnd
      End If
      
      .Refresh
      DoEvents
   End With

End Sub

' draws the moving thumbwheel effect
Private Sub WheelMove()

   If Increment = 0 Then Exit Sub
   
   m_Value = CheckValue(m_Value + Sgn((Increment And (m_Orientation = Horizontal))) - Sgn((Increment And (m_Orientation = Vertical))))
   
   If Not m_SpinOver Then If (m_Value = m_Min) Or (m_Value = m_Max) Then Exit Sub
   
   WheelPosition = WheelPosition + Sgn(Increment)
   WheelPosition = WheelPosition + (8 And (WheelPosition < 0)) - (9 And (WheelPosition > 8))
   
   Call DrawWheel
   
   RaiseEvent Change

End Sub

Private Sub picWheel_Click()

   RaiseEvent Click

End Sub

Private Sub picWheel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim rctTop     As Rect
Dim ptaTopLeft As PointAPI

   If Button <> vbLeftButton Then Exit Sub
   
   If m_MouseTrap Then
      With ptaTopLeft
         GetClientRect picWheel.hWnd, rctTop
         .X = rctTop.Left
         .Y = rctTop.Top
         ClientToScreen picWheel.hWnd, ptaTopLeft
         OffsetRect rctTop, .X, .Y
         ClipCursor rctTop
      End With
   End If
   
   Clicked = True
   LastX = X
   LastY = Y

End Sub

Private Sub picWheel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   picWheel.ToolTipText = Extender.ToolTipText
   
   If Button <> vbLeftButton Then Exit Sub
   
   If Clicked Then
      If m_Orientation = Horizontal Then
         Increment = X - LastX
         
      ' Vertical
      Else
         Increment = Y - LastY
      End If
      
      Call WheelMove
      
      LastX = X
      LastY = Y
   End If

End Sub

Private Sub picWheel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If Button = vbLeftButton Then Clicked = False
   
   ClipCursor ByVal 0

End Sub

Private Sub UserControl_InitProperties()

   m_ShadeControl = vbWhite
   m_ShadeWheel = vbWhite
   m_SpinOver = True

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

   With PropBag
      UserControl.Enabled = .ReadProperty("Enabled", True)
      m_Max = .ReadProperty("Max", 0)
      m_Min = .ReadProperty("Min", 0)
      m_MouseTrap = .ReadProperty("MouseTrap", False)
      m_Orientation = .ReadProperty("Orientation", Horizontal)
      m_ShadeControl = .ReadProperty("ShadeControl", vbWhite)
      m_ShadeWheel = .ReadProperty("ShadeWheel", vbWhite)
      m_SpinOver = .ReadProperty("SpinOver", True)
      m_Value = .ReadProperty("Value", 0)
      picWheel.Visible = True
   End With

End Sub

Private Sub UserControl_Resize()

Static blnBusy As Boolean

Dim imgImage   As Image

   If blnBusy Then Exit Sub
   
   With picMaskHorizontal
      .BackColor = m_ShadeWheel
      .Width = picTurnHorizontal.Width
      .Height = picTurnHorizontal.Width
   End With
   
   With picMaskVertical
      .BackColor = m_ShadeWheel
      .Width = picTurnVertical.Width
      .Height = picTurnVertical.Width
   End With
   
   With shpShadeControl
      .FillColor = m_ShadeControl
      .Width = ScaleWidth
      .Height = ScaleHeight
   End With
   
   m_Orientation = (Vertical And (Height > Width))
   
   If m_Orientation = Horizontal Then
      Set imgImage = imgHorizontal
      imgVertical.Visible = False
      
   ' Vertical
   Else
      Set imgImage = imgVertical
      imgHorizontal.Visible = False
   End If
   
   ' set wheel position
   With imgImage
      .Top = 0
      .Left = 0
      blnBusy = True
      Width = .Width * Screen.TwipsPerPixelX
      Height = .Height * Screen.TwipsPerPixelY
      blnBusy = False
      picWheel.Move 8, 8, .Width - 16, .Height - 16
      .Visible = True
   End With
   
   Set imgImage = Nothing
   
   Call DrawWheel

End Sub

Private Sub UserControl_Terminate()

   ClipCursor ByVal 0

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   With PropBag
      .WriteProperty "Enabled", UserControl.Enabled, True
      .WriteProperty "Max", m_Max, 0
      .WriteProperty "Min", m_Min, 0
      .WriteProperty "MouseTrap", m_MouseTrap, False
      .WriteProperty "Orientation", m_Orientation, Horizontal
      .WriteProperty "ShadeControl", m_ShadeControl, vbWhite
      .WriteProperty "ShadeWheel", m_ShadeWheel, vbWhite
      .WriteProperty "SpinOver", m_SpinOver, True
      .WriteProperty "Value", m_Value, 0
   End With

End Sub
