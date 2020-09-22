Attribute VB_Name = "mdlWinHook"
Option Explicit

' Private Constants
Private Const SC_CLOSE     As Long = &HF060&
Private Const SC_MOVE      As Long = &HF010&

' Private Class
Private Mouse              As clsMouseWheel

' Private Variables
Private PrevProcMouseDown  As Long
Private PrevProcMouseWheel As Long
Private PrevProcPopupMenu  As Long
Private PrevProcSysMenu    As Long

' Private API's
Private Declare Function CallWindowProc Lib "User32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetCursorPos Lib "User32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "User32" (ByVal hHook As Long) As Long

' centered the commondialog window
Public Function CenterCommonDialog(ByVal Message As Long, ByVal hWnd As Long, ByVal ThreadId As Long) As Long

Const HCBT_ACTIVATE As Long = 5
Const HWND_TOP      As Long = 0
Const SWP_NOZORDER  As Long = &H4

Dim rctParent       As Rect
Dim rctClient       As Rect

   If Message = HCBT_ACTIVATE Then
      GetWindowRect hWndParent, rctParent
      GetWindowRect hWnd, rctClient
      
      With rctParent
         .Left = (.Left + ((.Right - .Left) - (rctClient.Right - rctClient.Left)) / 2)
         .Top = (.Top + ((.Bottom - .Top) - (rctClient.Bottom - rctClient.Top)) / 2)
         SetWindowPos hWnd, HWND_TOP, .Left, .Top, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
      End With
      
      UnhookWindowsHookEx WindowHook
   End If

End Function

Public Function ProcSysMenu(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim blnExit As Boolean
Dim lngLeft As Long
Dim lngTop  As Long

   If wMsg = WM_SYSCOMMAND Then
      If wParam = SC_CLOSE Then
         blnExit = True
         
         Call frmMyTimeZones.EndMyTimeZones
         
      ElseIf wParam = SC_MOVE Then
         blnExit = True
         
         With frmMyTimeZones
            lngTop = .Top / Screen.TwipsPerPixelY + .imgMove.Item(4).Top + .imgMove.Item(4).Height
            lngLeft = .Left / Screen.TwipsPerPixelX + .imgMove.Item(4).Left + .imgMove.Item(4).Width / 2
            SetCursorPos lngLeft, lngTop - 12
            DoEvents
            SetCursorPos lngLeft, lngTop - 13
         End With
         
      ElseIf wParam = SC_SYSTRAY Then
         blnExit = True
         
         Call frmMyTimeZones.SysTrayDisplay(True)
      End If
   End If
   
   If blnExit Then
      ProcSysMenu = -1
      
   Else
      ProcSysMenu = CallWindowProc(PrevProcSysMenu, hWnd, wMsg, wParam, lParam)
   End If

End Function

' hook/unhook moue for tracking mouse wheel event
Public Sub SubclassMouseWheel(Optional ByRef IsMe As Object, Optional ByVal hWnd As Long)

   If Mouse Is Nothing Then
      Set Mouse = IsMe
      hWndParent = hWnd
      PrevProcMouseWheel = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf ProcMouseWheel)
      
   Else
      SetWindowLong hWndParent, GWL_WNDPROC, PrevProcMouseWheel
      Set Mouse = Nothing
   End If

End Sub

' subclass the clock specified by hWnd
Public Sub SubclassPopupMenu(ByVal hWnd As Long, Optional ByVal IsClock As Boolean, Optional ByVal MenuItems As Integer, Optional ByVal OpenName As String)

   If MenuItems Then
      PrevProcPopupMenu = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf ProcPopupMenu)
      
      Call PopupMenuCreate(IsClock, MenuItems, OpenName)
      
   Else
      SetWindowLong hWnd, GWL_WNDPROC, PrevProcPopupMenu
      
      Call PopupMenuDestroy
   End If

End Sub

' subclass the system menu from the Form specified by hWnd
Public Sub SubclassSystemMenu(ByVal hWnd As Long)

   If PrevProcSysMenu Then
      SetWindowLong hWnd, GWL_WNDPROC, PrevProcSysMenu
      PrevProcSysMenu = 0
      
   Else
      PrevProcSysMenu = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf ProcSysMenu)
   End If

End Sub

' subclass the textbox specified by hWnd
Public Sub SubclassTextBox(ByVal hWnd As Long)

   If PrevProcMouseDown Then
      SetWindowLong hWnd, GWL_WNDPROC, PrevProcMouseDown
      PrevProcMouseDown = 0
      
   Else
      PrevProcMouseDown = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf ProcMouseDown)
   End If

End Sub

' check if the right mouse button is down
Private Function ProcMouseDown(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

   If wMsg = WM_RBUTTONDOWN Then
      ProcMouseDown = -1
      
   Else
      ProcMouseDown = CallWindowProc(PrevProcMouseDown, hWnd, wMsg, wParam, lParam)
   End If

End Function

' check if mouse wheel is used
Private Function ProcMouseWheel(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

   If wMsg = WM_MOUSEWHEEL Then Call Mouse.WheelUsed(wParam > 0)
   
   ProcMouseWheel = CallWindowProc(PrevProcMouseWheel, hWnd, wMsg, wParam, lParam)

End Function

' check if popupmenu for clock is called
Private Function ProcPopupMenu(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

   ProcPopupMenu = CallWindowProc(PrevProcPopupMenu, hWnd, wMsg, wParam, lParam)

End Function

