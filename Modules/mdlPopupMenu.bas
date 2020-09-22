Attribute VB_Name = "mdlPopupMenu"
Option Explicit

' Private Constants
Private Const MF_SEPARATOR   As Long = &H800
Private Const MF_STRING      As Long = &H0
Private Const MIIM_DATA      As Long = &H20
Private Const MIIM_ID        As Long = &H2
Private Const MIIM_STATE     As Long = &H1
Private Const MIIM_TYPE      As Long = &H10
Private Const MIM_BACKGROUND As Long = &H2
Private Const SMNU_SYSTRAY   As Long = 3

' Private Types
Private Type MenuInfo
   cbSize                    As Long
   fMask                     As Long
   dwStyle                   As Long
   cyMax                     As Long
   hbrBack                   As Long
   dwContextHelpID           As Long
   dwMenuData                As Long
End Type

Private Type MenuItemInfo
   cbSize                    As Long
   fMask                     As Long
   fType                     As Long
   fState                    As Long
   wID                       As Long
   hSubMenu                  As Long
   hbmpChecked               As Long
   hbmpUnchecked             As Long
   dwItemData                As Long
   dwTypeData                As String
   cch                       As Long
End Type

' Private Variable
Private PopupMenuHandle      As Long

' Private API's
Private Declare Function CreatePopupMenu Lib "User32" () As Long
Private Declare Function DestroyMenu Lib "User32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemID Lib "User32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function InsertMenuItem Lib "User32" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MenuItemInfo) As Long
Private Declare Function RemoveMenu Lib "User32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function SetMenuInfo Lib "User32" (ByVal hMenu As Long, mi As MenuInfo) As Long
Private Declare Function SetMenuItemBitmaps Lib "User32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long

Public Function PopupMenuTrack(ByVal hWnd As Long, Optional ByVal hMenu As Long) As Long

Dim lngMenu    As Long
Dim ptaMouseXY As PointAPI
Dim rctMenu    As Rect

   If hMenu Then
      lngMenu = hMenu
      
   Else
      lngMenu = PopupMenuHandle
   End If
   
   GetCursorPos ptaMouseXY
   PopupMenuTrack = TrackPopupMenu(lngMenu, TPM_NONOTIFY Or TPM_RETURNCMD Or TPM_RIGHTBUTTON Or TPM_TOPALIGN, ptaMouseXY.X, ptaMouseXY.Y, 0, hWnd, rctMenu)

End Function

Public Function SystemMenuCreate(ByVal hWnd As Long, ByVal hPicture As Long) As Boolean

Const SMNU_MAXIMISE As Integer = 4
Const SMNU_SIZE     As Integer = 2
Const MF_REMOVE     As Long = &H1000&

Dim lngMenuItems    As Long
Dim lngSysMenu      As Long
Dim mniMenu         As MenuInfo
Dim miiItem         As MenuItemInfo

   lngSysMenu = GetSystemMenu(hWnd, 0)
   
   If lngSysMenu Then
      RemoveMenu lngSysMenu, SMNU_MAXIMISE, MF_BYPOSITION Or MF_REMOVE
      RemoveMenu lngSysMenu, SMNU_SIZE, MF_BYPOSITION Or MF_REMOVE
      
      With miiItem
         .cbSize = LenB(miiItem)
         .fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE Or MIIM_DATA
         .fType = MF_STRING
         .fState = MF_ENABLED
         .dwTypeData = AppText(91)
         .cch = Len(.dwTypeData)
         .wID = SC_SYSTRAY
      End With
      
      InsertMenuItem lngSysMenu, SMNU_SYSTRAY, MF_BYPOSITION, miiItem
      SetMenuItemBitmaps lngSysMenu, SC_SYSTRAY, 1, hPicture, hPicture
      
      With mniMenu
         .cbSize = LenB(mniMenu)
         .fMask = MIM_BACKGROUND
         .hbrBack = CreateSolidBrush(&HF9F0E3)
      End With
      
      SetMenuInfo lngSysMenu, mniMenu
      SystemMenuCreate = True
      
      Call SubclassSystemMenu(hWnd)
   End If

End Function

Public Sub PopupMenuCreate(ByVal IsClock As Boolean, ByVal MenuItems As Integer, ByVal OpenName As String)

Const MF_CHECKED   As Long = &H8&
Const MF_DEFAULT   As Long = &H1000
Const MF_SEPARATOR As Long = &H800

Dim intCount       As Integer
Dim lngState       As Long
Dim lngType        As Long
Dim mniMenu        As MenuInfo
Dim miiItem        As MenuItemInfo
Dim MenuText()     As String
Dim TimeType()     As String

   If Not IsClock Then
      If CheckTimeToGo Then
         MenuText = Split(AppText(217), ",")
         TimeType = Split(AppText(218), ",")
         MenuItems = MenuItems + UBound(MenuText) + UBound(TimeType)
         
      Else
         MenuItems = 0
      End If
   End If
   
   ' add a colon at the end of the first menu item
   PopupMenuHandle = CreatePopupMenu
   
   For intCount = 0 To MenuItems
      lngState = MF_ENABLED
      lngType = MF_STRING
      
      If intCount = 0 Then lngState = lngState Or MF_DEFAULT
      If (intCount = 1) Or (intCount = 5) Then lngType = MF_SEPARATOR
      
      If IsClock Then
         If ((intCount > 5) And (AppSettings(SET_LOCKFAVORITS))) Or ((intCount <> 6) And (OpenName = AppText(6))) Then lngState = lngState Or MF_GRAYED
         
      Else
         If ((intCount = 3) And (Format(TimeToGo(0), "yyyymmdd") > Format(Date, "yyyymmdd"))) Or ((intCount = 4) And (Format(TimeToGo(1), "yyyymmdd") < Format(Date, "yyyymmdd"))) Then lngState = lngState Or MF_GRAYED
         If TimeToGoShowType = intCount - 2 Then lngState = lngState Or MF_CHECKED
         If TimeToGoShow = intCount - 6 Then lngState = lngState Or MF_CHECKED
      End If
      
      With miiItem
         If IsClock Then
            .wID = Choose(intCount + 1, 1000, 1050, 1100, 1200, 1300, 1350, 1400, 1500)
            .dwTypeData = AppText(Choose(intCount + 1, 26, 49, 55, 142, 145, 49, 57, 58))
            
            If (intCount = 6) And (OpenName <> AppText(6)) Then lngState = lngState Or MF_GRAYED
            
         Else
            If intCount < 2 Then
               .wID = 2000 + (50 And intCount)
               .dwTypeData = AppText(26 + (23 And intCount))
               
            ElseIf intCount < 6 Then
               If intCount = 5 Then
                  .wID = .wID + 50
                  .dwTypeData = AppText(49)
                  
               Else
                  .wID = 1900 + intCount * 100
                  .dwTypeData = Trim(TimeType(intCount - 2))
               End If
               
            Else
               .wID = 1800 + intCount * 100
               .dwTypeData = Trim(MenuText(intCount - 6))
            End If
         End If
         
         If intCount = 0 Then .dwTypeData = .dwTypeData & ": " & OpenName
         
         .cbSize = LenB(miiItem)
         .fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE Or MIIM_DATA
         .fType = lngType
         .fState = lngState
         .cch = Len(.dwTypeData)
         InsertMenuItem PopupMenuHandle, intCount, 1, miiItem
         
         If IsClock Then
            If (intCount <> 1) And (intCount <> 5) Then SetMenuItemBitmaps PopupMenuHandle, .wID, 1, frmMyTimeZones.imgPopupMenu.Item(intCount).Picture, frmMyTimeZones.imgPopupMenu.Item(intCount).Picture
            
         ElseIf intCount = 0 Then
            SetMenuItemBitmaps PopupMenuHandle, .wID, MF_BYPOSITION, frmMyTimeZones.imgPopupMenu.Item(9).Picture, frmMyTimeZones.imgPopupMenu.Item(9).Picture
         End If
      End With
   Next 'intCount
   
   With mniMenu
      .cbSize = LenB(mniMenu)
      .fMask = MIM_BACKGROUND
      .hbrBack = CreateSolidBrush(&HF9F0E3)
   End With
   
   SetMenuInfo PopupMenuHandle, mniMenu
   Erase MenuText, TimeType

End Sub

Public Sub PopupMenuDestroy()

   DestroyMenu PopupMenuHandle

End Sub

Public Sub SystemMenuEnableSysTrayItem(ByVal hWnd As Long, ByVal Enabled As Boolean)

Dim lngID      As Long
Dim lngSysMenu As Long

   lngSysMenu = GetSystemMenu(hWnd, 0)
   lngID = GetMenuItemID(lngSysMenu, SMNU_SYSTRAY)
   EnableMenuItem lngSysMenu, lngID, Abs(Not Enabled)

End Sub

Public Sub SystemMenuTrack(ByVal hWnd As Long, ByRef MousePointer As MousePointerConstants)

Dim lngParam As Long

   MousePointer = vbDefault
   SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&
   lngParam = PopupMenuTrack(hWnd, GetSystemMenu(hWnd, 0))
   MousePointer = vbSizeAll
   ProcSysMenu hWnd, WM_SYSCOMMAND, lngParam, 0

End Sub
