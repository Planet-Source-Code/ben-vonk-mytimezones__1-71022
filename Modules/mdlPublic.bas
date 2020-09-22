Attribute VB_Name = "mdlPublic"
Option Explicit

' Public Constants
' ** Also used in controls and forms **
Public Const ALL_MESSAGES           As Long = -1
Public Const CB_FINDSTRING          As Integer = &H14C
Public Const LB_FINDSTRING          As Integer = &H18F
Public Const SOUND_ATTENTION        As Integer = 2
Public Const vbStop                 As Integer = &H50
Public Const BDR_EDGED              As Long = &H16
Public Const BDR_RAISED             As Long = &H5
Public Const BDR_RAISEDINNER        As Long = &H4
Public Const BDR_RAISEDOUTER        As Long = &H1
Public Const BDR_SUNKEN             As Long = &HA
Public Const BF_RIGHT               As Long = &H4
Public Const BF_TOP                 As Long = &H2
Public Const BF_LEFT                As Long = &H1
Public Const BF_BOTTOM              As Long = &H8
Public Const BF_RECT                As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const CODE_LEN               As Long = 200
Public Const DI_NORMAL              As Long = &H3
Public Const ERROR_SUCCESS          As Long = 0
Public Const GMEM_FIXED             As Long = 0
Public Const GWL_EXSTYLE            As Long = -20
Public Const GWL_STYLE              As Long = -16
Public Const GWL_WNDPROC            As Long = -4
Public Const HKEY_CURRENT_USER      As Long = &H80000001
Public Const HTCAPTION              As Long = &H2
Public Const HWND_TOPMOST           As Long = -1
Public Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Public Const KEY_NOTIFY             As Long = &H10
Public Const KEY_QUERY_VALUE        As Long = &H1
Public Const LOCALE_SENGLANGUAGE    As Long = &H1001
Public Const LOCALE_USER_DEFAULT    As Long = &H400
Public Const MF_BYPOSITION          As Long = &H400&
Public Const MF_ENABLED             As Long = &H0
Public Const MF_GRAYED              As Long = &H1
Public Const PATCH_01               As Long = 18
Public Const PATCH_02               As Long = 68
Public Const PATCH_03               As Long = 78
Public Const PATCH_04               As Long = 88
Public Const PATCH_05               As Long = 93
Public Const PATCH_06               As Long = 116
Public Const PATCH_07               As Long = 121
Public Const PATCH_08               As Long = 132
Public Const PATCH_09               As Long = 137
Public Const PATCH_0A               As Long = 186
Public Const REG_SZ                 As Long = 1
Public Const SC_SYSTRAY             As Long = &HF200&
Public Const SWP_FRAMECHANGED       As Long = &H20
Public Const SWP_NOACTIVATE         As Long = &H10
Public Const SWP_NOMOVE             As Long = &H2
Public Const SWP_NOSIZE             As Long = &H1
Public Const SWP_NOZORDER           As Long = &H4
Public Const TPM_NONOTIFY           As Long = &H80
Public Const TPM_RETURNCMD          As Long = &H100
Public Const TPM_RIGHTBUTTON        As Long = &H2
Public Const TPM_TOPALIGN           As Long = &H0
Public Const VER_PLATFORM_WIN32_NT  As Long = 2
Public Const WM_ACTIVATE            As Long = &H6
Public Const WM_LBUTTONDBLCLK       As Long = &H203
Public Const WM_LBUTTONDOWN         As Long = &H201
Public Const WM_LBUTTONUP           As Long = &H202
Public Const WM_NCLBUTTONDOWN       As Long = &HA1
Public Const WM_MOUSEMOVE           As Long = &H200
Public Const WM_MOUSEWHEEL          As Long = &H20A
Public Const WM_PAINT               As Long = &HF
Public Const WM_RBUTTONDOWN         As Long = &H204
Public Const WM_SYSCOMMAND          As Long = &H112&
Public Const WM_THEMECHANGED        As Long = &H31A
Public Const WM_TIMER               As Long = &H113
Public Const WS_BORDER              As Long = &H800000
Public Const WS_EX_CLIENTEDGE       As Long = &H200
Public Const PI_PART                As Single = 3.14159265358979 / 180
Public Const APP_PRODUCTNAME        As String = "MyTimeZones™"
Public Const APP_NUMBER             As String = "001"
Public Const DEFAULT_TIMESERVER     As String = "time-a.nist.gov"
Public Const DISABLE_AUTO_DLTS      As String = "DisableAutoDaylightTimeSet"
Public Const FUNC_CWP               As String = "CallWindowProcA"
Public Const FUNC_EBM               As String = "EbMode"
Public Const FUNC_SWL               As String = "SetWindowLongA"
Public Const MOD_USER               As String = "User32"
Public Const MOD_VBA5               As String = "vba5"
Public Const MOD_VBA6               As String = "vba6"
Public Const SKEY_TIMEZONE          As String = "System\CurrentControlSet\Control\TimeZoneInformation"
Public Const SPECIAL_DAYS           As String = "SpecialDays.dat"
Public Const SUBCLASS_ASM_CODE      As String = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D0000005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D000000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"

' Public Enumerations
Public Enum AppSetting
   SET_SHOWINTROSCREEN
   SET_ACTIVEBORDER
   SET_SHOWTIPTEXT
   SET_ASKCONFIRM
   SET_PLAYSOUND
   SET_MOUSEINTHUMBWHEEL
   SET_LOCKSYSTEMDATE
   SET_LOCKFAVORITS
   SET_CHECKDOUBLEALARMS
   SET_AUTOSTART
   SET_AUTOSYNCHRONISE
   SET_AUTOSAVE
   SET_HEMISPHERE
   SET_MODEZONEMAP
   SET_SHOWCLOCKIMAGE
   SET_AUTODELETETIMETOGO
End Enum

' ** Used in controls **
Public Enum BackStyles
   Transparent
   Opaque
End Enum

Public Enum BorderStyles
   [None]
   [Fixed Single]
End Enum

Public Enum MsgWhen
   MSG_BEFORE = 1
   MSG_AFTER = 2
   MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER
End Enum

Public Enum Orientations
   Horizontal
   Vertical
End Enum

Public Enum Speeds
   Slow
   Default
   Fast
End Enum

Public Enum TimeToGoShowTypes
   TotalTime
   PassedTime
   RemainingTime
End Enum

' Private Enumeration
Private Enum SoundConst
   SND_ASYNC = &H1
   SND_MEMORY = &H4
   SND_NODEFAULT = &H2
End Enum

' Public Types
' ** Also used in controls **
Public Type OSVersionInfo
   dwOSVersionInfoSize            As Long
   dwMajorVersion                 As Long
   dwMinorVersion                 As Long
   dwBuildNumber                  As Long
   dwPlatformId                   As Long
   szCSDVersion                   As String * 128
End Type

' ** Also used in controls **
Public Type PointAPI
   X                              As Long
   Y                              As Long
End Type

' ** Also used in controls **
Public Type Rect
   Left                           As Long
   Top                            As Long
   Right                          As Long
   Bottom                         As Long
End Type

' ** Also used in controls **
Public Type ComboBoxInfo
   cbSize                         As Long
   rcItem                         As Rect
   rcButton                       As Rect
   lStateButton                   As Long
   hWndCombo                      As Long
   hWndEdit                       As Long
   hWndList                       As Long
End Type

' ** Used in controls **
Public Type SubclassDataType
   hWnd                           As Long
   nAddrSclass                    As Long
   nAddrOrig                      As Long
   nMsgCountA                     As Long
   nMsgCountB                     As Long
   aMsgTabelA()                   As Long
   aMsgTabelB()                   As Long
End Type

Public Type SystemTime
   wYear                          As Integer
   wMonth                         As Integer
   wDayOfWeek                     As Integer
   wDay                           As Integer
   wHour                          As Integer
   wMinute                        As Integer
   wSecond                        As Integer
   wMilliseconds                  As Integer
End Type

Public Type TimeZonesInfo
   Bias                           As Long
   StandardBias                   As Long
   DaylightBias                   As Long
   StandardDate                   As SystemTime
   DaylightDate                   As SystemTime
   DisplayName                    As String
   StandardName                   As String
   DaylightName                   As String
   MapID                          As String
End Type

' Private Type
Private Type ClockInfo
   Index                          As Integer
   ZoneID                         As Integer
   DisplayName                    As String
   AlarmTime                      As String
   AlarmMessage                   As String
   AlarmTipText                   As String
   ImageFile                      As String
End Type

' Public Variables
Public AppSettings(15)            As Boolean
Public AutoDaylightTimeSet        As Boolean
Public InIntro                    As Boolean
Public InZoneInfo                 As Boolean
Public InZoneMap                  As Boolean
Public NoDaylightTimeSet          As Boolean
Public NoInternet                 As Boolean
Public ShowSettings               As Boolean
Public TimeZone13                 As Boolean
Public FavoritsInfo(4)            As ClockInfo
Public ZonesInfo(1)               As ClockInfo
Public CalendarDate               As Date
Public DisplaySpeed               As Speeds
Public FirstWeekDay               As Integer
Public IndexFromTimeZone          As Integer
Public IndexSystemTimeZone        As Integer
Public SelectedClock              As Integer
Public SelectedFavorit            As Integer
Public SelectedListIndex          As Integer
Public SelectedPopupMenu          As Integer
Public SelectedTimeZoneID         As Integer
Public SelectedTrayClock          As Integer
Public SelectedZone               As Integer
Public TimeToGoShow               As Integer
Public TimeToWait                 As Integer
Public TotalFavorits              As Integer
Public AnimationhDC               As Long
Public DescriptionWidth           As Long
Public hWndParent                 As Long
Public WindowHook                 As Long
Public GlobeXY                    As PointAPI
Public AddToFontSize              As Single
Public ScreenResize               As Single
Public AppError(51)               As String
Public AppText(261)               As String
Public AppVar(1)                  As String
Public DataPath                   As String
Public DefaultDateFormat          As String
Public ImageName                  As String
Public InfoCopyright              As String
Public InfoVersion                As String
Public LanguageText(7)            As String  ' used for Calendar usercontrol
Public LongDateFormat             As String
Public SpecialDays()              As String
Public TimeServerURL              As String
Public TimeToGo(1)                As String
Public TimeZoneRegKeyName()       As String
Public TimeToGoShowType           As TimeToGoShowTypes
Public AllZones()                 As TimeZonesInfo

' Public API's
' ** Also used in controls and forms **
Public Declare Function CreatePen Lib "GDI32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function BitBlt Lib "GDI32" (ByVal hDCDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CombineRgn Lib "GDI32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function CreateBitmap Lib "GDI32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function CreateCompatibleBitmap Lib "GDI32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hDC As Long) As Long
Public Declare Function CreateRectRgn Lib "GDI32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateRoundRectRgn Lib "GDI32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function CreateSolidBrush Lib "GDI32" (ByVal crColor As Long) As Long
Public Declare Function DeleteDC Lib "GDI32" (ByVal hDC As Long) As Long
Public Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Public Declare Function GetPixel Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function LineTo Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function MoveToEx Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As PointAPI) As Long
Public Declare Function SelectObject Lib "GDI32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function StretchBlt Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function FreeLibrary Lib "Kernel32" (ByVal hLibModule As Long) As Long
Public Declare Function GetModuleHandle Lib "Kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function GetProcAddress Lib "Kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function GetLocaleInfo Lib "Kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Public Declare Function GetVersionEx Lib "Kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVersionInfo) As Long
Public Declare Function GlobalAlloc Lib "Kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "Kernel32" (ByVal hMem As Long) As Long
Public Declare Function LoadLibrary Lib "Kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function OleTranslateColor Lib "OleAut32" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long
Public Declare Function CopyRect Lib "User32" (lpDestRect As Rect, lpSourceRect As Rect) As Long
Public Declare Function DrawEdge Lib "User32" (ByVal hDC As Long, qrc As Rect, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function DrawIconEx Lib "User32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function EnableMenuItem Lib "User32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Public Declare Function GetCaretPos Lib "User32" (lpPoint As PointAPI) As Long
Public Declare Function GetClientRect Lib "User32" (ByVal hWnd As Long, lpRect As Rect) As Long
Public Declare Function GetCursorPos Lib "User32" (lpPoint As PointAPI) As Long
Public Declare Function GetComboBoxInfo Lib "User32" (ByVal hWndCombo As Long, ByRef pcbi As ComboBoxInfo) As Long
Public Declare Function GetDC Lib "User32" (ByVal hWnd As Long) As Long
Public Declare Function GetParent Lib "User32" (ByVal hWnd As Long) As Long
Public Declare Function GetSystemMenu Lib "User32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Public Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowRect Lib "User32" (ByVal hWnd As Long, lpRect As Rect) As Long
Public Declare Function GetWindowThreadProcessId Lib "User32" (ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long
Public Declare Function InflateRect Lib "User32" (lpRect As Rect, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function KillTimer Lib "User32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Public Declare Function PostMessage Lib "User32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function PtInRect Lib "User32" (Rect As Rect, ByVal lPtX As Long, ByVal lPtY As Long) As Integer
Public Declare Function ReleaseCapture Lib "User32" () As Long
Public Declare Function ScreenToClient Lib "User32" (ByVal hWnd As Long, lpPoint As PointAPI) As Long
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
Public Declare Function SetCapture Lib "User32" (ByVal hWnd As Long) As Long
Public Declare Function SetForegroundWindow Lib "User32" (ByVal hWnd As Long) As Long
Public Declare Function SetRect Lib "User32" (lpRect As Rect, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetTimer Lib "User32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowLongA Lib "User32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SetWindowRgn Lib "User32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function TrackPopupMenu Lib "User32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hWnd As Long, lprc As Rect) As Long
Public Declare Function WindowFromPoint Lib "User32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function CloseThemeData Lib "UxTheme" (ByVal lngTheme As Long) As Long
Public Declare Function DrawThemeBackground Lib "UxTheme" (ByVal hTheme As Long, ByVal lhDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As Rect, pClipRect As Rect) As Long
Public Declare Function GetCurrentThemeName Lib "UxTheme" (ByVal pszThemeFileName As Long, ByVal cchMaxNameChars As Long, ByVal pszColorBuff As Long, ByVal cchMaxColorChars As Long, ByVal pszSizeBuff As Long, ByVal cchMaxSizeChars As Long) As Long
Public Declare Function GetThemeDocumentationProperty Lib "UxTheme" (ByVal pszThemeName As Long, ByVal pszPropertyName As Long, ByVal pszValueBuff As Long, ByVal cchMaxValChars As Long) As Long
Public Declare Function OpenThemeData Lib "UxTheme" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Public Declare Function SoundPlay Lib "WinMM" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

' Private API's
Private Declare Function GetSystemTime Lib "Kernel32" (lpSystemTime As SystemTime) As Long
Private Declare Function StrLen Lib "Kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function SafeArrayGetDim Lib "OleAut32" (ByRef saArray() As Any) As Long

' returns True if Windows is themed
Public Function CheckIsThemed(ByRef IsThemedWindows As Boolean) As Boolean

Dim lngLibrary              As Long
Dim osvInfo                 As OSVersionInfo
Dim strTheme                As String
Dim strName                 As String

   IsThemedWindows = False
   
   With osvInfo
      .dwOSVersionInfoSize = Len(osvInfo)
      GetVersionEx osvInfo
      
      If .dwPlatformId = VER_PLATFORM_WIN32_NT Then
         If ((.dwMajorVersion > 4) And .dwMinorVersion) Or (.dwMajorVersion > 5) Then
            IsThemedWindows = True
            lngLibrary = LoadLibrary("UXTheme")
            
            If lngLibrary Then
               strTheme = String(255, vbNullChar)
               GetCurrentThemeName StrPtr(strTheme), Len(strTheme), 0, 0, 0, 0
               strTheme = StripNull(strTheme)
               
               If Len(strTheme) Then
                  strName = String(255, vbNullChar)
                  GetThemeDocumentationProperty StrPtr(strTheme), StrPtr("ThemeName"), StrPtr(strName), Len(strName)
                  CheckIsThemed = (StripNull(strName) <> "")
               End If
               
               FreeLibrary lngLibrary
            End If
         End If
      End If
   End With

End Function

' sets text with first letter in caps
Public Function CapsText(ByVal Text As String) As String

   CapsText = UCase(Left(Text, 1)) & Mid(Text, 2)

End Function

' returns True if the dates are set
Public Function CheckTimeToGo() As Boolean

   CheckTimeToGo = (Len(TimeToGo(0)) And Len(TimeToGo(1)))

End Function

' create message for clock alarm
Public Function CreateAlarmMessage(ByVal Index As Integer, Optional ByVal Text As String) As String

Dim intCount  As Integer
Dim strBuffer As String

   If Index < 0 Then Exit Function
   If Text = "" Then Text = FavoritsInfo(Index).AlarmMessage
   
   For intCount = 0 To 1
      If intCount Then
         strBuffer = Format(frmMyTimeZones.clkFavorits.Item(Index).DateTime, "hh:mm:ss")
         
      Else
         strBuffer = CapsText(Format(frmMyTimeZones.clkFavorits.Item(Index).DateTime, "Long Date"))
      End If
      
      Text = Replace(Text, AppVar(intCount), strBuffer)
   Next 'intCount
   
   CreateAlarmMessage = Text

End Function

' returns the leddisplay tooltiptext for specified clock
Public Function GetDisplayToolTipText(ByVal ClockName As String) As String

   GetDisplayToolTipText = GetToolTipText(AppText(97) & " " & ClockName)

End Function

' returns the specified day of the month
Public Function GetGivenMonthDay(ByVal IsYear As Integer, ByVal IsDate As String) As Date

Dim dteDate    As Date
Dim intCount   As Integer
Dim intDay     As Integer
Dim intWeekDay As Integer

   intWeekDay = WeekDay("01-" & Right(IsDate, 2) & "-" & IsYear, vbMonday)
   intDay = (InStr("MOTUWETHFRSASU", UCase(Left(IsDate, 2))) + 1) / 2
   intCount = Val(Mid(IsDate, 3, 1))
   
   If (intCount = 0) Or (intCount > 4) Then
      dteDate = DateAdd("d", -1, DateSerial(IsYear, CInt(Right(IsDate, 2)) + 1, 1))
      intWeekDay = WeekDay(dteDate, vbMonday)
      intCount = intDay - intWeekDay - (7 And (intDay > intWeekDay))
      intDay = Day(dteDate) + intCount
      
   Else
      intDay = intDay + (intCount - (1 And (intDay >= intWeekDay))) * 7 - intWeekDay + 1
   End If
   
   GetGivenMonthDay = CDate(CStr(intDay) & "-" & Right(IsDate, 2) & "-" & CStr(IsYear))

End Function

' returns the icon for a button specified by Index
Public Function GetIcon(ByVal Index As Integer, Optional ByVal Locked As Boolean) As Integer

   If Index = 1 Then
      GetIcon = Index + (27 And (NoInternet Or (TimeServerURL = AppText(198))))
      
   ElseIf Index = 2 Then
      GetIcon = Index + (9 And AppSettings(SET_LOCKSYSTEMDATE))
      
   ElseIf Index = 3 Then
      GetIcon = Index + (9 And AutoDaylightTimeSet)
      
   ElseIf Index = 7 Then
      GetIcon = Index + (7 And (AppSettings(SET_LOCKFAVORITS) Or Locked))
      
   ElseIf Index = 8 Then
      GetIcon = Index + (13 And (AppSettings(SET_LOCKFAVORITS) Or Locked))
      
   Else
      GetIcon = Index
   End If

End Function

' returns the listindex of the specified listbox
Public Function GetListIndex(ByVal hWnd, ByVal SearchType As Long, ByVal SearchValue As String)

   GetListIndex = SendMessage(hWnd, SearchType, -1, ByVal SearchValue)

End Function

' returns the part from given source
Public Function GetNamePart(ByVal Source As String, ByVal Index As Integer) As String

   GetNamePart = Trim(Split(Source, ",")(Index - 1))

End Function

' returns the selected part of the date
Public Function GetSelectedDatePart(ByVal KeyCode As Integer, ByVal Part As Integer, ByVal Max As Integer, ByVal Min As Integer) As Integer

   If (KeyCode = vbKeyLeft) Or (KeyCode = vbKeyRight) Then
      Part = Part - (1 And (KeyCode = vbKeyLeft)) + (1 And (KeyCode = vbKeyRight))
      
      If Part > Max Then Part = Min
      If Part < Min Then Part = Max
      
      GetSelectedDatePart = Part
   End If

End Function

' returns the system date
Public Function GetSystemDate() As Date

Dim sysTime As SystemTime

   With sysTime
      GetSystemTime sysTime
      GetSystemDate = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
   End With

End Function

' returns the displayname of the selected timezone
Public Function GetTimeZoneText(ByVal Index As Integer, Optional ByVal NamePart As Boolean) As String

   If NamePart Then
      GetTimeZoneText = Trim(Split(AllZones(Index).DisplayName, ")", 2)(1))
      
   Else
      GetTimeZoneText = Mid(AllZones(Index).DisplayName, 2, InStr(AllZones(Index).DisplayName, ")") - 2)
   End If

End Function

' returns the tooltiptext given by text
Public Function GetToolTipText(ByVal Text As String) As String

   ' if SET_SHOWTIPTEXT = False then no tooltiptext
   If Not AppSettings(SET_SHOWTIPTEXT) Or (Text = "") Then Exit Function
   
   GetToolTipText = " " & Trim(Text) & " "

End Function

' check if <Alt> + <F4> keys are pressed
Public Function IsExit(ByVal KeyCode As Integer, ByVal Shift As Integer) As Boolean

   IsExit = (((Shift And vbAltMask) > 0) And (KeyCode = vbKeyF4))

End Function

' convert the local date to UTC
Public Function LocalDateToUTC(ByRef DateTime As Date, ByRef TimeZone As TimeZonesInfo) As Date

   LocalDateToUTC = DateSerial(Year(DateTime), Month(DateTime), Day(DateTime)) + TimeSerial(Hour(DateTime), Minute(DateTime) + (TimeZone.Bias + (TimeZone.DaylightBias And IsDayLight(DateTime, TimeZone))), Second(DateTime))

End Function

' check if array SpecialDays() exist
Public Function NoSpecialDays() As Boolean

   NoSpecialDays = (SafeArrayGetDim(SpecialDays) = 0)
   
   If Not NoSpecialDays Then NoSpecialDays = (SpecialDays(0) = "")

End Function

' strip nulls of given string
Public Function StripNull(ByVal Text As String) As String

   StripNull = Left(Text, StrLen(StrPtr(Text)))

End Function

' returns the name of the clock without brackets
Public Function TrimClockName(ByVal ClockName As String) As String

   If (Left(ClockName, 1) = "[") Or (Left(ClockName, 1) = "(") Then ClockName = Mid(ClockName, 2)
   If (Right(ClockName, 1) = "]") Or (Right(ClockName, 1) = ")") Then ClockName = Left(ClockName, Len(ClockName) - 1)
   
   TrimClockName = Trim(ClockName)

End Function

' convert UTC to local date
Public Function UTCToLocalDate(ByRef DateTime As Date, ByRef TimeZone As TimeZonesInfo) As Date

Dim dteDateTime As Date

   dteDateTime = DateSerial(Year(DateTime), Month(DateTime), Day(DateTime)) + TimeSerial(Hour(DateTime), Minute(DateTime) - TimeZone.Bias - TimeZone.DaylightBias, Second(DateTime))
   
   If Not IsDayLight(dteDateTime, TimeZone) Then dteDateTime = DateSerial(Year(DateTime), Month(DateTime), Day(DateTime)) + TimeSerial(Hour(DateTime), Minute(DateTime) - TimeZone.Bias, Second(DateTime))
   
   UTCToLocalDate = dteDateTime

End Function

' convert day of the week to a date
Public Function WeekDayToDate(ByVal Years As Integer, ByVal Months As Integer, ByVal WeekDays As Integer, ByVal Days As Integer) As Date

Dim dteDate  As Date
Dim intCount As Integer
Dim intDays  As Integer

   For intCount = 1 To 31
      dteDate = DateSerial(Years, Months, intCount)
      
      If Month(dteDate) > Months Then Exit For
      
      If WeekDays = WeekDay(dteDate) - 1 Then
         intDays = intDays + 1
         WeekDayToDate = dteDate
         
         If intDays = Days Then Exit For
      End If
   Next 'intCount

End Function

' check if date end is bigger than date begin
Public Sub CheckDatesTimeToGo(ByVal FillLabels As Boolean, Optional ByRef Label As Object)

Dim strBuffer As String

   If Len(TimeToGo(0)) And Len(TimeToGo(1)) Then
      If Format(TimeToGo(0), "yyyymmdd") > Format(TimeToGo(1), "yyyymmdd") Then
         strBuffer = TimeToGo(0)
         TimeToGo(0) = TimeToGo(1)
         TimeToGo(1) = strBuffer
         
         If FillLabels Then
            Label.Item(0).Caption = Format(TimeToGo(0), "d mmmm yyyy")
            Label.Item(1).Caption = Format(TimeToGo(1), "d mmmm yyyy")
         End If
         
      ElseIf TimeToGo(0) = TimeToGo(1) Then
         If FillLabels Then TimeToGo(1) = ""
      End If
   End If

End Sub

Public Sub DisableDisplay(ByRef Display As Object)

   With Display
      .BackColor = &HE8EAED
      .NoTextScrolling = False
      .Text = ""
      .ToolTipText = ""
   End With

End Sub

' sets date for the specified calendar
Public Sub SetCalendarDate(ByRef Calendar As Object, ByVal IsDate As Date)

   With Calendar
      .CalYear = Year(IsDate)
      .CalMonth = Month(IsDate)
      .CalDay = Day(IsDate)
   End With

End Sub

' play sound
Public Sub PlaySound(ByVal Index As Integer)

Dim strSoundBuffer As String

   If Not AppSettings(SET_PLAYSOUND) Then Exit Sub
   
   On Local Error GoTo GiveBeep
   strSoundBuffer = StrConv(LoadResData(Index, "Sounds"), vbUnicode)
   SoundPlay strSoundBuffer, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY
   
   GoTo ExitSub
   
GiveBeep:
   Beep
   
ExitSub:
   On Local Error GoTo 0

End Sub

' check timezones daylight setting
Private Function IsDayLight(ByRef DateTime As Date, ByRef TimeZone As TimeZonesInfo) As Boolean

Dim dteDayLightBegin As Date
Dim dteDayLightEnd   As Date

   With TimeZone.DaylightDate
      If .wYear Then
         dteDayLightBegin = DateSerial(Year(DateTime), .wMonth, .wDay)
         
      Else
         dteDayLightBegin = WeekDayToDate(Year(DateTime), .wMonth, .wDayOfWeek, .wDay)
      End If
      
      dteDayLightBegin = dteDayLightBegin + TimeSerial(.wHour, .wMinute, .wSecond)
   End With
   
   With TimeZone.StandardDate
      If .wYear Then
         dteDayLightEnd = DateSerial(Year(DateTime), .wMonth, .wDay)
         
      Else
         dteDayLightEnd = WeekDayToDate(Year(DateTime), .wMonth, .wDayOfWeek, .wDay)
      End If
      
      dteDayLightEnd = dteDayLightEnd + TimeSerial(.wHour, .wMinute, .wSecond)
   End With
   
   If dteDayLightBegin = dteDayLightEnd Then Exit Function
   
   If dteDayLightBegin < dteDayLightEnd Then
      If (DateTime > dteDayLightBegin) And (DateTime < dteDayLightEnd) Then IsDayLight = True
      
   ' Australia
   ElseIf (DateTime < dteDayLightEnd) Or (DateTime > dteDayLightBegin) Then
      IsDayLight = True
   End If

End Function
