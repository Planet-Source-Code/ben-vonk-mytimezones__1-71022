Attribute VB_Name = "mdlInitialize"
Option Explicit

' Private Constants
Private Const BUFFER_SIZE             As Long = 2048
Private Const HKEY_LOCAL_MACHINE      As Long = &H80000002
Private Const KEY_CREATE_LINK         As Long = &H20
Private Const KEY_CREATE_SUB_KEY      As Long = &H4
Private Const KEY_SET_VALUE           As Long = &H2
Private Const KEY_ALL_ACCESS          As Long = KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK
Private Const REG_DWORD               As Long = 4
Private Const REG_APP_KEY             As String = "Software\MyTimeZones"

' Private Types
Private Type TagInitCommonControlsEx
   lSize As Long
   lICC As Long
End Type

Private Type TimeZoneInfo
   Bias                               As Long
   StandardName(31)                   As Integer
   StandardDate                       As SystemTime
   StandardBias                       As Long
   DaylightName(31)                   As Integer
   DaylightDate                       As SystemTime
   DaylightBias                       As Long
End Type

' Private Variables
Private AutoStartIsChanged            As Boolean
Private RegSubKey(7)                  As String
Private RegVariable(31)               As String

' Private API's
Private Declare Function RegCloseKey Lib "AdvApi32" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "AdvApi32" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCreateKeyEx Lib "AdvApi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteValue Lib "AdvApi32" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegEnumKey Lib "AdvApi32" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegFlushKey Lib "AdvApi32" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "AdvApi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "AdvApi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "AdvApi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function InitCommonControlsEx Lib "ComCtl32" (iccex As TagInitCommonControlsEx) As Boolean
Private Declare Function GetTimeZoneInformation Lib "Kernel32" (lpTimeZoneInformation As TimeZoneInfo) As Long
Private Declare Function SetTimeZoneInformation Lib "Kernel32" (lpTimeZoneInformation As TimeZoneInfo) As Long

' returns the local current timezone
Public Function GetCurrentTimeZone() As String

Dim tziTempZone As TimeZoneInfo

   GetTimeZoneInformation tziTempZone
   GetCurrentTimeZone = SetName(tziTempZone.StandardName)

End Function

' returns the value of the key
Public Function GetRegKeyValue(ByVal SubKey As String, ByVal ValueName As String) As String

Dim lngKey    As Long
Dim strBuffer As String

   strBuffer = Space(BUFFER_SIZE)
   
   If RegOpenKeyEx(HKEY_LOCAL_MACHINE, SubKey, 0, KEY_ALL_ACCESS, lngKey) Then Exit Function
   If RegQueryValueEx(lngKey, ValueName, 0, REG_SZ, ByVal strBuffer, BUFFER_SIZE) = ERROR_SUCCESS Then GetRegKeyValue = Trim(StripNull(strBuffer))
   
   RegCloseKey lngKey

End Function

' store timezoneinfo into the system
Public Function SetTimeZoneInfo(ByVal Index As Integer) As Boolean

Const TIME_ZONE_ID_DAYLIGHT As Long = 2

Dim lngKey                  As Long
Dim lngValue                As Long
Dim strBuffer               As String
Dim tziTempZone             As TimeZoneInfo

   On Local Error GoTo ExitFunction
   SetTimeZoneInfo = False
   
   If RegOpenKeyEx(HKEY_LOCAL_MACHINE, SKEY_TIMEZONE, 0, KEY_ALL_ACCESS, lngKey) = ERROR_SUCCESS Then
      If RegQueryValueEx(lngKey, DISABLE_AUTO_DLTS, 0, REG_SZ, ByVal strBuffer, BUFFER_SIZE) = ERROR_SUCCESS Then RegDeleteValue lngKey, DISABLE_AUTO_DLTS
      
      With AllZones(Index)
         tziTempZone.Bias = .Bias
         tziTempZone.StandardDate = .StandardDate
         tziTempZone.StandardBias = .StandardBias
         tziTempZone.DaylightDate = .DaylightDate
         tziTempZone.DaylightBias = .DaylightBias
         SetName tziTempZone.StandardName, .StandardName
         SetName tziTempZone.DaylightName, .DaylightName
      End With
      
      SetTimeZoneInformation tziTempZone
      
      If Not AutoDaylightTimeSet And Not NoDaylightTimeSet Then
         strBuffer = Space(BUFFER_SIZE)
         RegQueryValueEx lngKey, DISABLE_AUTO_DLTS, 0, REG_SZ, ByVal strBuffer, BUFFER_SIZE
         
         If RegQueryValueEx(lngKey, DISABLE_AUTO_DLTS, 0, REG_SZ, ByVal strBuffer, BUFFER_SIZE) = TIME_ZONE_ID_DAYLIGHT Then
            lngValue = 1
            RegSetValueEx lngKey, DISABLE_AUTO_DLTS, 0, REG_DWORD, lngValue, LenB(lngValue)
         End If
      End If
      
      RegCloseKey lngKey
      SetTimeZoneInfo = True
   End If
   
ExitFunction:
   On Local Error GoTo 0

End Function

' initialised the CommonControls for GetOpenFileName
Public Sub InitialiseCommonControls()

Const ICC_USEREX_CLASSES As Long = &H200

Dim ticControlsEx        As TagInitCommonControlsEx

   With ticControlsEx
      .lSize = LenB(ticControlsEx)
      .lICC = ICC_USEREX_CLASSES
   End With
   
   On Error Resume Next
   InitCommonControlsEx ticControlsEx
   On Error GoTo 0

End Sub

' store all registry timezones
Public Sub GetAllTimeZones()

Dim strKey As String

   strKey = GetRegKeyCurrentVersion & "Time Zones"
   
   If Not GetTimeZoneCollection(strKey) Then
      ShowMessage AppError(1) & " [HKey_Local_Machine] " & strKey, vbStop, AppError(2), AppError(0), TimeToWait
      End
   End If

End Sub

' get the MyTimeZones registry values
Public Sub GetRegValues()

Dim intCount As Integer
Dim strKey   As String

   ' Settings (Boolean)
   strKey = REG_APP_KEY & "\" & RegSubKey(0)
   
   ' 0 = ShowIntroScreen,   1 = ActiveBorder,     2 = ShowTipText
   ' 3 = AskConfirm,        4 = PlaySound,        5 = MouseInThumbWheel
   ' 6 = LockSystemDate,    7 = LockFavorits,     8 = CheckDoubleAlarms
   ' 9 = AutoStart,        10 = AutoSynchronise, 11 = AutoSave
   '12 = Hemisphere,       13 = ModeZoneMap,     14 = ShowClockImage
   '15 = AutoDeleteTimeToGo
   For intCount = 0 To 15
      GetRegValue strKey, RegVariable(intCount), AppSettings(intCount), Not ((intCount = SET_MOUSEINTHUMBWHEEL) Or (intCount = SET_LOCKSYSTEMDATE) Or (intCount = SET_LOCKFAVORITS) Or (intCount = SET_AUTOSTART) Or (intCount = SET_AUTOSYNCHRONISE) Or (intCount = SET_HEMISPHERE) Or (intCount = SET_AUTODELETETIMETOGO))
   Next 'intCount
   
   ' Settings (Integer)
   ' FirstWeekDay, SelectedZone, SelectedFavorit, SelectedClock
   ' TimeServerID, TimeToGoShow, TimeToGoShowType
   AutoStartIsChanged = AppSettings(SET_AUTOSTART)
   GetRegValue strKey, RegVariable(16), FirstWeekDay, 1
   GetRegValue strKey, RegVariable(17), SelectedFavorit, 0
   GetRegValue strKey, RegVariable(18), SelectedZone, 0
   GetRegValue strKey, RegVariable(19), SelectedClock, 0
   GetRegValue strKey, RegVariable(20), SelectedTrayClock, 0
   GetRegValue strKey, RegVariable(21), TimeServerURL, DEFAULT_TIMESERVER
   GetRegValue strKey, RegVariable(22), TimeToGoShow, -1
   GetRegValue strKey, RegVariable(23), TimeToGoShowType, 2
   ' Settings (String)
   ' TimeToGo(0) = From, TimeToGo(1) = To
   GetRegValue strKey, RegVariable(24), TimeToGo(0), ""
   GetRegValue strKey, RegVariable(25), TimeToGo(1), ""
   
   Call CheckDatesTimeToGo(False)
   
   ' FromZone = 0, ToZone = 1
   For intCount = 0 To 1
      strKey = REG_APP_KEY & "\" & RegSubKey(intCount + 1)
      
      ' Index, ZoneID, ImageFile
      With ZonesInfo(intCount)
         GetRegValue strKey, RegVariable(26), .DisplayName, ""
         GetRegValue strKey, RegVariable(27), .ImageFile, ""
         .Index = GetClockIndex(.DisplayName, IndexSystemTimeZone, False)
         .ZoneID = GetClockIndex(.DisplayName, IndexSystemTimeZone, True)
      End With
   Next 'intCount
   
   ' Favorits\Clock_1 To Favorits\Clock_5
   For intCount = 0 To 4
      strKey = REG_APP_KEY & "\" & RegSubKey(intCount + 3)
      
      With FavoritsInfo(intCount)
         GetRegValue strKey, RegVariable(28), .DisplayName, ""
         GetRegValue strKey, RegVariable(29), .AlarmTime, ""
         GetRegValue strKey, RegVariable(30), .AlarmMessage, ""
         GetRegValue strKey, RegVariable(31), .ImageFile, ""
         .Index = GetClockIndex(.DisplayName, -1, False)
         .ZoneID = GetClockIndex(.DisplayName, -1, True)
         
         If InStr(.AlarmTime, ":") Then .AlarmTipText = AppText(0)
         If .Index > -1 Then TotalFavorits = TotalFavorits + 1
      End With
   Next 'intCount

End Sub

' Loads the special day file
Public Sub LoadSpecialDays()

Dim blnLoadData As Boolean
Dim intCount    As Integer
Dim intFileLoad As Integer
Dim strBuffer() As String
Dim strData     As String

   If Dir(DataPath & SPECIAL_DAYS) = "" Then Exit Sub
   
   On Local Error GoTo ExitSub
   intFileLoad = FreeFile
   
   Open DataPath & SPECIAL_DAYS For Input As #intFileLoad
      Do While Not EOF(intFileLoad)
         Line Input #intFileLoad, strData
         
         If blnLoadData Then
            strBuffer = Split(strData, "=", 2)
            
            If UCase(Trim(strBuffer(0))) = "DAY" Then
               ReDim Preserve SpecialDays(intCount) As String
               
               SpecialDays(intCount) = Trim(strBuffer(1))
               intCount = intCount + (1 And (intCount < 99))
            End If
         End If
         
         If Trim(UCase(strData)) = "[SPECIALDAYS]" Then
            blnLoadData = True
            intCount = 0
         End If
      Loop
   Close #intFileLoad
   
ExitSub:
   On Local Error GoTo 0
   Close #intFileLoad
   Erase strBuffer

End Sub

' set default parameters
Public Sub SetDefaults()

Const LOCALE_SDATE      As Long = &H1D
Const LOCALE_SLONGDATE  As Long = &H20
Const LOCALE_SSHORTDATE As Long = &H1F

Dim intCount            As Integer
Dim strBuffer()         As String
Dim strSeparator        As String

   AppVar(0) = "%Date%"
   AppVar(1) = "%Time%"
   TimeToWait = 5000
   LongDateFormat = Replace(UCase(Replace(GetLocaleInformation(LOCALE_SLONGDATE), "'", "")), "DE ", "") ' delete ' and 'de ' in spanish windows
   strSeparator = GetLocaleInformation(LOCALE_SDATE)
   strBuffer = Split(LCase(GetLocaleInformation(LOCALE_SSHORTDATE)), strSeparator, 3)
   DefaultDateFormat = String(2, Asc(Left(strBuffer(0), 1))) & strSeparator & String(2, Asc(Left(strBuffer(1), 1))) & strSeparator & String(4, Asc(Left(strBuffer(2), 1)))
   Erase strBuffer
   
   Call GetLanguage(GetLocaleInformation(LOCALE_SENGLANGUAGE))
   
   AppText(49) = "/separator/"
   
   For intCount = 1 To 12
      If intCount < 8 Then
         ' set the weekday names
         LanguageText(0) = LanguageText(0) & StrConv(WeekdayName(intCount, , vbSunday), vbProperCase) & IIf(intCount < 7, ",", "")
      End If
      
      ' set the month names
      LanguageText(2) = LanguageText(2) & StrConv(MonthName(intCount), vbProperCase) & IIf(intCount < 12, ",", "")
   Next 'intCount

End Sub

' fill registry keys and variables
Public Sub SetRegKeys()

Dim intCount As Integer

   ' RegSubKeys
   RegSubKey(0) = "Settings"
   RegSubKey(1) = "FromZone"
   RegSubKey(2) = "ToZone"
   
   For intCount = 3 To 7
      RegSubKey(intCount) = "Favorits\Clock_" & intCount - 2
   Next 'intCount
   
   ' Fields Settings (Boolean)
   RegVariable(0) = "ShowIntroScreen"
   RegVariable(1) = "ActiveBorder"
   RegVariable(2) = "ShowTipText"
   RegVariable(3) = "AskConfirm"
   RegVariable(4) = "PlaySound"
   RegVariable(5) = "LockSystemDate"
   RegVariable(6) = "MouseInThumbWheel"
   RegVariable(7) = "LockFavorits"
   RegVariable(8) = "CheckDoubleAlarms"
   RegVariable(9) = "AutoStart"
   RegVariable(10) = "AutoSynchronise"
   RegVariable(11) = "AutoSave"
   RegVariable(12) = "Hemispher"
   RegVariable(13) = "ModeZoneMap"
   RegVariable(14) = "ShowClockImage"
   RegVariable(15) = "AutoDeleteTimeToGo"
   ' Fields Settings (Integer)
   RegVariable(16) = "FirstWeekDay"
   RegVariable(17) = "SelectedZone"
   RegVariable(18) = "SelectedFavorit"
   RegVariable(19) = "SelectedClock"
   RegVariable(20) = "SelectedTrayClock"
   RegVariable(21) = "TimeServerID"
   RegVariable(22) = "TimeToGoShow"
   RegVariable(23) = "TimeToGoShowType"
   ' Fields Settings (String)
   RegVariable(24) = "TimeToGoFrom"
   RegVariable(25) = "TimeToGoTo"
   ' Fields FromZone And ToZone
   RegVariable(26) = "DisplayName"
   RegVariable(27) = "ImageFile"
   ' Fields Favorits\Clock_1 To Favorits\Clock_5
   RegVariable(28) = "DisplayName"
   RegVariable(29) = "AlarmTime"
   RegVariable(30) = "AlarmMessage"
   RegVariable(31) = "ImageFile"
   ' set ShowIntroScreen value
   GetRegValue REG_APP_KEY & "\" & RegSubKey(0), RegVariable(0), AppSettings(SET_SHOWINTROSCREEN), True

End Sub

' set the MyTimeZones registry values
Public Sub SetRegValues()

Dim intCount As Integer
Dim strKey   As String

   If Not ExistKey(REG_APP_KEY) Then CreateKey REG_APP_KEY
   
   For intCount = 0 To UBound(RegSubKey)
      strKey = REG_APP_KEY & "\" & RegSubKey(intCount)
      
      If Not ExistKey(strKey) Then CreateKey strKey
   Next 'intCount
   
   ' Settings (Boolean)
   strKey = REG_APP_KEY & "\" & RegSubKey(0)
   
   ' 0 = ShowIntroScreen,   1 = ActiveBorder,     2 = ShowTipText
   ' 3 = AskConfirm,        4 = PlaySound,        5 = MouseInThumbWheel
   ' 6 = LockSystemDate,    7 = LockFavorits,     8 = CheckDoubleAlarms
   ' 9 = AutoStart,        10 = AutoSynchronise, 11 = AutoSave
   '12 = Hemisphere,       13 = ModeZoneMap,     14 = ShowClockImage
   '15 = AutoDeleteTimeToGo
   For intCount = 0 To 15
      SetRegValue strKey, RegVariable(intCount), AppSettings(intCount)
   Next 'intCount
   
   ' Settings (Integer)
   ' FirstWeekDay, SelectedZone, SelectedFavorit, SelectedClock
   ' TimeServerID, TimeToGoShow, TimeToGoShowType
   SetRegValue strKey, RegVariable(16), FirstWeekDay
   SetRegValue strKey, RegVariable(17), SelectedFavorit
   SetRegValue strKey, RegVariable(18), SelectedZone
   SetRegValue strKey, RegVariable(19), SelectedClock
   SetRegValue strKey, RegVariable(20), SelectedTrayClock
   SetRegValue strKey, RegVariable(21), TimeServerURL
   SetRegValue strKey, RegVariable(22), TimeToGoShow
   SetRegValue strKey, RegVariable(23), TimeToGoShowType
   ' Settings (String)
   ' TimeToGo(0) = From, TimeToGo(1) = To
   SetRegValue strKey, RegVariable(24), TimeToGo(0)
   SetRegValue strKey, RegVariable(25), TimeToGo(1)
   
   ' FromZone = 0, ToZone = 1
   For intCount = 0 To 1
      strKey = REG_APP_KEY & "\" & RegSubKey(intCount + 1)
      
      ' Index, ZoneID, ImageFile
      With ZonesInfo(intCount)
         SetRegValue strKey, RegVariable(26), .DisplayName
         SetRegValue strKey, RegVariable(27), .ImageFile
      End With
   Next 'intCount
   
   If AppSettings(SET_LOCKFAVORITS) Then Exit Sub
   
   ' Favorits\Clock_1 To Favorits\Clock_5
   For intCount = 0 To 4
      strKey = REG_APP_KEY & "\" & RegSubKey(intCount + 3)
      
      With FavoritsInfo(intCount)
         SetRegValue strKey, RegVariable(28), .DisplayName
         SetRegValue strKey, RegVariable(29), .AlarmTime
         SetRegValue strKey, RegVariable(30), .AlarmMessage
         SetRegValue strKey, RegVariable(31), .ImageFile
      End With
   Next 'intCount
   
   ' Create or Delete HKLM_Software\Windows\CurrentVersion\Run\MyTimeZones
   If AppSettings(SET_AUTOSTART) <> AutoStartIsChanged Then AddDeleteAutoRunKey AppSettings(SET_AUTOSTART)

End Sub

' adds or delete the autorun key in the registry
Private Function AddDeleteAutoRunKey(ByVal Add As Boolean) As Boolean

Dim lngKey     As Long
Dim strAppPath As String

   On Local Error GoTo ExitFunction
   RegCreateKey HKEY_LOCAL_MACHINE, GetRegKeyCurrentVersion(False) & "Run", lngKey
   strAppPath = Chr(34) & App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & App.EXEName & ".exe" & Chr(34) & " /Silent"
   
   If Add Then
      RegSetValueEx lngKey, App.Title, 0, REG_SZ, ByVal strAppPath, Len(strAppPath)
      
   Else
      RegDeleteValue lngKey, App.Title
   End If
   
   AutoStartIsChanged = Add
   RegCloseKey lngKey
   AddDeleteAutoRunKey = True
   
ExitFunction:
   On Local Error GoTo 0

End Function

' create MyTimeZones registry key
Private Function CreateKey(ByVal NewSubKey As String) As Boolean

Const REG_CREATED_NEW_KEY     As Long = &H1
Const REG_OPTION_NON_VOLATILE As Long = &H0

Dim lngKey                    As Long
Dim lngResult                 As Long

   If RegCreateKeyEx(HKEY_CURRENT_USER, NewSubKey, 0, "0", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, lngKey, lngResult) Then Exit Function
   If RegFlushKey(lngKey) = ERROR_SUCCESS Then RegCloseKey lngKey
   
   CreateKey = (lngResult = REG_CREATED_NEW_KEY)

End Function

' check if MyTimeZones registry key exist
Private Function ExistKey(ByVal SubKey As String) As Boolean

Dim lngKey As Long

   If RegOpenKeyEx(HKEY_CURRENT_USER, SubKey, 0, KEY_ALL_ACCESS, lngKey) = ERROR_SUCCESS Then
      ExistKey = True
      RegCloseKey lngKey
   End If

End Function

' returns Index or ZoneId for the specified clock
Private Function GetClockIndex(ByVal DisplayName As String, ByVal Index As Integer, ByVal ZoneID As Boolean) As Integer

Dim intCount   As Integer
Dim intPointer As Integer

   GetClockIndex = Index
   
   If DisplayName = "" Then Exit Function
   
   intPointer = InStr(DisplayName, "*")
   
   If intPointer Then DisplayName = Replace(DisplayName, "*", "")
   
   intPointer = InStr(DisplayName, "[")
   
   If intPointer Then DisplayName = Left(DisplayName, intPointer - 3)
   
   With frmMyTimeZones.cmbTimeZones
      For intCount = 0 To .ListCount - 1
         If DisplayName = .List(intCount) Then
            If ZoneID Then
               GetClockIndex = .ItemData(intCount)
               
            Else
               GetClockIndex = intCount
            End If
            
            Exit For
         End If
      Next 'intCount
   End With

End Function

' returns the local information of the system
Private Function GetLocaleInformation(ByVal LocaleInfo As Long) As String

Dim strBuffer As String

   strBuffer = Space(100)
   GetLocaleInformation = Left(strBuffer, GetLocaleInfo(LOCALE_USER_DEFAULT, LocaleInfo, strBuffer, Len(strBuffer)) - 1)

End Function

' gets the version regkey
Private Function GetRegKeyCurrentVersion(Optional ByVal CheckWinNT As Boolean = True) As String

   GetRegKeyCurrentVersion = "SOFTWARE\Microsoft\Windows" & IIf(IsWinNT And CheckWinNT, " NT", "") & "\CurrentVersion\"

End Function

' get MyTimeZones registry value
Private Function GetRegValue(ByVal SubKey As String, ByVal ValueName As String, ByRef Value As Variant, ByRef Default As Variant) As Boolean

Dim lngBufferSize As Long
Dim lngKey        As Long
Dim lngType       As Long
Dim lngData       As Long
Dim strBuffer     As String

   Value = Default
   
   If RegOpenKeyEx(HKEY_CURRENT_USER, SubKey, 0, KEY_ALL_ACCESS, lngKey) Then Exit Function
   If RegQueryValueEx(lngKey, ValueName, 0, lngType, ByVal 0&, lngBufferSize) Then GoTo ExitFunction
   
   If lngType = REG_SZ Then
      strBuffer = Space(lngBufferSize + 1)
      
      If RegQueryValueEx(lngKey, ValueName, 0, lngType, ByVal strBuffer, lngBufferSize) = ERROR_SUCCESS Then
         Value = Mid(Left(" " & strBuffer, lngBufferSize), 2)
         GetRegValue = True
      End If
      
   ElseIf lngType = REG_DWORD Then
      lngBufferSize = 4
      
      If RegQueryValueEx(lngKey, ValueName, 0, lngType, lngData, lngBufferSize) = ERROR_SUCCESS Then
         Select Case VarType(Value)
            Case vbBoolean
               Value = CBool(lngData)
               
            Case vbInteger
               Value = CInt(lngData)
               
            Case vbLong
               Value = lngData
               
            Case vbSingle
               Value = CSng(lngData)
         End Select
         
         GetRegValue = True
      End If
   End If
   
ExitFunction:
   RegCloseKey lngKey

End Function

' get the registry timezone collection
Private Function GetTimeZoneCollection(ByVal SubKey As String) As Boolean

Const REG_BINARY As Long = 3

Dim intCount     As Integer
Dim lngIndex     As Long
Dim lngKey       As Long
Dim lngSubKey    As Long
Dim strKeyName   As String
Dim strKeyValue  As String

   RegOpenKeyEx HKEY_LOCAL_MACHINE, SubKey, 0, KEY_ALL_ACCESS, lngKey
   
   If lngKey = 0 Then Exit Function
   
   Do
      strKeyName = Space(BUFFER_SIZE)
      strKeyValue = Space(BUFFER_SIZE)
      
      If RegEnumKey(lngKey, lngIndex, strKeyName, BUFFER_SIZE) Then Exit Do
      
      strKeyName = StripNull(strKeyName)
      
      If RegOpenKeyEx(HKEY_LOCAL_MACHINE, SubKey & "\" & strKeyName, 0, KEY_ALL_ACCESS, lngSubKey) = ERROR_SUCCESS Then
         ReDim Preserve TimeZoneRegKeyName(intCount) As String
         ReDim Preserve AllZones(intCount) As TimeZonesInfo
         
         TimeZoneRegKeyName(intCount) = strKeyName
         RegQueryValueEx lngSubKey, "TZI", 0, REG_BINARY, AllZones(intCount), Len(AllZones(intCount))
         
         With AllZones(intCount)
            .DisplayName = TrimName(GetTimeZoneDetail(lngSubKey, "Display"))
            .StandardName = GetTimeZoneDetail(lngSubKey, "Std")
            .DaylightName = GetTimeZoneDetail(lngSubKey, "Dlt")
            .MapID = GetTimeZoneDetail(lngSubKey, "MapID")
            intCount = intCount - (.DisplayName <> "")
         End With
         
         RegCloseKey lngSubKey
      End If
      
      lngIndex = lngIndex + 1
   Loop
   
   RegCloseKey lngKey
   GetTimeZoneCollection = True

End Function

' get the detail of the timezone
Private Function GetTimeZoneDetail(ByVal hKey As Long, ByVal ValueName As String) As String

   GetTimeZoneDetail = Space(BUFFER_SIZE)
   RegQueryValueEx hKey, ValueName, 0, REG_SZ, ByVal GetTimeZoneDetail, BUFFER_SIZE
   GetTimeZoneDetail = Trim(StripNull(GetTimeZoneDetail))

End Function

' check if OS is NT or more
Private Function IsWinNT() As Boolean

Dim osvInfo As OSVersionInfo

   With osvInfo
      .dwOSVersionInfoSize = Len(osvInfo)
      GetVersionEx osvInfo
      IsWinNT = (.dwPlatformId = VER_PLATFORM_WIN32_NT) And (.dwMajorVersion >= 4)
   End With

End Function

' set the timezone displayname
Private Function SetName(ByRef ZoneName() As Integer, Optional ByVal NewName As String) As String

Dim intCount As Integer

   If NewName = "" Then
      For intCount = 0 To 31
         If ZoneName(intCount) = 0 Then Exit For
         
         SetName = SetName & Chr(ZoneName(intCount))
      Next 'intCount
      
   Else
      For intCount = 1 To Len(NewName)
         ZoneName(intCount - 1) = Asc(Mid(NewName, intCount, 1))
      Next 'intCount
      
      ZoneName(intCount - 1) = 0
   End If

End Function

' set MyTimeZones registry value
Private Function SetRegValue(ByVal SubKey As String, ByVal ValueName As String, ByVal Value As Variant) As Boolean

Dim lngKey    As Long
Dim lngResult As Long
Dim strValue  As String

   If RegOpenKeyEx(HKEY_CURRENT_USER, SubKey, 0, KEY_ALL_ACCESS, lngKey) Then Exit Function
   
   Select Case VarType(Value)
      Case vbBoolean, vbInteger, vbLong, vbSingle
         lngResult = RegSetValueEx(lngKey, ValueName, 0, REG_DWORD, CLng(Value), 4)
         
      Case vbString
         strValue = CStr(Value) & vbNullChar
         lngResult = RegSetValueEx(lngKey, ValueName, 0, REG_SZ, ByVal strValue, Len(strValue))
   End Select
   
   RegCloseKey lngKey
   SetRegValue = (lngResult = 0)

End Function

' trim string from '()' to fix bugs in Windows and check for timezone 13
Private Function TrimName(ByVal Text As String) As String

Dim intFirst As Integer
Dim intLast  As Integer
Dim strGMT   As String
Dim strTime  As String

   If Text = "" Then Exit Function
   
   intFirst = InStr(Text, "(") + 1
   intLast = InStr(Text, ")")
   strGMT = UCase(Mid(Text, intFirst, intLast - intFirst))
   
   If strGMT <> "GMT" Then
      strTime = Format(Mid(strGMT, 5), "hh:mm")
      strGMT = Left(strGMT, 4)
      
      If Hour(strTime) = 13 Then TimeZone13 = True
   End If
   
   TrimName = "(" & strGMT & strTime & ") " & Trim(Mid(Text, intLast + 1))

End Function

' set the language
Private Sub GetLanguage(ByVal Language As String)

Dim intIndex(4) As Integer

   Select Case UCase(Language)
      Case "SPANISH"
         intIndex(1) = 500
         intIndex(2) = 800
         intIndex(3) = 901
         
      Case "DUTCH"
         intIndex(1) = 1000
         intIndex(2) = 1300
         intIndex(3) = 1401
         
      Case Else
         intIndex(1) = 0
         intIndex(2) = 300
         intIndex(3) = 401
      End Select
   
   On Local Error Resume Next
   
   ' Application text
   For intIndex(0) = intIndex(1) To intIndex(1) + UBound(AppText)
      AppText(intIndex(0) - intIndex(1)) = LoadResString(intIndex(0))
   Next 'intIndex(0)
   
   ' Error messages
   For intIndex(0) = intIndex(2) To intIndex(2) + UBound(AppError)
      AppError(intIndex(0) - intIndex(2)) = LoadResString(intIndex(0))
   Next 'intIndex(0)
   
   intIndex(4) = Int(intIndex(3) / 100) * 100
   
   ' For Calendar UserControl
   For intIndex(0) = intIndex(3) To intIndex(3) + UBound(LanguageText)
      ' Miscellaneous  = 1
      ' QuarterNames   = 3
      ' SeasonNames    = 4
      ' MoonPhaseNames = 5
      ' MoonPhaseText  = 6
      ' ZodiacNames    = 7
      LanguageText(intIndex(0) - intIndex(4)) = LoadResString(intIndex(0))
   Next 'intIndex(0)
   
   On Local Error GoTo 0
   Erase intIndex

End Sub
