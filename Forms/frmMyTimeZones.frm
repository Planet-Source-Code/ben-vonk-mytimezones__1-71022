VERSION 5.00
Begin VB.Form frmMyTimeZones 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   6720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12192
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMyTimeZones.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMyTimeZones.frx":08CA
   ScaleHeight     =   560
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   Visible         =   0   'False
   Begin VB.PictureBox picSysTrayIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   192
      Left            =   9600
      Picture         =   "frmMyTimeZones.frx":944C
      ScaleHeight     =   192
      ScaleWidth      =   192
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2880
      Visible         =   0   'False
      Width           =   192
   End
   Begin MyTimeZones.ThemedComboBox tcbSkinner 
      Left            =   11280
      Top             =   5040
      _ExtentX        =   445
      _ExtentY        =   423
      BorderColorStyle=   1
      ComboBoxBorderColor=   13751252
      DriveListBoxBorderColor=   0
   End
   Begin VB.PictureBox picMasker 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3048
      Index           =   1
      Left            =   11640
      ScaleHeight     =   254
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   247
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   6240
      Visible         =   0   'False
      Width           =   2964
   End
   Begin MyTimeZones.FlatButton flbChoose 
      Height          =   1224
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   5508
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   2159
      BackStyle       =   0
      Picture         =   "frmMyTimeZones.frx":978E
   End
   Begin MyTimeZones.ThumbWheel twvDate 
      Height          =   708
      Left            =   5556
      TabIndex        =   17
      Top             =   2316
      Width           =   336
      _ExtentX        =   593
      _ExtentY        =   1249
      Max             =   100
      Orientation     =   1
      ShadeControl    =   15264493
      ShadeWheel      =   15264493
      Value           =   50
   End
   Begin VB.PictureBox picMasker 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4116
      Index           =   0
      Left            =   10920
      Picture         =   "frmMyTimeZones.frx":D3D8
      ScaleHeight     =   343
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   555
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   6240
      Visible         =   0   'False
      Width           =   6660
   End
   Begin VB.PictureBox picAnimation 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   744
      Left            =   9480
      Picture         =   "frmMyTimeZones.frx":EBA6
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1283
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5400
      Visible         =   0   'False
      Width           =   15396
   End
   Begin MyTimeZones.LEDDisplay ledDisplay 
      Height          =   264
      Left            =   2856
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4608
      Width           =   3756
      _ExtentX        =   6625
      _ExtentY        =   466
      BackColor       =   15264493
      BorderStyle     =   0
      DisplayColor    =   13751252
      ForeColor       =   14239232
      Size            =   1
   End
   Begin MyTimeZones.FlatButton flbControlBox 
      Height          =   204
      Index           =   0
      Left            =   1140
      TabIndex        =   7
      Top             =   6144
      Width           =   204
      _ExtentX        =   360
      _ExtentY        =   360
      BackStyle       =   0
      Icon            =   "frmMyTimeZones.frx":2B603
      OnlyIconClick   =   -1  'True
      Picture         =   "frmMyTimeZones.frx":2B99D
   End
   Begin MyTimeZones.Clock clkTimeZone 
      Height          =   2892
      Index           =   0
      Left            =   456
      TabIndex        =   1
      Top             =   1980
      Width           =   2892
      _ExtentX        =   5101
      _ExtentY        =   5101
      BackColor       =   14737632
      BackStyle       =   0
      ClockBold       =   -1  'True
      ClockHoursColor =   14239232
      ClockPlateBorderColor=   13751252
      ClockPlateColor =   13795392
      ClockPlateGradientColor=   16777184
      ClockPlateGradientStyle=   1
      DateColor       =   12591040
      DateStyle       =   2
      DigitsAmPm      =   0   'False
      DigitsBackStyle =   0
      DigitsForeColor =   8388736
      DigitsPosition  =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HandsHourMinuteBorderColor=   12632064
      HandsHourMinuteColor=   12139264
      MarksHourColor  =   16777215
      MarksMinuteColor=   12632064
      MarksStyle      =   3
      NameBackStyle   =   0
      NameBold        =   -1  'True
      NameForeColor   =   12591040
      NamePosition    =   1
   End
   Begin VB.PictureBox picDateTimeButton 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1224
      Left            =   9600
      ScaleHeight     =   102
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.PictureBox picChild 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4680
      Index           =   1
      Left            =   10200
      Picture         =   "frmMyTimeZones.frx":2BD63
      ScaleHeight     =   390
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   6240
      Visible         =   0   'False
      Width           =   6000
   End
   Begin VB.PictureBox picImages 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   768
      Left            =   11640
      Picture         =   "frmMyTimeZones.frx":3B43B
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   80
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox picChild 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6060
      Index           =   0
      Left            =   9480
      Picture         =   "frmMyTimeZones.frx":3BDC7
      ScaleHeight     =   505
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   685
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6240
      Visible         =   0   'False
      Width           =   8220
   End
   Begin VB.Timer tmrAlarmOff 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   9600
      Top             =   5040
   End
   Begin VB.ListBox lstSort 
      Height          =   264
      Left            =   9600
      Sorted          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Timer tmrClock 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9600
      Top             =   3600
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  'Center
      BackColor       =   &H00E8EAED&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D94600&
      Height          =   228
      Left            =   3648
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   3
      Top             =   2424
      Width           =   1824
   End
   Begin VB.TextBox txtTime 
      Alignment       =   2  'Center
      BackColor       =   &H00E8EAED&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D94600&
      Height          =   228
      Left            =   3648
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   4
      Top             =   2688
      Width           =   1824
   End
   Begin VB.ComboBox cmbTimeZones 
      BackColor       =   &H00E8EAED&
      ForeColor       =   &H00D94600&
      Height          =   312
      Left            =   1524
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   4980
      Width           =   6456
   End
   Begin VB.Timer tmrScrews 
      Enabled         =   0   'False
      Interval        =   55
      Left            =   9600
      Top             =   4560
   End
   Begin VB.Timer tmrAnimation 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   9600
      Top             =   4080
   End
   Begin VB.PictureBox picToDay 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EBEBEB&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   384
      Left            =   9600
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   384
   End
   Begin MyTimeZones.FlatButton flbControlBox 
      Height          =   204
      Index           =   1
      Left            =   8160
      TabIndex        =   8
      Top             =   6144
      Width           =   204
      _ExtentX        =   360
      _ExtentY        =   360
      BackStyle       =   0
      Icon            =   "frmMyTimeZones.frx":540FE
      OnlyIconClick   =   -1  'True
      Picture         =   "frmMyTimeZones.frx":54498
   End
   Begin MyTimeZones.Clock clkFavorits 
      Height          =   1812
      Index           =   0
      Left            =   144
      TabIndex        =   2
      Top             =   240
      Width           =   1812
      _ExtentX        =   3196
      _ExtentY        =   3196
      BackColor       =   15264493
      BackStyle       =   0
      ClockHours      =   0   'False
      ClockPlateBorderColor=   13751252
      ClockPlateColor =   10790823
      ClockPlateGradientColor=   16448250
      ClockPlateGradientPosition=   1
      ClockPlateGradientStyle=   1
      DigitsAmPm      =   0   'False
      DigitsBackStyle =   0
      DigitsForeColor =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HandSecondColor =   14737632
      HandSecondScroll=   0
      HandsHourMinuteBorderColor=   12632064
      HandsHourMinuteColor=   4210752
      HandsHourMinuteStyle=   2
      MarksHourColor  =   8421504
      MarksShows      =   2
      MarksStyle      =   2
      NameBackStyle   =   0
      NameForeColor   =   12591040
      NamePosition    =   2
   End
   Begin VB.Shape shpDisplay 
      BorderColor     =   &H00D1C3C4&
      Height          =   288
      Left            =   2844
      Top             =   4596
      Width           =   3780
   End
   Begin VB.Image imgPopupMenu 
      Height          =   192
      Index           =   4
      Left            =   9840
      Picture         =   "frmMyTimeZones.frx":5485E
      Top             =   2160
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Shape shpActiveClock 
      BorderColor     =   &H00FF9D6C&
      Height          =   2892
      Left            =   456
      Shape           =   3  'Circle
      Top             =   1980
      Width           =   2892
   End
   Begin VB.Shape shpClockBorder 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   4
      Height          =   2880
      Index           =   0
      Left            =   468
      Shape           =   3  'Circle
      Tag             =   "&H00E0E0E0&"
      Top             =   1992
      Width           =   2880
   End
   Begin VB.Shape shpAlarm 
      BorderColor     =   &H00C0C0FF&
      BorderWidth     =   2
      Height          =   1812
      Left            =   144
      Shape           =   3  'Circle
      Top             =   240
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.Image imgPopupMenu 
      Height          =   192
      Index           =   9
      Left            =   9600
      Picture         =   "frmMyTimeZones.frx":54BA0
      Top             =   2640
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   31
      Left            =   10680
      Picture         =   "frmMyTimeZones.frx":550E2
      Top             =   4920
      Width           =   384
   End
   Begin VB.Image imgPopupMenu 
      Height          =   192
      Index           =   7
      Left            =   9840
      Picture         =   "frmMyTimeZones.frx":559AC
      Top             =   2400
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Image imgPopupMenu 
      Height          =   192
      Index           =   6
      Left            =   9600
      Picture         =   "frmMyTimeZones.frx":55CEE
      Top             =   2400
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Image imgPopupMenu 
      Height          =   192
      Index           =   3
      Left            =   9600
      Picture         =   "frmMyTimeZones.frx":56030
      Top             =   2160
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Image imgPopupMenu 
      Height          =   192
      Index           =   2
      Left            =   9840
      Picture         =   "frmMyTimeZones.frx":56572
      Top             =   1920
      Width           =   192
   End
   Begin VB.Image imgPopupMenu 
      Height          =   192
      Index           =   0
      Left            =   9600
      Picture         =   "frmMyTimeZones.frx":56AB4
      Top             =   1920
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Image imgScrews 
      Height          =   384
      Index           =   4
      Left            =   11640
      Picture         =   "frmMyTimeZones.frx":56DF6
      Top             =   2880
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgScrews 
      Height          =   384
      Index           =   0
      Left            =   11640
      Picture         =   "frmMyTimeZones.frx":576C0
      Top             =   960
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgScrews 
      Height          =   384
      Index           =   1
      Left            =   11640
      Picture         =   "frmMyTimeZones.frx":57F8A
      Top             =   1440
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgScrews 
      Height          =   384
      Index           =   2
      Left            =   11640
      Picture         =   "frmMyTimeZones.frx":58854
      Top             =   1920
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgScrews 
      Height          =   384
      Index           =   3
      Left            =   11640
      Picture         =   "frmMyTimeZones.frx":5911E
      Top             =   2400
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   30
      Left            =   10200
      Picture         =   "frmMyTimeZones.frx":599E8
      Top             =   4920
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   29
      Left            =   11160
      Picture         =   "frmMyTimeZones.frx":5A2B2
      Top             =   4440
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   28
      Left            =   10680
      Picture         =   "frmMyTimeZones.frx":5AB7C
      Top             =   4440
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   27
      Left            =   10200
      Picture         =   "frmMyTimeZones.frx":5B446
      Tag             =   "8"
      Top             =   4440
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgMove 
      Height          =   396
      Index           =   4
      Left            =   3456
      MousePointer    =   15  'Size All
      Top             =   3732
      Width           =   2520
   End
   Begin VB.Image imgMove 
      Height          =   384
      Index           =   3
      Left            =   9000
      MousePointer    =   15  'Size All
      Top             =   6252
      Width           =   384
   End
   Begin VB.Image imgMove 
      Height          =   384
      Index           =   2
      Left            =   96
      MousePointer    =   15  'Size All
      Top             =   6252
      Width           =   384
   End
   Begin VB.Image imgMove 
      Height          =   384
      Index           =   1
      Left            =   9000
      MousePointer    =   15  'Size All
      Top             =   84
      Width           =   384
   End
   Begin VB.Image imgMove 
      Height          =   384
      Index           =   0
      Left            =   96
      MousePointer    =   15  'Size All
      Top             =   72
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   26
      Left            =   11160
      Picture         =   "frmMyTimeZones.frx":5BD10
      Tag             =   "27"
      Top             =   3960
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   25
      Left            =   10680
      Picture         =   "frmMyTimeZones.frx":5C5DA
      Tag             =   "26"
      Top             =   3960
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   24
      Left            =   10200
      Picture         =   "frmMyTimeZones.frx":5CEA4
      Tag             =   "25"
      Top             =   3960
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   23
      Left            =   11160
      Picture         =   "frmMyTimeZones.frx":5D76E
      Tag             =   "24"
      Top             =   3480
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   22
      Left            =   10680
      Picture         =   "frmMyTimeZones.frx":5E038
      Tag             =   "23"
      Top             =   3480
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   21
      Left            =   10200
      Picture         =   "frmMyTimeZones.frx":5E902
      Top             =   3480
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   20
      Left            =   11160
      Picture         =   "frmMyTimeZones.frx":5F1CC
      Tag             =   "7"
      Top             =   3000
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   19
      Left            =   10680
      Picture         =   "frmMyTimeZones.frx":5FA96
      Tag             =   "20"
      Top             =   3000
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   18
      Left            =   10200
      Picture         =   "frmMyTimeZones.frx":60360
      Tag             =   "19"
      Top             =   3000
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   17
      Left            =   11160
      Picture         =   "frmMyTimeZones.frx":60C2A
      Tag             =   "18"
      Top             =   2520
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   16
      Left            =   10680
      Picture         =   "frmMyTimeZones.frx":614F4
      Tag             =   "17"
      Top             =   2520
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   15
      Left            =   10200
      Picture         =   "frmMyTimeZones.frx":61DBE
      Tag             =   "16"
      Top             =   2520
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   14
      Left            =   11160
      Picture         =   "frmMyTimeZones.frx":62688
      Top             =   2040
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   13
      Left            =   10680
      Picture         =   "frmMyTimeZones.frx":62F52
      Top             =   2040
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   12
      Left            =   10200
      Picture         =   "frmMyTimeZones.frx":6381C
      Top             =   2040
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   11
      Left            =   11160
      Picture         =   "frmMyTimeZones.frx":640E6
      Top             =   1560
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   10
      Left            =   10680
      Picture         =   "frmMyTimeZones.frx":649B0
      Top             =   1560
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   9
      Left            =   10200
      Picture         =   "frmMyTimeZones.frx":6527A
      Top             =   1560
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   8
      Left            =   11160
      Picture         =   "frmMyTimeZones.frx":65B44
      Tag             =   "22"
      Top             =   1080
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   7
      Left            =   10680
      Picture         =   "frmMyTimeZones.frx":6640E
      Tag             =   "15"
      Top             =   1080
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   6
      Left            =   10200
      Picture         =   "frmMyTimeZones.frx":66CD8
      Top             =   1080
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   5
      Left            =   11160
      Picture         =   "frmMyTimeZones.frx":675A2
      Top             =   600
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   4
      Left            =   10680
      Picture         =   "frmMyTimeZones.frx":67E6C
      Top             =   600
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   3
      Left            =   10200
      Picture         =   "frmMyTimeZones.frx":68B36
      Top             =   600
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   2
      Left            =   11160
      Picture         =   "frmMyTimeZones.frx":69400
      Top             =   120
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   1
      Left            =   10680
      Picture         =   "frmMyTimeZones.frx":69CCA
      Top             =   120
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgImages 
      Height          =   384
      Index           =   0
      Left            =   10200
      Picture         =   "frmMyTimeZones.frx":6A594
      Top             =   120
      Visible         =   0   'False
      Width           =   384
   End
End
Attribute VB_Name = "frmMyTimeZones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MyTimeZones™ program
'
'Author Ben Vonk
'11-07-1997 First version (Named; 'WKlok')
'17-02-2004 Second version (Named; 'TijdZonesConversie' (based on Ark's 'TimeZone Info/Converter' at http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=11825&lngWId=1)
'New version 1.0.0 and changing name in: MyTimeZones
'26-12-2005 1.0.0 complete rebuild program based on the idea of the second version (TijdZonesConversie)
'08-01-2007 1.0.1 some changes and bugfixes are made
'31-01-2007 1.1.0 some changes and bugfixes are made
'19-02-2007 1.1.1 some bugs and WindowsXP timezone layout bug are fixed
'11-03-2007 1.2.0 now also support MouseEvents for Date/Time setting (ThumbWheel) and Calendar listings
'19-03-2007 1.3.0 some icons are changed and some bugs are fixed
'23-03-2007 1.4.0 disable popupmenus for date/time textbox with subclassing
'27-03-2007 1.4.1 fixed some bugs in SpecialDays module
'31-03-2007 2.0.0 change some layout items and delete some menuitems from controlmenu
'05-04-2007 2.1.0 change some icons and some colors in the worldmap
'12-04-2007 2.1.1 changed AtomicSynchronise and twvDate_Change procedures, set the english text in the resourcefile: Resources.res
'15-04-2007 2.1.2 split mdlPublic in 5 modules and change some icons
'18-04-2007 2.2.0 change functionality of setting favorit, restore the selected timezone for timezone clock
'30-04-2007 2.3.0 add PopupMenu for selected Clock and fixed the IntroScreen bug
'12-05-2007 3.0.0 change layout items and make some changes and also project updated to VB6
'23-06-2007 3.1.0 add ThemedScrollBar control in Calendar module and add balloon tooltiptext in systray
'12-08-2007 4.0.0 change layout, clock gradient, error messages, some icons, some skin parts, and message display
'26-12-2007 4.1.0 add TimeToGo
'14-01-2008 4.1.1 fixed some bugs
'22-01-2008 4.2.0 change some icons
'12-03-2008 4.3.0 change layout a little
'18-05-2008 4.4.0 change some icons and move language files to the resource file
'26-08-2008 4.5.0 implement ChrystalButton and change layout a little
'10-10-2008 5.0.0 add ThemedComboBox control, SystemTray class, change the SystemMenu and make changes in the project layout

Option Explicit

' Private Constants
Private Const RAS_MAXDEVICENAME    As Integer = 128
Private Const RAS_MAXDEVICETYPE    As Integer = 16
Private Const RAS_MAXENTRYNAME     As Integer = 256
Private Const SOUND_SIGNAL         As Integer = 1

' Private Types
Private Type InternetConnectionData
   Configured                      As Boolean
   Modem                           As Boolean
   OffLine                         As Boolean
End Type

Private Type RasConnection
   dwSize                          As Long
   hRasConn                        As Long
   szEntryName(RAS_MAXENTRYNAME)   As Byte
   szDeviceType(RAS_MAXDEVICETYPE) As Byte
   szDeviceName(RAS_MAXDEVICENAME) As Byte
End Type

' Private Classes with Events
Private WithEvents clsMouse        As clsMouseWheel
Attribute clsMouse.VB_VarHelpID = -1
Private WithEvents clsSysTray      As clsSystemTray
Attribute clsSysTray.VB_VarHelpID = -1

' Public Variable
Public AlarmIndex                  As Integer

' Private Variables
Private CalendarButtonIsClicked    As Boolean
Private ChangeDateTime             As Boolean
Private InPopupMenu                As Boolean
Private IsBusy                     As Boolean
Private IsLeftButton               As Boolean
Private IsRemove                   As Boolean
Private IsRightButton              As Boolean
Private IsSynchronise              As Boolean
Private SetChanges                 As Boolean
Private NoDaylightChanges          As Boolean
Private OnSysTray                  As Boolean
Private RunSilent                  As Boolean
Private SysMenuHooked              As Boolean
Private SystemDate                 As Date
Private UserTimer                  As Date
Private AllFavorits                As Integer
Private Animation                  As Integer
Private ButtonAnimation            As Integer
Private IndexToTimeZone            As Integer
Private ScrewItem                  As Integer
Private SelectedDatePart           As Integer
Private SubclassedTextBox          As Integer
Private InternetConnection         As InternetConnectionData
Private FormLeft                   As Long
Private FormTop                    As Long
Private hWndPopupMenu              As Long
Private AutoSyncTime               As String
Private UpdateClock                As String

' Private API's
Private Declare Function SetSystemTime Lib "Kernel32" (lpSystemTime As SystemTime) As Long
Private Declare Function RasEnumConnections Lib "RasApi32" Alias "RasEnumConnectionsA" (lpRasConn As Any, lpcb As Long, lpcConnections As Long) As Long
Private Declare Function RasHangUp Lib "RasApi32" Alias "RasHangUpA" (ByVal hRasConn As Long) As Long
Private Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindow Lib "User32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function OpenIcon Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function InternetCloseHandle Lib "WinInet" (ByVal hInet As Long) As Integer
Private Declare Function InternetGetConnectedState Lib "WinInet" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function InternetGetConnectedStateEx Lib "WinInet" (ByRef lpdwFlags As Long, ByVal lpszConnectionName As String, ByVal dwNameLen As Integer, ByVal dwReserved As Long) As Long
Private Declare Function InternetOpen Lib "WinInet" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetOpenUrl Lib "WinInet" Alias "InternetOpenUrlA" (ByVal hOpen As Long, ByVal sUrl As String, ByVal sHeaders As String, ByVal lLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetReadFile Lib "WinInet" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer

Public Sub CheckFavoritIsSet()

Dim blnLocked(1)  As Boolean
Dim strTipText(1) As String

   If AppSettings(SET_LOCKFAVORITS) Then Exit Sub
   
   If SelectedClock < 2 Then
      blnLocked(0) = True
      blnLocked(1) = True
      strTipText(0) = AppText(73)
      strTipText(1) = strTipText(0)
      
   Else
      blnLocked(1) = clkFavorits.Item(SelectedFavorit).NameClock = AppText(6)
      strTipText(0) = AppText(57)
      strTipText(1) = AppText(58 - (51 And blnLocked(1)))
      
      If FavoritsInfo(SelectedFavorit).Index = cmbTimeZones.ListIndex Then
         blnLocked(0) = True
         strTipText(0) = AppText(29)
      End If
   End If
   
   Call SetButton(7, GetIcon(7, blnLocked(0)), Me)
   Call SetButton(8, GetIcon(8, blnLocked(1)), Me)
   
   flbChoose.Item(7).ToolTipText = GetToolTipText(strTipText(0))
   flbChoose.Item(8).ToolTipText = GetToolTipText(strTipText(1))

End Sub

Public Sub EndMyTimeZones(Optional ByVal DoUnload As Boolean = True)

   DoEvents
   ledDisplay.Active = False
   MousePointer = vbHourglass
   
   Call SetRegValues
   Call PopupMenuHooking(0)
   Call MouseUnhook
   Call TextBoxUnhook
   
   If SysMenuHooked Then Call SubclassSystemMenu(hWnd)
   If DoUnload Then Unload Me
   
   Set frmMyTimeZones = Nothing

End Sub

Public Sub SetAlarmOff()

   If AlarmIndex = -1 Then Exit Sub
   
   If shpAlarm.Top = clkFavorits.Item(AlarmIndex).Top Then
      shpAlarm.Visible = False
      tmrAlarmOff.Enabled = False
      
      If Len(CreateAlarmMessage(AlarmIndex)) And Not CheckTimeToGo Then ledDisplay.Text = ""
   End If
   
   If Not SetTimeToGoDisplay(True) Then ledDisplay.BackColor = &HE8EAED
   
   Call PlaySound(SOUND_SIGNAL)

End Sub

Public Sub SysTrayDisplay(ByVal ShowIcon As Boolean, Optional ByVal State As Boolean = True)

Static sngLeft As Single
Static sngTop  As Single

   If Visible Then
      sngTop = Top
      sngLeft = Left
   End If
   
   If RunSilent Then
      RunSilent = False
      sngTop = (Screen.Height - Height) \ 2
      sngLeft = (Screen.Width - Width) \ 2
   End If
   
   If ShowIcon Then
      Load frmPopupMenu
      Set clsSysTray = New clsSystemTray
      
      With clsSysTray
         .Icon = Icon.Handle
         .Parent = hWnd
         .TipText = Caption
         
         Call .AddIcon
         
         If AppSettings(SET_SHOWTIPTEXT) Then Call .ShowBalloon(Caption, AppText(40) & vbCrLf & AppText(41), , 5000, AppSettings(SET_PLAYSOUND))
      End With
      
      Call SysTrayModifyTipText
      
   Else
      SetForegroundWindow hWnd
      clsSysTray.DeleteIcon
      Set clsSysTray = Nothing
   End If
   
   If State Then Visible = Not ShowIcon
   
   OnSysTray = ShowIcon

End Sub

Public Sub ToggleControls(ByVal State As Boolean)

   If State Then
      With ledDisplay
         If AlarmIndex > -1 Then
            If .ToolTipText <> GetDisplayToolTipText(clkFavorits.Item(AlarmIndex).NameClock) Then SetTimeToGoDisplay True
            
         ElseIf CheckTimeToGo Then
            ShowTimeToGo ledDisplay
            .NoTextScrolling = True
            .BackColor = &HA8ABAB
            .ToolTipText = GetTimeToGoToolTipText
            
         ElseIf Not SetTimeToGoDisplay(True) Then
            .Text = ""
            .NoTextScrolling = False
            .BackColor = &HE8EAED
         End If
      End With
      
      Call tmrClock_Timer
      
      MousePointer = vbDefault
      DoEvents
   End If
   
   txtDate.Visible = State
   txtTime.Visible = State
   ledDisplay.Active = State

End Sub

Private Function GetAlarmMessage(ByVal Index As Integer) As String

   If Len(FavoritsInfo(Index).AlarmMessage) Then GetAlarmMessage = " - " & AppText(78) & " " & CreateAlarmMessage(Index)

End Function

Private Function GetDayLightSetting() As Boolean

   GetDayLightSetting = GetRegKeyValue(SKEY_TIMEZONE, DISABLE_AUTO_DLTS) = ""

End Function

Private Function GetInternetConnection() As Boolean

Const INTERNET_CONNECTION_CONFIGURED As Long = &H40
Const INTERNET_CONNECTION_LAN        As Long = &H2
Const INTERNET_CONNECTION_MODEM      As Long = &H1
Const INTERNET_CONNECTION_OFFLINE    As Long = &H20
Const INTERNET_CONNECTION_PROXY      As Long = &H4

Dim lngReturn                        As Long
Dim lngConnection                    As Long

   lngReturn = InternetGetConnectedState(lngConnection, 0)
   
   With InternetConnection
      .Modem = (lngConnection And INTERNET_CONNECTION_MODEM) = INTERNET_CONNECTION_MODEM
      .Configured = (lngConnection And INTERNET_CONNECTION_CONFIGURED) = INTERNET_CONNECTION_CONFIGURED
      .OffLine = (lngConnection And INTERNET_CONNECTION_OFFLINE) = INTERNET_CONNECTION_OFFLINE
      
      If Not .Configured Then If ((lngConnection And INTERNET_CONNECTION_LAN) = INTERNET_CONNECTION_LAN) Or ((lngReturn And INTERNET_CONNECTION_PROXY) = INTERNET_CONNECTION_PROXY) Then .Configured = True
   End With
   
   GetInternetConnection = (lngReturn = 1)

End Function

Private Function GetTimeToGoToolTipText() As String

Dim strText As String
Dim strWord As String

   strText = Replace(AppText(98), "$", Split(AppText(218), ",")(TimeToGoShowType), , 1)
   strWord = Split(AppText(99), ",")(0 + (1 And ((TimeToGoShowType = 1) And (Format(Date, "yyyymmdd") > Format(TimeToGo(1), "yyyymmdd")))) + (2 And ((TimeToGoShowType = 2) And (Format(Date, "yyyymmdd") < Format(TimeToGo(0), "yyyymmdd")))))
   strText = Replace(Replace(strText, "$", LCase(Trim(Split(AppText(217), ",")(TimeToGoShow))), , 1), "$", strWord, , 1)
   GetTimeToGoToolTipText = GetToolTipText(Replace(Replace(strText, "#", CapsText(Format(TimeToGo(0), "Long Date")), , 1), "#", CapsText(Format(TimeToGo(1), "Long Date")), , 1))

End Function

Private Function SetTimeToGoDisplay(ByVal State As Boolean) As Boolean

   If State And (CheckTimeToGo And (TimeToGoShow > -1)) Then
      With ledDisplay
         .NoTextScrolling = True
         .BackColor = &HA8ABAB
         ShowTimeToGo ledDisplay
         .ToolTipText = GetTimeToGoToolTipText
         .Active = True
         SetTimeToGoDisplay = True
      End With
      
   Else
      Call DisableDisplay(ledDisplay)
   End If
   
   DoEvents

End Function

Private Sub ActivatePrevInstance()

Const GW_HWNDPREV       As Long = 3

Dim bytBuffer(1 To 255) As Byte
Dim cdsData             As CopyDataStruct
Dim lngWindow           As Long
Dim strAppTitle         As String

   With App
      lngWindow = FindWindow(vbNullString, CLASS_NAME_HIDDEN & .Title)
      strAppTitle = .Title
      .Title = .Title & "_Terminate"
   End With
   
   If lngWindow Then
      Call CopyMemory(bytBuffer(1), ByVal RECEIVED_DATA, Len(RECEIVED_DATA))
      
      With cdsData
         .dwData = 3
         .cbData = Len(RECEIVED_DATA) + 1
         .lpData = VarPtr(bytBuffer(1))
      End With
      
      SendMessage lngWindow, WM_COPYDATA, WM_ACTIVATE, cdsData
      
   Else
      If FindWindow("ThunderRTMain", strAppTitle) = 0 Then lngWindow = FindWindow("ThunderRT6Main", strAppTitle)
      If lngWindow = 0 Then Exit Sub
      
      lngWindow = GetWindow(lngWindow, GW_HWNDPREV)
      OpenIcon lngWindow
      SetForegroundWindow lngWindow
   End If
   
   End

End Sub

Private Sub AtomicSynchronise(ByVal Automatic As Boolean)

Const INTERNET_FLAG_RELOAD         As Long = &H80000000
Const INTERNET_OPEN_TYPE_PRECONFIG As Long = &H0

Dim blnIsConnected                 As Boolean
Dim dteDateTime                    As Date
Dim intErrCode                     As Integer
Dim intMiliSeconds                 As Integer
Dim intPointer                     As Integer
Dim lngInternet                    As Long
Dim lngHTTP                        As Long
Dim lngBytesRead                   As Long
Dim strConnectionType              As String * 255
Dim strPrompt                      As String
Dim strURLData                     As String * 1024
Dim sysUniversal                   As SystemTime

   If NoInternet Or IsSynchronise Or (TimeServerURL = "") Or (TimeServerURL = AppText(198)) Then Exit Sub
   
   blnIsConnected = GetInternetConnection
   
   If Automatic Then
      If Not blnIsConnected Or InternetConnection.OffLine Then Exit Sub
      
   Else
      If InternetConnection.OffLine And AppSettings(SET_ASKCONFIRM) Then
         ShowMessage AppError(36), vbStop, AppError(13), AppError(14), TimeToWait
         Exit Sub
      End If
      
      If AppSettings(SET_ASKCONFIRM) Then If ShowMessage(AppError(15), vbQuestion, AppError(13), AppError(14), TimeToWait) = vbNo Then Exit Sub
   End If
   
   Refresh
   DoEvents
   IsSynchronise = True
   MousePointer = vbHourglass
   On Local Error GoTo ErrorOpen
   lngInternet = InternetOpen(App.Title, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
   Refresh
   DoEvents
   
   If lngInternet = 0 Then
      intErrCode = 23
      
      GoTo ErrorOpen
   End If
   
   lngHTTP = InternetOpenUrl(lngInternet, "http://" & TimeServerURL & ":13/", vbNullString, 0, INTERNET_FLAG_RELOAD, 0)
   Refresh
   DoEvents
   
   If InternetGetConnectedStateEx(lngInternet, strConnectionType, 254, 0) <> 1 Then
      intErrCode = 37
      
      GoTo ErrorOpen
   End If
   
   If lngHTTP = 0 Then
      intErrCode = 24
      
      GoTo ErrorOpen
   End If
   
   strURLData = vbNullString
   InternetReadFile lngHTTP, strURLData, Len(strURLData), lngBytesRead
   Refresh
   DoEvents
   
   If lngHTTP Then InternetCloseHandle lngHTTP
   If lngInternet Then InternetCloseHandle lngInternet
   
   intErrCode = 8
   intPointer = InStr(strURLData, ":")
   
   If intPointer Then
      dteDateTime = DateSerial(Val(Mid(strURLData, intPointer - 11, 2)), Val(Mid(strURLData, intPointer - 8, 2)), Val(Mid(strURLData, intPointer - 5, 2))) & " " & CDate(Mid(strURLData, intPointer - 2, 8))
      intErrCode = 0
      
      If Mid(strURLData, intPointer + 12, 1) = "0" Then intMiliSeconds = Mid(strURLData, intPointer + 14, 3) & Mid(strURLData, intPointer + 18, 1)
   End If
   
   If intErrCode Then GoTo ErrorOpen
   If intMiliSeconds Then dteDateTime = DateAdd("s", -1, dteDateTime)
   
   With sysUniversal
      .wYear = Year(dteDateTime)
      .wMonth = Month(dteDateTime)
      .wDay = Day(dteDateTime)
      .wHour = Hour(dteDateTime)
      .wMinute = Minute(dteDateTime)
      .wSecond = Second(dteDateTime)
      .wMilliseconds = 0 + (((10000 - intMiliSeconds) / 10) And intMiliSeconds)
   End With
   
   SetSystemTime sysUniversal
   
   GoTo ExitSub
   
ErrorOpen:
   If Not Automatic And AppSettings(SET_ASKCONFIRM) Then
      If intErrCode = 22 Then
         strPrompt = AppError(22) & vbCrLf & "[#" & Err.Number & " - " & Err.Description & "]"
         
      Else
         strPrompt = Replace(AppError(intErrCode), "$", TimeServerURL)
      End If
      
      ShowMessage strPrompt, vbInformation, AppError(25), AppError(0), TimeToWait
   End If
   
ExitSub:
   On Local Error GoTo 0
   IsSynchronise = False
   MousePointer = vbDefault
   
   If Not blnIsConnected And InternetConnection.Modem Then If InternetGetConnectedStateEx(lngInternet, strConnectionType, 254, 0) = 1 Then If ShowMessage(AppError(35), vbQuestion, AppError(34), AppError(14), TimeToWait, vbYes) = vbYes Then Call ModemHangUp
   If Not Automatic Then Call ResetAllClocks

End Sub

Private Sub ChangeDaylightTimeSet(ByVal CheckSystem As Boolean)

Dim intDaylightIcon As Integer

   If CheckSystem Then
      If IndexSystemTimeZone = IndexFromTimeZone Then
         AutoDaylightTimeSet = GetDayLightSetting
         
      Else
         AutoDaylightTimeSet = AllZones(cmbTimeZones.ItemData(IndexFromTimeZone)).DaylightDate.wMonth
      End If
   End If
   
   If AllZones(cmbTimeZones.ItemData(IndexFromTimeZone)).DaylightDate.wMonth Then
      intDaylightIcon = 3 + (9 And AutoDaylightTimeSet)
      NoDaylightTimeSet = False
      
   Else
      intDaylightIcon = 13
      NoDaylightTimeSet = True
      AutoDaylightTimeSet = False
   End If
   
   Call SetButton(3, intDaylightIcon, Me)
   
   flbChoose.Item(3).ToolTipText = GetToolTipText(AppText(intDaylightIcon + 50))

End Sub

Private Sub CheckDaylightChanges()

   If (NoDaylightChanges + NoDaylightTimeSet + AppSettings(SET_LOCKSYSTEMDATE)) Or (IndexSystemTimeZone <> cmbTimeZones.ListIndex) Then Exit Sub
   
   If AutoDaylightTimeSet <> GetDayLightSetting Then
      AutoDaylightTimeSet = GetDayLightSetting
      
      Call ChangeDaylightTimeSet(False)
   End If

End Sub

Private Sub CreateForm()

Const WS_CLIPCHILDREN As Long = &H2000000
Const WS_CLIPSIBLINGS As Long = &H4000000
Const WS_MINIMIZEBOX  As Long = &H20000
Const WS_OVERLAPPED   As Long = &H0&
Const WS_SYSMENU      As Long = &H80000

   SetWindowLong hWnd, GWL_STYLE, WS_OVERLAPPED Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_SYSMENU Or WS_MINIMIZEBOX
   SysMenuHooked = SystemMenuCreate(hWnd, picSysTrayIcon.Picture.Handle)

End Sub

Private Sub FavoritAdd()

Dim intCount  As Integer
Dim strPrompt As String

   If SelectedClock - 2 <> SelectedFavorit Then Exit Sub
   If FavoritsInfo(SelectedFavorit).Index = cmbTimeZones.ListIndex Then Exit Sub
   
   For intCount = 0 To 4
      If FavoritsInfo(intCount).Index = cmbTimeZones.ListIndex Then
         If AppSettings(SET_ASKCONFIRM) Then
            strPrompt = Replace(AppError(6), "$", AllZones(cmbTimeZones.ItemData(cmbTimeZones.ListIndex)).DisplayName, , 1)
            strPrompt = Replace(strPrompt, "$", clkFavorits.Item(intCount).NameClock, , 1)
            ShowMessage strPrompt, vbStop, AppError(5), AppError(3), TimeToWait
         End If
         
         Exit Sub
      End If
   Next 'intCount
   
   With clkFavorits.Item(SelectedFavorit)
      If .Active And AppSettings(SET_ASKCONFIRM) Then
         strPrompt = Replace(AppError(4), "$", AllZones(FavoritsInfo(SelectedFavorit).ZoneID).DisplayName, , 1)
         strPrompt = Replace(strPrompt, "$", cmbTimeZones.Text, , 1)
         
         If ShowMessage(strPrompt, vbQuestion, AppError(5), AppText(57), TimeToWait) = vbNo Then Exit Sub
      End If
      
      TotalFavorits = TotalFavorits + 1
      tmrAnimation.Enabled = True
      .Picture = Nothing
      .ClockPlateGradientStyle = InToOut
      
      With FavoritsInfo(SelectedFavorit)
         .Index = cmbTimeZones.ListIndex
         SelectedTimeZoneID = cmbTimeZones.ItemData(.Index)
         .ZoneID = SelectedTimeZoneID
         .DisplayName = "*" & cmbTimeZones.Text
         .AlarmTime = ""
         .AlarmMessage = ""
         .AlarmTipText = ""
      End With
      
      .ClockPlateColor = &HEAD199
   End With
   
   Call SetFavoritClockDateTime(SelectedFavorit, SelectedTimeZoneID)
   Call FavoritActivate(SelectedFavorit, True, GetTimeZoneText(SelectedTimeZoneID))
   Call FavoritSelected(SelectedFavorit, True)

End Sub

Private Sub FavoritActivate(ByVal Index As Integer, ByVal Activate As Boolean, ByVal NameClock As String)

   With clkFavorits.Item(Index)
      .Locked = True
      .ClockPlateGradientStyle = OutToIn
      .DigitsBold = False
      .HandsHourMinuteStyle = Thin
      .MarksStyle = CircleMarks
      .NameBold = False
      .NameClock = NameClock
      
      If Activate Then
         .AlarmTime = FavoritsInfo(Index).AlarmTime
         .AlarmToolTipText = GetToolTipText(FavoritsInfo(Index).AlarmTipText & GetAlarmMessage(Index))
         .ClockPlateColor = &HE4A75A
         .ClockPlateGradientColor = &HFFFFE0
         .DigitsForeColor = &HC01FC0
         .HandSecondColor = &HFF&
         .HandsHourMinuteColor = &HB93B00
         .MarksHourColor = &HFF&
         .MarksShows = Hours
         .NameForeColor = &HC01FC0
         
      Else
         .AlarmTime = ""
         .AlarmToolTipText = ""
         .ClockPlateColor = &HA4A7A7
         .ClockPlateGradientColor = &HFAFDFD
         .DigitsForeColor = &H808080
         .HandSecondColor = &HE0E0E0
         .HandsHourMinuteColor = &H404040
         .MarksHourColor = &H808080
         .MarksShows = Quarters
         .NameForeColor = &H808080
      End If
      
      .Active = Activate
      .Locked = False
   End With

End Sub

Private Sub FavoritDelete()

   With clkFavorits.Item(SelectedFavorit)
      If Not .Active Or (SelectedClock < 2) Then Exit Sub
      If AppSettings(SET_ASKCONFIRM) Then If ShowMessage(Replace(AppError(7), "$", AllZones(FavoritsInfo(SelectedFavorit).ZoneID).DisplayName, , 1), vbQuestion, AppError(5), AppText(58), TimeToWait) = vbNo Then Exit Sub
      
      TotalFavorits = TotalFavorits - 1
      tmrAnimation.Enabled = True
      .Locked = True
      .Picture = Nothing
      .ClockPlateGradientStyle = InToOut
      .Locked = False
      
      With FavoritsInfo(SelectedFavorit)
         .Index = -1
         .ZoneID = -1
         .DisplayName = ""
         .AlarmTime = ""
         .AlarmMessage = ""
         .AlarmTipText = ""
         .ImageFile = ""
      End With
      
      If SelectedTrayClock = SelectedClock Then SelectedTrayClock = 0
      If TotalFavorits = 0 Then flbChoose.Item(8).ToolTipText = GetToolTipText(AppText(7))
      
      Call FavoritActivate(SelectedFavorit, False, AppText(6))
      
      .ClockPlateColor = &H757878
   End With

End Sub

Private Sub FavoritSelected(ByVal Index As Integer, ByVal Selected As Boolean)

   With clkFavorits.Item(Index)
      If Not .Active Then
         If Selected Then
            .ClockPlateColor = &H757878
            
         Else
            .ClockPlateColor = &HA4A7A7
         End If
         
         Exit Sub
      End If
      
      .Locked = True
      .DigitsBold = Selected
      .NameBold = Selected
      
      If Selected Then
         .ClockPlateColor = &HB5782B
         .HandsHourMinuteStyle = Bold
         .HandSecondScroll = Smooth
         .MarksShows = HoursMinutes
         
      Else
         .ClockPlateColor = &HE4A75A
         .HandsHourMinuteStyle = Thin
         .HandSecondScroll = ByTime
         .MarksShows = Hours
      End If
      
      .Locked = False
   End With

End Sub

Private Sub LoadCalendar(ByVal TimeToGoOnly As Boolean)

   MousePointer = vbHourglass
   Load frmCalendar
   frmCalendar.TimeToGoOnly = TimeToGoOnly
   
   If TimeToGoOnly Then Call frmCalendar.CreateTimeToGo
   
   Call ToggleControls(False)
   
   frmCalendar.Show vbModal, Me

End Sub

Private Sub ModemHangUp()

Const RAS_RASCONNSIZE   As Integer = 412

Dim lngCount            As Long
Dim lngBuffer           As Long
Dim rasConnections(255) As RasConnection

   rasConnections(0).dwSize = RAS_RASCONNSIZE
   lngBuffer = RAS_MAXENTRYNAME * rasConnections(0).dwSize
   
   If RasEnumConnections(rasConnections(0), lngBuffer, lngCount) = 0 Then
      For lngCount = 0 To lngCount - 1
         RasHangUp ByVal rasConnections(lngCount).hRasConn
      Next 'lngCount
   End If
   
   Erase rasConnections

End Sub

Private Sub MouseHook()

   ChangeDateTime = True
   Set clsMouse = New clsMouseWheel
   
   Call SelectDateTimePart
   Call clsMouse.Hook(hWnd)

End Sub

Private Sub MouseUnhook()

   If clsMouse Is Nothing Then Exit Sub
   
   Call clsMouse.Unhook
   
   Set clsMouse = Nothing
   ChangeDateTime = False
   DoEvents

End Sub

Private Sub PopupMenuHooking(ByVal hWnd As Long, Optional ByVal IsClock As Boolean, Optional ByVal MenuItems As Integer, Optional ByVal OpenName As String)

   If hWndPopupMenu <> hWnd Then
      Call SubclassPopupMenu(hWndPopupMenu)
      
      hWndPopupMenu = hWnd
      
      ' check for no add favorit and no delete favorit!
      If MenuItems Then Call SubclassPopupMenu(hWndPopupMenu, IsClock, MenuItems, OpenName)
   End If
   
   DoEvents

End Sub

Private Sub PopupMenuTracking()

Dim lngMenuClick As Long

   InPopupMenu = True
   lngMenuClick = PopupMenuTrack(hWndPopupMenu)
   Refresh
   DoEvents
   InPopupMenu = False
   IsRightButton = False
   
   Select Case lngMenuClick
      Case 1000 ' show clock info
         Call ShowInfoOrMap(True, True)
         
      Case 1100 ' show timezone map
         Call ShowInfoOrMap(True, False)
         
      Case 1200 ' get clock image
         Call OpenClockImage(Me)
         
         ImageName = ""
         
      Case 1300 ' set alarm message
         Call OpenAlarmMessage(Me)
         
      Case 1400, 1500 ' add or delete favorit
         ' 1400 (7) = add favorit, 1500 (8) = delete favorit
         Call flbChoose_Click(7 + (lngMenuClick Mod 1400) / 100)
         
      Case 2000 ' show time between dates
         Call LoadCalendar(True)
         
      Case 2100 To 3100
         If lngMenuClick < 2400 Then
            ' set show type to show time to go
            ' 2100 (0) = total time,    2200 (1) = passed time
            ' 2300 (2) = remaining time
            TimeToGoShowType = (lngMenuClick Mod 2100) / 100
            
         Else
            ' set time type to show differences
            ' 2400 (0) = seconds,  2500 (1) = minutes, 2600 (2) = hours
            ' 2700 (3) = days,     2800 (4) = weeks,   2900 (5) = months
            ' 3000 (6) = quarters, 3100 (7) = years
            TimeToGoShow = (lngMenuClick Mod 2400) / 100
         End If
         
         ledDisplay.ToolTipText = GetTimeToGoToolTipText
         ShowTimeToGo ledDisplay
   End Select
   
   Call PopupMenuHooking(0)

End Sub

Private Sub ResetAllClocks()

Const HWND_BROADCAST As Long = &HFFFF&
Const WM_TIMECHANGE  As Long = &H1E

Dim intCount         As Integer

   UserTimer = Now
   UpdateClock = UserTimer
   txtTime.Text = Format(UserTimer, "hh:mm:ss")
   
   Call SetNewCalendarDay(True)
   Call SetDateButton
   Call SetTimeButton
   
   txtDate.Text = Format(UserTimer, "d mmmm yyyy")
   
   For intCount = 0 To 4
      If clkFavorits.Item(intCount).Active Then Call SetFavoritClockDateTime(intCount, FavoritsInfo(intCount).ZoneID)
   Next 'intCount
   
   SendMessage HWND_BROADCAST, WM_TIMECHANGE, 0, ByVal 0&

End Sub

Private Sub SelectDateTimePart()

Dim intPointer As Integer

   If Not ChangeDateTime Then Exit Sub
   
   If SelectedDatePart < 4 Then
      With txtDate
         intPointer = InStr(.Text, " ")
         
         If SelectedDatePart = 1 Then
            .SelStart = 0
            .SelLength = intPointer - 1
            
         ElseIf SelectedDatePart = 2 Then
            .SelStart = intPointer
            .SelLength = InStr(intPointer + 1, .Text, " ") - intPointer - 1
            
         ElseIf SelectedDatePart = 3 Then
            .SelStart = Len(.Text) - 4
            .SelLength = 4
         End If
         
         .SetFocus
      End With
      
   Else
      With txtTime
         .SelStart = Choose(SelectedDatePart - 3, 0, 3, Len(.Text) - 2)
         .SelLength = 2
         .SetFocus
      End With
   End If
   
   DoEvents

End Sub

Private Sub SetDateButton()

Dim intCount As Integer
Dim strDay   As String

   With flbChoose.Item(6)
      .BackStyle = Transparent
      .IconX = 9
      .IconY = 47
      .Icon = imgImages.Item(6).Picture
      picDateTimeButton.DrawWidth = 1
      picDateTimeButton.Picture = .Picture
      
      With picToDay
         strDay = Right(" " & Day(Date), 2)
         .Cls
         .CurrentX = (.ScaleWidth - .TextWidth(strDay)) \ (Len(strDay) * (10 - (7 And (ScreenResize > 1))))
         
         For intCount = 1 To 2
            picToDay.Print Mid(strDay, intCount, 1);
            .CurrentX = .CurrentX - 1
         Next 'intCount
         
         For intCount = 2 To 16
            StretchBlt picDateTimeButton.hDC, 14 + intCount * 0.53, 56 + intCount, 17, 1, .hDC, 0, intCount, 18, 1, vbSrcCopy
         Next 'intCount
      End With
      
      .BackStyle = Opaque
      .Picture = picDateTimeButton.Image
      .Icon = Nothing
      picDateTimeButton.DrawWidth = 2
   End With

End Sub

Private Sub SetFavoritAlarmToolTip(ByVal Index As Integer)

   If AppSettings(SET_SHOWTIPTEXT) Then
      clkFavorits.Item(Index).AlarmToolTipText = GetToolTipText(FavoritsInfo(Index).AlarmTipText & GetAlarmMessage(Index))
      
   Else
      clkFavorits.Item(Index).AlarmToolTipText = ""
   End If

End Sub

Private Sub SetFavoritClockDateTime(ByVal Index As Integer, ByVal ZoneID As Integer)

Dim dteDate As Date

   With clkFavorits.Item(Index)
      dteDate = UTCToLocalDate(GetSystemDate, AllZones(ZoneID))
      .SetDate = Format(dteDate, DefaultDateFormat)
      .SetTime = Format(dteDate, "hh:mm:ss")
   End With

End Sub

Private Sub SetFavoritClocks()

Dim intCount     As Integer
Dim strClockName As String

   For intCount = 0 To 4
      strClockName = Replace(FavoritsInfo(intCount).DisplayName & ",", ")", "),")
      
      If strClockName <> "," Then
         If InStr(strClockName, "*") Then
            strClockName = Split(strClockName, "*", 2)(1)
            
         Else
            strClockName = Split(strClockName, ")", 2)(1)
         End If
         
         Call FavoritActivate(intCount, (FavoritsInfo(intCount).ZoneID <> -1), TrimClockName(Trim(Split(strClockName, ",")(0))))
         
         If Len(FavoritsInfo(intCount).ImageFile) Then Call SetClockImage(intCount + 2, FavoritsInfo(intCount).ImageFile)
      End If
   Next 'intCount

End Sub

Private Sub SetNewCalendarDay(ByVal Force As Boolean)

   If Not Force And (DateDiff("d", SystemDate, UserTimer) = 0) Then Exit Sub
   
   txtDate.Text = Format(UserTimer, "d mmmm yyyy")
   SystemDate = Now
   
   Call SetNewDateTime
   Call SetDateButton

End Sub

Private Sub SetNewDateTime()

Dim dteDate As Date

   clkTimeZone.Item(SelectedZone).SetDate = CDate(txtDate.Text)
   clkTimeZone.Item(SelectedZone).SetTime = CDate(txtTime.Text)
   dteDate = UTCToLocalDate(LocalDateToUTC(CDate(txtDate.Text) + CDate(txtTime.Text), AllZones(cmbTimeZones.ItemData(IndexFromTimeZone))), AllZones(cmbTimeZones.ItemData(IndexToTimeZone)))
   clkTimeZone.Item(1 - SelectedZone).SetDate = DateSerial(Year(dteDate), Month(dteDate), Day(dteDate))
   clkTimeZone.Item(1 - SelectedZone).SetTime = TimeSerial(Hour(dteDate), Minute(dteDate), Second(dteDate))

End Sub

Private Sub SetScrews(ByVal Remove As Boolean)

   ScrewItem = -2
   IsRemove = Remove
   tmrScrews.Enabled = True

End Sub

Private Sub SetTimeButton()

Dim intCenter  As Integer
Dim sngHours   As Single
Dim sngMinutes As Single

   With flbChoose.Item(0)
      .BackStyle = Transparent
      .IconX = 9
      .IconY = 47
      .Icon = imgImages.Item(0).Picture
      picDateTimeButton.Picture = .Picture
      intCenter = picDateTimeButton.ScaleWidth \ 2 - 1
      sngHours = (360 - (Hour(Time) + Minute(Time) / 60) * 30) * PI_PART
      sngMinutes = (360 - Minute(Time) * 6) * PI_PART
      picDateTimeButton.Line (intCenter, 62)-(intCenter - (8 * Sin(sngMinutes)), 62 - (8 * Cos(sngMinutes))), QBColor(1)
      picDateTimeButton.Line (intCenter, 62)-(intCenter - (6 * Sin(sngHours)), 62 - (6 * Cos(sngHours))), QBColor(1)
      picDateTimeButton.Circle (intCenter, 62), 1, &HFF&
      .BackStyle = Opaque
      .Picture = picDateTimeButton.Image
      .Icon = Nothing
      DoEvents
   End With

End Sub

Private Sub SetTimeZones()

Dim intCount    As Integer
Dim intTotal    As Integer
Dim intValue    As Integer
Dim strName     As String
Dim strTimeZone As String

   With lstSort
      For intCount = 0 To UBound(AllZones)
         .AddItem AllZones(intCount).DisplayName
         .ItemData(.NewIndex) = intCount
      Next 'intCount
   End With
   
   With cmbTimeZones
      For intCount = UBound(AllZones) To 0 Step -1
         strTimeZone = Mid(lstSort.List(intCount), 5, 6)
         
         For intValue = intCount To 0 Step -1
            If Mid(lstSort.List(intValue), 5, 6) <> strTimeZone Then
               intValue = intValue + 1
               intTotal = intValue
               Exit For
            End If
         Next 'intValue
         
         For intValue = intValue To intCount
            strName = lstSort.List(intValue)
            
            If Left(strName, 5) = "(GMT+" Then
               intTotal = 0
               Exit For
            End If
            
            intCount = lstSort.ItemData(intValue)
            .AddItem strName
            .ItemData(.NewIndex) = intCount
         Next 'intValue
         
         intCount = intTotal
      Next 'intCount
      
      For intCount = 0 To UBound(AllZones)
         strName = lstSort.List(intCount)
         
         If Left(strName, 5) = "(GMT-" Then Exit For
         
         intValue = lstSort.ItemData(intCount)
         .AddItem strName
         .ItemData(.NewIndex) = intValue
      Next 'intCount
      
      strTimeZone = GetCurrentTimeZone
      
      For intCount = 0 To .ListCount - 1
         If AllZones(.ItemData(intCount)).StandardName = strTimeZone Then
            .ListIndex = intCount
            IndexSystemTimeZone = intCount
            IndexFromTimeZone = intCount
            IndexToTimeZone = intCount
            Exit For
         End If
      Next 'intCount
   End With
   
   lstSort.Clear

End Sub

Private Sub SetToolTipText()

Dim intCount   As Integer
Dim intTipText As Integer

   For intCount = 0 To flbChoose.Count - 1
      intTipText = 50 + GetIcon(intCount) - (7 And ((intCount = 8) And (AppSettings(SET_LOCKFAVORITS)))) - (17 And ((intCount = 1) And (AppSettings(SET_LOCKSYSTEMDATE))))
      
      If intCount = 1 Then If Not AppSettings(SET_LOCKSYSTEMDATE) Then intTipText = intTipText + (17 And NoInternet) + (18 And (TimeServerURL = AppText(198)))
      If intCount < imgMove.Count Then imgMove.Item(intCount).ToolTipText = GetToolTipText(AppText(80))
      If intCount < clkFavorits.Count Then Call SetFavoritAlarmToolTip(intCount)
      
      flbChoose.Item(intCount).ToolTipText = GetToolTipText(AppText(intTipText))
   Next 'intCount
   
   cmbTimeZones.ToolTipText = GetToolTipText(AppText(1))
   twvDate.ToolTipText = GetToolTipText(AppText(81))
   txtDate.ToolTipText = GetToolTipText(AppText(82))
   txtTime.ToolTipText = GetToolTipText(AppText(24))
   flbControlBox.Item(0).ToolTipText = GetToolTipText(AppText(90))
   flbControlBox.Item(1).ToolTipText = GetToolTipText(AppText(91))
   
   Call CheckFavoritIsSet

End Sub

Private Sub SetZoneClockToolTipText(ByVal Index As Integer)

   clkTimeZone.Item(Index).ToolTipText = GetToolTipText(GetTimeZoneText(cmbTimeZones.ItemData(ZonesInfo(Index).Index), True))

End Sub

Private Sub ShowActiveClock()

Dim objClock As Object

   If SelectedClock < 2 Then
      Set objClock = clkTimeZone.Item(SelectedClock)
      
   Else
      Set objClock = clkFavorits.Item(SelectedClock - 2)
   End If
   
   With shpActiveClock
      .Top = objClock.Top
      .Left = objClock.Left
      .Width = objClock.Width
      .Height = objClock.Height
      Set objClock = Nothing
   End With

End Sub

Private Sub ShowInfoOrMap(ByVal Settings As Boolean, ByVal ShowInfo As Boolean)

   DoEvents
   ShowSettings = Settings
   
   Call TextBoxUnhook
   
   If SelectedClock > 1 Then
      If Not clkFavorits.Item(SelectedFavorit).Active Then Exit Sub
      
      SelectedListIndex = FavoritsInfo(SelectedFavorit).Index
      
   Else
      SelectedListIndex = IndexFromTimeZone
   End If
   
   MousePointer = vbHourglass
   
   If ShowInfo Then
      frmInfo.Show vbModal, Me
      
   Else
      frmMap.Show vbModal, Me
   End If

End Sub

Private Sub SysTrayModifyTipText()

Dim strDate As String
Dim strName As String

   If SelectedTrayClock < 2 Then
      strName = clkTimeZone.Item(SelectedTrayClock).NameClock
      strDate = clkTimeZone.Item(SelectedTrayClock).DateTime
      
   Else
      strName = clkFavorits.Item(SelectedTrayClock - 2).NameClock
      strDate = clkFavorits.Item(SelectedTrayClock - 2).DateTime
   End If
   
   clsSysTray.TipText = strName & ": " & CapsText(FormatDateTime(strDate, vbLongDate)) & " - " & Format(strDate, "hh:mm")

End Sub

Private Sub TextBoxUnhook()

   If SubclassedTextBox = 1 Then
      Call SubclassTextBox(txtDate.hWnd)
      
   ElseIf SubclassedTextBox = 2 Then
      Call SubclassTextBox(txtTime.hWnd)
   End If
   
   SubclassedTextBox = 0

End Sub

Private Sub TimeZoneActivate(ByVal Index As Integer)

   With clkTimeZone.Item(Index)
      .Locked = True
      .NameClock = GetTimeZoneText(ZonesInfo(Index).ZoneID)
      .Active = True
      .Locked = False
   End With

End Sub

Private Sub TimeZoneSelected(ByVal Selected As Boolean)

   With clkTimeZone.Item(SelectedZone)
      .Locked = True
      
      If Selected Then
         .ClockPlateColor = &HA35111
         .MarksMinuteColor = &HB0B000
         
      Else
         .ClockPlateColor = &HD28040
         .MarksMinuteColor = &HC0C000
      End If
      
      Call SetNewDateTime
      
      .Locked = False
   End With
   
   SetChanges = False
   cmbTimeZones.ListIndex = IndexFromTimeZone
   SetChanges = True

End Sub

Private Sub clkFavorits_Alarm(Index As Integer)

Static blnBusy As Boolean

Dim strMessage As String

   Do While blnBusy
      DoEvents
   Loop
   
   Call SetAlarmOff
   
   blnBusy = True
   
   With shpAlarm
      SendKeys "{Esc}" ' close popupmenu if it is open
      strMessage = CreateAlarmMessage(Index)
      
      Do While tmrAlarmOff.Enabled
         DoEvents
      Loop
      
      ' clear display first
      If Len(strMessage) Then ledDisplay.Text = ""
      
      .Top = clkFavorits.Item(Index).Top
      .Left = clkFavorits.Item(Index).Left
      .Visible = True
   End With
   
   With ledDisplay
      If Len(strMessage) Then
         .Text = strMessage
         .NoTextScrolling = False
         .BackColor = &HA8ABAB
         .ToolTipText = GetDisplayToolTipText(clkFavorits.Item(Index).NameClock)
         .Active = True
         
      ElseIf CheckTimeToGo Then
         .Active = True
         
      Else
         .Text = ""
      End If
      
      tmrAlarmOff.Enabled = True
      AlarmIndex = Index
   End With
   
   Call PlaySound(SOUND_SIGNAL)
   
   blnBusy = False

End Sub

Private Sub clkFavorits_DblClick(Index As Integer)

   If (Index <> SelectedFavorit) Or IsRightButton Or InPopupMenu Then Exit Sub
   
   If IsLeftButton Then
      IsLeftButton = False
      
      Call ShowInfoOrMap(True, True)
   End If

End Sub

Private Sub clkFavorits_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

   IsLeftButton = (Button = vbLeftButton)

End Sub

Private Sub clkFavorits_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

   Call PopupMenuHooking(clkFavorits.Item(Index).hWnd, True, 7, clkFavorits.Item(Index).NameClock)

End Sub

Private Sub clkFavorits_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Static intSelectedClock As Integer

   If tmrAnimation.Enabled Then Exit Sub
   
   intSelectedClock = SelectedClock
   SelectedClock = Index + 2
   SelectedTimeZoneID = FavoritsInfo(Index).ZoneID
   IsRightButton = (Button = vbRightButton)
   
   Call ShowActiveClock
   
   If Index <> SelectedFavorit Then
      Call FavoritSelected(SelectedFavorit, False)
      Call FavoritSelected(Index, True)
      
      SelectedFavorit = Index
   End If
   
   Call CheckFavoritIsSet
   
   If Not IsRightButton Or (intSelectedClock <> SelectedClock) Then
      If FavoritsInfo(Index).Index > -1 Then
         If cmbTimeZones.ListIndex <> FavoritsInfo(Index).Index Then
            cmbTimeZones.ListIndex = FavoritsInfo(Index).Index
            SelectedTimeZoneID = FavoritsInfo(Index).ZoneID
         End If
         
      Else
         SelectedTimeZoneID = cmbTimeZones.ItemData(IndexFromTimeZone)
      End If
   End If
   
   If IsRightButton Then Call PopupMenuTracking

End Sub

Private Sub clkTimeZone_DblClick(Index As Integer)

   If (Index <> SelectedZone) Or IsRightButton Or InPopupMenu Then Exit Sub
   
   If IsLeftButton Then
      IsLeftButton = False
      
      Call ShowInfoOrMap(True, True)
   End If

End Sub

Private Sub clkTimeZone_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

   IsLeftButton = (Button = vbLeftButton)

End Sub

Private Sub clkTimeZone_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

   Call PopupMenuHooking(clkTimeZone.Item(Index).hWnd, True, 3, clkTimeZone.Item(Index).NameClock)
   
End Sub

Private Sub clkTimeZone_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

   If tmrAnimation.Enabled Then Exit Sub
   If cmbTimeZones.ListIndex <> IndexFromTimeZone Then cmbTimeZones.ListIndex = IndexFromTimeZone
   
   SelectedClock = Index
   SelectedTimeZoneID = cmbTimeZones.ItemData(IndexFromTimeZone)
   IsRightButton = (Button = vbRightButton)
   
   Call ShowActiveClock
   
   If Index <> SelectedZone Then
      Call TimeZoneSelected(False)
      
      IndexToTimeZone = IndexToTimeZone Xor IndexFromTimeZone
      IndexFromTimeZone = IndexToTimeZone Xor IndexFromTimeZone
      IndexToTimeZone = IndexToTimeZone Xor IndexFromTimeZone
      SelectedZone = Index
      SelectedTimeZoneID = cmbTimeZones.ItemData(IndexFromTimeZone)
      
      Call TimeZoneSelected(True)
   End If
   
   Call CheckFavoritIsSet
   
   If IsRightButton Then Call PopupMenuTracking

End Sub

Private Sub clsMouse_Wheel(ScrollLines As Integer)

   twvDate.ScrollValue = twvDate.ScrollValue + ScrollLines

End Sub

Private Sub clsSysTray_Click(Button As Integer)

   If Visible Then
      Call TextBoxUnhook
      Call PopupMenuHooking(0)
      
   Else
      With frmPopupMenu
         If Button = vbLeftButton Then
            If .Visible Then Call .EndPopupMenu
            
            Call SysTrayDisplay(False)
         End If
         
         If Button = vbRightButton Then
            Call .ActivatePopupMenu
            
            .Show vbModal, Me
            
            If SelectedPopupMenu = 7 Then Call SysTrayDisplay(False)
            
            If SelectedPopupMenu = 8 Then
               Call SysTrayDisplay(False, False)
               Call EndMyTimeZones
            End If
         End If
      End With
      
      SelectedPopupMenu = -1
   End If

End Sub

Private Sub clsSysTray_MouseMove()

   clsSysTray.HideBalloon

End Sub

Private Sub clsSysTray_ReceivedData(Data As String)

   Call SysTrayDisplay(False)

End Sub

Private Sub cmbTimeZones_Click()

   If Not Visible Then Exit Sub
   
   SelectedListIndex = cmbTimeZones.ListIndex
   
   If InZoneInfo Or InZoneMap Then
      SelectedTimeZoneID = cmbTimeZones.ItemData(SelectedListIndex)
      
      If InZoneMap Then
         Call frmMap.CreateWindow
         
      Else
         Call frmInfo.CreateWindow
      End If
      
      Exit Sub
   End If
   
   If SelectedClock < 2 Then
      With ZonesInfo(SelectedZone)
         IndexFromTimeZone = cmbTimeZones.ListIndex
         .DisplayName = cmbTimeZones.Text
         .Index = IndexFromTimeZone
         .ZoneID = cmbTimeZones.ItemData(IndexFromTimeZone)
         AutoDaylightTimeSet = True
         NoDaylightChanges = False
      End With
      
      Call SetZoneClockToolTipText(SelectedZone)
      Call SetNewDateTime
      Call ChangeDaylightTimeSet(True)
      
      If SetChanges Then Call TimeZoneActivate(SelectedZone)
      
      Call clkTimeZone_MouseUp(SelectedClock, (vbRightButton And IsRightButton), 0, 0, 0)
   End If
   
   Call SetToolTipText

End Sub

Private Sub flbChoose_Click(Index As Integer)

Dim intPrevListIndex As Integer

   ' for slower systems first check if all events are handled,
   ' if so then blnInUse = False
   If IsBusy Then Exit Sub
   
   IsBusy = True
   
   Select Case Index
      Case 0 ' reset all clocks with system time
         NoDaylightChanges = False
         
         Call CheckDaylightChanges
         Call ResetAllClocks
         
      Case 1 ' synchronise system clock with internet
         If NoInternet Or (TimeServerURL = AppText(198)) Then
            If AppSettings(SET_ASKCONFIRM) Then ShowMessage AppText(96 + NoInternet), vbStop, AppError(13), AppError(14), TimeToWait
            
            GoTo ExitSub
         End If
         
         Call AtomicSynchronise(False)
         
      Case 2 ' set system date/time with clock date/time
         If AppSettings(SET_LOCKSYSTEMDATE) Then
            If AppSettings(SET_ASKCONFIRM) Then ShowMessage AppText(61), vbStop, AppError(13), AppError(16), TimeToWait
            
            GoTo ExitSub
         End If
         
         If AppSettings(SET_ASKCONFIRM) Then If ShowMessage(AppError(17), vbQuestion, AppError(13), AppError(16), TimeToWait) = vbNo Then GoTo ExitSub
         
         NoDaylightChanges = False
         IndexSystemTimeZone = IndexFromTimeZone
         Time = CDate(txtTime.Text)
         Date = CDate(txtDate.Text)
         
         If ledDisplay.Active Then ShowTimeToGo ledDisplay
         
         Call SetTimeZoneInfo(cmbTimeZones.ItemData(IndexFromTimeZone))
         Call ResetAllClocks
         
      Case 3 ' set daylighttimeset on/off
         If NoDaylightTimeSet + AppSettings(SET_LOCKSYSTEMDATE) Then GoTo ExitSub
         
         AutoDaylightTimeSet = Not AutoDaylightTimeSet
         NoDaylightChanges = True
         
         Call ChangeDaylightTimeSet(False)
         
      Case 4, 5 ' 4 = show timzone information, 5 = show timzone map
         If (SelectedClock > 1) And (FavoritsInfo(SelectedFavorit).DisplayName = "") Then
            ShowMessage AppError(43), vbStop, AppError(5), AppError(0), TimeToWait
            
         Else
            intPrevListIndex = cmbTimeZones.ListIndex
            
            Call ShowInfoOrMap(False, (True And (Index = 4)))
            
            cmbTimeZones.ListIndex = intPrevListIndex
            
            Call cmbTimeZones_Click
         End If
         
      Case 6 ' show calendar
         Call LoadCalendar(False)
         
      Case 7, 8 ' 7 = add favorit, 8 = delete favorit
         If tmrAnimation.Enabled Then GoTo ExitSub
         
         If AppSettings(SET_LOCKFAVORITS) And (SelectedClock > 1) Then
            If AppSettings(SET_ASKCONFIRM) Then ShowMessage AppText(64), vbStop, AppError(5), AppText(50 + Index), TimeToWait
            
            GoTo ExitSub
         End If
         
         If AppSettings(SET_ASKCONFIRM) Then
            If Trim(flbChoose.Item(Index).ToolTipText) = AppText(73) Then
               ShowMessage AppText(73), vbStop, AppError(5), AppText(50 + Index), TimeToWait
               GoTo ExitSub
               
            ElseIf (Index = 7) And FavoritsInfo(SelectedFavorit).Index = cmbTimeZones.ListIndex Then
               ShowMessage AppError(51), vbStop, AppError(5), AppText(50 + Index), TimeToWait
               GoTo ExitSub
               
            ElseIf (Index = 8) And (Trim(flbChoose.Item(Index).ToolTipText) = AppText(7)) Then
               ShowMessage AppError(50), vbStop, AppError(5), AppText(50 + Index), TimeToWait
               GoTo ExitSub
            End If
         End If
         
         Animation = Val(imgImages.Item(Index).Tag)
         ButtonAnimation = Index
         
         If Index = 7 Then
            Call FavoritAdd
            
         Else
            Call FavoritDelete
         End If
         
      Case 9 ' open setting page
         MousePointer = vbHourglass
         Load frmSettings
         
         Call ToggleControls(False)
         
         frmSettings.Show vbModal, Me
         
         Call SetToolTipText
         Call ChangeDaylightTimeSet(True)
         
      Case 10 ' save settings and exit MyTimeZones
         Call EndMyTimeZones
   End Select
   
ExitSub:
   IsBusy = False

End Sub

Private Sub flbChoose_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

   If Index = 6 Then CalendarButtonIsClicked = True

End Sub

Private Sub flbChoose_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

   If Index = 6 Then CalendarButtonIsClicked = False

End Sub

Private Sub flbControlBox_Click(Index As Integer)

   If Index Then
      Call SysTrayDisplay(True)
      
   Else
      WindowState = vbMinimized
   End If

End Sub

Private Sub Form_Initialize()

   If App.PrevInstance Then Call ActivatePrevInstance
   
   Call InitialiseCommonControls
   
   RunSilent = (UCase(Command) = "/SILENT")

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

   If IsExit(KeyCode, Shift) Then Call EndMyTimeZones

End Sub

Private Sub Form_Load()

Dim intCount As Integer
Dim sngTimer As Single

   IsBusy = True
   sngTimer = Timer
   Caption = App.Title
   DataPath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "Data\"
   AutoSyncTime = DateAdd("n", 2, Time)
   ScreenResize = Screen.TwipsPerPixelX / 12
   AlarmIndex = -1
   AnimationhDC = picAnimation.hDC
   
   Call SetDefaults
   Call SetRegKeys
   
   With App
      InfoVersion = AppText(10) & " " & .Major & "." & .Minor & .Revision
      InfoCopyright = Replace(.LegalCopyright, "by", AppText(11))
   End With
   
   If AppSettings(SET_SHOWINTROSCREEN) And Not RunSilent Then
      InIntro = True
      Load frmIntro
   End If
   
   With Screen
      Height = Height * ScreenResize - (Height - ScaleHeight * .TwipsPerPixelY) * (ScreenResize - Int(ScreenResize))
      Width = 9480 * ScreenResize
      Top = (.Height - Height) \ 2 - 16 * .TwipsPerPixelY
      Left = (.Width - Width) \ 2
      FormTop = Top
      FormLeft = Left
   End With
   
   Call ResizeAllControls(Me, True)
   Call GetAllTimeZones
   Call SetTimeZones
   Call GetRegValues
   
   With twvDate
      txtDate.Height = 19
      txtDate.Top = .Top + .Height / 2 - txtDate.Height - 3
      txtTime.Height = txtDate.Height
      txtTime.Top = .Top + .Height / 2 + 2
   End With
   
   With picDateTimeButton
      .Height = .Height * ScreenResize
      .Width = .Width * ScreenResize
   End With
   
   twvDate.MouseTrap = AppSettings(SET_MOUSEINTHUMBWHEEL)
   clkTimeZone.Item(0).Font.Size = 9 / ScreenResize
   clkTimeZone.Item(0).Locked = True
   Load clkTimeZone.Item(1)
   Load shpClockBorder.Item(1)
   clkTimeZone.Item(1).Left = 511
   clkTimeZone.Item(1).Visible = True
   shpClockBorder.Item(1).Left = 512
   shpClockBorder.Item(1).Visible = True
   CalendarDate = Date
   SelectedDatePart = 5
   clkFavorits.Item(0).Font.Size = 8 / ScreenResize
   clkFavorits.Item(0).Locked = True
   GetInternetConnection
   NoInternet = Not InternetConnection.Configured
   AddToFontSize = ((3 And (ScreenResize > 1)) / 2)
   picToDay.FontSize = 10 + AddToFontSize
   
   Call CreateForm
   Call RoundForm(hWnd, ScaleWidth, ScaleHeight, 65)
   Call LoadSpecialDays
   
   For intCount = 0 To 10
      If intCount Then
         Load flbChoose.Item(intCount)
         
         With flbChoose.Item(intCount)
            .Top = flbChoose.Item(0).Top
            .Left = flbChoose.Item(intCount - 1).Left + .Width
            .Visible = True
         End With
      End If
      
      Call SetButton(intCount, GetIcon(intCount), Me)
   Next 'intCount
   
   For intCount = 0 To 4
      If intCount Then
         Load clkFavorits.Item(intCount)
         
         With clkFavorits.Item(intCount)
            .Left = clkFavorits.Item(intCount - 1).Left + 154
            .Visible = True
         End With
      End If
      
      Call FavoritActivate(intCount, False, AppText(6))
   Next 'intCount
   
   For intCount = 0 To 1
      Call TimeZoneActivate(intCount)
      Call SetZoneClockToolTipText(intCount)
      
      If Len(ZonesInfo(intCount).ImageFile) Then Call SetClockImage(intCount, ZonesInfo(intCount).ImageFile)
   Next 'intCount
   
   If SelectedClock < 2 Then
      SelectedTimeZoneID = cmbTimeZones.ItemData(IndexFromTimeZone)
      
   Else
      SelectedTimeZoneID = FavoritsInfo(SelectedFavorit).ZoneID
   End If
   
   IndexFromTimeZone = ZonesInfo(SelectedZone).Index
   IndexToTimeZone = ZonesInfo(1 And Not SelectedZone).Index
   shpActiveClock.Visible = AppSettings(SET_ACTIVEBORDER)
   
   Call SetFavoritClocks
   Call ResetComboBox(False)
   Call ResetAllClocks
   Call TimeZoneSelected(True)
   Call FavoritSelected(SelectedFavorit, True)
   Call ShowActiveClock
   Call SetToolTipText
   Call ChangeDaylightTimeSet(True)
   
   If InIntro Then
      InIntro = False
      
   ElseIf Not RunSilent Then
      Show
      DoEvents
   End If
   
   If (SelectedClock > 1) And (FavoritsInfo(SelectedFavorit).Index > -1) Then cmbTimeZones.ListIndex = FavoritsInfo(SelectedFavorit).Index
   If RunSilent Then Call SysTrayDisplay(True)
   If AppSettings(SET_AUTOSYNCHRONISE) Then Call AtomicSynchronise(True)
   
   If Timer - sngTimer < 0.5 Then
      DisplaySpeed = Slow
      
   ElseIf Timer - sngTimer > 3.5 Then
      DisplaySpeed = Fast
      
   Else
      DisplaySpeed = Default
   End If
   
   Call SetNewDateTime
   
   With ledDisplay
      .Width = 313
      .Speed = DisplaySpeed
      shpDisplay.Top = .Top - 1
      shpDisplay.Left = .Left - 1
      shpDisplay.Width = .Width + 2
      shpDisplay.Height = .Height + 2
   End With
   
   tmrClock.Enabled = True
   SelectedListIndex = cmbTimeZones.ListIndex
   SetTimeToGoDisplay True
   IsBusy = False

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

   If UnloadMode > vbFormCode Then Call EndMyTimeZones(False)

End Sub

Private Sub Form_Resize()

   Call SystemMenuEnableSysTrayItem(hWnd, True And (WindowState <> vbMinimized))

End Sub

Private Sub imgMove_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

   If tmrScrews.Enabled Then Exit Sub
   If Button = vbLeftButton Then Call SetScrews(True)

End Sub

Private Sub imgMove_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Const HTCLIENT As Long = &H1

Dim intCount   As Integer

   If Button = vbLeftButton Then
      ScrewItem = -1
      tmrScrews.Enabled = False
      
      For intCount = 0 To 3
         imgMove.Item(intCount).Picture = imgScrews.Item(4).Picture
      Next 'intCount
      
      ReleaseCapture
      SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&
      SetCapture hWnd
      ReleaseCapture
      PostMessage hWnd, WM_LBUTTONUP, HTCLIENT, X + Y * &H10000
      FormTop = Top
      FormLeft = Left
      
      Call SetScrews(False)
   End If

End Sub

Private Sub imgMove_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

   If Button = vbLeftButton Then
      Call SetScrews(False)
      
   ElseIf Button = vbRightButton Then
      Call SystemMenuEnableSysTrayItem(hWnd, True)
      Call SystemMenuTrack(hWnd, imgMove.Item(Index).MousePointer)
   End If

End Sub

Private Sub ledDisplay_DblClick()

   If IsRightButton Or tmrAlarmOff.Enabled Then Exit Sub
   
   If IsLeftButton Then
      IsLeftButton = False
      
      Call LoadCalendar(True)
   End If

End Sub

Private Sub ledDisplay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   IsLeftButton = (Button = vbLeftButton)

End Sub

Private Sub ledDisplay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

   IsRightButton = (Button = vbRightButton)
   
   If tmrAlarmOff.Enabled Then Exit Sub
   
   If IsRightButton Then
      Call PopupMenuHooking(ledDisplay.hWnd, False, 4, AppText(210))
      Call PopupMenuTracking
   End If

End Sub

Private Sub tmrAlarmOff_Timer()

   Call SetAlarmOff

End Sub

Private Sub tmrAnimation_Timer()

   Animation = Val(imgImages.Item(Animation).Tag)
   
   Call SetButton(ButtonAnimation, Animation, Me)
   
   If Animation = ButtonAnimation Then
      tmrAnimation.Enabled = False
      
      Call CheckFavoritIsSet
   End If

End Sub

Private Sub tmrClock_Timer()

Static blnNoInternet       As Boolean
Static strUpdateDateButton As String
Static strUpdateSyncButton As String
Static strUpdateTimeButton As String

Dim strTimerText           As String

   If strUpdateDateButton <> Date Then
      strUpdateDateButton = Date
      
      If Not CalendarButtonIsClicked Then Call SetDateButton
   End If
   
   If strUpdateTimeButton < Time Then
      strUpdateTimeButton = DateAdd("n", 1, Time)
      
      Call SetTimeButton
   End If
   
   If strUpdateSyncButton < Time Then
      GetInternetConnection
      NoInternet = Not InternetConnection.Configured
      
      If blnNoInternet <> NoInternet Then
         blnNoInternet = NoInternet
         
         If Not AppSettings(SET_LOCKSYSTEMDATE) Then
            Call SetButton(1, GetIcon(1), Me)
            Call SetToolTipText
         End If
      End If
   End If
   
   If AppSettings(SET_AUTOSYNCHRONISE) And (AutoSyncTime < Time) Then
      AutoSyncTime = DateAdd("n", 60, Time)
      
      Call AtomicSynchronise(True)
   End If
   
   For AllFavorits = 0 To 4
      If Len(FavoritsInfo(AllFavorits).DisplayName) Then
         strTimerText = GetTimeZoneText(FavoritsInfo(AllFavorits).ZoneID) & " - " & CapsText(Format(clkFavorits.Item(AllFavorits).DateTime, LongDateFormat))
         
         Call SetFavoritClockDateTime(AllFavorits, FavoritsInfo(AllFavorits).ZoneID)
         
      Else
         strTimerText = AppText(7)
      End If
      
      clkFavorits.Item(AllFavorits).ToolTipText = GetToolTipText(strTimerText)
      
      Call SetFavoritAlarmToolTip(AllFavorits)
   Next 'AllFavorits
   
   If ledDisplay.Active Then
      If tmrAlarmOff.Enabled And Len(CreateAlarmMessage(AlarmIndex)) Then
         ledDisplay.Text = CreateAlarmMessage(AlarmIndex)
         
      ElseIf AppSettings(SET_AUTODELETETIMETOGO) And (Format(Date, "yyyymmdd") >= Format(TimeToGo(1), "yyyymmdd")) Then
         TimeToGo(0) = ""
         TimeToGo(1) = ""
         SetTimeToGoDisplay False
         
      Else
         ShowTimeToGo ledDisplay
      End If
   End If
   
   If UpdateClock <> Now Then
      UserTimer = DateAdd("s", DateDiff("s", UpdateClock, Now), CDate(txtDate.Text) + CDate(txtTime.Text))
      txtTime.Text = Format(UserTimer, "hh:mm:ss")
      UpdateClock = Now
   End If
   
   Call SetNewCalendarDay(False)
   Call CheckDaylightChanges
   
   If ChangeDateTime Then Call SelectDateTimePart
   If OnSysTray Then Call SysTrayModifyTipText

End Sub

Private Sub tmrScrews_Timer()

Dim intScrews As Integer

   If ScrewItem = -2 Then ScrewItem = 3 - (3 And IsRemove)
   
   For intScrews = 0 To 3
      imgMove.Item(intScrews).Picture = imgScrews.Item(ScrewItem).Picture
      
      If ScrewItem + IsRemove = False Then imgMove.Item(intScrews).Picture = Nothing
   Next 'intScrews
   
   ScrewItem = ScrewItem + (1 And IsRemove) - (1 And Not IsRemove)
   
   If (ScrewItem < 0) Or (ScrewItem > 4) Then tmrScrews.Enabled = False

End Sub

Private Sub twvDate_Change()

Dim intCount As Integer

   With twvDate
      intCount = Sgn(.Value - Val(.Tag))
      .Tag = .Value
   End With
   
   If SelectedDatePart < 4 Then
      With txtDate
         .Text = Format(DateSerial(Year(.Text) + (intCount And (SelectedDatePart = 3)), Month(.Text) + (intCount And (SelectedDatePart = 2)), Day(.Text) + (intCount And (SelectedDatePart = 1))), "d mmmm yyyy")
         
         If Year(.Text) < 1900 Then .Text = Format(DateSerial(1900, Month(.Text), Day(.Text)), "d mmmm yyyy")
      End With
      
   Else
      With txtTime
         .Text = Format(TimeSerial(Hour(.Text) + (intCount And (SelectedDatePart = 4)) + (24 And (Hour(.Text) = 0) And (SelectedDatePart = 4)), Minute(.Text) + (intCount And (SelectedDatePart = 5)), Second(.Text) + (intCount And (SelectedDatePart = 6))), "hh:mm:ss")
      End With
   End If
   
   Call SelectDateTimePart
   Call SetNewDateTime

End Sub

Private Sub twvDate_Click()

   Call MouseUnhook
   Call MouseHook

End Sub

Private Sub twvDate_GotFocus()

   Call MouseHook

End Sub

Private Sub twvDate_LostFocus()

   Call MouseUnhook

End Sub

Private Sub txtDate_DblClick()

   Call SelectDateTimePart

End Sub

Private Sub txtDate_GotFocus()

   Call MouseHook

End Sub

Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)

Dim intDatePart As Integer

   If (KeyCode = vbKeyUp) Or (KeyCode = vbKeyDown) Then
      twvDate.ScrollValue = twvDate.ScrollValue + (1 And (KeyCode = vbKeyUp)) - (1 And (KeyCode = vbKeyDown))
      KeyCode = vbEmpty
      
   Else
      intDatePart = GetSelectedDatePart(KeyCode, SelectedDatePart, 3, 1)
      KeyCode = vbEmpty
      
      If intDatePart Then SelectedDatePart = intDatePart
      
      Call SelectDateTimePart
   End If

End Sub

Private Sub txtDate_KeyUp(KeyCode As Integer, Shift As Integer)

   Call SelectDateTimePart

End Sub

Private Sub txtDate_LostFocus()

   Call TextBoxUnhook
   Call MouseUnhook

End Sub

Private Sub txtDate_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim intPointer As Integer

   With txtDate
      intPointer = InStr(.Text, " ")
      SelectedDatePart = 2 - (1 And (.SelStart < intPointer)) + (1 And (.SelStart > InStr(intPointer + 1, .Text, " ")))
   End With
   
   Call SelectDateTimePart

End Sub

Private Sub txtDate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If SubclassedTextBox = 2 Then Call TextBoxUnhook
   
   If SubclassedTextBox = 0 Then
      Call SubclassTextBox(txtDate.hWnd)
      
      SubclassedTextBox = 1
   End If

End Sub

Private Sub txtTime_DblClick()

   Call SelectDateTimePart

End Sub

Private Sub txtTime_GotFocus()

   Call MouseHook

End Sub

Private Sub txtTime_KeyDown(KeyCode As Integer, Shift As Integer)

Dim intDatePart As Integer

   If (KeyCode = vbKeyUp) Or (KeyCode = vbKeyDown) Then
      twvDate.ScrollValue = twvDate.ScrollValue + (1 And (KeyCode = vbKeyUp)) - (1 And (KeyCode = vbKeyDown))
      KeyCode = vbEmpty
      
   Else
      intDatePart = GetSelectedDatePart(KeyCode, SelectedDatePart, 6, 4)
      KeyCode = vbEmpty
      
      If intDatePart Then SelectedDatePart = intDatePart
      
      Call SelectDateTimePart
   End If

End Sub

Private Sub txtTime_KeyUp(KeyCode As Integer, Shift As Integer)

   Call SelectDateTimePart

End Sub

Private Sub txtTime_LostFocus()

   Call TextBoxUnhook
   Call MouseUnhook

End Sub

Private Sub txtTime_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   SelectedDatePart = 5 - (1 And (txtTime.SelStart < 4)) + (1 And (txtTime.SelStart > 5))
   
   Call SelectDateTimePart

End Sub

Private Sub txtTime_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If SubclassedTextBox = 1 Then Call TextBoxUnhook
   
   If SubclassedTextBox = 0 Then
      Call SubclassTextBox(txtTime.hWnd)
      
      SubclassedTextBox = 2
   End If

End Sub
