VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Scopeform 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   Caption         =   "Scope"
   ClientHeight    =   5670
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   8940
   ControlBox      =   0   'False
   Icon            =   "Scopeform.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   378
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   596
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command21 
      Caption         =   "+/-"
      Height          =   375
      Left            =   120
      TabIndex        =   65
      Top             =   3480
      Width           =   375
   End
   Begin VB.TextBox Text18 
      BackColor       =   &H00FF0000&
      Height          =   195
      Left            =   2280
      TabIndex        =   63
      Top             =   4080
      Width           =   150
   End
   Begin VB.TextBox Text17 
      BackColor       =   &H000000FF&
      Height          =   195
      Left            =   2280
      TabIndex        =   62
      Top             =   3600
      Width           =   150
   End
   Begin VB.CommandButton Command20 
      Caption         =   "DC"
      Height          =   375
      Left            =   1680
      TabIndex        =   61
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton Command19 
      Caption         =   "AC"
      Height          =   375
      Left            =   1680
      TabIndex        =   60
      Top             =   3480
      Width           =   495
   End
   Begin VB.TextBox Text16 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   360
      TabIndex        =   59
      Top             =   4680
      Width           =   255
   End
   Begin VB.TextBox Text15 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3480
      TabIndex        =   57
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CURSORS"
      Height          =   255
      Left            =   3840
      TabIndex        =   56
      Top             =   4920
      Width           =   975
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      Left            =   5280
      Max             =   4155
      SmallChange     =   5
      TabIndex        =   55
      Top             =   5280
      Value           =   3000
      Width           =   3495
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   120
      Max             =   4155
      SmallChange     =   5
      TabIndex        =   54
      Top             =   5280
      Value           =   1000
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ZERO"
      Height          =   255
      Left            =   8040
      TabIndex        =   53
      Top             =   3720
      Width           =   615
   End
   Begin VB.VScrollBar VScroll4 
      Height          =   3015
      Left            =   8400
      Max             =   1500
      Min             =   -1500
      TabIndex        =   48
      Top             =   360
      Width           =   255
   End
   Begin VB.VScrollBar VScroll3 
      Height          =   3015
      Left            =   7920
      Max             =   1500
      Min             =   -1500
      TabIndex        =   47
      Top             =   360
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   10
      Left            =   2880
      Max             =   2000
      SmallChange     =   5
      TabIndex        =   46
      Top             =   4440
      Width           =   2775
   End
   Begin VB.TextBox Text14 
      BackColor       =   &H00FF0000&
      Height          =   195
      Left            =   1680
      TabIndex        =   44
      Top             =   4800
      Width           =   150
   End
   Begin VB.TextBox Text13 
      BackColor       =   &H00FF0000&
      Height          =   195
      Left            =   960
      TabIndex        =   43
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox Text12 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   6240
      TabIndex        =   40
      Top             =   4680
      Width           =   855
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1200
      TabIndex        =   37
      Text            =   "1"
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00FF0000&
      Height          =   225
      Left            =   7680
      TabIndex        =   36
      Top             =   4320
      Width           =   150
   End
   Begin VB.CommandButton Command18 
      Caption         =   "REF CLR"
      Height          =   375
      Left            =   7920
      TabIndex        =   35
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton Command17 
      Caption         =   "REF SET"
      Height          =   375
      Left            =   7920
      TabIndex        =   34
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0C0C0&
      Caption         =   "v"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   33
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0C0&
      Caption         =   "^"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   32
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   6240
      TabIndex        =   30
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton Command10 
      Caption         =   "SMOOTH"
      Height          =   375
      Left            =   720
      TabIndex        =   29
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      Caption         =   "AUTO RANGE"
      Height          =   495
      Left            =   600
      TabIndex        =   28
      Top             =   1440
      Width           =   1095
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3015
      LargeChange     =   10
      Left            =   7320
      Max             =   1000
      SmallChange     =   10
      TabIndex        =   26
      Top             =   360
      Value           =   1000
      Width           =   255
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00FF0000&
      Height          =   195
      Left            =   360
      TabIndex        =   25
      Top             =   4080
      Width           =   150
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00FF0000&
      Height          =   195
      Left            =   1320
      TabIndex        =   24
      Top             =   4080
      Width           =   150
   End
   Begin VB.CommandButton Command8 
      Caption         =   "L     R"
      Height          =   375
      Left            =   600
      TabIndex        =   23
      Top             =   3960
      Width           =   615
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FF0000&
      Height          =   195
      Left            =   1320
      TabIndex        =   22
      Top             =   3600
      Width           =   150
   End
   Begin VB.CommandButton Command7 
      Caption         =   "TRIG"
      Height          =   375
      Left            =   600
      TabIndex        =   21
      Top             =   3480
      Width           =   615
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FF0000&
      Height          =   195
      Left            =   1200
      TabIndex        =   20
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "D"
      Height          =   495
      Left            =   1200
      TabIndex        =   19
      Top             =   2880
      Width           =   375
   End
   Begin VB.PictureBox ScopeBuff 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C000C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000002&
      Height          =   336
      Left            =   7320
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton StartButton 
      BackColor       =   &H00C0C0C0&
      Caption         =   "START"
      Height          =   450
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   17
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton StopButton 
      BackColor       =   &H00C0C0C0&
      Caption         =   "STOP"
      Enabled         =   0   'False
      Height          =   450
      Left            =   1440
      TabIndex        =   16
      Top             =   840
      Width           =   735
   End
   Begin VB.ComboBox DevicesBox 
      BackColor       =   &H80000009&
      Height          =   315
      Left            =   6960
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FF0000&
      Height          =   195
      Left            =   1800
      TabIndex        =   14
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FF0000&
      Height          =   195
      Left            =   720
      TabIndex        =   13
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "R"
      Height          =   495
      Left            =   1800
      TabIndex        =   11
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "M"
      Height          =   495
      Left            =   720
      TabIndex        =   10
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "L"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      TabIndex        =   8
      Text            =   "Oscilloscope Plot from PC soundcard"
      Top             =   0
      Width           =   3975
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      Height          =   3015
      Left            =   2760
      ScaleHeight     =   2955
      ScaleWidth      =   4155
      TabIndex        =   6
      Top             =   360
      Width           =   4215
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C0C0C0&
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00C0C0C0&
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00C0C0C0&
      Caption         =   "^"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00C0C0C0&
      Caption         =   "v"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox Text19 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4080
      TabIndex        =   1
      Text            =   "4"
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox Text20 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Text            =   "1"
      Top             =   2040
      Width           =   495
   End
   Begin MSComDlg.CommonDialog dlgfile 
      Left            =   7200
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label15 
      Caption         =   "COUPLING"
      Height          =   255
      Left            =   2400
      TabIndex        =   64
      Top             =   3840
      Width           =   855
   End
   Begin VB.Line Line6 
      X1              =   576
      X2              =   600
      Y1              =   344
      Y2              =   344
   End
   Begin VB.Line Line5 
      X1              =   344
      X2              =   232
      Y1              =   320
      Y2              =   320
   End
   Begin VB.Line Line4 
      X1              =   232
      X2              =   232
      Y1              =   344
      Y2              =   320
   End
   Begin VB.Line Line3 
      X1              =   344
      X2              =   344
      Y1              =   344
      Y2              =   320
   End
   Begin VB.Line Line2 
      X1              =   344
      X2              =   576
      Y1              =   344
      Y2              =   344
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   232
      Y1              =   344
      Y2              =   344
   End
   Begin VB.Label Label14 
      Caption         =   "POS"
      Height          =   255
      Left            =   5400
      TabIndex        =   58
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label13 
      Caption         =   "H"
      Height          =   255
      Left            =   5280
      TabIndex        =   52
      Top             =   4800
      Width           =   135
   End
   Begin VB.Label Label12 
      Caption         =   " V POS"
      Height          =   255
      Left            =   8040
      TabIndex        =   51
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label Label11 
      Caption         =   "  R"
      Height          =   255
      Left            =   8400
      TabIndex        =   50
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label10 
      Caption         =   "  L"
      Height          =   255
      Left            =   7920
      TabIndex        =   49
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "10ms/div"
      Height          =   255
      Left            =   5760
      TabIndex        =   45
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "R"
      Height          =   255
      Left            =   6000
      TabIndex        =   42
      Top             =   4800
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "L"
      Height          =   255
      Left            =   6000
      TabIndex        =   41
      Top             =   4320
      Width           =   135
   End
   Begin VB.Label Label6 
      Caption         =   "dBr"
      Height          =   255
      Left            =   7200
      TabIndex        =   39
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "dBr"
      Height          =   255
      Left            =   7200
      TabIndex        =   38
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "TRIG"
      Height          =   255
      Left            =   7200
      TabIndex        =   31
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "DELAY"
      Height          =   255
      Left            =   7200
      TabIndex        =   27
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "DAZYWEB LABS  VB-4000"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   2055
   End
   Begin VB.Menu menuFile 
      Caption         =   "File"
      Begin VB.Menu menuLoad 
         Caption         =   "Load"
      End
      Begin VB.Menu menuSave 
         Caption         =   "Save"
      End
   End
   Begin VB.Menu menuPrint 
      Caption         =   "Print"
   End
   Begin VB.Menu menuOptions 
      Caption         =   "Options"
      Begin VB.Menu menuReadout 
         Caption         =   "Amplitude Readout"
         Begin VB.Menu menuDB 
            Caption         =   "dB"
            Checked         =   -1  'True
         End
         Begin VB.Menu menuVolts 
            Caption         =   "Volts"
         End
      End
      Begin VB.Menu menuTimebase 
         Caption         =   "Timebase Readout"
         Begin VB.Menu menuTime 
            Caption         =   "Time"
            Checked         =   -1  'True
         End
         Begin VB.Menu menuFreq 
            Caption         =   "Frequency"
         End
      End
      Begin VB.Menu menuSmooth 
         Caption         =   "Smooth coefficient"
         Begin VB.Menu menuSmooth2 
            Caption         =   "2"
            Checked         =   -1  'True
         End
         Begin VB.Menu menuSmooth3 
            Caption         =   "3"
         End
         Begin VB.Menu menuSmooth4 
            Caption         =   "4"
         End
      End
      Begin VB.Menu menuGrid 
         Caption         =   "Grid"
      End
      Begin VB.Menu menuColor 
         Caption         =   "Color"
      End
      Begin VB.Menu mneuTracecolor 
         Caption         =   "Trace color"
      End
   End
   Begin VB.Menu menuHelp 
      Caption         =   "Help"
   End
   Begin VB.Menu menuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Scopeform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  DazyWeb Laboratories VB-4000  Oscilloscope Module  v1.2   build June 24, 2001
'      copyright 2001
'
Option Explicit

Dim error As String
Dim playflag As Integer
Dim trigdelay As Integer
Dim trigpol As Integer
Dim starttrig As Integer
Dim endtrig As Integer
Dim counterL As Long
Dim counterR As Long
Dim couplingflag As Integer
Dim acoffsetL As Double
Dim acoffsetR As Double
Dim timeflag As Integer
Dim cursor1 As Integer
Dim cursor2 As Integer
Dim cursorflag As Integer
Dim doscalarflag As Integer
Dim hpos As Integer
Dim vposL As Integer
Dim vposR As Integer
Dim msperdiv
Dim smoothcoeff As Integer
Dim smoothflag As Integer
Dim voltsdbflag As Integer
Dim amplitudeL As Double
Dim amplitudeR As Double
Dim refsetflag
Dim calvalueL As Single
Dim calvalueR As Single
Dim calvalueflag As Integer
Dim monoflag As Integer
Dim trigwindow As Single
Dim dif As Integer
Dim newdif As Double
Dim avedif As Double
Dim avecount As Double
Dim domfreq As Double
Dim domfreqst As String
Dim n1L As Integer
Dim n2L As Integer
Dim n1R As Integer
Dim n2R As Integer
Dim zerocross As Integer
Dim peakdataL As Long
Dim peakdataL2 As Long
Dim peakdataR As Long
Dim peakdataR2 As Long
Dim hrangeflag As Integer
Dim vrangeflag As Integer
Dim trigoffset As Long
Dim trigchannel As Integer
Dim trigflag As Integer
Dim trigsetflag As Integer
Dim linecolor2
Dim dualflag As Integer
Dim Wave As WaveHdr
Dim InData(16384) As Integer
Dim InData2(16384) As Integer
Dim fname As String
Dim fnum1 As Integer
Dim cntr1 As Integer
Dim aveflag As Integer
Dim legend As String
Dim displayflag As String
Dim gridflag As Integer
Dim colorflag As Integer
Dim linecolor
Dim v As Single
Dim vr As String
Dim scopedata(19384) As Single
Dim scopedata2(19384) As Single
Dim scopedata3(19384) As Single
Dim data_L(10000) As Long
Dim data_R(10000) As Long
Dim scalar As Single
Dim Hscalar As Single
Dim n As Integer
Dim lvlscalarL As Single
Dim lvlscalarR As Single
Dim logvol1(8160) As Variant
Dim sr As Integer
Dim loadedsamples As Long
Dim X As Long
Dim pwid As Single
Dim phgt As Single
Dim xmid As Single
Dim ymid As Single
Dim factor1 As Long
Dim y As Long
Dim lplotold As Long
Dim lplotnew As Long
Dim texttoprint As String
Dim Y1 As Double
Dim Y2 As Double
Dim channels As Integer
Dim channelflag As String

Private DevHandle As Long 'Handle of the open audio device

Private Visualizing As Boolean
Private Divisor As Long

Private ScopeHeight As Long 'Saves time because hitting up a Long is faster
                            'than a property.
                            
Private Type WaveFormatEx
    FormatTag As Integer
    channels As Integer
    SamplesPerSec As Long
    AvgBytesPerSec As Long
    BlockAlign As Integer
    BitsPerSample As Integer
    ExtraDataSize As Integer
End Type

Private Type WaveHdr
    lpData As Long
    dwBufferLength As Long
    dwBytesRecorded As Long
    dwUser As Long
    dwFlags As Long
    dwLoops As Long
    lpNext As Long 'wavehdr_tag
    Reserved As Long
End Type

Private Type WaveInCaps
    ManufacturerID As Integer      'wMid
    ProductID As Integer       'wPid
    DriverVersion As Long       'MMVERSIONS vDriverVersion
    ProductName(1 To 32) As Byte 'szPname[MAXPNAMELEN]
    Formats As Long
    channels As Integer
    Reserved As Integer
End Type

Private Const WAVE_INVALIDFORMAT = &H0&                 '/* invalid format */
Private Const WAVE_FORMAT_1M08 = &H1&                   '/* 11.025 kHz, Mono,   8-bit
Private Const WAVE_FORMAT_1S08 = &H2&                   '/* 11.025 kHz, Stereo, 8-bit
Private Const WAVE_FORMAT_1M16 = &H4&                   '/* 11.025 kHz, Mono,   16-bit
Private Const WAVE_FORMAT_1S16 = &H8&                   '/* 11.025 kHz, Stereo, 16-bit
Private Const WAVE_FORMAT_2M08 = &H10&                  '/* 22.05  kHz, Mono,   8-bit
Private Const WAVE_FORMAT_2S08 = &H20&                  '/* 22.05  kHz, Stereo, 8-bit
Private Const WAVE_FORMAT_2M16 = &H40&                  '/* 22.05  kHz, Mono,   16-bit
Private Const WAVE_FORMAT_2S16 = &H80&                  '/* 22.05  kHz, Stereo, 16-bit
Private Const WAVE_FORMAT_4M08 = &H100&                 '/* 44.1   kHz, Mono,   8-bit
Private Const WAVE_FORMAT_4S08 = &H200&                 '/* 44.1   kHz, Stereo, 8-bit
Private Const WAVE_FORMAT_4M16 = &H400&                 '/* 44.1   kHz, Mono,   16-bit
Private Const WAVE_FORMAT_4S16 = &H800&                 '/* 44.1   kHz, Stereo, 16-bit

Private Const WAVE_FORMAT_PCM = 1

Private Const WHDR_DONE = &H1&              '/* done bit */
Private Const WHDR_PREPARED = &H2&          '/* set if this header has been prepared */
Private Const WHDR_BEGINLOOP = &H4&         '/* loop start block */
Private Const WHDR_ENDLOOP = &H8&           '/* loop end block */
Private Const WHDR_INQUEUE = &H10&          '/* reserved for driver */

Private Const WIM_OPEN = &H3BE
Private Const WIM_CLOSE = &H3BF
Private Const WIM_DATA = &H3C0

Private Declare Function waveInAddBuffer Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long
Private Declare Function waveInPrepareHeader Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long
Private Declare Function waveInUnprepareHeader Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long

Private Declare Function waveInGetNumDevs Lib "winmm" () As Long
Private Declare Function waveInGetDevCaps Lib "winmm" Alias "waveInGetDevCapsA" (ByVal uDeviceID As Long, ByVal WaveInCapsPointer As Long, ByVal WaveInCapsStructSize As Long) As Long

Private Declare Function waveInOpen Lib "winmm" (WaveDeviceInputHandle As Long, ByVal WhichDevice As Long, ByVal WaveFormatExPointer As Long, ByVal CallBack As Long, ByVal CallBackInstance As Long, ByVal Flags As Long) As Long
Private Declare Function waveInClose Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long

Private Declare Function waveInStart Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Private Declare Function waveInReset Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Private Declare Function waveInStop Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long




Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

'''sndPlaySound Constants
Const SND_ALIAS = &H10000
Const SND_ASYNC = &H1
Const SND_FILENAME = &H20000
Const SND_LOOP = &H8
Const SND_MEMORY = &H4
Const SND_NODEFAULT = &H2
Const SND_NOSTOP = &H10
Const SND_SYNC = &H0
Const SND_PURGE = &H40

Dim SoundFile As String

Private Sub Command1_Click()  'ZERO VERT POSITION

vposL = 0
vposR = 0
VScroll3.Value = 0
VScroll4.Value = 0
doscope

End Sub

Private Sub Command19_Click()  'AC COUPLING

couplingflag = 0
Text17.BackColor = vbRed
Text18.BackColor = vbBlue

End Sub

Private Sub Command2_Click()  'CURSORS ON/OFF

If cursorflag = 0 Then
cursorflag = 1
Else
cursorflag = 0
End If

doscope

End Sub

Private Sub Command20_Click()  'DC COUPLING

couplingflag = 1
Text17.BackColor = vbBlue
Text18.BackColor = vbRed


End Sub

Private Sub Command21_Click() 'trig +/-

If trigpol = 0 Then
trigpol = 1
Else
trigpol = 0
End If

End Sub

Private Sub HScroll2_Change() 'first cursor

cursor1 = HScroll2.Value
doscope

End Sub

Private Sub HScroll3_Change() 'second cursor

cursor2 = HScroll3.Value
doscope

End Sub

Private Sub menuFreq_Click()  'frequency with cursors

timeflag = 0
menuTime.Checked = False
menuFreq.Checked = True
doscope

End Sub

Private Sub menuTime_Click()  'time with cursors

timeflag = 1
menuTime.Checked = True
menuFreq.Checked = False
doscope

End Sub

Private Sub VScroll3_Change()  'L vert position

vposL = VScroll3.Value
doscope

End Sub

Private Sub VScroll4_Change()  ' R vert position

vposR = VScroll4.Value
doscope

End Sub

Private Sub HScroll1_Change()  'Horiz position

hpos = HScroll1.Value
doscope

End Sub

Private Sub Command10_Click()   'SMOOTH


If smoothflag = 0 Then
smoothflag = 1
Text14.BackColor = vbRed
Else
smoothflag = 0
Text14.BackColor = vbBlue
End If

doscope

End Sub

Private Sub Command11_Click()  'RIGHT LVLSCALAR UP

lvlscalarR = lvlscalarR * 2
If lvlscalarR > 256 Then
lvlscalarR = 256
End If
doscope
Text11.Text = CStr(lvlscalarR)


End Sub

Private Sub Command12_Click()  'RIGHT LVLSCALAR DOWN

lvlscalarR = lvlscalarR / 2
If lvlscalarR < 0.015625 Then
lvlscalarR = 0.015625
End If
doscope
Text11.Text = CStr(lvlscalarR)

End Sub

Private Sub Command17_Click()  'REF SET

If refsetflag = 0 Then
refsetflag = 1
Text10.BackColor = vbRed
End If
calvalueflag = 1

End Sub

Private Sub Command18_Click()  'REF CLR

calvalueL = 1
calvalueR = 1
refsetflag = 0
Text10.BackColor = vbBlue


End Sub

Private Sub Command9_Click()  'autorange vert

If vrangeflag = 0 Then
vrangeflag = 1
Text13.BackColor = vbRed
Else
vrangeflag = 0
Text13.BackColor = vbBlue
End If

End Sub

Private Sub Command6_Click()  'dual mode

If dualflag = 0 Then
dualflag = 1
Text5.BackColor = vbRed
Else
dualflag = 0
Text5.BackColor = vbBlue
End If

End Sub

Private Sub Command7_Click()  'TRIG L

If trigflag = 0 Then
trigflag = 1
Text6.BackColor = vbRed
Else
trigflag = 0
Text6.BackColor = vbBlue
End If

End Sub

Private Sub Command8_Click() 'L/R

If trigchannel = 0 Then
trigchannel = 1
Text7.BackColor = vbRed
Text8.BackColor = vbBlue
Else
trigchannel = 0
Text7.BackColor = vbBlue
Text8.BackColor = vbRed
End If

End Sub



Private Sub menuColor_Click()  'DISPLAY COLOR

If colorflag = 0 Then
colorflag = 1
Else
colorflag = 0
End If

doscope

End Sub

Private Sub menuDB_Click()  'db readout

menuDB.Checked = True
menuVolts.Checked = False
voltsdbflag = 0
Label3.Caption = "dBr"
Label6.Caption = "dBr"
doscope

End Sub

Private Sub menuGrid_Click()  'GRID TOGGLE

If gridflag = 0 Then
gridflag = 1
Else
gridflag = 0
End If

doscope


End Sub

Private Sub menuSmooth2_Click()  ' smooth = 2

smoothcoeff = 2
menuSmooth2.Checked = True
menuSmooth3.Checked = False
menuSmooth4.Checked = False
Text16.Text = CStr(smoothcoeff)
doscope

End Sub

Private Sub menuSmooth3_Click()  'smooth = 3

smoothcoeff = 3
menuSmooth2.Checked = False
menuSmooth3.Checked = True
menuSmooth4.Checked = False
Text16.Text = CStr(smoothcoeff)
doscope

End Sub

Private Sub menuSmooth4_Click()  'smooth = 4

smoothcoeff = 4
menuSmooth2.Checked = False
menuSmooth3.Checked = False
menuSmooth4.Checked = True
Text16.Text = CStr(smoothcoeff)
doscope

End Sub

Private Sub menuVolts_Click()  'volts readout

voltsdbflag = 1
menuDB.Checked = False
menuVolts.Checked = True
Label3.Caption = "mV"
Label6.Caption = "mV"
doscope

End Sub

Private Sub mneuTracecolor_Click()  'TRACECOLOR

If monoflag = 0 Then
monoflag = 1
Else
monoflag = 0
End If
doscope

End Sub

Private Sub VScroll1_Change()  'trig delay slider

trigdelay = VScroll1.Value

End Sub



Private Sub menuHelp_Click()  'Help

helpform.Show

End Sub

Private Sub StartButton_Click()            'START BUTTON
    Static WaveFormat As WaveFormatEx
    With WaveFormat
        .FormatTag = WAVE_FORMAT_PCM
        .channels = 2         ' 1 = right only, 2 = stereo
        .SamplesPerSec = 44100   '11025  22050  44100
        .BitsPerSample = 16
        .BlockAlign = (.channels * .BitsPerSample) \ 8
        .AvgBytesPerSec = .BlockAlign * .SamplesPerSec
        .ExtraDataSize = 0
    End With
    
    Debug.Print "waveInOpen:"; waveInOpen(DevHandle, DevicesBox.ListIndex, VarPtr(WaveFormat), 0, 0, 0)
    If DevHandle = 0 Then
        Call MsgBox("Wave input device didn't open!", vbExclamation, "Ack!")
        Exit Sub
    End If
    Debug.Print " "; DevHandle
    Call waveInStart(DevHandle)
    
    StopButton.Enabled = True
    StartButton.Enabled = False
    'DevicesBox.Enabled = False
    
    playflag = 1
    doscope2
    doscope
    
End Sub


Private Sub StopButton_Click()          'STOP BUTTON
    
    Call DoStop
    playflag = 0
    
End Sub


Private Sub DoStop()
    Call waveInReset(DevHandle)
    Call waveInClose(DevHandle)
    DevHandle = 0
    StopButton.Enabled = False
    StartButton.Enabled = True
    DevicesBox.Enabled = True
    
End Sub



Private Sub doscope2()



With ScopeBuff 'Save some time referencing it...
    
        Do
            Wave.lpData = VarPtr(InData(0))
            Wave.dwBufferLength = 16384
            Wave.dwFlags = 0
            Call waveInPrepareHeader(DevHandle, VarPtr(Wave), Len(Wave))
            Call waveInAddBuffer(DevHandle, VarPtr(Wave), Len(Wave))
            
            Do
                'Just wait for the blocks to be done or the device to close
            Loop Until ((Wave.dwFlags And WHDR_DONE) = WHDR_DONE) Or DevHandle = 0
            If DevHandle = 0 Then Exit Do 'Cut out if the device is closed
            
            Call waveInUnprepareHeader(DevHandle, VarPtr(Wave), Len(Wave))
              
           doscope
              
           DoEvents
            
            Loop While DevHandle <> 0
    
    End With
    
           
End Sub

              
              
              
Public Function doscope()


If trigflag = 1 Then
trigsetflag = 1
End If

'parse for stereo data
cntr1 = 0

If smoothflag = 0 Then
For n = 0 To 16380 / 2
data_R(cntr1) = (InData(n * 2) + 0)
data_L(cntr1) = (InData((n * 2) + 1))
cntr1 = cntr1 + 1
Next n
End If

If smoothflag = 1 And smoothcoeff = 2 Then
For n = 0 To 16380 / 2
data_R(cntr1) = ((InData((n * 2) + 0)) / 2) + ((InData((n * 2) + 2)) / 2)
data_L(cntr1) = ((InData((n * 2) + 1)) / 2) + ((InData((n * 2) + 3)) / 2)
cntr1 = cntr1 + 1
Next n
End If

If smoothflag = 1 And smoothcoeff = 3 Then
For n = 0 To 16378 / 2
data_R(cntr1) = ((InData((n * 2) + 0)) / 3) + ((InData((n * 2) + 2)) / 3) + ((InData((n * 2) + 4)) / 3)
data_L(cntr1) = ((InData((n * 2) + 1)) / 3) + ((InData((n * 2) + 3)) / 3) + ((InData((n * 2) + 5)) / 3)
cntr1 = cntr1 + 1
Next n
End If

If smoothflag = 1 And smoothcoeff = 4 Then
For n = 0 To 16376 / 2
data_R(cntr1) = ((InData((n * 2) + 0)) / 4) + ((InData((n * 2) + 2)) / 4) + ((InData((n * 2) + 4)) / 4) + ((InData((n * 2) + 6)) / 4)
data_L(cntr1) = ((InData((n * 2) + 1)) / 4) + ((InData((n * 2) + 3)) / 4) + ((InData((n * 2) + 5)) / 4) + ((InData((n * 2) + 7)) / 4)
cntr1 = cntr1 + 1
Next n
End If

'end of stereo data parse



'left
For n = 0 To (16383 / 2)
scopedata2(n) = data_L(n)

If peakdataL < scopedata2(n) Then
peakdataL = scopedata2(n)
If n < 200 Then
n1L = n
End If
End If
If peakdataL2 > scopedata2(n) Then
peakdataL2 = scopedata2(n)
n2L = n
End If
Next n
For n = (16383 / 2) To 16383
scopedata2(n) = scopedata2(n / 2)
Next n



'right
For n = 0 To (16383 / 2)
scopedata3(n) = data_R(n)

If peakdataR < scopedata3(n) Then
peakdataR = scopedata3(n)
n1R = n
End If
If peakdataR2 > scopedata3(n) Then
peakdataR2 = scopedata3(n)
n2R = n
End If
Next n
For n = (16383 / 2) To 16383
scopedata3(n) = scopedata3(n / 2)
Next n

If channels = 2 Then  'force sum left and right in stereo
For n = 0 To (16383 / 2)
scopedata2(n) = data_L(n) + data_R(n)
If peakdataL < scopedata2(n) Then
peakdataL = scopedata2(n)
End If
Next n
For n = (16383 / 2) To 16383
scopedata2(n) = scopedata2(n / 2)
Next n

End If


If vrangeflag = 1 And ((peakdataL - peakdataL2) * lvlscalarL) > 3000 And lvlscalarL > 0.01 Then
lvlscalarL = lvlscalarL / 2
Text20.Text = CStr(lvlscalarL)
End If

If vrangeflag = 1 And ((peakdataL - peakdataL2) * lvlscalarL) < 2500 And lvlscalarL < 257 Then
lvlscalarL = lvlscalarL * 2
Text20.Text = CStr(lvlscalarL)
End If

If vrangeflag = 1 And ((peakdataR - peakdataR2) * lvlscalarR) > 3000 And lvlscalarR > 0.01 Then
lvlscalarR = lvlscalarR / 2
Text11.Text = CStr(lvlscalarR)
End If

If vrangeflag = 1 And ((peakdataR - peakdataR2) * lvlscalarR) < 2500 And lvlscalarR < 257 Then
lvlscalarR = lvlscalarR * 2
Text11.Text = CStr(lvlscalarR)
End If



Update


 Picture1.Cls


If colorflag = 1 Then
Picture1.BackColor = vbBlack
linecolor = vbGreen
Else
Picture1.BackColor = vbWhite
linecolor = vbBlack
End If


If gridflag = 1 Then

Picture1.ForeColor = &H555555
  
For n = 1 To 9
Picture1.Line (0, (n * 295.5))-(4155, (n * 295.5))
Picture1.Line ((n * 415.5), 0)-((n * 415.5), 2955)
Next n
  
Picture1.ForeColor = vbBlack
End If
  
  If couplingflag = 0 Then  'ac coupling
  counterL = 1
  counterR = 1
  For n = 1 To 4095
  counterL = counterL + data_L(n)
  counterR = counterR + data_R(n)
  Next n
  acoffsetL = 2955 * (counterL / 4095) / (16384)
  acoffsetL = lvlscalarL * acoffsetL / (-0.54)
  
  acoffsetR = 2955 * (counterR / 4095) / (16384)
  acoffsetR = lvlscalarR * acoffsetR / (-0.54)
  Else
  acoffsetL = 0
  acoffsetR = 0
  End If
  
  
  If trigflag = 0 Then
  starttrig = 1
  End If
  
  If trigflag = 1 And trigchannel = 0 And trigpol = 0 Then
  starttrig = n1L
  End If
  
  If trigflag = 1 And trigchannel = 0 And trigpol = 1 Then
  starttrig = n2L
  End If
  
  If trigflag = 1 And trigchannel = 1 And trigpol = 0 Then
  starttrig = n1R
  End If
  
  If trigflag = 1 And trigchannel = 1 And trigpol = 1 Then
  starttrig = n2R
  End If
  
  
  
  If trigflag = 1 Then
  starttrig = starttrig + trigdelay
  End If
  
  If starttrig > 8190 Then
  starttrig = 8190
  End If
  
  If starttrig < 0 Then
  starttrig = 0
  End If
  
  
  
  endtrig = (starttrig + trigdelay + Int(8190 / scalar))

  If endtrig > 8190 Then
  endtrig = 8190
  End If
  
  
  
    For n = starttrig To endtrig   'plot dataset here
    'For n = 1 To Int(8190 / scalar)
    'For n = 1 To Int(8190 / 2)
    If monoflag = 1 And colorflag = 1 Then
    linecolor2 = vbGreen
    Else
    linecolor2 = vbBlue
    End If
    
    If trigflag = 0 Then
    If dualflag = 0 Then
    Picture1.Line ((Int((n * Hscalar))) + hpos, (Int(1500 - (lvlscalarL * (scopedata2(n) / 3)))) + vposL - acoffsetL)-((Int(((n + 1) * Hscalar))) + hpos, (Int(1500 - (lvlscalarL * (scopedata2(n + 1) / 3)))) + vposL - acoffsetL), linecolor
    Else
    Picture1.Line ((Int((n * Hscalar))) + hpos, (Int(1500 - (lvlscalarL * (data_L(n) / 3)))) + vposL - acoffsetL)-((Int(((n + 1) * Hscalar))) + hpos, (Int(1500 - (lvlscalarL * (data_L(n + 1) / 3)))) + vposL - acoffsetL), linecolor
    Picture1.Line ((Int((n * Hscalar))) + hpos, (Int(1500 - (lvlscalarR * (data_R(n) / 3)))) + vposR - acoffsetR)-((Int(((n + 1) * Hscalar))) + hpos, (Int(1500 - (lvlscalarR * (data_R(n + 1) / 3)))) + vposR - acoffsetR), linecolor2
    End If
    End If
    
    
    If trigchannel = 0 Then  'left channel trigger
    '  If (data_L(n) > 0) And (data_L(n) > (trigslideval - trigwindow)) And (data_L(n) < (trigslideval + trigwindow)) And (data_L(n + 1) > data_L(n)) And trigsetflag = 1 Then
    'If (data_L(n) > 0) And (data_L(n) < data_L(n + 1)) And trigsetflag = 1 Then
      trigsetflag = 0
      trigoffset = starttrig
      End If
    
     '  If (data_L(n) < 0) And (data_L(n) < (trigslideval + trigwindow)) And (data_L(n) > (trigslideval - trigwindow)) And (data_L(n + 1) < data_L(n)) And trigsetflag = 1 Then
    'If (data_L(n) < 0) And (data_L(n) > data_L(n + 1)) And trigsetflag = 1 Then
    '   trigsetflag = 0
    '   trigoffset = n2L
    '   End If
    'End If   'end left channel trigger
    
    
    
    
   If trigchannel = 1 Then   'right channel trigger
     ' If (data_R(n) > 0) And (data_R(n) > (trigslideval - trigwindow)) And (data_R(n) < (trigslideval + trigwindow)) And (data_R(n + 1) > data_R(n)) And trigsetflag = 1 Then
      trigsetflag = 0
      trigoffset = starttrig
      End If
    
      'If (data_R(n) < 0) And (data_R(n) < (trigslideval + trigwindow)) And (data_R(n) > (trigslideval - trigwindow)) And (data_R(n + 1) < data_R(n)) And trigsetflag = 1 Then
      'trigsetflag = 0
      'trigoffset = n
      'End If
    'End If   'end right channel trigger
    
    
    
    If trigflag = 1 And Int(((n - trigoffset) * Hscalar) + hpos) > 0 And Int(((n - trigoffset) * Hscalar) + hpos) < 8190 Then
      If trigsetflag = 0 Then
        If dualflag = 0 Then
        Picture1.Line ((Int(((n - trigoffset) * Hscalar))) + hpos, (Int(1500 - (lvlscalarL * (scopedata2(n) / 3)))) + vposL - acoffsetL)-((Int(((n + 1 - trigoffset) * Hscalar))) + hpos, (Int(1500 - (lvlscalarL * (scopedata2(n + 1) / 3)))) + vposL - acoffsetL), linecolor
        Else
        Picture1.Line ((Int(((n - trigoffset) * Hscalar))) + hpos, (Int(1500 - (lvlscalarL * (data_L(n) / 3)))) + vposL - acoffsetL)-((Int(((n + 1 - trigoffset) * Hscalar))) + hpos, (Int(1500 - (lvlscalarL * (data_L(n + 1) / 3)))) + vposL - acoffsetL), linecolor
        Picture1.Line ((Int(((n - trigoffset) * Hscalar))) + hpos, (Int(1500 - (lvlscalarR * (data_R(n) / 3)))) + vposR - acoffsetR)-((Int(((n + 1 - trigoffset) * Hscalar))) + hpos, (Int(1500 - (lvlscalarR * (data_R(n + 1) / 3)))) + vposR - acoffsetR), linecolor2
        End If
      End If
    End If
    
    

    Next n
    
   
    
    If calvalueL = 0 Then
    calvalueL = 1
    End If
    
    If calvalueR = 0 Then
    calvalueR = 1
    End If
    
    If calvalueflag = 1 Then
    calvalueL = Abs(peakdataL - peakdataL2)
    calvalueR = Abs(peakdataR - peakdataR2)
    calvalueflag = 0
    End If

    
    If (Abs(peakdataL - peakdataL2) / calvalueL) <> 0 Then
    amplitudeL = (20 * ((Log(Abs(peakdataL - peakdataL2) / calvalueL)) / Log(10)))
    If amplitudeL < 0.1 And amplitudeL > -0.1 Then
    amplitudeL = 0
    End If
    If voltsdbflag = 0 Then
    Text9.Text = CStr(Left$(amplitudeL, 6))
    Else
    amplitudeL = 1000 * (Abs(peakdataL - peakdataL2) / calvalueL)
    Text9.Text = CStr(Left$(amplitudeL, 6))
    End If
    End If
    
    If (Abs(peakdataR - peakdataR2) / calvalueR) <> 0 Then
    amplitudeR = (20 * ((Log(Abs(peakdataR - peakdataR2) / calvalueR)) / Log(10)))
    If amplitudeR < 0.1 And amplitudeR > -0.1 Then
    amplitudeR = 0
    End If
    If voltsdbflag = 0 Then
    Text12.Text = CStr(Left$(amplitudeR, 6))
    Else
    amplitudeR = 1000 * (Abs(peakdataR - peakdataR2) / calvalueR)
    Text12.Text = CStr(Left$(amplitudeR, 6))
    End If
    End If
    
    
    n1L = 0
    n2L = 0
    n1R = 0
    n2R = 0

    If cursorflag = 1 Then
    Picture1.Line (cursor1, 0)-(cursor1, 2955), vbRed
    Picture1.Line (cursor2, 0)-(cursor2, 2955), vbRed

If timeflag = 0 Then 'freq readout
If (cursor1 - cursor2) <> 0 Then
Text15.Text = Left$(CStr(1000 / (Abs(cursor1 - cursor2) * msperdiv * 10 / 4155)), 6) + " Hz"
End If
Else                 ' time readout
Text15.Text = Left$(CStr(Abs(cursor1 - cursor2) * msperdiv * 10 / 4155), 6) + " ms"
End If
Else
Text15.Text = " "
End If




peakdataL = 0
peakdataL2 = 0
peakdataR = 0
peakdataR2 = 0

trigsetflag = 0
trigoffset = 0
zerocross = 0

End Function


Private Sub Command3_Click()  'Display left channel

displayflag = "L"
Text2.BackColor = vbRed
Text4.BackColor = vbBlue
Text3.BackColor = vbBlue
channels = 1
doscope
End Sub

Private Sub Command4_Click()  'Display mono channel

displayflag = "M"
Text2.BackColor = vbBlue
Text4.BackColor = vbBlue
Text3.BackColor = vbRed
channels = 2
doscope
End Sub

Private Sub Command5_Click()  'Display right channel

displayflag = "R"
Text2.BackColor = vbBlue
Text4.BackColor = vbRed
Text3.BackColor = vbBlue
channels = 0
doscope
End Sub

Private Sub menuExit_Click() 'EXIT
error = "Press STOP before exiting."
If playflag = 1 Then
Call MsgBox(error)
Else

fname = "c:/vb3scope2init"
fnum1 = FreeFile
    
On Error Resume Next
Open fname For Output As #fnum1
On Error Resume Next

Write #fnum1, scalar
Write #fnum1, lvlscalarL
Write #fnum1, gridflag
Write #fnum1, colorflag
Write #fnum1, displayflag
Write #fnum1, dualflag
Write #fnum1, trigflag
Write #fnum1, trigchannel
Write #fnum1, trigdelay
Write #fnum1, trigwindow
Write #fnum1, lvlscalarR
Write #fnum1, refsetflag
Write #fnum1, monoflag
Write #fnum1, calvalueL
Write #fnum1, calvalueR
Write #fnum1, vrangeflag
Write #fnum1, voltsdbflag
Write #fnum1, smoothflag
Write #fnum1, smoothcoeff
Write #fnum1, hpos
Write #fnum1, vposL
Write #fnum1, vposR
Write #fnum1, cursorflag
Write #fnum1, cursor1
Write #fnum1, cursor2
Write #fnum1, timeflag
Write #fnum1, Hscalar
Write #fnum1, couplingflag
Write #fnum1, trigpol
Close fnum1

Close

Call waveInReset(DevHandle)
    Call waveInClose(DevHandle)
    DevHandle = 0
Unload Scopeform

End If

End Sub

Private Sub Form_Load()  'scope


cursor1 = 1000
cursor2 = 3000
smoothcoeff = 2
calvalueL = 1
calvalueR = 1
avecount = 1
scalar = 5
Hscalar = 5
lvlscalarL = 1
lvlscalarR = 1
linecolor2 = vbBlue

fname = "c:/vb3scope2init"
fnum1 = FreeFile
    
On Error Resume Next
Open fname For Input As #fnum1
On Error Resume Next

Input #fnum1, scalar
Input #fnum1, lvlscalarL
Input #fnum1, gridflag
Input #fnum1, colorflag
Input #fnum1, displayflag
Input #fnum1, dualflag
Input #fnum1, trigflag
Input #fnum1, trigchannel
Input #fnum1, trigdelay
Input #fnum1, trigwindow
Input #fnum1, lvlscalarR
Input #fnum1, refsetflag
Input #fnum1, monoflag
Input #fnum1, calvalueL
Input #fnum1, calvalueR
Input #fnum1, vrangeflag
Input #fnum1, voltsdbflag
Input #fnum1, smoothflag
Input #fnum1, smoothcoeff
Input #fnum1, hpos
Input #fnum1, vposL
Input #fnum1, vposR
Input #fnum1, cursorflag
Input #fnum1, cursor1
Input #fnum1, cursor2
Input #fnum1, timeflag
Input #fnum1, Hscalar
Input #fnum1, couplingflag
Input #fnum1, trigpol
Close fnum1

Text19.Text = CStr(scalar)
Text20.Text = CStr(lvlscalarL)
Text11.Text = CStr(lvlscalarR)
Text16.Text = CStr(smoothcoeff)

VScroll1.Value = trigdelay
VScroll3.Value = vposL
VScroll4.Value = vposR
HScroll1.Value = hpos
HScroll2.Value = cursor1
HScroll3.Value = cursor2

If couplingflag = 1 Then
Text17.BackColor = vbBlue
Text18.BackColor = vbRed
Else
Text17.BackColor = vbRed
Text18.BackColor = vbBlue
End If


If timeflag = 1 Then
menuTime.Checked = True
menuFreq.Checked = False
Else
menuTime.Checked = False
menuFreq.Checked = True
End If


If smoothcoeff = 2 Then
menuSmooth2.Checked = True
End If

If smoothcoeff = 3 Then
menuSmooth3.Checked = True
menuSmooth2.Checked = False
End If

If smoothcoeff = 4 Then
menuSmooth4.Checked = True
menuSmooth2.Checked = False
End If

If smoothflag = 1 Then
Text14.BackColor = vbRed
End If

If voltsdbflag = 0 Then
menuDB.Checked = True
menuVolts.Checked = False
Label3.Caption = "dBr"
Label6.Caption = "dBr"
Else
menuDB.Checked = False
menuVolts.Checked = True
Label3.Caption = "mV"
Label6.Caption = "mV"
End If

If vrangeflag = 1 Then
Text13.BackColor = vbRed
End If

If refsetflag = 1 Then
Text10.BackColor = vbRed
End If

If trigflag = 1 Then
Text6.BackColor = vbRed
Else
Text6.BackColor = vbBlue
End If

If trigchannel = 1 Then
Text7.BackColor = vbRed
Text8.BackColor = vbBlue
Else
Text7.BackColor = vbBlue
Text8.BackColor = vbRed
End If


If dualflag = 1 Then
Text5.BackColor = vbRed
End If


If displayflag = "L" Then
Text2.BackColor = vbRed
Text4.BackColor = vbBlue
Text3.BackColor = vbBlue
End If


If displayflag = "R" Then
Text2.BackColor = vbBlue
Text4.BackColor = vbRed
Text3.BackColor = vbBlue
End If


If displayflag = "M" Then
Text2.BackColor = vbBlue
Text4.BackColor = vbBlue
Text3.BackColor = vbRed
End If


'add scale lines here

For n = 0 To 10
Scopeform.CurrentY = 18 + (n * 20)
Scopeform.CurrentX = 155
v = (5 - n)
If v > 0 Then
vr = "+" + CStr(v)
Else
vr = " " + CStr(v)
End If
Scopeform.Print vr + "   --"
    
Scopeform.CurrentY = 235
Scopeform.CurrentX = 183 + (n * 27.4)
Scopeform.Print CStr(n)
Scopeform.Line ((187 + (n * 27.5)), 230)-((187 + (n * 27.5)), 222)
Next n


doscope


End Sub


Private Sub Command14_Click()  'scalar up

doscalarflag = 1

If scalar <> 1 And scalar <> 2.5 And scalar <> 5 And scalar <> 10 And scalar <> 25 And scalar <> 50 And scalar <> 100 And scalar <> 250 Then
scalar = 5
End If

If scalar = 1 And doscalarflag = 1 Then
scalar = 2.5
doscalarflag = 0
End If

If scalar = 2.5 And doscalarflag = 1 Then
scalar = 5
doscalarflag = 0
End If

If scalar = 5 And doscalarflag = 1 Then
scalar = 10
doscalarflag = 0
End If

If scalar = 10 And doscalarflag = 1 Then
scalar = 25
doscalarflag = 0
End If

If scalar = 25 And doscalarflag = 1 Then
scalar = 50
doscalarflag = 0
End If

If scalar = 50 And doscalarflag = 1 Then
scalar = 100
doscalarflag = 0
End If

If scalar = 100 And doscalarflag = 1 Then
scalar = 250
doscalarflag = 0
End If

Hscalar = scalar * 0.932 'fudge correction factor


doscope
Text19.Text = CStr(scalar)
End Sub

Private Sub Command13_Click() 'scalar down

doscalarflag = 1

If scalar = 2.5 And doscalarflag = 1 Then
scalar = 1
doscalarflag = 0
End If

If scalar = 5 And doscalarflag = 1 Then
scalar = 2.5
doscalarflag = 0
End If

If scalar = 10 And doscalarflag = 1 Then
scalar = 5
doscalarflag = 0
End If

If scalar = 25 And doscalarflag = 1 Then
scalar = 10
doscalarflag = 0
End If

If scalar = 50 And doscalarflag = 1 Then
scalar = 25
doscalarflag = 0
End If

If scalar = 100 And doscalarflag = 1 Then
scalar = 50
doscalarflag = 0
End If

If scalar = 250 And doscalarflag = 1 Then
scalar = 100
doscalarflag = 0
End If

Hscalar = scalar * 0.932 'fudge correction factor

doscope
Text19.Text = CStr(scalar)
End Sub

Private Sub Command15_Click() 'L scope level up

lvlscalarL = lvlscalarL * 2
If lvlscalarL > 256 Then
lvlscalarL = 256
End If
doscope
Text20.Text = CStr(lvlscalarL)

End Sub

Private Sub Command16_Click()  'L scope level down

lvlscalarL = lvlscalarL / 2
If lvlscalarL < 0.015625 Then
lvlscalarL = 0.015625
End If
doscope
Text20.Text = CStr(lvlscalarL)

End Sub

Private Sub menuSave_Click()  'save data

dlgfile.Flags = _
        cdlOFNFileMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames
    
    On Error Resume Next
    dlgfile.ShowSave
    If Err.Number = cdlCancel Then Exit Sub
    If Err.Number <> 0 Then
        MsgBox "Error" & Str$(Err.Number) & _
            " selecting file." & vbCrLf & _
            Err.Description
    End If
    On Error GoTo 0

    SaveFile1 dlgfile.FileName, dlgfile.FileTitle
    
    




End Sub


Sub SaveFile1(fname As String, fTitle As String)


    
fnum1 = FreeFile
    
On Error Resume Next
Open fname For Output As #fnum1
On Error Resume Next


For n = 0 To 16383
Write #fnum1, InData(n)
Next n


Close fnum1


End Sub


Private Sub menuLoad_Click()  'load data

dlgfile.Flags = _
        cdlOFNFileMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames
    
    On Error Resume Next
    dlgfile.ShowOpen
    If Err.Number = cdlCancel Then Exit Sub
    If Err.Number <> 0 Then
        MsgBox "Error" & Str$(Err.Number) & _
            " selecting file." & vbCrLf & _
            Err.Description
    End If
    On Error GoTo 0

    LoadFile1 dlgfile.FileName, dlgfile.FileTitle


End Sub

Sub LoadFile1(fname As String, fTitle As String)


    
fnum1 = FreeFile
    
On Error Resume Next
Open fname For Input As #fnum1
On Error Resume Next


For n = 0 To 16383
Input #fnum1, InData(n)  'use Indata2(n) later for overlay
Next n
Close fnum1

Update

doscope

End Sub




Private Sub menuPrint_Click()



If channels = 1 Or channels = 0 Then  'parse for stereo data
cntr1 = 0

For n = 0 To 16381 / 2
data_R(cntr1) = InData(n * 2)
data_L(cntr1) = InData((n * 2) + 1)
cntr1 = cntr1 + 1
Next n

End If 'end of stereo data parse



If channels = 1 Then 'left
For n = 0 To (16383 / 2)
scopedata2(n) = data_L(n)
Next n
For n = (16383 / 2) To 16383
scopedata2(n) = scopedata2(n / 2)
Next n
End If


If channels = 0 Then 'right
For n = 0 To (16383 / 2)
scopedata2(n) = data_R(n)
Next n
For n = (16383 / 2) To 16383
scopedata2(n) = scopedata2(n / 2)
Next n
End If


If channels = 2 Then 'force sum left and right in stereo
For n = 0 To (16383 / 2)
scopedata2(n) = data_L(n) + data_R(n)
Next n
For n = (16383 / 2) To 16383
scopedata2(n) = scopedata2(n / 2)
Next n

End If



doplot


End Sub



Private Function doplot()  'for printer



    ' Get the printer's dimensions in twips.
    pwid = Printer.ScaleX(Printer.ScaleWidth, Printer.ScaleMode, vbTwips)
    phgt = Printer.ScaleY(Printer.ScaleHeight, Printer.ScaleMode, vbTwips)
    
    ' Convert the printer's dimensions into the
    ' object's coordinates.
    pwid = Picture1.ScaleX(pwid, vbTwips, Picture1.ScaleMode)
    phgt = Picture1.ScaleY(phgt, vbTwips, Picture1.ScaleMode)
    
    ' Compute the center of the object.
    xmid = Picture1.ScaleLeft + Picture1.ScaleWidth / 2
    ymid = Picture1.ScaleTop + Picture1.ScaleHeight / 2
    
    ' Pass the coordinates of the upper left and
    ' lower right corners into the Scale method.
    Printer.Scale _
        (xmid - pwid / 2, ymid - phgt / 2)- _
        (xmid + pwid / 2, ymid + phgt / 2)



  'print scopedata




If displayflag = "L" Or dualflag = 1 And displayflag <> "M" Then
    For n = 1 To Int(4095 / scalar)
    
   Y1 = (Int(1500 - (lvlscalarL * (data_L(n) / 3))))
   
   If Y1 < 0 Then
   Y1 = 0
   End If
   
   If Y1 > 3000 Then
   Y1 = 3000
   End If
   
   Y2 = (Int(1500 - (lvlscalarL * (data_L(n + 1) / 3))))
  
   If Y2 < 0 Then
   Y2 = 0
   End If
   
   If Y2 > 3000 Then
   Y2 = 3000
   End If
       
    Printer.Line ((Int((n * scalar))), Y1)-((Int(((n + 1) * scalar))), Y2), vbBlack
      
    Next n
End If


If displayflag = "R" Or dualflag = 1 And displayflag <> "M" Then
 For n = 1 To Int(4095 / scalar)
    
   Y1 = (Int(1500 - (lvlscalarR * (data_R(n) / 3))))
   
   If Y1 < 0 Then
   Y1 = 0
   End If
   
   If Y1 > 3000 Then
   Y1 = 3000
   End If
   
   Y2 = (Int(1500 - (lvlscalarR * (data_R(n + 1) / 3))))
  
   If Y2 < 0 Then
   Y2 = 0
   End If
   
   If Y2 > 3000 Then
   Y2 = 3000
   End If
       
    Printer.Line ((Int((n * scalar))), Y1)-((Int(((n + 1) * scalar))), Y2), vbBlack
      
    Next n
End If

If displayflag = "M" Then
    For n = 1 To Int(4095 / scalar)
    
   Y1 = (Int(1500 - (lvlscalarL * (scopedata2(n) / 3))))
   
   If Y1 < 0 Then
   Y1 = 0
   End If
   
   If Y1 > 3000 Then
   Y1 = 3000
   End If
   
   Y2 = (Int(1500 - (lvlscalarL * (scopedata2(n + 1) / 3))))
  
   If Y2 < 0 Then
   Y2 = 0
   End If
   
   If Y2 > 3000 Then
   Y2 = 3000
   End If
       
    Printer.Line ((Int((n * scalar))), Y1)-((Int(((n + 1) * scalar))), Y2), vbBlack
      
    Next n
End If




Printer.DrawWidth = 2
Printer.Line (0, 0)-(0, 3000)
Printer.Line (4155, 0)-(4155, 3000)
Printer.Line (0, 0)-(4155, 0)
Printer.Line (0, 3000)-(4155, 3000)
Printer.DrawWidth = 1



For y = 0 To 10   'make gain lines
If y / 5 = Int(y / 5) Then
Printer.DrawWidth = 4
Printer.Line (4155, (y * 300))-(4305, (y * 300))
Printer.Line (-150, (y * 300))-(0, (y * 300))
Else
Printer.DrawWidth = 2
Printer.Line (4155, (y * 300))-(4255, (y * 300))
Printer.Line (-100, (y * 300))-(0, (y * 300))
End If
Next y
Printer.DrawWidth = 1


For y = 0 To 10   'make timebase lines
If y / 5 = Int(y / 5) Then
Printer.DrawWidth = 4
Printer.Line ((y * 415.5), 3000)-((y * 415.5), 3150)
Else
Printer.DrawWidth = 2
Printer.Line ((y * 415.5), 3000)-((y * 415.5), 3100)
End If
Next y

Printer.DrawWidth = 1



For X = 0 To 10   'print gain numbers

Printer.CurrentY = (X * 300) - 100
Printer.CurrentX = -400
texttoprint = CStr(Abs((X * 2) - 10))
Printer.Print texttoprint
Printer.CurrentX = 4455
Printer.CurrentY = (X * 300) - 100
Printer.Print texttoprint
Next X


For X = 0 To 10   'print timebase numbers

Printer.CurrentY = (3200)
Printer.CurrentX = (X * 415.5) - 75
texttoprint = CStr(X)
Printer.Print texttoprint
Next X

Printer.CurrentY = 3500   'timebase value
Printer.CurrentX = 1525
texttoprint = (CStr(msperdiv)) + " ms/div"
Printer.Print texttoprint

Printer.CurrentY = 3700  'left level value
Printer.CurrentX = 1525
texttoprint = "Left = " + Text9.Text + Label3.Caption
Printer.Print texttoprint

Printer.CurrentY = 3900   'right level value
Printer.CurrentX = 1525
texttoprint = "Right = " + Text12.Text + Label3.Caption
Printer.Print texttoprint


Printer.CurrentY = -500
Printer.CurrentX = 700
texttoprint = Text1.Text  'print legend
Printer.Print texttoprint

Printer.CurrentY = -700
Printer.CurrentX = 1525
texttoprint = Now     'print time and date
Printer.Print texttoprint


Printer.EndDoc


End Function


Sub Update()


'        #div   ms to sec  spl size   scalar
msperdiv = (16.5 * (10 * 1000 / 16384) / scalar)
msperdiv = Left$(msperdiv, 4)

Label8.Caption = (CStr(msperdiv)) + " ms/div"


End Sub

