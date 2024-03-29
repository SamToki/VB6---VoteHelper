VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form FormMainWindow 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F5F5F5&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VoteHelper v1.03eng"
   ClientHeight    =   9450
   ClientLeft      =   45
   ClientTop       =   795
   ClientWidth     =   15030
   FillColor       =   &H000000FF&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   Icon            =   "FormMainWindow.frx":0000
   LinkTopic       =   "FormMainWindow"
   MaxButton       =   0   'False
   MouseIcon       =   "FormMainWindow.frx":23D2
   MousePointer    =   99  'Custom
   ScaleHeight     =   9450
   ScaleWidth      =   15030
   StartUpPosition =   2  '徳鳥嶄伉
   Begin VB.Timer TimerProgressbarAnimation 
      Interval        =   1
      Left            =   14700
      Top             =   9135
   End
   Begin VB.Timer TimerMaxQuanBlink 
      Interval        =   500
      Left            =   96
      Top             =   1056
   End
   Begin VB.CommandButton CmdTotalQuan 
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   876
      Left            =   13056
      MouseIcon       =   "FormMainWindow.frx":2524
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   192
      Width           =   1740
   End
   Begin VB.TextBox TextboxItemTitle6 
      Appearance      =   0  'Flat
      BackColor       =   &H00FAFAFA&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   684
      Left            =   1056
      MousePointer    =   3  'I-Beam
      TabIndex        =   14
      Text            =   "Candidate Name"
      Top             =   6720
      Width           =   4150
   End
   Begin VB.TextBox TextboxInput 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00709000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   590
      Left            =   14145
      MouseIcon       =   "FormMainWindow.frx":2676
      MousePointer    =   99  'Custom
      TabIndex        =   30
      Top             =   8544
      Width           =   660
   End
   Begin VB.TextBox TextboxVoteInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FAFAFA&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   570
      Left            =   190
      MousePointer    =   3  'I-Beam
      TabIndex        =   27
      Text            =   "Enter More Information Here"
      Top             =   7770
      Width           =   14600
   End
   Begin VB.TextBox TextboxItemTitle5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FAFAFA&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   684
      Left            =   1056
      MousePointer    =   3  'I-Beam
      TabIndex        =   13
      Text            =   "Candidate Name"
      Top             =   5664
      Width           =   4150
   End
   Begin VB.TextBox TextboxItemTitle4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FAFAFA&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   684
      Left            =   1056
      MousePointer    =   3  'I-Beam
      TabIndex        =   12
      Text            =   "Candidate Name"
      Top             =   4608
      Width           =   4150
   End
   Begin VB.TextBox TextboxItemTitle3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FAFAFA&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   684
      Left            =   1056
      MousePointer    =   3  'I-Beam
      TabIndex        =   11
      Text            =   "Candidate Name"
      Top             =   3552
      Width           =   4150
   End
   Begin VB.TextBox TextboxItemTitle2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FAFAFA&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   684
      Left            =   1056
      MousePointer    =   3  'I-Beam
      TabIndex        =   10
      Text            =   "Candidate Name"
      Top             =   2496
      Width           =   4150
   End
   Begin VB.TextBox TextboxItemTitle1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FAFAFA&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   684
      Left            =   1056
      MousePointer    =   3  'I-Beam
      TabIndex        =   9
      Text            =   "Candidate Name"
      Top             =   1440
      Width           =   4150
   End
   Begin VB.TextBox TextboxVoteTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00FAFAFA&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   876
      Left            =   190
      MousePointer    =   3  'I-Beam
      TabIndex        =   0
      Text            =   "Enter Vote Topic Here"
      Top             =   190
      Width           =   11052
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   435
      Left            =   1575
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   435
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   767
      _cy             =   767
   End
   Begin VB.Shape ShapeProgressbar 
      BackColor       =   &H00FF9000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   120
      Left            =   0
      Top             =   9345
      Width           =   14820
   End
   Begin VB.Label LabelTotalQuanTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "&Quantity:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   585
      Left            =   11430
      TabIndex        =   1
      Top             =   370
      Width           =   1455
   End
   Begin VB.Label LabelItemPerc6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   675
      Left            =   11925
      TabIndex        =   26
      Top             =   6720
      Width           =   2775
   End
   Begin VB.Label LabelItemPerc5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   675
      Left            =   11925
      TabIndex        =   25
      Top             =   5670
      Width           =   2775
   End
   Begin VB.Label LabelItemPerc4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   675
      Left            =   11925
      TabIndex        =   24
      Top             =   4605
      Width           =   2775
   End
   Begin VB.Label LabelItemPerc3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   675
      Left            =   11925
      TabIndex        =   23
      Top             =   3555
      Width           =   2775
   End
   Begin VB.Label LabelItemPerc2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   675
      Left            =   11925
      TabIndex        =   22
      Top             =   2490
      Width           =   2775
   End
   Begin VB.Label LabelItemPerc1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   675
      Left            =   11925
      TabIndex        =   21
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label LabelItemQuan6 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   680
      Left            =   5472
      TabIndex        =   20
      Top             =   6720
      Width           =   2775
   End
   Begin VB.Label LabelItemQuan5 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   680
      Left            =   5472
      TabIndex        =   19
      Top             =   5664
      Width           =   2775
   End
   Begin VB.Label LabelItemQuan4 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   680
      Left            =   5472
      TabIndex        =   18
      Top             =   4608
      Width           =   2775
   End
   Begin VB.Label LabelItemQuan3 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   680
      Left            =   5472
      TabIndex        =   17
      Top             =   3552
      Width           =   2775
   End
   Begin VB.Label LabelItemQuan2 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   680
      Left            =   5472
      TabIndex        =   16
      Top             =   2496
      Width           =   2775
   End
   Begin VB.Label LabelItemQuan1 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   675
      Left            =   5475
      TabIndex        =   15
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H000000FF&
      BorderColor     =   &H0000B0B0&
      BorderWidth     =   2
      FillColor       =   &H000000FF&
      Height          =   684
      Left            =   5376
      Top             =   6720
      Width           =   9420
   End
   Begin VB.Label LabelItemNum6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00B0FFFF&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   50.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000B0B0&
      Height          =   972
      Left            =   96
      TabIndex        =   8
      Top             =   6528
      Width           =   876
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      FillColor       =   &H000000FF&
      Height          =   684
      Left            =   5376
      Top             =   5664
      Width           =   9420
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00B0B000&
      BorderWidth     =   2
      FillColor       =   &H000000FF&
      Height          =   684
      Left            =   5376
      Top             =   4608
      Width           =   9420
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000000FF&
      BorderColor     =   &H000070FF&
      BorderWidth     =   2
      FillColor       =   &H000000FF&
      Height          =   684
      Left            =   5376
      Top             =   3552
      Width           =   9420
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000000FF&
      BorderColor     =   &H0000B000&
      BorderWidth     =   2
      FillColor       =   &H000000FF&
      Height          =   684
      Left            =   5376
      Top             =   2496
      Width           =   9420
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      FillColor       =   &H000000FF&
      Height          =   684
      Left            =   5376
      Top             =   1440
      Width           =   9420
   End
   Begin VB.Label LabelInputCommand 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "&Press key:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   585
      Left            =   12285
      TabIndex        =   29
      Top             =   8610
      Width           =   1695
   End
   Begin VB.Label LabelStatusbar 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F0F0&
      Caption         =   "Welcome! Press F5 to start voting, F6 to change quantity."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   525
      Left            =   195
      TabIndex        =   28
      Top             =   8610
      Width           =   11895
   End
   Begin VB.Label LabelItemNum5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFB0FF&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   50.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   972
      Left            =   96
      TabIndex        =   7
      Top             =   5472
      Width           =   876
   End
   Begin VB.Label LabelItemNum4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFB0&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   50.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B0B000&
      Height          =   972
      Left            =   96
      TabIndex        =   6
      Top             =   4416
      Width           =   876
   End
   Begin VB.Label LabelItemNum3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00B0D0FF&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   50.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000070FF&
      Height          =   972
      Left            =   96
      TabIndex        =   5
      Top             =   3360
      Width           =   876
   End
   Begin VB.Label LabelItemNum2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00B0FFB0&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   50.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000B000&
      Height          =   972
      Left            =   96
      TabIndex        =   4
      Top             =   2304
      Width           =   876
   End
   Begin VB.Label LabelItemNum1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00B0B0FF&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   50.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   972
      Left            =   96
      TabIndex        =   3
      Top             =   1248
      Width           =   876
   End
   Begin VB.Shape ShapeItemBar1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00B0B0FF&
      FillStyle       =   0  'Solid
      Height          =   690
      Left            =   5370
      Top             =   1440
      Width           =   9255
   End
   Begin VB.Shape ShapeItemBar2 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00B0FFB0&
      FillStyle       =   0  'Solid
      Height          =   690
      Left            =   5370
      Top             =   2490
      Width           =   9255
   End
   Begin VB.Shape ShapeItemBar3 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00B0D0FF&
      FillStyle       =   0  'Solid
      Height          =   690
      Left            =   5370
      Top             =   3555
      Width           =   9255
   End
   Begin VB.Shape ShapeItemBar4 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFB0&
      FillStyle       =   0  'Solid
      Height          =   684
      Left            =   5376
      Top             =   4608
      Width           =   9255
   End
   Begin VB.Shape ShapeItemBar5 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFB0FF&
      FillStyle       =   0  'Solid
      Height          =   684
      Left            =   5376
      Top             =   5664
      Width           =   9255
   End
   Begin VB.Shape ShapeItemBar6 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00B0FFFF&
      FillStyle       =   0  'Solid
      Height          =   684
      Left            =   5376
      Top             =   6720
      Width           =   9255
   End
   Begin VB.Menu Menu 
      Caption         =   "&Menu"
      Begin VB.Menu MenuSoundSwitch 
         Caption         =   "S&ound"
         Checked         =   -1  'True
         Shortcut        =   {F2}
      End
      Begin VB.Menu MenuEXIT 
         Caption         =   "E&XIT"
         Shortcut        =   {F4}
      End
      Begin VB.Menu Menu1_ 
         Caption         =   "-"
      End
      Begin VB.Menu MenuVote 
         Caption         =   "&Vote"
         Begin VB.Menu MenuVoteStartStopClear 
            Caption         =   "&Start"
            Shortcut        =   {F5}
         End
         Begin VB.Menu MenuVoteTotalQuan 
            Caption         =   "&Quantity: 50"
            Shortcut        =   {F6}
         End
         Begin VB.Menu MenuVote1_ 
            Caption         =   "-"
         End
         Begin VB.Menu MenuVoteVoteCand1 
            Caption         =   "Vote for Candidate &1"
            Enabled         =   0   'False
            Shortcut        =   ^{F1}
         End
         Begin VB.Menu MenuVoteVoteCand2 
            Caption         =   "Vote for Candidate &2"
            Enabled         =   0   'False
            Shortcut        =   ^{F2}
         End
         Begin VB.Menu MenuVoteVoteCand3 
            Caption         =   "Vote for Candidate &3"
            Enabled         =   0   'False
            Shortcut        =   ^{F3}
         End
         Begin VB.Menu MenuVoteVoteCand4 
            Caption         =   "Vote for Candidate &4"
            Enabled         =   0   'False
            Shortcut        =   ^{F4}
         End
         Begin VB.Menu MenuVoteVoteCand5 
            Caption         =   "Vote for Candidate &5"
            Enabled         =   0   'False
            Shortcut        =   ^{F5}
         End
         Begin VB.Menu MenuVoteVoteCand6 
            Caption         =   "Vote for Candidate &6"
            Enabled         =   0   'False
            Shortcut        =   ^{F6}
         End
      End
      Begin VB.Menu Menu2_ 
         Caption         =   "-"
      End
      Begin VB.Menu MenuAbout 
         Caption         =   "&About"
         Begin VB.Menu MenuAboutDownload 
            Caption         =   "&Download latest version..."
         End
         Begin VB.Menu MenuAboutUpdate 
            Caption         =   "Check for &update manually..."
         End
         Begin VB.Menu MenuAboutGitHub 
            Caption         =   "&GitHub @SamToki user page..."
         End
         Begin VB.Menu MenuAboutLicense 
            Caption         =   "Released under &license GNU GPL v3..."
         End
         Begin VB.Menu MenuAboutCopyright 
            Caption         =   "TM && (C) 2015-2022 SAM TOKI STUDIO"
            Enabled         =   0   'False
         End
         Begin VB.Menu MenuAboutDate 
            Caption         =   "2022/01/28"
            Enabled         =   0   'False
         End
      End
   End
End
Attribute VB_Name = "FormMainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  ---------------------------------------------------------------------------------------------------------------------
'
'  === INFORMATION ===
'
'  SAM TOKI STUDIO
'  This is a .frm source code file.
'
'  VoteHelper
'
'  Powered by Sam Toki
'  Version: v1.03eng
'  Date:    2022/01/28 (Fri)
'  History: First version v0.10 was built on 2017/10/21.
'
'  WARNING: Commercial use of this computer software is strictly prohibited.
'           Open source license:      GNU GPL v3
'           Creative Commons license: CC BY-NC 3.0
'
'  Copyright: TM & (C) 2015-2022 SAM TOKI STUDIO. All rights reserved.
'             SAM TOKI STUDIO is a trademark of CZJ Software Technologies (CZJST) Inc. in the P.R.C and other countries.
'
'  ---------------------------------------------------------------------------------------------------------------------
'
'  === NOTES FOR REFERENCE ===
'
'  ...
'
'  ---------------------------------------------------------------------------------------------------------------------

'[] DECLARATIONS []

Option Explicit

'Declare Menu...
Public soundswitch As Boolean
Public inputnumberdigits As Integer

'Declare Vote...
Public status As Integer
Public totalquan As Integer
Public currentquan As Integer
Public maxquan As Integer
Public itemquan1 As Integer
Public itemquan2 As Integer
Public itemquan3 As Integer
Public itemquan4 As Integer
Public itemquan5 As Integer
Public itemquan6 As Integer
Public itemperc1 As Single
Public itemperc2 As Single
Public itemperc3 As Single
Public itemperc4 As Single
Public itemperc5 As Single
Public itemperc6 As Single

Public blinkorder As Integer

Public maxquanJudgeLoop As Integer  'MAX QUANTITY JUDGE, CODES FROM INTERNET
Public Arr As Variant  'MAX QUANTITY JUDGE, CODES FROM INTERNET

'Declare Animation...
Public progressanimationtarget As Integer  'Range: 0~15120
Public itembar1animationtarget As Integer  'Range: 0~9420
Public itembar2animationtarget As Integer  'Range: 0~9420
Public itembar3animationtarget As Integer  'Range: 0~9420
Public itembar4animationtarget As Integer  'Range: 0~9420
Public itembar5animationtarget As Integer  'Range: 0~9420
Public itembar6animationtarget As Integer  'Range: 0~9420

'Declare Dialog...
Public answer

'Declare Others...
Public setanimationswitch As Boolean

'Declare Special...
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
     ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOW = 5

'  ---------------------------------------------------------------------------------------------------------------------

'[] LOAD []

    Sub Form_Load()
        'Load and Initialization...

        soundswitch = True: inputnumberdigits = 4

        status = 0: totalquan = 50: currentquan = 0: maxquan = 0: blinkorder = 1
        itemquan1 = 0: itemquan2 = 0: itemquan3 = 0: itemquan4 = 0: itemquan5 = 0: itemquan6 = 0
        itemperc1 = 0: itemperc2 = 0: itemperc3 = 0: itemperc4 = 0: itemperc5 = 0: itemperc6 = 0

        progressanimationtarget = 0
        itembar1animationtarget = 0: itembar2animationtarget = 0: itembar3animationtarget = 0: itembar4animationtarget = 0: itembar5animationtarget = 0: itembar6animationtarget = 0

        setanimationswitch = True

        MenuVoteTotalQuan.Enabled = True
        MenuVoteVoteCand1.Enabled = False: MenuVoteVoteCand2.Enabled = False: MenuVoteVoteCand3.Enabled = False: MenuVoteVoteCand4.Enabled = False: MenuVoteVoteCand5.Enabled = False: MenuVoteVoteCand6.Enabled = False
        CmdTotalQuan.Enabled = True
        TextboxInput.Enabled = False

        Call Refresher: Call TimerMaxQuanBlink_Timer

        MenuVoteStartStopClear.Caption = "&Start"
        LabelStatusbar.Caption = "Welcome! Press F5 to start voting, F6 to change quantity."
    End Sub

'  ---------------------------------------------------------------------------------------------------------------------

'[] TIMERS []

    'CODES FROM INTERNET
    Public Function MaxQuanJudge(Arr As Variant)
        MaxQuanJudge = Arr(0)
        For maxquanJudgeLoop = 0 To UBound(Arr)
        If Arr(maxquanJudgeLoop) > MaxQuanJudge Then MaxQuanJudge = Arr(maxquanJudgeLoop)
        Next
    End Function

    Public Sub Refresher()
        'Refresh totalquan...
        MenuVoteTotalQuan.Caption = "&Quantity: " & totalquan
        CmdTotalQuan.Caption = totalquan

        'Refresh itemquan...
        LabelItemQuan1.Caption = itemquan1: LabelItemQuan2.Caption = itemquan2: LabelItemQuan3.Caption = itemquan3: LabelItemQuan4.Caption = itemquan4: LabelItemQuan5.Caption = itemquan5: LabelItemQuan6.Caption = itemquan6

        'Calculate maxquan...
        Arr = Array(itemquan1, itemquan2, itemquan3, itemquan4, itemquan5, itemquan6)
        maxquan = MaxQuanJudge(Arr)

        'Calculate percents...
        If Not ((itemquan1 + itemquan2 + itemquan3 + itemquan4 + itemquan5 + itemquan6) = 0) Then
            itemperc1 = 100 * itemquan1 / (itemquan1 + itemquan2 + itemquan3 + itemquan4 + itemquan5 + itemquan6)
            itemperc2 = 100 * itemquan2 / (itemquan1 + itemquan2 + itemquan3 + itemquan4 + itemquan5 + itemquan6)
            itemperc3 = 100 * itemquan3 / (itemquan1 + itemquan2 + itemquan3 + itemquan4 + itemquan5 + itemquan6)
            itemperc4 = 100 * itemquan4 / (itemquan1 + itemquan2 + itemquan3 + itemquan4 + itemquan5 + itemquan6)
            itemperc5 = 100 * itemquan5 / (itemquan1 + itemquan2 + itemquan3 + itemquan4 + itemquan5 + itemquan6)
            itemperc6 = 100 * itemquan6 / (itemquan1 + itemquan2 + itemquan3 + itemquan4 + itemquan5 + itemquan6)
        End If
            LabelItemPerc1.Caption = Format(itemperc1, "0.00") & "%": LabelItemPerc2.Caption = Format(itemperc2, "0.00") & "%": LabelItemPerc3.Caption = Format(itemperc3, "0.00") & "%": LabelItemPerc4.Caption = Format(itemperc4, "0.00") & "%": LabelItemPerc5.Caption = Format(itemperc5, "0.00") & "%": LabelItemPerc6.Caption = Format(itemperc6, "0.00") & "%"

        If Not ((itemquan1 + itemquan2 + itemquan3 + itemquan4 + itemquan5 + itemquan6) = 0) Then
            itemperc1 = 100 * itemquan1 / maxquan: itemperc2 = 100 * itemquan2 / maxquan: itemperc3 = 100 * itemquan3 / maxquan: itemperc4 = 100 * itemquan4 / maxquan: itemperc5 = 100 * itemquan5 / maxquan: itemperc6 = 100 * itemquan6 / maxquan
        End If

        'Check if vote ends...
        If currentquan > totalquan Then
            currentquan = totalquan: status = 2

            MenuVoteTotalQuan.Enabled = False
            MenuVoteVoteCand1.Enabled = False: MenuVoteVoteCand2.Enabled = False: MenuVoteVoteCand3.Enabled = False: MenuVoteVoteCand4.Enabled = False: MenuVoteVoteCand5.Enabled = False: MenuVoteVoteCand6.Enabled = False
            CmdTotalQuan.Enabled = False
            TextboxInput.Enabled = False

            MenuVoteStartStopClear.Caption = "&Clear"
            LabelStatusbar.Caption = "Vote finished! Press F5 to clear vote."

            If soundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Print Complete.wav"
        End If
    End Sub

    Public Sub TimerMaxQuanBlink_Timer()
        If maxquan = 0 Then
            LabelItemNum1.BackStyle = 0: LabelItemNum2.BackStyle = 0: LabelItemNum3.BackStyle = 0: LabelItemNum4.BackStyle = 0: LabelItemNum5.BackStyle = 0: LabelItemNum6.BackStyle = 0
            Exit Sub
        End If

        If itemquan1 = maxquan Then
            If blinkorder = 1 Then LabelItemNum1.BackStyle = 1 Else LabelItemNum1.BackStyle = 0
            Else: LabelItemNum1.BackStyle = 0
        End If
        If itemquan2 = maxquan Then
            If blinkorder = 1 Then LabelItemNum2.BackStyle = 1 Else LabelItemNum2.BackStyle = 0
            Else: LabelItemNum2.BackStyle = 0
        End If
        If itemquan3 = maxquan Then
            If blinkorder = 1 Then LabelItemNum3.BackStyle = 1 Else LabelItemNum3.BackStyle = 0
            Else: LabelItemNum3.BackStyle = 0
        End If
        If itemquan4 = maxquan Then
            If blinkorder = 1 Then LabelItemNum4.BackStyle = 1 Else LabelItemNum4.BackStyle = 0
            Else: LabelItemNum4.BackStyle = 0
        End If
        If itemquan5 = maxquan Then
            If blinkorder = 1 Then LabelItemNum5.BackStyle = 1 Else LabelItemNum5.BackStyle = 0
            Else: LabelItemNum5.BackStyle = 0
        End If
        If itemquan6 = maxquan Then
            If blinkorder = 1 Then LabelItemNum6.BackStyle = 1 Else LabelItemNum6.BackStyle = 0
            Else: LabelItemNum6.BackStyle = 0
        End If

        If blinkorder = 1 Then blinkorder = 0 Else blinkorder = 1
    End Sub

'[] COMMANDS []

    'CMD General...
    Public Sub MenuSoundSwitch_Click()
        Select Case soundswitch
            Case True
                soundswitch = False
                MenuSoundSwitch.Checked = False
            Case False
                soundswitch = True
                MenuSoundSwitch.Checked = True
        End Select
    End Sub
    Public Sub MenuEXIT_Click()
        End
    End Sub
    Public Sub MenuAboutDownload_Click()
        Call ShellExecute(Me.hWnd, "open", "https://github.com/SamToki/VB6---VoteHelper/raw/main/RELEASE/VoteHelper.exe", "", "", SW_SHOW)
    End Sub
    Public Sub MenuAboutUpdate_Click()
        Call ShellExecute(Me.hWnd, "open", "https://github.com/SamToki/VB6---VoteHelper", "", "", SW_SHOW)
    End Sub
    Public Sub MenuAboutGitHub_Click()
        Call ShellExecute(Me.hWnd, "open", "https://github.com/SamToki", "", "", SW_SHOW)
    End Sub
    Public Sub MenuAboutLicense_Click()
        Call ShellExecute(Me.hWnd, "open", "https://www.gnu.org/licenses/gpl-3.0.html", "", "", SW_SHOW)
    End Sub

    'CMD Vote...
    Public Sub MenuVoteTotalQuan_Click()
        FormInputNumber.currentinputnumber = 1
        FormInputNumber.LabelInputNumber1.Caption = ">": FormInputNumber.LabelInputNumber2.Caption = ">": FormInputNumber.LabelInputNumber3.Caption = ">": FormInputNumber.LabelInputNumber4.Caption = ">"
        FormMainWindow.Enabled = False

        FormInputNumber.Move (Screen.Width / 2), (Screen.Height / 2), 0, 0
        FormInputNumber.windowanimationtargetleft = (Screen.Width / 2) - (6210 / 2)
        FormInputNumber.windowanimationtargettop = (Screen.Height / 2) - (5895 / 2)
        FormInputNumber.windowanimationtargetwidth = 6210
        FormInputNumber.windowanimationtargetheight = 5895
        FormInputNumber.Show
    End Sub
    Public Sub CmdTotalQuan_Click()
        Call MenuVoteTotalQuan_Click
    End Sub

    Private Sub MenuVoteStartStopClear_Click()
        Select Case status
            Case 0
                status = 1: currentquan = 1
                FormInputNumber.Hide

                MenuVoteTotalQuan.Enabled = False
                MenuVoteVoteCand1.Enabled = True: MenuVoteVoteCand2.Enabled = True: MenuVoteVoteCand3.Enabled = True: MenuVoteVoteCand4.Enabled = True: MenuVoteVoteCand5.Enabled = True: MenuVoteVoteCand6.Enabled = True
                CmdTotalQuan.Enabled = False
                TextboxInput.Enabled = True
                TextboxInput.SetFocus

                MenuVoteStartStopClear.Caption = "Sto&p"
                LabelStatusbar.Caption = "Vote started!　" & currentquan & " / " & totalquan
            Case 1
                status = 2

                MenuVoteTotalQuan.Enabled = False
                MenuVoteVoteCand1.Enabled = False: MenuVoteVoteCand2.Enabled = False: MenuVoteVoteCand3.Enabled = False: MenuVoteVoteCand4.Enabled = False: MenuVoteVoteCand5.Enabled = False: MenuVoteVoteCand6.Enabled = False
                CmdTotalQuan.Enabled = False
                TextboxInput.Enabled = False

                MenuVoteStartStopClear.Caption = "&Clear"
                LabelStatusbar.Caption = "Vote stopped. Press F5 to clear vote."
            Case 2
                status = 0: currentquan = 0: maxquan = 0: blinkorder = 1
                itemquan1 = 0: itemquan2 = 0: itemquan3 = 0: itemquan4 = 0: itemquan5 = 0: itemquan6 = 0
                itemperc1 = 0: itemperc2 = 0: itemperc3 = 0: itemperc4 = 0: itemperc5 = 0: itemperc6 = 0
        
                MenuVoteTotalQuan.Enabled = True
                MenuVoteVoteCand1.Enabled = False: MenuVoteVoteCand2.Enabled = False: MenuVoteVoteCand3.Enabled = False: MenuVoteVoteCand4.Enabled = False: MenuVoteVoteCand5.Enabled = False: MenuVoteVoteCand6.Enabled = False
                CmdTotalQuan.Enabled = True
                TextboxInput.Enabled = False
        
                MenuVoteStartStopClear.Caption = "&Start"
                LabelStatusbar.Caption = "Vote cleared. Press F5 to start a new vote, F6 to change quantity."
        End Select

        If soundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Pop-up Blocked.wav"
        Call Refresher: Call TimerMaxQuanBlink_Timer
    End Sub

    Public Sub MenuVoteVoteCand1_Click()
        itemquan1 = itemquan1 + 1: currentquan = currentquan + 1
        LabelStatusbar.Caption = "Vote ongoing...　" & currentquan & " / " & totalquan & "　A new vote for Candidate 1 !"
        If soundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Ding.wav"
        Call Refresher
    End Sub
    Public Sub MenuVoteVoteCand2_Click()
        itemquan2 = itemquan2 + 1: currentquan = currentquan + 1
        LabelStatusbar.Caption = "Vote ongoing...　" & currentquan & " / " & totalquan & "　A new vote for Candidate 2 !"
        If soundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Ding.wav"
        Call Refresher
    End Sub
    Public Sub MenuVoteVoteCand3_Click()
        itemquan3 = itemquan3 + 1: currentquan = currentquan + 1
        LabelStatusbar.Caption = "Vote ongoing...　" & currentquan & " / " & totalquan & "　A new vote for Candidate 3 !"
        If soundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Ding.wav"
        Call Refresher
    End Sub
    Public Sub MenuVoteVoteCand4_Click()
        itemquan4 = itemquan4 + 1: currentquan = currentquan + 1
        LabelStatusbar.Caption = "Vote ongoing...　" & currentquan & " / " & totalquan & "　A new vote for Candidate 4 !"
        If soundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Ding.wav"
        Call Refresher
    End Sub
    Public Sub MenuVoteVoteCand5_Click()
        itemquan5 = itemquan5 + 1: currentquan = currentquan + 1
        LabelStatusbar.Caption = "Vote ongoing...　" & currentquan & " / " & totalquan & "　A new vote for Candidate 5 !"
        If soundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Ding.wav"
        Call Refresher
    End Sub
    Public Sub MenuVoteVoteCand6_Click()
        itemquan6 = itemquan6 + 1: currentquan = currentquan + 1
        LabelStatusbar.Caption = "Vote ongoing...　" & currentquan & " / " & totalquan & "　A new vote for Candidate 6 !"
        If soundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Ding.wav"
        Call Refresher
    End Sub

    Private Sub TextboxInput_Change()
        Select Case TextboxInput.Text
            Case "1"
                Call MenuVoteVoteCand1_Click
            Case "2"
                Call MenuVoteVoteCand2_Click
            Case "3"
                Call MenuVoteVoteCand3_Click
            Case "4"
                Call MenuVoteVoteCand4_Click
            Case "5"
                Call MenuVoteVoteCand5_Click
            Case "6"
                Call MenuVoteVoteCand6_Click
            Case ""
                Call Refresher
            Case Else
                LabelStatusbar.Caption = "Vote ongoing...　" & currentquan & " / " & totalquan & "　Invalid input!"
                If soundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\chord.wav"
        End Select

        TextboxInput.Text = ""
        Call Refresher
    End Sub

'  ---------------------------------------------------------------------------------------------------------------------

'[] ANIMATIONS []

    Public Sub TimerProgressbarAnimation_Timer()
        progressanimationtarget = 15120 * (currentquan / totalquan)
        itembar1animationtarget = 120 + 93 * itemperc1
        itembar2animationtarget = 120 + 93 * itemperc2
        itembar3animationtarget = 120 + 93 * itemperc3
        itembar4animationtarget = 120 + 93 * itemperc4
        itembar5animationtarget = 120 + 93 * itemperc5
        itembar6animationtarget = 120 + 93 * itemperc6

        If ShapeProgressbar.Width = progressanimationtarget Then GoTo TimerProgressbarAnimation_Skip1_
        If ShapeProgressbar.Width > progressanimationtarget Then ShapeProgressbar.Width = ShapeProgressbar.Width - Abs(ShapeProgressbar.Width - progressanimationtarget) / 4
        If ShapeProgressbar.Width < progressanimationtarget Then ShapeProgressbar.Width = ShapeProgressbar.Width + Abs(ShapeProgressbar.Width - progressanimationtarget) / 4
        If Abs(ShapeProgressbar.Width - progressanimationtarget) < 10 Then ShapeProgressbar.Width = progressanimationtarget
TimerProgressbarAnimation_Skip1_:

        If ShapeItemBar1.Width = itembar1animationtarget Then GoTo TimerProgressbarAnimation_Skip2_
        If ShapeItemBar1.Width > itembar1animationtarget Then ShapeItemBar1.Width = ShapeItemBar1.Width - Abs(ShapeItemBar1.Width - itembar1animationtarget) / 4
        If ShapeItemBar1.Width < itembar1animationtarget Then ShapeItemBar1.Width = ShapeItemBar1.Width + Abs(ShapeItemBar1.Width - itembar1animationtarget) / 4
        If Abs(ShapeItemBar1.Width - itembar1animationtarget) < 10 Then ShapeItemBar1.Width = itembar1animationtarget
TimerProgressbarAnimation_Skip2_:

        If ShapeItemBar2.Width = itembar2animationtarget Then GoTo TimerProgressbarAnimation_Skip3_
        If ShapeItemBar2.Width > itembar2animationtarget Then ShapeItemBar2.Width = ShapeItemBar2.Width - Abs(ShapeItemBar2.Width - itembar2animationtarget) / 4
        If ShapeItemBar2.Width < itembar2animationtarget Then ShapeItemBar2.Width = ShapeItemBar2.Width + Abs(ShapeItemBar2.Width - itembar2animationtarget) / 4
        If Abs(ShapeItemBar2.Width - itembar2animationtarget) < 10 Then ShapeItemBar2.Width = itembar2animationtarget
TimerProgressbarAnimation_Skip3_:

        If ShapeItemBar3.Width = itembar3animationtarget Then GoTo TimerProgressbarAnimation_Skip4_
        If ShapeItemBar3.Width > itembar3animationtarget Then ShapeItemBar3.Width = ShapeItemBar3.Width - Abs(ShapeItemBar3.Width - itembar3animationtarget) / 4
        If ShapeItemBar3.Width < itembar3animationtarget Then ShapeItemBar3.Width = ShapeItemBar3.Width + Abs(ShapeItemBar3.Width - itembar3animationtarget) / 4
        If Abs(ShapeItemBar3.Width - itembar3animationtarget) < 10 Then ShapeItemBar3.Width = itembar3animationtarget
TimerProgressbarAnimation_Skip4_:

        If ShapeItemBar4.Width = itembar4animationtarget Then GoTo TimerProgressbarAnimation_Skip5_
        If ShapeItemBar4.Width > itembar4animationtarget Then ShapeItemBar4.Width = ShapeItemBar4.Width - Abs(ShapeItemBar4.Width - itembar4animationtarget) / 4
        If ShapeItemBar4.Width < itembar4animationtarget Then ShapeItemBar4.Width = ShapeItemBar4.Width + Abs(ShapeItemBar4.Width - itembar4animationtarget) / 4
        If Abs(ShapeItemBar4.Width - itembar4animationtarget) < 10 Then ShapeItemBar4.Width = itembar4animationtarget
TimerProgressbarAnimation_Skip5_:

        If ShapeItemBar5.Width = itembar5animationtarget Then GoTo TimerProgressbarAnimation_Skip6_
        If ShapeItemBar5.Width > itembar5animationtarget Then ShapeItemBar5.Width = ShapeItemBar5.Width - Abs(ShapeItemBar5.Width - itembar5animationtarget) / 4
        If ShapeItemBar5.Width < itembar5animationtarget Then ShapeItemBar5.Width = ShapeItemBar5.Width + Abs(ShapeItemBar5.Width - itembar5animationtarget) / 4
        If Abs(ShapeItemBar5.Width - itembar5animationtarget) < 10 Then ShapeItemBar5.Width = itembar5animationtarget
TimerProgressbarAnimation_Skip6_:

        If ShapeItemBar6.Width = itembar6animationtarget Then GoTo TimerProgressbarAnimation_Skip7_
        If ShapeItemBar6.Width > itembar6animationtarget Then ShapeItemBar6.Width = ShapeItemBar6.Width - Abs(ShapeItemBar6.Width - itembar6animationtarget) / 4
        If ShapeItemBar6.Width < itembar6animationtarget Then ShapeItemBar6.Width = ShapeItemBar6.Width + Abs(ShapeItemBar6.Width - itembar6animationtarget) / 4
        If Abs(ShapeItemBar6.Width - itembar6animationtarget) < 10 Then ShapeItemBar6.Width = itembar6animationtarget
TimerProgressbarAnimation_Skip7_:

    End Sub
