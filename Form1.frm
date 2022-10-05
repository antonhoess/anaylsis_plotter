VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Analysis"
   ClientHeight    =   12555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12990
   ControlBox      =   0   'False
   ForeColor       =   &H00FF0000&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8.719
   ScaleMode       =   0  'User
   ScaleWidth      =   9.021
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TmrMouseCoordinates 
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer TmrCoefIncrDecr 
      Left            =   360
      Top             =   0
   End
   Begin VB.PictureBox FrmControl 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8415
      Left            =   240
      ScaleHeight     =   561
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   777
      TabIndex        =   0
      Top             =   600
      Width           =   11655
      Begin VB.Frame Frame3 
         BackColor       =   &H0080C0FF&
         Caption         =   "Precision"
         Height          =   615
         Left            =   9600
         TabIndex        =   139
         Top             =   2280
         Width           =   1935
         Begin MSComCtl2.FlatScrollBar ScrPrecision 
            Height          =   255
            Left            =   600
            TabIndex        =   140
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            Arrows          =   65536
            Max             =   8
            Orientation     =   1638401
            Value           =   3
         End
         Begin VB.Label LblPrecision 
            BackColor       =   &H0080C0FF&
            Caption         =   "?"
            Height          =   255
            Left            =   240
            TabIndex        =   141
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.ListBox LstTip 
         BackColor       =   &H00FF0000&
         Height          =   2400
         Left            =   10680
         TabIndex        =   129
         Top             =   5760
         Width           =   615
      End
      Begin VB.ListBox LstHop 
         BackColor       =   &H0000FF00&
         Height          =   2400
         Left            =   10080
         TabIndex        =   128
         Top             =   5760
         Width           =   615
      End
      Begin VB.ListBox LstNullsFdd 
         BackColor       =   &H00E0E0E0&
         Height          =   2400
         Left            =   9480
         TabIndex        =   127
         Top             =   5760
         Width           =   615
      End
      Begin VB.ListBox LstNullsFd 
         BackColor       =   &H00C0C0C0&
         Height          =   2400
         Left            =   8880
         TabIndex        =   126
         Top             =   5760
         Width           =   615
      End
      Begin VB.ListBox List8 
         Enabled         =   0   'False
         Height          =   2400
         Left            =   8280
         TabIndex        =   125
         TabStop         =   0   'False
         Top             =   5760
         Width           =   615
      End
      Begin VB.ListBox LstNullsDenMultiFactors 
         Height          =   2400
         Left            =   7680
         TabIndex        =   124
         TabStop         =   0   'False
         Top             =   5760
         Width           =   615
      End
      Begin VB.ListBox LstNullsDenMulti 
         Height          =   2400
         Left            =   7080
         TabIndex        =   123
         TabStop         =   0   'False
         Top             =   5760
         Width           =   615
      End
      Begin VB.ListBox LstNullsDen 
         BackColor       =   &H80000002&
         Height          =   2400
         Left            =   6480
         TabIndex        =   122
         TabStop         =   0   'False
         Top             =   5760
         Width           =   615
      End
      Begin VB.Frame FrmColor 
         BackColor       =   &H0080C0FF&
         Caption         =   "Color"
         Height          =   1095
         Left            =   600
         TabIndex        =   97
         Top             =   5280
         Width           =   675
         Begin VB.PictureBox PicColorSelArea 
            Height          =   778
            Left            =   670
            ScaleHeight     =   48
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   96
            TabIndex        =   99
            Top             =   240
            Width           =   1500
            Begin VB.PictureBox PicColorPalette 
               BackColor       =   &H00404040&
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   9
               Left            =   720
               ScaleHeight     =   16
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   16
               TabIndex        =   111
               TabStop         =   0   'False
               Top             =   240
               Width           =   240
            End
            Begin VB.PictureBox PicColorPalette 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   11
               Left            =   1200
               ScaleHeight     =   16
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   16
               TabIndex        =   110
               TabStop         =   0   'False
               Top             =   240
               Width           =   240
            End
            Begin VB.PictureBox PicColorPalette 
               BackColor       =   &H00004080&
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   10
               Left            =   960
               ScaleHeight     =   16
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   16
               TabIndex        =   109
               TabStop         =   0   'False
               Top             =   240
               Width           =   240
            End
            Begin VB.PictureBox PicColorPalette 
               BackColor       =   &H00C0C0FF&
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   8
               Left            =   480
               ScaleHeight     =   16
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   16
               TabIndex        =   108
               TabStop         =   0   'False
               Top             =   240
               Width           =   240
            End
            Begin VB.PictureBox PicColorPalette 
               BackColor       =   &H000080FF&
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   7
               Left            =   240
               ScaleHeight     =   16
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   16
               TabIndex        =   107
               TabStop         =   0   'False
               Top             =   240
               Width           =   240
            End
            Begin VB.PictureBox PicColorPalette 
               BackColor       =   &H00FF8080&
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   6
               Left            =   0
               ScaleHeight     =   16
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   16
               TabIndex        =   106
               TabStop         =   0   'False
               Top             =   240
               Width           =   240
            End
            Begin VB.PictureBox PicColorPalette 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   5
               Left            =   1200
               ScaleHeight     =   16
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   16
               TabIndex        =   105
               TabStop         =   0   'False
               Top             =   0
               Width           =   240
            End
            Begin VB.PictureBox PicColorPalette 
               BackColor       =   &H00C000C0&
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   4
               Left            =   960
               ScaleHeight     =   16
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   16
               TabIndex        =   104
               TabStop         =   0   'False
               Top             =   0
               Width           =   240
            End
            Begin VB.PictureBox PicColorPalette 
               BackColor       =   &H0000FF00&
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   3
               Left            =   720
               ScaleHeight     =   16
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   16
               TabIndex        =   103
               TabStop         =   0   'False
               Top             =   0
               Width           =   240
            End
            Begin VB.PictureBox PicColorPalette 
               BackColor       =   &H0080FFFF&
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   2
               Left            =   480
               ScaleHeight     =   16
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   16
               TabIndex        =   102
               TabStop         =   0   'False
               Top             =   0
               Width           =   240
            End
            Begin VB.PictureBox PicColorPalette 
               BackColor       =   &H000000FF&
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   0
               Left            =   0
               ScaleHeight     =   16
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   16
               TabIndex        =   101
               TabStop         =   0   'False
               Top             =   0
               Width           =   240
            End
            Begin VB.PictureBox PicColorPalette 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   1
               Left            =   240
               ScaleHeight     =   16
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   16
               TabIndex        =   100
               TabStop         =   0   'False
               Top             =   0
               Width           =   240
            End
            Begin VB.Label LblColorSelCustom 
               Alignment       =   2  'Center
               BackColor       =   &H000735BC&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "SELECT  "
               Height          =   240
               Left            =   0
               TabIndex        =   112
               Top             =   480
               Width           =   1440
            End
         End
         Begin VB.PictureBox PicColorMain 
            BackColor       =   &H00FF0000&
            Height          =   778
            Left            =   80
            ScaleHeight     =   0.5
            ScaleLeft       =   1
            ScaleMode       =   0  'User
            ScaleWidth      =   0.344
            TabIndex        =   98
            TabStop         =   0   'False
            Top             =   240
            Width           =   550
         End
      End
      Begin VB.ListBox LstNullsNumMultiFactors 
         Height          =   2400
         Left            =   5880
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   5760
         Width           =   615
      End
      Begin VB.ListBox LstNullsNumMulti 
         Height          =   2400
         Left            =   5280
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   5760
         Width           =   615
      End
      Begin VB.Frame FrmCalcFuncValue 
         BackColor       =   &H0080C0FF&
         Caption         =   "Funktionswert"
         Height          =   1095
         Left            =   2040
         TabIndex        =   89
         Top             =   5880
         Width           =   2175
         Begin VB.TextBox TxtCalcFuncValueY 
            Height          =   285
            Left            =   1140
            Locked          =   -1  'True
            TabIndex        =   94
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox TxtCalcFuncValueX 
            Height          =   285
            Left            =   360
            TabIndex        =   92
            Top             =   720
            Width           =   495
         End
         Begin VB.CommandButton BtnCalcFuncValue 
            BackColor       =   &H0080C0FF&
            Caption         =   "Funktionswert errechnen"
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   90
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   ") ="
            Height          =   255
            Left            =   900
            TabIndex        =   93
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "f ("
            Height          =   255
            Left            =   180
            TabIndex        =   91
            Top             =   720
            Width           =   255
         End
      End
      Begin VB.Frame FrmPlotSettings 
         BackColor       =   &H0080C0FF&
         Caption         =   "Plot-Settings"
         Height          =   1455
         Left            =   2040
         TabIndex        =   84
         Top             =   4320
         Width           =   2175
         Begin VB.CheckBox ChkAlwaysInForeground 
            BackColor       =   &H0080C0FF&
            Caption         =   "Immer im Vordergrund"
            Height          =   255
            Left            =   120
            TabIndex        =   88
            Top             =   1080
            Width           =   1935
         End
         Begin VB.CheckBox ChkAxisLabels 
            BackColor       =   &H0080C0FF&
            Caption         =   "Koordinaten"
            Height          =   255
            Left            =   120
            TabIndex        =   87
            Top             =   840
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CheckBox ChkAxes 
            BackColor       =   &H0080C0FF&
            Caption         =   "Achsenkreuz"
            Height          =   195
            Left            =   120
            TabIndex        =   86
            Top             =   600
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox ChkGrid 
            BackColor       =   &H0080C0FF&
            Caption         =   "Raster"
            Height          =   195
            Left            =   120
            TabIndex        =   85
            Top             =   360
            Value           =   1  'Checked
            Width           =   855
         End
      End
      Begin VB.Frame FmCalcIntegral 
         BackColor       =   &H0080C0FF&
         Caption         =   "Integral"
         Height          =   1815
         Left            =   9600
         TabIndex        =   73
         Top             =   360
         Width           =   1935
         Begin VB.TextBox TxtIntAbs 
            Height          =   375
            Left            =   1080
            TabIndex        =   82
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox TxtIntSum 
            Height          =   375
            Left            =   120
            TabIndex        =   80
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox TxtIntUpperBound 
            Height          =   285
            Left            =   1200
            TabIndex        =   78
            Text            =   "0"
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox TxtIntLowerBound 
            Height          =   285
            Left            =   1200
            TabIndex        =   76
            Text            =   "0"
            Top             =   360
            Width           =   615
         End
         Begin VB.CommandButton BtnCalcIntegral 
            BackColor       =   &H0080C0FF&
            Caption         =   "Integral errechnen"
            Height          =   735
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   74
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Betrag"
            Height          =   195
            Left            =   1080
            TabIndex        =   81
            Top             =   1080
            Width           =   465
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Summe"
            Height          =   195
            Left            =   120
            TabIndex        =   79
            Top             =   1080
            Width           =   525
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "b"
            Height          =   195
            Left            =   1080
            TabIndex        =   77
            Top             =   720
            Width           =   90
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "a"
            Height          =   195
            Left            =   1080
            TabIndex        =   75
            Top             =   375
            Width           =   90
         End
      End
      Begin VB.CheckBox ChkGridSpacingLock 
         BackColor       =   &H0080C0FF&
         Caption         =   "Lock"
         Height          =   375
         Left            =   2160
         TabIndex        =   72
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton BtnExtremum 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Extrema"
         Height          =   375
         Left            =   9600
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   4440
         Width           =   1935
      End
      Begin VB.CheckBox Check7 
         BackColor       =   &H0080C0FF&
         Caption         =   "Graph erklicken"
         Height          =   375
         Left            =   9600
         TabIndex        =   42
         Top             =   4800
         Width           =   1695
      End
      Begin VB.ListBox LstNullsNum 
         BackColor       =   &H80000002&
         Height          =   2400
         Left            =   4680
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   5760
         Width           =   615
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H0080C0FF&
         Caption         =   "Eigenschaften"
         Height          =   4935
         Left            =   4680
         TabIndex        =   64
         Top             =   240
         Width           =   4815
         Begin VB.CheckBox ChkRtf 
            BackColor       =   &H0080C0FF&
            Caption         =   "RTF"
            Height          =   255
            Left            =   3000
            Style           =   1  'Graphical
            TabIndex        =   137
            Top             =   600
            Value           =   1  'Checked
            Width           =   615
         End
         Begin VB.CheckBox ChkLinDecompShortForm 
            BackColor       =   &H0080C0FF&
            Caption         =   "Short"
            Height          =   375
            Left            =   3000
            Style           =   1  'Graphical
            TabIndex        =   135
            Top             =   240
            Value           =   1  'Checked
            Width           =   615
         End
         Begin VB.ListBox LstMainDefGapZero 
            Height          =   2400
            Left            =   2520
            TabIndex        =   134
            Top             =   2400
            Width           =   375
         End
         Begin VB.CommandButton BtnNewton 
            BackColor       =   &H0080C0FF&
            Caption         =   "Newton-Verfahren (Lücken, Pole und Nullstellen) + Linearfaktorzerlegung"
            Height          =   615
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   240
            Width           =   2895
         End
         Begin VB.ListBox LstMainPolesOrder 
            Height          =   2400
            Left            =   3840
            TabIndex        =   40
            Top             =   2400
            Width           =   615
         End
         Begin VB.ListBox LstMainPoles 
            Height          =   2400
            Left            =   3120
            TabIndex        =   39
            Top             =   2400
            Width           =   615
         End
         Begin VB.ListBox LstMainNullsMultiFactors 
            Height          =   2400
            Left            =   840
            TabIndex        =   38
            Top             =   2400
            Width           =   615
         End
         Begin VB.ListBox LstMainNullsMulti 
            Height          =   2400
            Left            =   120
            TabIndex        =   37
            Top             =   2400
            Width           =   615
         End
         Begin VB.ListBox LstMainDefGap 
            Height          =   2400
            Left            =   1800
            TabIndex        =   36
            Top             =   2400
            Width           =   615
         End
         Begin VB.CommandButton BtnNewtonShow 
            BackColor       =   &H0080C0FF&
            Caption         =   "Anzeigen"
            Height          =   615
            Left            =   3720
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   240
            Width           =   975
         End
         Begin RichTextLib.RichTextBox RtbLinFacNum 
            Height          =   375
            Left            =   120
            TabIndex        =   136
            Top             =   960
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   661
            _Version        =   393217
            BorderStyle     =   0
            Enabled         =   0   'False
            MultiLine       =   0   'False
            ReadOnly        =   -1  'True
            TextRTF         =   $"Form1.frx":0442
         End
         Begin RichTextLib.RichTextBox RtbLinFacDen 
            Height          =   375
            Left            =   120
            TabIndex        =   138
            Top             =   1560
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   661
            _Version        =   393217
            BorderStyle     =   0
            Enabled         =   0   'False
            MultiLine       =   0   'False
            ReadOnly        =   -1  'True
            TextRTF         =   $"Form1.frx":04C4
         End
         Begin VB.Label Label2941 
            BackStyle       =   0  'Transparent
            Caption         =   "Ordnung"
            Height          =   255
            Left            =   3840
            TabIndex        =   133
            Top             =   2160
            Width           =   615
         End
         Begin VB.Label Label294 
            BackStyle       =   0  'Transparent
            Caption         =   "Vielfachheit"
            Height          =   255
            Left            =   720
            TabIndex        =   132
            Top             =   2160
            Width           =   975
         End
         Begin VB.Label Label292 
            BackStyle       =   0  'Transparent
            Caption         =   "NS"
            Height          =   255
            Left            =   240
            TabIndex        =   130
            Top             =   2160
            Width           =   375
         End
         Begin VB.Label Label29412 
            BackStyle       =   0  'Transparent
            Caption         =   "Pole"
            Height          =   255
            Left            =   3240
            TabIndex        =   66
            Top             =   2160
            Width           =   375
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "Def.-lücken"
            Height          =   255
            Left            =   1920
            TabIndex        =   65
            Top             =   2160
            Width           =   855
         End
         Begin VB.Line Line1 
            X1              =   240
            X2              =   4560
            Y1              =   1440
            Y2              =   1440
         End
      End
      Begin VB.CommandButton BtnAsymptote 
         BackColor       =   &H0080C0FF&
         Caption         =   "Asymtote"
         Enabled         =   0   'False
         Height          =   375
         Left            =   9600
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   3960
         Width           =   1935
      End
      Begin VB.CommandButton BtnIntegrate 
         BackColor       =   &H0080C0FF&
         Caption         =   "Integrieren"
         Height          =   375
         Left            =   9600
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   3000
         Width           =   1935
      End
      Begin VB.TextBox TxtLineWidth 
         Height          =   285
         Left            =   1320
         TabIndex        =   25
         Text            =   "1"
         Top             =   5520
         Width           =   615
      End
      Begin MSComCtl2.FlatScrollBar ScrLineWidth 
         Height          =   495
         Left            =   1320
         TabIndex        =   62
         Top             =   5880
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   2
         Min             =   1
         Max             =   10
         Orientation     =   1638400
         Value           =   10
      End
      Begin VB.CommandButton BtnDifferentiate 
         BackColor       =   &H0080C0FF&
         Caption         =   "Differentieren"
         Height          =   375
         Left            =   9600
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   3480
         Width           =   1935
      End
      Begin VB.TextBox TxtDegreeDenominator 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   12
         Text            =   "0"
         Top             =   3480
         Width           =   495
      End
      Begin VB.CheckBox ChkRationalFunction 
         BackColor       =   &H0080C0FF&
         Caption         =   "Gebr. Rat. Fkt"
         Height          =   435
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox TxtDegreeNumerator 
         Height          =   285
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   10
         Text            =   "0"
         Top             =   2760
         Width           =   495
      End
      Begin VB.Frame FrmMainMenu 
         BackColor       =   &H0080C0FF&
         Caption         =   "Hauptmenü"
         Height          =   1095
         Left            =   600
         TabIndex        =   6
         Top             =   1680
         Width           =   1335
         Begin VB.CommandButton CmdDraw 
            BackColor       =   &H0080C0FF&
            Caption         =   "GO!"
            Height          =   255
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdClear 
            BackColor       =   &H0080C0FF&
            Caption         =   "Clear"
            Height          =   255
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   480
            Width           =   1095
         End
         Begin VB.CommandButton BtnTrace 
            BackColor       =   &H0080C0FF&
            Caption         =   "Trace"
            Height          =   255
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   720
            Width           =   1095
         End
      End
      Begin VB.Frame FrmDegCoef 
         BackColor       =   &H0080C0FF&
         Caption         =   "Grad = Koeffizient"
         Height          =   975
         Left            =   600
         TabIndex        =   54
         Top             =   240
         Width           =   1455
         Begin VB.OptionButton OptDenominator 
            BackColor       =   &H0080C0FF&
            Caption         =   "N"
            Enabled         =   0   'False
            Height          =   255
            Left            =   720
            TabIndex        =   4
            Top             =   600
            Width           =   375
         End
         Begin VB.OptionButton OptNumerator 
            BackColor       =   &H0080C0FF&
            Caption         =   "Z"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   600
            Value           =   -1  'True
            Width           =   375
         End
         Begin VB.TextBox TxtSetCoefficient 
            Height          =   285
            Left            =   720
            TabIndex        =   2
            Text            =   "0"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox TxtDegToSetCoef 
            Height          =   285
            Left            =   120
            TabIndex        =   1
            Text            =   "0"
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "="
            Height          =   255
            Left            =   480
            TabIndex        =   55
            Top             =   280
            Width           =   255
         End
      End
      Begin VB.CommandButton BtnCoefDecr 
         Caption         =   "-"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   8040
         Width           =   375
      End
      Begin VB.CommandButton BtnCoefIncr 
         Caption         =   "+"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton BtnCoefficients 
         BackColor       =   &H0080C0FF&
         Caption         =   "Koeffizienten"
         Height          =   255
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Plot-Skalierung"
         Height          =   1575
         Left            =   2040
         TabIndex        =   20
         Top             =   2640
         Width           =   2175
         Begin VB.CheckBox ChkProportional 
            BackColor       =   &H0080C0FF&
            Caption         =   "Proportional"
            Height          =   375
            Left            =   120
            TabIndex        =   23
            Top             =   1080
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.TextBox TxtUnitsHeight 
            Height          =   285
            Left            =   1440
            TabIndex        =   22
            Text            =   "10"
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox TxtUnitsWidth 
            Height          =   285
            Left            =   1440
            TabIndex        =   21
            Text            =   "13"
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Höheneinheiten"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Breiteneinheiten"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.CommandButton BtnExit 
         BackColor       =   &H000080FF&
         Caption         =   "Beenden"
         Height          =   615
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   7560
         Width           =   2175
      End
      Begin VB.CommandButton BtnSaveCoefficients 
         BackColor       =   &H0080C0FF&
         Caption         =   "Kfz. speichern"
         Height          =   375
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   4440
         Width           =   1335
      End
      Begin VB.CommandButton BtnLoadCoefficients 
         BackColor       =   &H0080C0FF&
         Caption         =   "Kfz. laden"
         Height          =   375
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   4800
         Width           =   1335
      End
      Begin VB.CommandButton BtnOffsetCoordSystem 
         BackColor       =   &H0080C0FF&
         Caption         =   "Verschieben (Koordinatensystem)"
         Height          =   615
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox TxtOffsetCoordSystemX 
         Height          =   285
         Left            =   2280
         TabIndex        =   15
         Text            =   "0"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox TxtOffsetCoordSystemY 
         Height          =   285
         Left            =   2280
         TabIndex        =   16
         Text            =   "0"
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox TxtGridSpacingX 
         Height          =   285
         Left            =   3720
         TabIndex        =   18
         Text            =   "1"
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox TxtGridSpacingY 
         Height          =   285
         Left            =   3720
         TabIndex        =   19
         Text            =   "1"
         Top             =   1800
         Width           =   735
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Intervall"
         Height          =   1815
         Left            =   600
         TabIndex        =   43
         Top             =   6480
         Width           =   1335
         Begin VB.CommandButton BtnCalcCodomain 
            BackColor       =   &H0080C0FF&
            Caption         =   "Wertebereich errechnen"
            Height          =   495
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   83
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox TxtCodomainUpperBound 
            Height          =   285
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   1440
            Width           =   495
         End
         Begin VB.TextBox TxtCodomainLowerBound 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   1440
            Width           =   495
         End
         Begin VB.TextBox TxtIntvLowerBound 
            Height          =   285
            Left            =   120
            TabIndex        =   26
            Text            =   "-1000"
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox TxtIntvUpperBound 
            Height          =   285
            Left            =   720
            TabIndex        =   27
            Text            =   "1000"
            Top             =   840
            Width           =   495
         End
         Begin VB.Label LblCodomain 
            BackStyle       =   0  'Transparent
            Caption         =   "Wertebereich:"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   1200
            Width           =   1215
         End
      End
      Begin VB.CommandButton BtnHide 
         BackColor       =   &H0000C0C0&
         Caption         =   "Ausblenden"
         Height          =   375
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   7080
         Width           =   2175
      End
      Begin ComctlLib.Slider SldCoef 
         Height          =   7695
         Left            =   120
         TabIndex        =   50
         Top             =   360
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   13573
         _Version        =   327682
         BorderStyle     =   1
         Orientation     =   1
         Min             =   -10000
         Max             =   10000
         TickStyle       =   3
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Ordnung"
         Height          =   255
         Left            =   6960
         TabIndex        =   131
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label11 
         BackColor       =   &H0080C0FF&
         Caption         =   "???"
         Height          =   255
         Left            =   8400
         TabIndex        =   121
         Top             =   5520
         Width           =   450
      End
      Begin VB.Label Label117 
         BackColor       =   &H0080C0FF&
         Caption         =   "NS Nx"
         Height          =   255
         Left            =   7800
         TabIndex        =   120
         Top             =   5520
         Width           =   570
      End
      Begin VB.Label Label116 
         BackColor       =   &H0080C0FF&
         Caption         =   "NS N*"
         Height          =   255
         Left            =   7200
         TabIndex        =   119
         Top             =   5520
         Width           =   450
      End
      Begin VB.Label Label115 
         BackColor       =   &H0080C0FF&
         Caption         =   "NS Zx"
         Height          =   255
         Left            =   6000
         TabIndex        =   118
         Top             =   5520
         Width           =   570
      End
      Begin VB.Label Label114 
         BackColor       =   &H0080C0FF&
         Caption         =   "NS Z*"
         Height          =   255
         Left            =   5400
         TabIndex        =   117
         Top             =   5520
         Width           =   450
      End
      Begin VB.Label Label113 
         BackColor       =   &H0080C0FF&
         Caption         =   "NS N"
         Height          =   255
         Left            =   6600
         TabIndex        =   116
         Top             =   5520
         Width           =   450
      End
      Begin VB.Label Label111 
         BackColor       =   &H0080C0FF&
         Caption         =   "NS Z"
         Height          =   255
         Left            =   4800
         TabIndex        =   115
         Top             =   5520
         Width           =   450
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080C0FF&
         Caption         =   "NS f''"
         Height          =   255
         Left            =   9600
         TabIndex        =   114
         Top             =   5520
         Width           =   375
      End
      Begin VB.Label Label23 
         BackColor       =   &H0080C0FF&
         Caption         =   "NS f'"
         Height          =   255
         Left            =   9000
         TabIndex        =   113
         Top             =   5520
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Y:"
         Height          =   255
         Left            =   2120
         TabIndex        =   71
         Top             =   1120
         Width           =   135
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "Y:"
         Height          =   255
         Left            =   840
         TabIndex        =   70
         Top             =   4200
         Width           =   135
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tip"
         Height          =   195
         Left            =   10800
         TabIndex        =   69
         Top             =   5520
         Width           =   225
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hop"
         Height          =   195
         Left            =   10200
         TabIndex        =   68
         Top             =   5520
         Width           =   300
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Dicke"
         Height          =   255
         Left            =   1320
         TabIndex        =   63
         Top             =   5280
         Width           =   495
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Grad(N)"
         Enabled         =   0   'False
         Height          =   255
         Left            =   720
         TabIndex        =   61
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label LblMouseCoordsX 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   60
         Top             =   3960
         Width           =   615
      End
      Begin VB.Label LblMouseCoordsY 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   59
         Top             =   4200
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Grad(Z)"
         Height          =   255
         Left            =   720
         TabIndex        =   58
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "X:"
         Height          =   255
         Left            =   840
         TabIndex        =   57
         Top             =   3960
         Width           =   135
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Koordinaten:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   56
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label LblMoveMenu 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Menü verschieben"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   49
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "X:"
         Height          =   255
         Left            =   2120
         TabIndex        =   48
         Top             =   750
         Width           =   135
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Grid Spacing X:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   47
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Grid Spacing Y:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   46
         Top             =   1800
         Width           =   1455
      End
   End
   Begin MSComDlg.CommonDialog Cdg1 
      Left            =   840
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "gps"
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim X, Y, X1, Y1, Y2, X2, V, G, B As Boolean, W, KSX, KSY, SFX, SFY, STPX, STPY
Dim DragX As Integer, DragY As Integer ' For drag&drop the control panel
Dim WXK As Boolean 'Wiederholte X-Koordinate
Dim CoefIncr As Boolean
Dim aX, aY, aY2, dx, dy
Dim Precision As Long

Private Sub ChkRtf_Click()
    Call BtnNewton_Click
End Sub

Private Sub ChkGridSpacingLock_Click()
    If ChkGridSpacingLock.Value = 1 Then
        TxtGridSpacingY.Text = TxtGridSpacingX.Text
        Call GridSpacing
    End If
End Sub

Private Sub ChkLinDecompShortForm_Click()
    Call BtnNewton_Click
    ChkRtf.Enabled = ChkLinDecompShortForm.Value
End Sub

Private Sub ChkProportional_Click()
    If ChkProportional.Value = 0 Then
        FrmMain.ScaleHeight = Int(TxtUnitsWidth.Text / STPX * STPY * 100) / 100
    Else
        TxtUnitsHeight.Text = Int(TxtUnitsWidth.Text / STPX * STPY * 100) / 100
        FrmMain.ScaleHeight = TxtUnitsHeight.Text
    End If
    
    Draw
End Sub

Private Sub Draw(Optional Clear As Boolean = True)
    If Clear Then FrmMain.Cls
    If ChkGrid.Value = 1 Then Call Raster
    If ChkAxisLabels.Value = 1 Then Call Koordinaten
    If ChkAxes.Value = 1 Then Call Nullpunkt
    Call Graph
End Sub

Private Sub ChkGrid_Click()
    Draw
End Sub

Private Sub ChkAxisLabels_Click()
    Draw
End Sub

Private Sub ChkAxes_Click()
    Draw
End Sub

Private Sub ChkAlwaysInForeground_Click()
    If ChkAlwaysInForeground.Value = 0 Then
        'Form in den Normalzustand zurücksetzen
        Call SetWindowPos(Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, 3)
    Else
        'Form dauerhaft in den Vordergrund setzen
        Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, 3)
    End If
End Sub

Private Sub ChkRationalFunction_Click()
    GRF = ChkRationalFunction.Value
    
    If ChkRationalFunction.Value = 1 Then
        IsRationalFunction = True
        OptDenominator.Enabled = True
        Label21.Enabled = True
        TxtDegreeDenominator.Enabled = True
        BtnAsymptote.Enabled = True
    Else
        IsRationalFunction = False
        OptDenominator.Enabled = False
        OptNumerator.Value = True
        Label21.Enabled = False
        TxtDegreeDenominator.Enabled = False
        BtnAsymptote.Enabled = False
    End If
End Sub

Private Sub Check7_Click()
    FrmMainMenu.Enabled = True
    'If Check7.Value = 0 Then DegNum = -1
End Sub

Private Sub GridSpacing()
    SFX = TxtGridSpacingX.Text
    SFY = TxtGridSpacingY.Text
    Draw
End Sub

Private Sub BtnHide_Click()
    FrmControl.Visible = False
End Sub

Private Sub BtnCalcCodomain_Click()
    If Not IsRationalFunction Then
        TxtCodomainLowerBound.Text = GetFuncValByX(TxtIntvLowerBound.Text, CoefNum)
        TxtCodomainUpperBound.Text = GetFuncValByX(TxtIntvUpperBound.Text, CoefNum)
    Else
        TxtCodomainLowerBound.Text = _
        GetFuncValByX(TxtIntvLowerBound.Text, CoefNum) / _
        GetFuncValByX(TxtIntvLowerBound.Text, CoefDen)

        TxtCodomainUpperBound.Text = _
        GetFuncValByX(TxtIntvUpperBound.Text, CoefNum) / _
        GetFuncValByX(TxtIntvUpperBound.Text, CoefDen)
    End If
End Sub

Private Sub BtnCalcFuncValue_Click()
    Dim Value As Double
    
    On Error GoTo CalcValueError

    Value = GetFuncValByX(CDbl(TxtCalcFuncValueX.Text), CoefNum)
    
    If IsRationalFunction Then
        Value = Value / GetFuncValByX(CDbl(TxtCalcFuncValueX.Text), CoefDen)
    End If
    
    TxtCalcFuncValueY.Text = Value
    
    Exit Sub
    
CalcValueError:
    TxtCalcFuncValueY.Text = "Undefined"
End Sub


Private Sub BtnCoefIncr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CoefIncr = True
    TmrCoefIncrDecr.Interval = 250
End Sub


Private Sub BtnCoefIncr_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    TmrCoefIncrDecr.Interval = 0
End Sub


Private Sub BtnCoefDecr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CoefIncr = False
    TmrCoefIncrDecr.Interval = 250
End Sub


Private Sub BtnCoefDecr_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    TmrCoefIncrDecr.Interval = 0
End Sub


Private Sub BtnAsymptote_Click()
    If IsRationalFunction Then
        Dim AllDenCoefsAreZero As Boolean
        Dim DegCur As Integer
        Dim DegAsymptote As Integer
        Dim CoefAsymptote() As Double
        
        ' Check if all denominator coefficients are zero and exit sub if so
        ' Not sure if logic is correct
        AllDenCoefsAreZero = True
        
        For DegCur = 0 To DegDen
            If CoefDen(DegCur) <> 0 Then
                AllDenCoefsAreZero = False
                Exit For
            End If
        Next DegCur
        
        If AllDenCoefsAreZero Then
            MsgBox "Asymptote nicht definiert!", , "Hinweis"
            Exit Sub
        End If
        
        DegAsymptote = DegNum - DegDen
        
        ' If there is an asymptote
        If DegAsymptote >= 0 Then
            Dim CoefNumAsymptote() As Double, CoefDenAsymptote() As Double
            ReDim CoefNumAsymptote(0 To DegNum)
            ReDim CoefDenAsymptote(0 To DegDen)
            ReDim CoefAsymptote(0 To DegAsymptote)
            
            Dim I As Integer
            Dim K As Integer
            Dim CoefTmp As Double
            
            CoefNumAsymptote = CoefNum
            CoefDenAsymptote = CoefDen
            
            ' Perform polynomial division
            For I = 0 To DegAsymptote
                CoefTmp = CoefNumAsymptote(DegNum - I) / CoefDenAsymptote(DegDen)
                CoefAsymptote(DegAsymptote - I) = CoefTmp
                For K = 0 To DegDen
                    CoefNumAsymptote(DegNum - I - K) = CoefNumAsymptote(DegNum - I - K) - CoefTmp * CoefDenAsymptote(DegDen - K)
                Next K
            Next I
            
            ' XXX Fill structure and pass it to the Graph routine - might help to unify all GraphN-rountines into one
            Dim Func As RationalFunction
            With Func
                .IsRational = False
                .DegNum = DegAsymptote
                .DegDen = Empty
                ReDim .CoefNum(0 To DegAsymptote)
                .CoefNum = CoefAsymptote
                Erase .CoefDen
            End With
            
            ' Draw the graph
            Call GraphInternal(Func)
        End If
    End If
End Sub


Private Function GetLinearFactorString(Nulls As NewtonResult, Degree As Integer, Optional ShortForm As Boolean = False, Optional Rtf As Boolean = False) As String
    Dim I As Integer
    Dim Sign As String
    Dim LFS As String
        
    If Degree > 0 And Nulls.NullCnt = Degree Then
        ' Write main factor
        LFS = Round(Nulls.Factor, Precision)
    
        ' Write linear factors to text box
        If ShortForm Then
            Dim MultiNumbers() As Double
    
            MultiNumbers = MergeNumbersToMultiNumbers(Nulls.Nulls)
            
            If Not IsArrayEmpty(MultiNumbers) Then
                For I = 0 To CInt((UBound(MultiNumbers) - 1) \ 2)
                    If MultiNumbers(I * 2) < 0 Then
                        Sign = "+"
                    Else
                        Sign = "-"
                    End If
                    
                    If MultiNumbers(I * 2 + 1) > 1 Then
                        If Not Rtf Then
                            LFS = LFS & " (x " & Sign & Str(Abs(MultiNumbers(I * 2))) & ")^" & MultiNumbers(I * 2 + 1)
                        Else
                            LFS = LFS & " · (x " & Sign & Str(Abs(MultiNumbers(I * 2))) & "){\super " & MultiNumbers(I * 2 + 1) & "}"
                        End If
                    Else
                        If Not Rtf Then
                            LFS = LFS & " (x " & Sign & Str(Abs(MultiNumbers(I * 2))) & ")"
                        Else
                            LFS = LFS & " · (x " & Sign & Str(Abs(MultiNumbers(I * 2))) & ")"
                        End If
                    End If
                Next I
                
                If Rtf Then
                    LFS = "{\rtf1\ansi\deff0 {\fonttbl {\f0 Arial;}}" & LFS & "}"
                End If
            End If
        Else
            For I = 0 To Nulls.NullCnt - 1
                If Nulls.Nulls(I) < 0 Then
                    Sign = "+"
                Else
                    Sign = "-"
                End If
                LFS = LFS & " (x " + Sign + Str(Abs(Nulls.Nulls(I))) + ")"
            Next I
        End If
    Else
        LFS = "Cannot calculate linear factors!"
    End If
        
    GetLinearFactorString = LFS
End Function


Private Function MergeNumbersToMultiNumbers(Numbers() As Double) As Double()
    Dim I As Integer, K As Integer
    Dim Found As Boolean
    Dim ResCount As Integer
    Dim Result() As Double
    
    If Not IsArrayEmpty(Numbers) Then
        ' We ist this result list for the Numbers and their multiplicity in this format: Null0, Mult0, Null1, Mult1, ...
        ReDim Result(0 To UBound(Numbers) * 2 + 1)
    
        ' Check all numbers in list
        For I = 0 To UBound(Numbers)
            Found = False
            ' Check the current number against possible entries in the result list
            For K = 0 To ResCount - 1
                If Numbers(I) = Result(K * 2) Then
                    Found = True
                    Exit For
                End If
            Next K
            
            If Not Found Then
                ' Add new result entry (number + multiplicity=1)
                Result(ResCount * 2) = Numbers(I)
                Result(ResCount * 2 + 1) = 1
                ResCount = ResCount + 1
            Else
                ' Increment multiplicity of existing result
                Result(K * 2 + 1) = Result(K * 2 + 1) + 1
            End If
        Next I
        
        ' Shrink array to results only
        ReDim Preserve Result(0 To ResCount * 2 - 1)
    End If
    
    MergeNumbersToMultiNumbers = Result
End Function


Private Sub MergeNullsToMultiNulls(LstNulls As ListBox, LstNullsMulti As ListBox, LstNullsMultiFactors As ListBox)
    Dim I As Integer
    Dim MultiNumbers() As Double
    Dim Numbers() As Double
    
    ' Fill list with numbers from list box
    ReDim Numbers(0 To LstNulls.ListCount - 1)
    For I = 0 To LstNulls.ListCount - 1
        Numbers(I) = CDbl(LstNulls.List(I))
    Next I
    
    MultiNumbers = MergeNumbersToMultiNumbers(Numbers)

    If Not IsArrayEmpty(MultiNumbers) Then
        For I = 0 To CInt((UBound(MultiNumbers) - 1) \ 2)
            LstNullsMulti.AddItem (MultiNumbers(I * 2))
            LstNullsMultiFactors.AddItem (MultiNumbers(I * 2 + 1))
        Next I
    End If
End Sub


Private Sub AddNullsToList(Nulls As NewtonResult, LstNulls As ListBox)
    Dim I As Integer
    
    For I = 0 To Nulls.NullCnt - 1
        LstNulls.AddItem (Nulls.Nulls(I))
    Next I
End Sub


Private Sub DetermineNulls()
    Dim N As Integer, D As Integer
    Dim Found As Boolean
    
    ' Determine nulls
    For N = 0 To LstNullsNumMulti.ListCount - 1 ' Check all numerator nulls
        ' Check if the current null is also contained in the denominator nulls
        Found = False
        For D = 0 To LstNullsDenMulti.ListCount - 1
            If LstNullsNumMulti.List(N) = LstNullsDenMulti.List(D) Then
                Found = True
                Exit For
            End If
        Next D
        
        If Not Found Then
            LstMainNullsMulti.AddItem (LstNullsNumMulti.List(N))
            LstMainNullsMultiFactors.AddItem (LstNullsNumMultiFactors.List(N))
        End If
    Next N
End Sub

Private Sub DetermineDefinitionGaps()
    Dim Found As Boolean
    Dim N As Integer, D As Integer
    Dim MultiplicityDiff As Integer
    
    ' Determine definition gaps
    For D = 0 To LstNullsDenMulti.ListCount - 1 ' Check all denominator nulls
        Found = False
        
        ' Check if the current null is also contained in the numerator nulls
        For N = 0 To LstNullsNumMulti.ListCount - 1
            If LstNullsDenMulti.List(D) = LstNullsNumMulti.List(N) Then
                Found = True
                Exit For
            End If
        Next N

        If Found Then
            LstMainDefGap.AddItem (LstNullsDenMulti.List(D))
            
            ' Differentiate between definition gaps with value 0 or p(x)/q(x)
            If LstNullsNumMultiFactors.List(N) = LstNullsDenMultiFactors.List(D) Then
                LstMainDefGapZero.AddItem (LstNullsNumMultiFactors.List(N))
            Else
                If LstNullsNumMultiFactors.List(N) > LstNullsDenMultiFactors.List(D) Then
                    LstMainDefGapZero.AddItem (0)
                End If
            End If
        Else
            LstMainPoles.AddItem (LstNullsDenMulti.List(D))
            LstMainPolesOrder.AddItem (LstNullsDenMultiFactors.List(D))
        End If
    Next D
End Sub


Private Sub SetLinFacTxt(RtbLinFac As RichTextBox, Nulls As NewtonResult, Degree As Integer)
    ' Write linear factor decomposition to (RTF) text box
    If ChkRtf.Value = 1 And ChkLinDecompShortForm.Value = 1 Then
        RtbLinFac.TextRTF = GetLinearFactorString(Nulls, Degree, ChkLinDecompShortForm.Value = 1, True)
    Else
        ' Reset Font
        RtbLinFac.Font.Name = "MS Sans Serif"
        RtbLinFac.Font.Size = 8
        
        RtbLinFac.Text = GetLinearFactorString(Nulls, Degree, ChkLinDecompShortForm.Value = 1, False)
    End If
End Sub


Private Sub BtnNewton_Click()
    Dim I As Integer
    Dim NullsNum As NewtonResult, NullsDen As NewtonResult
    Dim Found As Boolean
    Dim Removable As Boolean
    Dim MultiplicityDiff As Integer
    Dim NumValueAtX As Double
    Dim DenValueAtX As Double
    'XXX Call HornerSchema
    
    If IsArrayEmpty(CoefNum) Then Exit Sub
    
    If Not IsRationalFunction Then
        NullsNum = Newton((CoefNum), True, Precision)
    Else
        NullsNum = Newton((CoefNum), True, Precision)
        NullsDen = Newton((CoefDen), False, Precision)
    End If
    
    If Not IsRationalFunction Then
        If NullsNum.NullCnt = 0 Then Exit Sub
    Else
        If NullsNum.NullCnt = 0 Then Exit Sub
        If NullsDen.NullCnt = 0 Then Exit Sub
    End If

    ' Clear controls
    LstNullsNum.Clear
    LstNullsNumMulti.Clear
    LstNullsNumMultiFactors.Clear
    LstNullsDen.Clear
    LstNullsDenMulti.Clear
    LstNullsDenMultiFactors.Clear
    
    LstMainDefGap.Clear
    LstMainDefGapZero.Clear
    LstMainNullsMulti.Clear
    LstMainNullsMultiFactors.Clear
    LstMainPoles.Clear
    LstMainPolesOrder.Clear
    
    If IsRationalFunction Then
        ' f(x) = p(x) / q(x)
        '
        ' Nullstelle + Vielfachheit -> p(x) = 0 && q(x) != 0
        ' Definitionslücken, zwei Arten:
        ' -> Hebbare Definitionslücke -> mult(null_den(x)) <= mult(null_num(x))
        ' -> Polstelle + Ordnung -> p(x) != 0 && q(x) = 0
        
        ' Determine base values
        ' > Add nulls to list
        Call AddNullsToList(NullsNum, LstNullsNum)
        Call AddNullsToList(NullsDen, LstNullsDen)
        
        ' > Merge nulls to multi-nulls where possible
        Call MergeNullsToMultiNulls(LstNullsNum, LstNullsNumMulti, LstNullsNumMultiFactors)
        Call MergeNullsToMultiNulls(LstNullsDen, LstNullsDenMulti, LstNullsDenMultiFactors)
        
        ' Write linear factor decomposition to (RTF) text box
        Call SetLinFacTxt(RtbLinFacNum, NullsNum, DegNum)
        Call SetLinFacTxt(RtbLinFacDen, NullsDen, DegDen)
        
        ' Determine nulls
        Call DetermineNulls
        
        ' Determine definition gaps
        Call DetermineDefinitionGaps
    
    Else ' Non-rational functions
        ' Determine base values
        ' > Add nulls to list
        Call AddNullsToList(NullsNum, LstNullsNum)
        
        ' > Merge nulls to multi-nulls where possible
        Call MergeNullsToMultiNulls(LstNullsNum, LstNullsNumMulti, LstNullsNumMultiFactors)
        
        ' Write linear factor decomposition to (RTF) text box
        Call SetLinFacTxt(RtbLinFacNum, NullsNum, DegNum)
        
        RtbLinFacNum.TextRTF = GetLinearFactorString(NullsNum, DegNum, ChkLinDecompShortForm.Value = 1, True)
        
        ' Copy list entries - there is no more to evaluate like it is the case with rational functions
        For I = 0 To LstNullsNumMulti.ListCount - 1
            LstMainNullsMulti.AddItem (LstNullsNumMulti.List(I))
            LstMainNullsMultiFactors.AddItem (LstNullsNumMultiFactors.List(I))
        Next I
    End If
End Sub


Private Sub BtnDifferentiate_Click()
    Dim I As Integer, K As Integer
    If Not IsRationalFunction Then
        ' Differentiate
        If DegNum > 0 Then
            For I = 1 To DegNum
                CoefNum(I - 1) = CoefNum(I) * I
            Next I
            DegNum = DegNum - 1
            ReDim Preserve CoefNum(0 To DegNum)
            TxtDegreeNumerator.Text = DegNum
        Else
            MsgBox "Funktion kann nicht mehr weiter differenziert werden."
        End If
    Else
        ' Quotient rule, see http://www.netalive.org/rationale-funktionen/chapters/3.5.html
        ' f(x) = g(x) / h(x)
        ' f'(x) = (h(x)*g'(x) - g(x)*h'(x)) / [h(x)]^2
        Dim CoefNumDiff() As Double
        Dim CoefDenDiff() As Double
        ReDim CoefNumDiff(0 To DegNum - 1) ' XXX Stürzt ab, wenn DegNum = 0 ist - wohl auch selbiges Problem mit DegDen
        ReDim CoefDenDiff(0 To DegDen - 1)
        
        ' Differentiate numerator and denominator polynomes
        For I = 1 To DegNum
            CoefNumDiff(I - 1) = CoefNum(I) * I
        Next I
        
        For I = 1 To DegDen
            CoefDenDiff(I - 1) = CoefDen(I) * I
        Next I
        
        ' Calculate numerator-part of differentiated function
        Dim CoefResNum() As Double
        Dim CoefResDen() As Double
        ReDim CoefResNum(0 To DegNum + DegDen - 1)
        ReDim CoefResDen(0 To DegDen * 2)
        
        ' Numerator-part
        ' h(x) * g'(x)
        For I = 0 To DegDen
            For K = 0 To DegNum - 1
                CoefResNum(I + K) = CoefResNum(I + K) + CoefDen(I) * CoefNumDiff(K)
            Next K
        Next I
        
        ' g(x) * h'(x)
        For I = 0 To DegNum
            For K = 0 To DegDen - 1
                CoefResNum(I + K) = CoefResNum(I + K) - CoefNum(I) * CoefDenDiff(K)
            Next K
        Next I
        
        ' Denominator-part
        For I = 0 To DegDen
            For K = 0 To DegDen
                CoefResDen(I + K) = CoefResDen(I + K) + CoefDen(I) * CoefDen(K)
            Next K
        Next I
    
        ReDim CoefNum(0 To DegNum * 2)
        ReDim CoefDen(0 To DegDen * 2 + 1)
        
        CoefNum = CoefResNum
        CoefDen = CoefResDen
    End If
End Sub


Private Sub BtnIntegrate_Click()
    If Not IsRationalFunction Then
        Dim I As Integer
        ' Integrate
        ReDim Preserve CoefNum(0 To DegNum + 1)
        For I = DegNum To 0 Step -1
            CoefNum(I + 1) = CoefNum(I) / (I + 1)
        Next I
        CoefNum(0) = 0
        DegNum = DegNum + 1
        TxtDegreeNumerator.Text = DegNum
    Else
        MsgBox "Integration not yet implemented for rational functions!"
    End If
End Sub


Private Sub BtnNewtonShow_Click()
    Dim I As Integer, J As Integer
    Dim X As Double
    Dim FormDrawSettings As DrawSettings
    Dim CoefNumTmp() As Double
    Dim CoefDenTmp() As Double
    
    FormDrawSettings = GetDrawSettings(FrmMain)
    
    FrmMain.DrawWidth = 3
    
    If IsRationalFunction Then
        ' Draw nulls
        For I = 0 To LstMainNullsMulti.ListCount - 1
            FrmMain.Circle (LstMainNullsMulti.List(I) + FrmMain.ScaleWidth / 2, FrmMain.ScaleHeight / 2), 0.1, RGB(255, 0, 0)
        Next I
        
        ' Draw definition gaps
        For I = 0 To LstMainDefGap.ListCount - 1
            If CInt(LstMainDefGapZero.List(I)) = 0 Then ' Definition gap with value 0
                FrmMain.Circle (LstMainDefGap.List(I) + FrmMain.ScaleWidth / 2, FrmMain.ScaleHeight / 2), 0.1, RGB(255, 0, 0)
            Else ' Reduce fraction to prevent calculation of 0/0
                ' Create temp. helper coefficient arrays
                ReDim CoefNumTmp(0 To UBound(CoefNum))
                ReDim CoefNumTmp(0 To UBound(CoefDen))
                CoefNumTmp = CoefNum
                CoefDenTmp = CoefDen
                
                ' Divide polynomials N times by the null in the current definition gap
                For J = 0 To CInt(LstMainDefGapZero.List(J)) - 1 '  Repeat N times where N is the multiplicity of this null
                    Call Nullstellendivision(CoefNumTmp, CInt(LstMainDefGap.List(I)))
                    Call Nullstellendivision(CoefDenTmp, CInt(LstMainDefGap.List(I)))
                Next J
                
                Y1 = GetFuncValByX(LstMainDefGap.List(I), CoefNumTmp)
                Y2 = GetFuncValByX(LstMainDefGap.List(I), CoefDenTmp)
                
                FrmMain.Circle (LstMainDefGap.List(I) + FrmMain.ScaleWidth / 2, FrmMain.ScaleHeight / 2 - Y1 / Y2), 0.1, RGB(255, 0, 0)
            End If
        Next I
        
        FrmMain.DrawStyle = 2
        
        ' Draw poles
        For I = 0 To LstMainPoles.ListCount - 1
            FrmMain.Line (LstMainPoles.List(I) + FrmMain.ScaleWidth / 2, 0)-(LstMainPoles.List(I) + FrmMain.ScaleWidth / 2, FrmMain.ScaleHeight), RGB(255, 0, 0)
            FrmMain.DrawStyle = 0
        Next I
    Else
        ' Draw nulls
        For I = 0 To LstNullsNumMulti.ListCount - 1
            X = CDbl(LstNullsNumMulti.List(I))
            Y1 = GetFuncValByX(X, CoefNum)
            FrmMain.Circle (X + FrmMain.ScaleWidth / 2, FrmMain.ScaleHeight / 2), 0.1, RGB(255, 0, 0)
        Next I
        
        ' Draw Hops and Tips
        For I = 0 To LstNullsFd.ListCount - 1
            X = CDbl(LstNullsFd.List(I))
            Y1 = GetFuncValByX(X, GetDiffFuncCoefs(GetDiffFuncCoefs(CoefNum)))
            ' Hops
            If Y1 < 0 Then
                Y1 = GetFuncValByX(X, CoefNum)
                FrmMain.Circle (X + FrmMain.ScaleWidth / 2, -Y1 + FrmMain.ScaleHeight / 2), 0.1, RGB(0, 255, 0)
            Else
                ' Tips
                If Y1 > 0 Then
                    Y1 = GetFuncValByX(X, CoefNum)
                    FrmMain.Circle (X + FrmMain.ScaleWidth / 2, -Y1 + FrmMain.ScaleHeight / 2), 0.1, RGB(0, 0, 255)
                End If
            End If
        Next I
        
        ' Draw Inflection Points and Saddle Points
        For I = 0 To LstNullsFdd.ListCount - 1
            X = CDbl(LstNullsFdd.List(I))
            Y1 = GetFuncValByX(X, GetDiffFuncCoefs(GetDiffFuncCoefs(GetDiffFuncCoefs(CoefNum))))
            If Y1 <> 0 Then
                Y1 = GetFuncValByX(X, CoefNum)
                If GetFuncValByX(X, GetDiffFuncCoefs(CoefNum)) <> 0 Then
                    ' Inflection Point
                    FrmMain.Circle (X + FrmMain.ScaleWidth / 2, -Y1 + FrmMain.ScaleHeight / 2), 0.1, RGB(255, 255, 0)
                Else
                    ' Saddle Point
                    FrmMain.Circle (X + FrmMain.ScaleWidth / 2, -Y1 + FrmMain.ScaleHeight / 2), 0.1, RGB(0, 255, 255)
                End If
            End If
        Next I
    End If
    
    Call SetDrawSettings(FrmMain, FormDrawSettings)
End Sub

Private Sub BtnCalcIntegral_Click()
    If Not IsRationalFunction Then  ' Integration gebrochen rationaler Funktionen ist viel komplizierter. Siehe z.B.: https://www.youtube.com/watch?v=AOaRHMoYaRw
        Dim I As Integer
        Dim IntLowerBound As Double, IntUpperBound As Double
        IntLowerBound = CDbl(TxtIntLowerBound.Text)
        IntUpperBound = CDbl(TxtIntUpperBound.Text)
        
        Dim IntValDiff As Double, IntValLowerBound As Double, IntValUpperBound As Double
        
        ' Integrate and calculate values at lower and upper bound
        IntValDiff = 0
        IntValLowerBound = 0
        IntValUpperBound = 0
        For I = 0 To DegNum
            IntValLowerBound = IntValLowerBound + CoefNum(I) / (I + 1) * IntLowerBound ^ (I + 1)
            IntValUpperBound = IntValUpperBound + CoefNum(I) / (I + 1) * IntUpperBound ^ (I + 1)
        Next I
        
        ' Calculate the integral itself (not piecewise calculated) and write the results to the text boxes
        IntValDiff = IntValUpperBound - IntValLowerBound
        TxtIntSum.Text = IntValDiff
        TxtIntAbs.Text = Abs(IntValLowerBound) + Abs(IntValUpperBound)
        
        ' Draw the integral
        Dim LinesPerUnit As Integer
        Dim Y As Double
        Dim IntegralRangeWidth As Double
        LinesPerUnit = 10
        IntegralRangeWidth = CDbl(TxtIntUpperBound.Text) - CDbl(TxtIntLowerBound.Text)
        
        For I = 0 To IntegralRangeWidth * LinesPerUnit + 1
            Y = 0
            
            ' Differenciate between all lines except the last one and the last one,
            ' as the last one is not necessarily at the same line-disctance like all previous ones
            If I < IntegralRangeWidth * LinesPerUnit Then
                For U = 0 To DegNum
                    Y = Y + CoefNum(U) * (IntLowerBound + I / LinesPerUnit) ^ U
                Next U
                FrmMain.Line (IntLowerBound + Me.ScaleWidth / 2 + I / LinesPerUnit, Me.ScaleHeight / 2 - Y)-(IntLowerBound + Me.ScaleWidth / 2 + I / LinesPerUnit, Me.ScaleHeight / 2), RGB(255, 0, 255)
            Else
                Y = 0
                For U = 0 To DegNum
                    Y = Y + CoefNum(U) * IntUpperBound ^ U
                Next U
                FrmMain.Line (IntUpperBound + Me.ScaleWidth / 2, Me.ScaleHeight / 2 - Y)-(IntUpperBound + Me.ScaleWidth / 2, Me.ScaleHeight / 2), RGB(255, 0, 255)
            End If
        Next I
    Else
        MsgBox "Calculation of integral not yet implemented for rational functions!"
    End If
End Sub

Private Sub BtnExtremum_Click()
    Dim I As Integer
    Dim Newton1 As NewtonResult
    Dim ZAbl1() As Double, ZAbl2() As Double
    
    ' Clear lists
    LstNullsFd.Clear
    LstNullsFdd.Clear
    LstHop.Clear
    LstTip.Clear
    
    ' Calculate 1st derivative and add nulls to list
    ZAbl1 = GetDiffFuncCoefs(CoefNum)
    Newton1 = Newton((ZAbl1), True, Precision)
    
    For I = 0 To Newton1.NullCnt - 1
        LstNullsFd.AddItem (Newton1.Nulls(I))
    Next I
    
    ' Calculate 2nd derivative and add nulls to list
    ZAbl2 = GetDiffFuncCoefs(ZAbl1)
    Newton1 = Newton((ZAbl2), True, Precision)
    
    For I = 0 To Newton1.NullCnt - 1
        LstNullsFdd.AddItem (Newton1.Nulls(I))
    Next I
    
    ' XXX
    For I = 0 To LstNullsFd.ListCount - 1
        If GetFuncValByX(LstNullsFd.List(I) + 10 ^ -5, CoefNum) < GetFuncValByX(LstNullsFd.List(I), CoefNum) Then
            LstHop.AddItem (LstNullsFd.List(I))
        Else
            LstTip.AddItem (LstNullsFd.List(I))
        End If
    Next I
End Sub

Private Sub BtnTrace_Click()
    If B = True Then
        B = False
        'TmrMouseCoordinates.Interval = 0
    Else
        B = True
        'TmrMouseCoordinates.Interval = 1
    End If
End Sub

Private Sub CmdDraw_Click()
    X = -100 '-1
    Draw (False)
End Sub

Private Sub CmdClear_Click()
    FrmMain.Cls
    
    Call Raster
    Call Nullpunkt
    Call Koordinaten
End Sub

Private Sub BtnCoefficients_Click()
    FrmCoefficients.Show (1)
End Sub

Private Sub ScalePlot()
    FrmMain.ScaleWidth = TxtUnitsWidth.Text
    If ChkProportional.Value = 0 Then
        FrmMain.ScaleHeight = TxtUnitsHeight.Text
    Else
        FrmMain.ScaleHeight = Int(TxtUnitsWidth.Text / STPX * (STPY) * 100) / 100
    End If
    
    Draw
End Sub

Private Sub BtnExit_Click()
    If MsgBox("Sind Sie sicher, dass Sie das Programm beenden möchten?", vbQuestion + vbYesNo + vbDefaultButton2, "Programm beenden") = vbYes Then End
End Sub

Private Sub BtnSaveCoefficients_Click()
    Dim I As Integer
    Dim Filename As String
    
    On Error GoTo ShowSaveError
    With Cdg1
        .Flags = .Flags Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist
        .CancelError = True
        .ShowSave
        Filename = .Filename
    End With
    On Error GoTo 0
    
    Dim FileNum
    Dim GRP1 As GRP
    Dim CoefficientsZ As String, CoefficientsN As String
    
    GRP1.ZCoefficients = ""
    GRP1.NCoefficients = ""
    CoefficientsZ = ""
    CoefficientsN = ""
    FileNum = FreeFile
    
    GRP1.GRF = ChkRationalFunction.Value
    GRP1.ZG = TxtDegreeNumerator.Text
    GRP1.NG = TxtDegreeDenominator.Text
    GRP1.DefL = TxtIntvLowerBound.Text
    GRP1.DefR = TxtIntvUpperBound.Text
    GRP1.IntL = TxtIntLowerBound.Text
    GRP1.IntR = TxtIntUpperBound.Text
    GRP1.Width = TxtLineWidth.Text
    GRP1.Color = PicColorMain.BackColor
    For I = 0 To TxtDegreeNumerator.Text
        CoefficientsZ = CoefficientsZ & ";" & Str(CoefNum(I))
    Next I
    If IsRationalFunction Then
        For I = 0 To TxtDegreeDenominator.Text
            CoefficientsN = CoefficientsN & ";" & Str(CoefDen(I))
        Next I
    End If
    GRP1.ZCoefficients = CoefficientsZ
    GRP1.NCoefficients = CoefficientsN
    
    Open Filename For Binary As FileNum
    'Print #FileNum, GRP1.GRF & ";" & GRP1.ZG & ";" & GRP1.NG & ";" & GRP1.DefL & ";" & GRP1.DefR & ";" & GRP1.IntL & ";" & GRP1.IntR & ";" & GRP1.Width & ";" & GRP1.Color & GRP1.ZCoefficients & GRP1.NCoefficients
    'Print #FileNum, GRP1.GRF & GRP1.ZG & GRP1.NG & GRP1.DefL & GRP1.DefR & GRP1.IntL & GRP1.IntR & GRP1.Width & GRP1.Color & GRP1.ZCoefficients & GRP1.NCoefficients
    Put #FileNum, , GRP1
    Close FileNum
    
    GRP1.ZCoefficients = ""
    GRP1.NCoefficients = ""
    CoefficientsZ = ""
    CoefficientsN = ""
    
ShowSaveError:
End Sub

'Private Sub BtnSaveCoefficients_Click()
'   SaveStringArray App.Path & "\Test.dat", CoefNum()
'End Sub
'
'Private Sub BtnLoadCoefficients_Click()
' ReadStringArray App.Path & "\Test.dat", CoefNum
'End Sub

Private Sub BtnLoadCoefficients_Click()
    Dim Filename As String
    
    On Error GoTo ShowOpenError
    With Cdg1
        .Filter = "Graphen (*.gps)|*.gps"
        .CancelError = True
        .ShowOpen
        Filename = .Filename
    End With
    On Error GoTo 0
    
    Dim FileNum, Länge, GRP1 As GRP
    FileNum = FreeFile
    Länge = Len(GRP1)
    Open Filename For Random As FileNum Len = Länge
    Get #FileNum, , GRP1

    ChkRationalFunction.Value = -CInt(GRP1.GRF)
    IsRationalFunction = GRP1.GRF
    TxtDegreeNumerator.Text = GRP1.ZG
    DegNum = GRP1.ZG
    DegNum = GRP1.ZG
    DegDen = GRP1.NG
    TxtDegreeDenominator.Text = GRP1.NG
    TxtIntvLowerBound.Text = GRP1.DefL
    TxtIntvUpperBound.Text = GRP1.DefR
    TxtIntLowerBound.Text = GRP1.IntL
    TxtLineWidth.Text = GRP1.Width
    TxtIntUpperBound.Text = GRP1.IntR
    PicColorMain.BackColor = GRP1.Color
    
    ReDim CoefNum(0 To GRP1.ZG)
    Dim Fields() As String
    Fields = Split(Mid(Trim(GRP1.ZCoefficients), 2), ";")
    Dim F As Integer
    For F = 0 To UBound(Fields)
        CoefNum(F) = CDbl(Fields(F))
    Next F
    
    If GRP1.GRF Then
        ReDim CoefDen(0 To GRP1.NG)
        Fields = Split(Mid(Trim(GRP1.NCoefficients), 2), ";")
        For F = 0 To UBound(Fields)
            CoefDen(F) = CDbl(Fields(F))
        Next F
    End If
   
    Close FileNum
    FrmMainMenu.Enabled = True
    
ShowOpenError:
End Sub

Private Sub BtnOffsetCoordSystem_Click()
    KSX = -TxtOffsetCoordSystemX.Text
    KSY = TxtOffsetCoordSystemY.Text
    Draw
End Sub


Private Sub ScrPrecision_Change()
    Precision = ScrPrecision.Value
    LblPrecision.Caption = Precision
    
    Call BtnNewton_Click
End Sub

Private Sub Form_DragDrop(source As Control, X As Single, Y As Single)
    Dim ScaleModeTmp
    Dim rect As POINTAPI
    GetCursorPos rect

    ScaleModeTmp = Me.ScaleMode
    Me.ScaleMode = vbPixels
    FrmControl.Move FrmControl.Left + rect.X - DragX, FrmControl.Top + rect.Y - DragY
    FrmControl.Visible = True
    Me.ScaleMode = ScaleModeTmp
End Sub


Private Sub ScrLineWidth_Change()
    TxtLineWidth.Text = -ScrLineWidth.Value + 11
End Sub

Private Sub Form_Activate()
    TxtUnitsWidth.Text = Int(Me.ScaleWidth * 100) / 100
    TxtUnitsHeight.Text = Int(Me.ScaleHeight * 100) / 100
    
    Dimension = 0
    ReDim C1(0 To 1000)
    ReDim D1(0 To 1000)
End Sub

Private Sub Form_Click()
'XXX
'On Error Resume Next
'If Check7.Value = 1 Then
'
'C1(Dimension) = LblMouseCoordsX.Caption
'D1(Dimension) = LblMouseCoordsY.Caption
'
'ReDim M(1 To Dimension + 2, 1 To Dimension + 1)
'
'For U = 1 To Dimension + 1
'For I = 1 To Dimension + 1
'M(Dimension + 2 - I, U) = C1(U - 1) ^ (I - 1)
'Next I
'Next U
'
'For I = 1 To Dimension + 1
'M(Dimension + 2, I) = D1(I - 1)
'Next I
'
'Dimension = Dimension + 1
'
'
'For I = 1 To Dimension + 1
'For U = I + 1 To Dimension
'Factor = -(M(I, U) / M(I, I)) '  -(CoefNum(U, I)
'For S = 1 To Dimension + 1 ' Eigentlich nicht 1 to sondern i to !
'M(S, U) = M(S, I) * Factor + M(S, U)
'Next S
'Next U
'Next I
'
'For U = 1 To Dimension
'Factor = M(U, U)
'For I = 1 To Dimension + 1
'M(I, U) = M(I, U) / Factor
'Next I
'Next U
'
'
'ReDim O(Dimension + 1, Dimension)
'For I = 1 To Dimension
'For U = 1 To Dimension
'O(Dimension + 1 - I, Dimension + 1 - U) = M(I, U)
'Next U
'Next I
'
'For U = 1 To Dimension
'O(Dimension + 1, Dimension + 1 - U) = M(Dimension + 1, U)
'Next U
'
'For I = 1 To Dimension + 1
'For U = 1 To Dimension
'M(I, U) = O(I, U)
'Next U
'Next I
'
'For I = 1 To Dimension '+ 1
'For U = I + 1 To Dimension
'Factor = -(M(I, U) / M(I, I)) '  -(m(U, I)
'For S = 1 To Dimension + 1 ' Eigentlich nicht 1 to sondern i to !
'M(S, U) = M(S, I) * Factor + M(S, U)
'Next S
'Next U
'Next I
'
'For I = 1 To Dimension
'O(Dimension + 1, I) = M(Dimension + 1, I)
'Next I
'
'For I = 1 To Dimension
'O(Dimension + 1, Dimension + 1 - I) = M(Dimension + 1, I)
'Next I
'
'TxtDegreeNumerator.Text = Dimension - 1
'ReDim CoefNum(0 To Dimension - 1)
'For I = 1 To Dimension '+ 1
''CoefNum(I - 1) = M(Dimension + 1, Dimension + 1 - I)
'CoefNum(I - 1) = M(Dimension + 1, I)
'Next I
'
'IsRationalFunction = False ' *** Für definitionslückenüberprüfung
'FrmMain.Cls
'Call Raster
'Call Koordinaten
'Call Nullpunkt
'Call Graph
'End If
End Sub


Private Sub Form_Load()
    Me.WindowState = vbMaximized
    Call ChkAlwaysInForeground_Click
    Call ScrPrecision_Change
    
    PicColorMain.Tag = "Collapsed"
    
    KSX = 0
    KSY = 0
    SFX = 1
    SFY = 1
    
    STPX = Screen.Width / Screen.TwipsPerPixelX
    STPY = Screen.Height / Screen.TwipsPerPixelY - 22
 
    Call Nullpunkt

    Y = FrmMain.ScaleHeight / 2
    X = 0

 
    B = False

    Call Raster
    Call Koordinaten

    CoefIncr = True

    If Command() <> "" Then
        Dim FileNum, Länge, GRP1 As GRP
        FileNum = FreeFile
        Länge = Len(GRP1)
        Open Mid(Command(), 2, Len(Command()) - 2) For Random As FileNum Len = Länge
        Get #FileNum, 1, GRP1
    
        ChkRationalFunction.Value = Trim(GRP1.GRF)
        IsRationalFunction = Int(Trim(GRP1.GRF))
        TxtDegreeNumerator.Text = Trim(GRP1.ZG)
        DegNum = Trim(GRP1.ZG)
        DegNum = Trim(GRP1.ZG)
        DegDen = Trim(GRP1.NG)
        TxtDegreeDenominator.Text = Trim(GRP1.NG)
        TxtIntvLowerBound.Text = Trim(GRP1.DefL)
        TxtIntvUpperBound.Text = Trim(GRP1.DefR)
        TxtIntLowerBound.Text = Trim(GRP1.IntL)
        TxtLineWidth.Text = Trim(GRP1.Width)
        TxtIntUpperBound.Text = Trim(GRP1.IntR)
        PicColorMain.BackColor = Trim(GRP1.Color)
    
        ' XXX Nachfolgender Block stammt 1:1 aus BtnLoadCoefficients_Click - in Function auslagern?
        Dim Fields() As String
        ReDim CoefNum(0 To CInt(GRP1.ZG))
        Fields = Split(Mid(Trim(GRP1.ZCoefficients), 2), ";")
        Dim F As Integer
        For F = 0 To UBound(Fields)
            CoefNum(F) = CDbl(Fields(F))
        Next F
   
        If GRP1.GRF = 1 Then
            ReDim CoefDen(0 To CInt(GRP1.NG))
            Fields = Split(Mid(Trim(GRP1.NCoefficients), 2), ";")
            For F = 0 To UBound(Fields)
                CoefDen(F) = CDbl(Fields(F))
            Next F
        End If
    
        FrmMainMenu.Enabled = True
    
    Else
        DegNum = 0
        DegDen = 0
        ReDim CoefNum(0 To DegNum)
        ReDim CoefDen(0 To DegDen)
        IsRationalFunction = False
    End If
    
    FrmColor.Tag = FrmColor.Width
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LblMouseCoordsX.Caption = Int((X - FrmMain.ScaleWidth / 2 + KSX) * 100) / 100
    LblMouseCoordsY.Caption = -Int((Y - FrmMain.ScaleHeight / 2 + KSY) * 100) / 100
    Dim Pt As POINTAPI

    Call GetCursorPos(Pt)
    'aX = Pt.X

    W = X - FrmMain.ScaleWidth / 2 + KSX

    If B = True Then
        If IsRationalFunction Then
'            DegNum = TxtDegreeNumerator.Text
            For G = 0 To DegNum
                aY = aY + CoefNum(G) * W ^ G
            Next G
            
'            DegNum = TxtDegreeDenominator.Text
            For G = 0 To DegDen 'DegNum
                aY2 = aY2 + CoefDen(G) * W ^ G
            Next G
            aY = aY / aY2
        Else
            'If Not IsRationalFunction Then
'            DegNum = TxtDegreeNumerator.Text
            For G = 0 To DegNum
                aY = aY + CoefNum(G) * W ^ G
            Next G
            'End If
        End If
        
        aY = -aY + FrmMain.ScaleHeight / 2
    End If

    If B = True Then
        Call SetCursorPos(X / FrmMain.ScaleWidth * STPX, (aY - KSY) / FrmMain.ScaleHeight * (STPY) + 20)
        Call GetCursorPos(Pt)
        'LblMouseCoordsX.Caption = Pt.X
        'LblMouseCoordsY.Caption = Pt.Y
        LblMouseCoordsX.Caption = Int((X - FrmMain.ScaleWidth / 2 + KSX) * 100) / 100
        LblMouseCoordsY.Caption = -Int((Y - FrmMain.ScaleHeight / 2 + KSY) * 100 + 1) / 100 '***
        aY = 0
        aY2 = 0
    End If
    DoEvents
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim I As Integer
    If Button = vbRightButton Then
        If FrmControl.Visible = False Then 'XXX
            FrmControl.Visible = True
        End If
    End If

    If Button = vbLeftButton Then
        If Check7.Value = 1 Then
            WXK = False
            'I = 0
            For I = 0 To UBound(C1) - LBound(C1)
                If LblMouseCoordsX.Caption = C1(I) Then WXK = True
            Next I
            
            If WXK = False Then
                If Dimension > 0 Then
                    If LblMouseCoordsX.Caption <> 0 Then
                        C1(Dimension) = C1(Dimension - 1)
                        D1(Dimension) = D1(Dimension - 1)
                        C1(Dimension - 1) = LblMouseCoordsX.Caption
                        D1(Dimension - 1) = LblMouseCoordsY.Caption
                    Else
                        C1(Dimension) = LblMouseCoordsX.Caption
                        D1(Dimension) = LblMouseCoordsY.Caption
                    End If
                Else
                    C1(Dimension) = LblMouseCoordsX.Caption
                    D1(Dimension) = LblMouseCoordsY.Caption
                End If
            
                'If WXK = False Then
                'LstNullsFd.AddItem (LblMouseCoordsX.Caption)
                'LstNullsFdd.AddItem (LblMouseCoordsY.Caption)
                'If LblMouseCoordsX.Caption <> 0 Then
                'If Dimension > 0 Then
                'C1(Dimension) = C1(Dimension - 1)
                'D1(Dimension) = D1(Dimension - 1)
                'C1(Dimension - 1) = LblMouseCoordsX.Caption
                'D1(Dimension - 1) = LblMouseCoordsY.Caption
                'Else
                'C1(Dimension) = LblMouseCoordsX.Caption
                'D1(Dimension) = LblMouseCoordsY.Caption
                'End If
                'Label33.Caption = D1(Dimension)
                'End If
                
                ReDim M(1 To Dimension + 2, 1 To Dimension + 1)
                
                For U = 1 To Dimension + 1
                    For I = 1 To Dimension + 1
                        M(Dimension + 2 - I, U) = C1(U - 1) ^ (I - 1)
                    Next I
                Next U
                
                For I = 1 To Dimension + 1
                    M(Dimension + 2, I) = D1(I - 1)
                Next I
                
                Dimension = Dimension + 1
                
                For I = 1 To Dimension + 1
                    For U = I + 1 To Dimension
                        Factor = -(M(I, U) / M(I, I)) '  -(CoefNum(U, I)
                        For S = 1 To Dimension + 1 ' Eigentlich nicht 1 to sondern i to !
                            'If M(I, U) = 0 Then
                            'Exit For
                            'Else
                            M(S, U) = M(S, I) * Factor + M(S, U)
                            'End If
                        Next S
                    Next U
                Next I
                
                For U = 1 To Dimension
                    Factor = M(U, U)
                    For I = 1 To Dimension + 1
                        M(I, U) = M(I, U) / Factor
                    Next I
                Next U
                
                
                ReDim O(Dimension + 1, Dimension)
                For I = 1 To Dimension
                    For U = 1 To Dimension
                        O(Dimension + 1 - I, Dimension + 1 - U) = M(I, U)
                    Next U
                Next I
                
                For U = 1 To Dimension
                    O(Dimension + 1, Dimension + 1 - U) = M(Dimension + 1, U)
                Next U
                
                For I = 1 To Dimension + 1
                    For U = 1 To Dimension
                        M(I, U) = O(I, U)
                    Next U
                Next I
                
                For I = 1 To Dimension '+ 1
                    For U = I + 1 To Dimension
                        Factor = -(M(I, U) / M(I, I)) '  -(m(U, I)
                        For S = 1 To Dimension + 1 ' Eigentlich nicht 1 to sondern i to !
                            M(S, U) = M(S, I) * Factor + M(S, U)
                        Next S
                    Next U
                Next I
                
                For I = 1 To Dimension
                    O(Dimension + 1, I) = M(Dimension + 1, I)
                Next I
                
                For I = 1 To Dimension
                    O(Dimension + 1, Dimension + 1 - I) = M(Dimension + 1, I)
                Next I
                
                TxtDegreeNumerator.Text = Dimension - 1
                ReDim CoefNum(0 To Dimension - 1)
                For I = 1 To Dimension '+ 1
                    'CoefNum(I - 1) = M(Dimension + 1, Dimension + 1 - I)
                    CoefNum(I - 1) = M(Dimension + 1, I)
                Next I
                
                IsRationalFunction = False ' *** Für definitionslückenüberprüfung
                Draw
            End If
        End If
    End If
End Sub

Private Sub Form_Resize()
    Draw
End Sub

Private Sub LblMoveMenu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim rect As POINTAPI
    GetCursorPos rect
    
    DragX = rect.X '(LblMoveMenu.Left + X) / Screen.TwipsPerPixelX / 1280 * FrmMain.ScaleWidth 'X - FrmControl.Left 'X - FrmControl.Left
    DragY = rect.Y '(LblMoveMenu.Top + Y) / Screen.TwipsPerPixelY / 1024 * FrmMain.ScaleHeight 'Y - FrmControl.Top 'Y - FrmControl.Top
    'MsgBox rect.X & ", " & rect.Y
    FrmControl.Visible = False
    FrmControl.Drag 1
End Sub

Private Sub ToggleOptNumDen()
    Dim Deg As Integer
    
    If OptNumerator.Value Then
        Deg = CInt(TxtDegreeNumerator.Text)
        
         ' XXX Dieser Block kommt so ähnlich mehrfach vor. Zusamenfassen und in Sub auslagern?
        If CInt(TxtDegToSetCoef.Text) > Deg Then TxtDegToSetCoef.Text = Deg
        Deg = CInt(TxtDegToSetCoef.Text)
        SldCoef.Value = CoefNum(Deg) * -100
        TxtSetCoefficient.Text = CoefNum(Deg)
    Else
        Deg = CInt(TxtDegreeDenominator.Text)
                
        If CInt(TxtDegToSetCoef.Text) > Deg Then TxtDegToSetCoef.Text = Deg
        Deg = CInt(TxtDegToSetCoef.Text)
        SldCoef.Value = CoefDen(Deg) * -100
        TxtSetCoefficient.Text = CoefDen(Deg)
    End If
End Sub


Private Sub OptNumerator_Click()
    Call ToggleOptNumDen
End Sub


Private Sub OptDenominator_Click()
    Call ToggleOptNumDen
End Sub

Private Sub PicColorPalette_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim I As Integer
    For I = 0 To PicColorPalette.Count - 1
        PicColorPalette(I).BorderStyle = 0
    Next I
    PicColorPalette(Index).BorderStyle = 1
End Sub

Private Sub PicColorPalette_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    PicColorMain.BackColor = PicColorPalette(Index).BackColor
    FrmColor.Width = FrmColor.Tag
    PicColorMain.Tag = "Collapsed"
End Sub

Private Sub PicColorMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FrmColor.Width = 148
    
    If PicColorMain.Tag = "Collapsed" Then
        PicColorMain.Tag = "Expanded"
        FrmColor.Width = 148
    Else
        PicColorMain.Tag = "Collapsed"
        FrmColor.Width = FrmColor.Tag
    End If
End Sub

Private Sub LblColorSelCustom_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ShowColorError
    
    With Cdg1
        .CancelError = True
        .ShowColor
        PicColorMain.BackColor = .Color
    End With
    On Error GoTo 0
    
    FrmColor.Width = FrmColor.Tag
    PicColorMain.Tag = "Collapsed"
    
    Dim I As Integer
    For I = 0 To PicColorPalette.Count - 1
        PicColorPalette(I).BorderStyle = 0
    Next I
    
ShowColorError:
End Sub

Private Sub SldCoef_Scroll()
    Dim GradToSetCoefficient As Integer
    GradToSetCoefficient = CInt(TxtDegToSetCoef.Text)
    
    TxtSetCoefficient.Text = -SldCoef.Value / 100
    
    If OptNumerator.Value = True Then
        CoefNum(GradToSetCoefficient) = -SldCoef.Value / 100
    Else
        CoefDen(GradToSetCoefficient) = -SldCoef.Value / 100
    End If
    
    Draw
End Sub


Private Sub TxtDegreeNumerator_GotFocus()
    TxtDegreeNumerator.SelStart = 0
    TxtDegreeNumerator.SelLength = Len(TxtDegreeNumerator.Text)
    
    TxtDegreeNumerator.Tag = TxtDegreeNumerator.Text
End Sub


Private Sub TxtDegreeNumerator_LostFocus()
    ' Reset to previous value if no valid number got entered
    If TxtDegreeNumerator.Tag <> "" Then
        TxtDegreeNumerator.Text = TxtDegreeNumerator.Tag
    End If
End Sub


Private Sub TxtDegreeNumerator_Validate(Cancel As Boolean)
    Cancel = True
    
    If IsNumeric(TxtDegreeNumerator.Text) Then
        If CInt(TxtDegreeNumerator.Text) = TxtDegreeNumerator.Text Then
            If CInt(TxtDegreeNumerator.Text) >= 0 Then
                Cancel = False
            End If
        End If
    End If
End Sub


Private Sub TxtDegreeNumerator_KeyPress(KeyAscii As Integer)
    ' The (valid) needs to be set by pressing Return
    If KeyAscii = vbKeyReturn Then
        Dim InputInvalid As Boolean
        Call TxtDegreeNumerator_Validate(InputInvalid)
        If Not InputInvalid Then
            DegNum = CInt(TxtDegreeNumerator.Text)
            ReDim Preserve CoefNum(0 To DegNum)
            
            If OptNumerator.Value And CInt(TxtDegToSetCoef.Text) > DegNum Then
                TxtDegToSetCoef.Text = DegNum
                TxtSetCoefficient.Text = CoefNum(DegNum)
                SldCoef.Value = CoefNum(DegNum) * -100
            End If
        
            TxtDegreeNumerator.Tag = ""
        End If
    End If
End Sub


Private Sub TxtDegreeDenominator_GotFocus()
    TxtDegreeDenominator.SelStart = 0
    TxtDegreeDenominator.SelLength = Len(TxtDegreeDenominator.Text)
    
    TxtDegreeDenominator.Tag = TxtDegreeDenominator.Text
End Sub


Private Sub TxtDegreeDenominator_LostFocus()
    ' Reset to previous value if no valid number got entered
    If TxtDegreeDenominator.Tag <> "" Then
        TxtDegreeDenominator.Text = TxtDegreeDenominator.Tag
    End If
End Sub


Private Sub TxtDegreeDenominator_Validate(Cancel As Boolean)
    Cancel = True
    
    If IsNumeric(TxtDegreeDenominator.Text) Then
        If CInt(TxtDegreeDenominator.Text) = TxtDegreeDenominator.Text Then
            If CInt(TxtDegreeDenominator.Text) >= 0 Then
                Cancel = False
            End If
        End If
    End If
End Sub


Private Sub TxtDegreeDenominator_KeyPress(KeyAscii As Integer)
    ' The (valid) needs to be set by pressing Return
    If KeyAscii = vbKeyReturn Then
        Dim InputInvalid As Boolean
        Call TxtDegreeDenominator_Validate(InputInvalid)
        If Not InputInvalid Then
            DegDen = CInt(TxtDegreeDenominator.Text)
            ReDim Preserve CoefDen(0 To DegDen)
            
            If Not OptNumerator.Value And CInt(TxtDegToSetCoef.Text) > DegDen Then
                TxtDegToSetCoef.Text = DegDen
                TxtSetCoefficient.Text = CoefDen(DegDen)
                SldCoef.Value = CoefDen(DegDen) * -100
            End If
            
            TxtDegreeDenominator.Tag = ""
        End If
    End If
End Sub




Private Sub TxtIntvLowerBound_GotFocus()
    TxtIntvLowerBound.SelStart = 0
    TxtIntvLowerBound.SelLength = Len(TxtIntvLowerBound.Text)

    TxtIntvLowerBound.Tag = TxtIntvLowerBound.Text
End Sub


Private Sub TxtIntvLowerBound_LostFocus()
    ' Reset to previous value if no valid number got entered
    If TxtIntvLowerBound.Tag <> "" Then
        TxtIntvLowerBound.Text = TxtIntvLowerBound.Tag
    End If
End Sub


Private Sub TxtIntvLowerBound_Validate(Cancel As Boolean)
    Cancel = True
    
    If IsNumeric(TxtIntvLowerBound.Text) Then
        If CDbl(TxtIntvLowerBound.Text) < CDbl(TxtIntvUpperBound.Text) Then
            Cancel = False
        End If
    End If
End Sub


Private Sub TxtIntvLowerBound_KeyPress(KeyAscii As Integer)
    ' The (valid) needs to be set by pressing Return
    If KeyAscii = vbKeyReturn Then
        Dim InputInvalid As Boolean
        Call TxtIntvLowerBound_Validate(InputInvalid)
        If Not InputInvalid Then
            TxtIntvLowerBound.Tag = ""
        End If
    End If
End Sub


Private Sub TxtIntvUpperBound_GotFocus()
    TxtIntvUpperBound.SelStart = 0
    TxtIntvUpperBound.SelLength = Len(TxtIntvUpperBound.Text)

    TxtIntvUpperBound.Tag = TxtIntvUpperBound.Text
End Sub


Private Sub TxtIntvUpperBound_LostFocus()
    ' Reset to previous value if no valid number got entered
    If TxtIntvUpperBound.Tag <> "" Then
        TxtIntvUpperBound.Text = TxtIntvUpperBound.Tag
    End If
End Sub


Private Sub TxtIntvUpperBound_Validate(Cancel As Boolean)
    Cancel = True
    
    If IsNumeric(TxtIntvUpperBound.Text) Then
        If CDbl(TxtIntvUpperBound.Text) > CDbl(TxtIntvLowerBound.Text) Then
            Cancel = False
        End If
    End If
End Sub


Private Sub TxtIntvUpperBound_KeyPress(KeyAscii As Integer)
    ' The (valid) needs to be set by pressing Return
    If KeyAscii = vbKeyReturn Then
        Dim InputInvalid As Boolean
        Call TxtIntvUpperBound_Validate(InputInvalid)
        If Not InputInvalid Then
            TxtIntvUpperBound.Tag = ""
        End If
    End If
End Sub



Private Function Koordinaten()
    If ChkAxisLabels.Value = 1 Then
        Dim I As Integer
        FrmMain.ForeColor = RGB(0, 100, 0)
      
        For I = -Int(FrmMain.ScaleWidth / 2 - KSX) To Int(FrmMain.ScaleWidth / 2 + KSX)
            FrmMain.CurrentY = FrmMain.ScaleHeight / 2 - KSY
            FrmMain.CurrentX = FrmMain.ScaleWidth / 2 - KSX + I
            FrmMain.Print Int(I)
        Next I
        
        For I = -Int(FrmMain.ScaleHeight / 2 - KSY) To Int(FrmMain.ScaleHeight / 2 + KSY)
            FrmMain.CurrentX = FrmMain.ScaleWidth / 2 - KSX
            FrmMain.CurrentY = FrmMain.ScaleHeight / 2 - KSY + I
            FrmMain.Print Int(-I)
        Next I
        
        FrmMain.ForeColor = RGB(0, 0, 255)
    End If
End Function

Private Function Raster()
    If ChkGrid.Value = 1 Then
        Dim I As Integer
        FrmMain.DrawWidth = 1
        FrmMain.DrawStyle = 2
        FrmMain.ForeColor = RGB(255, 0, 0)
         
        For I = Int(-1 / SFX) To Int(FrmMain.ScaleWidth / SFX) + 1
            FrmMain.Line ((FrmMain.ScaleWidth / 2 / SFX - (Int(FrmMain.ScaleWidth / 2 / SFX))) * SFX + I * SFX - (KSX - SFX * Int(KSX / SFX)), 0)-((FrmMain.ScaleWidth / 2 / SFX - (Int(FrmMain.ScaleWidth / 2 / SFX))) * SFX + I * SFX - (KSX - SFX * Int(KSX / SFX)), FrmMain.ScaleHeight) '(KSX - Int(KSX) --> beim Strecken anpassen
        Next I
        
        For I = Int(-1 / SFY) To Int(FrmMain.ScaleHeight / SFY) + 1
            FrmMain.Line (0, (FrmMain.ScaleHeight / 2 / SFY - (Int(FrmMain.ScaleHeight / 2 / SFY))) * SFY + I * SFY - (KSY - SFY * Int(KSY / SFY)))-(FrmMain.ScaleWidth, (FrmMain.ScaleHeight / 2 / SFY - (Int(FrmMain.ScaleHeight / 2 / SFY))) * SFY + I * SFY - (KSY - SFY * Int(KSY / SFY)))
        Next I
        
        FrmMain.ForeColor = RGB(0, 0, 255)
        FrmMain.DrawStyle = 0
    End If
End Function

Private Function Nullpunkt()
    If ChkAxes.Value = 1 Then
        FrmMain.DrawWidth = 3
        FrmMain.ForeColor = 0
        
        FrmMain.Line (FrmMain.ScaleWidth / 2 - KSX, 0)-(FrmMain.ScaleWidth / 2 - KSX, FrmMain.ScaleHeight)
        FrmMain.Line (0, FrmMain.ScaleHeight / 2 - KSY)-(FrmMain.ScaleWidth, FrmMain.ScaleHeight / 2 - KSY)
        
        FrmMain.DrawWidth = 1
        FrmMain.ForeColor = RGB(0, 0, 255)
    End If
End Function


Private Function GraphInternal(ByRef Func As RationalFunction)
    Dim I As Double
    Dim DivByZero As Boolean
    Dim DrawWidthOri As Integer
    DrawWidthOri = FrmMain.DrawWidth
    
    X = -100
    FrmMain.DrawWidth = TxtLineWidth.Text
    FrmMain.ForeColor = PicColorMain.BackColor
    
    For X1 = 1 + KSX * STPX / FrmMain.ScaleWidth To STPX + KSX * STPX / FrmMain.ScaleWidth
        ' Überprüfung auf Definitionslücke: wenn Nenner gleich 0 ist, wäre es Division durch 0 und daher eine Definitionslücke
        If Func.IsRational Then
            V = (X1 / STPX * FrmMain.ScaleWidth - FrmMain.ScaleWidth / 2)
            If GetFuncValByX(V, Func.CoefDen) = 0 Then DivByZero = True
        End If
        
        ' Only draw at the current X-position if there is no definition gap
        If Not DivByZero Then
            V = (X1 / STPX * FrmMain.ScaleWidth - FrmMain.ScaleWidth / 2)
            I = X1 / STPX * FrmMain.ScaleWidth
            Y1 = GetFuncValByX(V, Func.CoefNum)
            
            If Func.IsRational Then
                Y2 = GetFuncValByX(V, Func.CoefDen)
                Y1 = Y1 / Y2
            End If
            
            Y1 = FrmMain.ScaleHeight / 2 - Y1
            
            If Y < FrmMain.ScaleHeight + KSY + 100 Then ' +1
                If Y > -FrmMain.ScaleHeight / 2 + KSY - 100 Then '+ 1 Then
                    If FrmMain.ScaleWidth / 2 + TxtIntvLowerBound.Text < I Then
                        If FrmMain.ScaleWidth / 2 + TxtIntvUpperBound.Text > I Then
                            FrmMain.Line (X - KSX, Y - KSY)-(I - KSX, Y1 - KSY)
                        End If
                    End If
                End If
            End If
            
            Y = Y1
            X = (X1 - 0) / STPX * FrmMain.ScaleWidth
        End If
    Next X1
    
    FrmMain.DrawWidth = DrawWidthOri
End Function

Private Function Graph()
    Dim FuncBase As RationalFunction
    With FuncBase
        .IsRational = IsRationalFunction
        .DegNum = DegNum
        .DegDen = DegDen
        ReDim .CoefNum(0 To DegNum)
        .CoefNum = CoefNum
        ReDim .CoefDen(0 To DegDen)
        .CoefDen = CoefDen
    End With
    
    Call GraphInternal(FuncBase)
End Function


Private Sub TxtCalcFuncValueX_GotFocus()
    TxtCalcFuncValueX.SelStart = 0
    TxtCalcFuncValueX.SelLength = Len(TxtCalcFuncValueX.Text)
End Sub

Private Sub TxtIntLowerBound_GotFocus()
    TxtIntLowerBound.SelStart = 0
    TxtIntLowerBound.SelLength = Len(TxtIntLowerBound.Text)
End Sub


Private Sub TxtLineWidth_GotFocus()
    TxtLineWidth.SelStart = 0
    TxtLineWidth.SelLength = Len(TxtLineWidth.Text)
End Sub

Private Sub TxtLineWidth_LostFocus()
    TxtLineWidth.Text = Int(TxtLineWidth.Text)
'    FrmMain.DrawWidth = TxtLineWidth.Text
End Sub

Private Sub TxtLineWidth_Validate(Cancel As Boolean)
    Cancel = True
    
    If IsNumeric(TxtLineWidth.Text) Then
        If CInt(TxtLineWidth.Text) = TxtLineWidth.Text Then
            If CInt(TxtLineWidth.Text) > 0 Then
                Cancel = False
            End If
        End If
    End If
End Sub

Private Sub TxtIntUpperBound_GotFocus()
    TxtIntUpperBound.SelStart = 0
    TxtIntUpperBound.SelLength = Len(TxtIntUpperBound.Text)
End Sub

Private Sub TxtIntSum_GotFocus()
    TxtIntSum.SelStart = 0
    TxtIntSum.SelLength = Len(TxtIntSum.Text)
End Sub

Private Sub TxtIntAbs_GotFocus()
    TxtIntAbs.SelStart = 0
    TxtIntAbs.SelLength = Len(TxtIntAbs.Text)
End Sub



Private Sub TxtDegToSetCoef_GotFocus()
    TxtDegToSetCoef.SelStart = 0
    TxtDegToSetCoef.SelLength = Len(TxtDegToSetCoef.Text)

    TxtDegToSetCoef.Tag = TxtDegToSetCoef.Text
End Sub

Private Sub TxtDegToSetCoef_LostFocus()
    ' Reset to previous value if no valid number got entered
    If TxtDegToSetCoef.Tag <> "" Then
        TxtDegToSetCoef.Text = TxtDegToSetCoef.Tag
    End If
End Sub


Private Sub TxtDegToSetCoef_Validate(Cancel As Boolean)
    Cancel = True
    
    Dim Deg As Integer
    
    If OptNumerator.Value = True Then
        Deg = CInt(TxtDegreeNumerator.Text)
    Else
        Deg = CInt(TxtDegreeDenominator.Text)
    End If
    
    
    If IsNumeric(TxtDegToSetCoef.Text) Then
        If CInt(TxtDegToSetCoef.Text) = TxtDegToSetCoef.Text Then
            If CInt(TxtDegToSetCoef.Text) >= 0 And CInt(TxtDegToSetCoef.Text) <= Deg Then
                Cancel = False
            End If
        End If
    End If
End Sub


Private Sub TxtDegToSetCoef_KeyPress(KeyAscii As Integer)
    ' The (valid) needs to be set by pressing Return
    If KeyAscii = vbKeyReturn Then
        Dim InputInvalid As Boolean
        Call TxtDegToSetCoef_Validate(InputInvalid)
        If Not InputInvalid Then
            Dim Deg As Integer
            Deg = CInt(TxtDegToSetCoef.Text)
            TxtDegToSetCoef.Text = Deg
            If OptNumerator.Value = True Then
                TxtSetCoefficient.Text = CoefNum(Deg)
                SldCoef.Value = CoefNum(Deg) * -100 ' Invert because of the inverted direction of the slider
            Else
                TxtSetCoefficient.Text = CoefDen(Deg)
                SldCoef.Value = CoefDen(Deg) * -100
            End If
            TxtDegToSetCoef.Tag = ""
        End If
    End If
End Sub









Private Sub TxtSetCoefficient_GotFocus()
    TxtSetCoefficient.SelStart = 0
    TxtSetCoefficient.SelLength = Len(TxtSetCoefficient.Text)

    TxtSetCoefficient.Tag = TxtSetCoefficient.Text
End Sub

Private Sub TxtSetCoefficient_LostFocus()
    ' Reset to previous value if no valid number got entered
    If TxtSetCoefficient.Tag <> "" Then
        TxtSetCoefficient.Text = TxtSetCoefficient.Tag
    End If
End Sub


Private Sub TxtSetCoefficient_Validate(Cancel As Boolean)
    Cancel = True
    
'    If OptNumerator.Value = True Then
'        DegNum = TxtDegreeNumerator.Text
'    Else
'        DegNum = TxtDegreeDenominator.Text
'    End If
    
    If IsNumeric(TxtSetCoefficient.Text) Then
        Cancel = False
    End If
End Sub


Private Sub TxtSetCoefficient_KeyPress(KeyAscii As Integer)
    ' The (valid) needs to be set by pressing Return
    If KeyAscii = vbKeyReturn Then
        Dim InputInvalid As Boolean
        Call TxtSetCoefficient_Validate(InputInvalid)
        If Not InputInvalid Then
            SldCoef.Value = TxtSetCoefficient.Text * -100 ' evtl. mit Int()
        
            If OptNumerator.Value = True Then
                CoefNum(CInt(TxtDegToSetCoef.Text)) = CDbl(TxtSetCoefficient.Text)
            Else
                CoefDen(CInt(TxtDegToSetCoef.Text)) = CDbl(TxtSetCoefficient.Text)
            End If
            
            Draw

            TxtSetCoefficient.Tag = ""
        End If
    End If
End Sub


Private Sub TxtUnitsWidth_GotFocus()
    TxtUnitsWidth.SelStart = 0
    TxtUnitsWidth.SelLength = Len(TxtUnitsWidth.Text)
    
    TxtUnitsWidth.Tag = TxtUnitsWidth.Text
End Sub


Private Sub TxtUnitsWidth_LostFocus()
    ' Reset to previous value if no valid number got entered
    If TxtUnitsWidth.Tag <> "" Then
        TxtUnitsWidth.Text = TxtUnitsWidth.Tag
    End If
End Sub


Private Sub TxtUnitsWidth_Validate(Cancel As Boolean)
    ' Check if a valid number got entered
    Cancel = True
    If IsNumeric(TxtUnitsWidth.Text) Then
        If CDbl(TxtUnitsWidth.Text) > 0 Then
            Cancel = False
        End If
    End If
End Sub


Private Sub TxtUnitsWidth_KeyPress(KeyAscii As Integer)
    ' The (valid) value needs to be set by pressing Return
    If KeyAscii = vbKeyReturn Then
        Dim InputInvalid As Boolean
        Call TxtUnitsWidth_Validate(InputInvalid)
        If Not InputInvalid Then
            If ChkProportional.Value = 1 Then
                TxtUnitsHeight.Text = Int(TxtUnitsWidth.Text / STPX * (STPY) * 100) / 100
            End If
            Call ScalePlot
            TxtUnitsWidth.Tag = ""
        End If
    End If
End Sub


Private Sub TxtUnitsHeight_GotFocus()
    TxtUnitsHeight.SelStart = 0
    TxtUnitsHeight.SelLength = Len(TxtUnitsHeight.Text)
    
    TxtUnitsHeight.Tag = TxtUnitsHeight.Text
End Sub


Private Sub TxtUnitsHeight_LostFocus()
    ' Reset to previous value if no valid number got entered
    If TxtUnitsHeight.Tag <> "" Then
        TxtUnitsHeight.Text = TxtUnitsHeight.Tag
    End If
End Sub


Private Sub TxtUnitsHeight_Validate(Cancel As Boolean)
    ' Check if a valid number got entered
    Cancel = True
    If IsNumeric(TxtUnitsHeight.Text) Then
        If CDbl(TxtUnitsHeight.Text) > 0 Then
            Cancel = False
        End If
    End If
End Sub


Private Sub TxtUnitsHeight_KeyPress(KeyAscii As Integer)
    ' The (valid) value needs to be set by pressing Return
    If KeyAscii = vbKeyReturn Then
        Dim InputInvalid As Boolean
        Call TxtUnitsHeight_Validate(InputInvalid)
        If Not InputInvalid Then
            If ChkProportional.Value = 1 Then
                TxtUnitsWidth.Text = Int(TxtUnitsHeight.Text / STPY * STPX * 100) / 100
            End If
            Call ScalePlot
            TxtUnitsHeight.Tag = ""
        End If
    End If
End Sub


Private Sub TxtGridSpacingX_GotFocus()
    TxtGridSpacingX.SelStart = 0
    TxtGridSpacingX.SelLength = Len(TxtGridSpacingX.Text)
    
    TxtGridSpacingX.Tag = TxtGridSpacingX.Text
End Sub


Private Sub TxtGridSpacingX_LostFocus()
    ' Reset to previous value if no valid number got entered
    If TxtGridSpacingX.Tag <> "" Then
        TxtGridSpacingX.Text = TxtGridSpacingX.Tag
    End If
End Sub


Private Sub TxtGridSpacingX_Validate(Cancel As Boolean)
    ' Check if a valid number got entered
    Cancel = True
    If IsNumeric(TxtGridSpacingX.Text) Then
        If CDbl(TxtGridSpacingX.Text) > 0 Then
            Cancel = False
        End If
    End If
End Sub


Private Sub TxtGridSpacingX_KeyPress(KeyAscii As Integer)
    ' The (valid) value needs to be set by pressing Return
    If KeyAscii = vbKeyReturn Then
        Dim InputInvalid As Boolean
        Call TxtGridSpacingX_Validate(InputInvalid)
        If Not InputInvalid Then
            If ChkGridSpacingLock.Value = 1 Then
                TxtGridSpacingY.Text = TxtGridSpacingX.Text
            End If
            
            Call GridSpacing
            TxtGridSpacingX.Tag = ""
        End If
    End If
End Sub


Private Sub TxtGridSpacingY_GotFocus()
    TxtGridSpacingY.SelStart = 0
    TxtGridSpacingY.SelLength = Len(TxtGridSpacingY.Text)
    
    TxtGridSpacingY.Tag = TxtGridSpacingY.Text
End Sub


Private Sub TxtGridSpacingY_LostFocus()
    ' Reset to previous value if no valid number got entered
    If TxtGridSpacingY.Tag <> "" Then
        TxtGridSpacingY.Text = TxtGridSpacingY.Tag
    End If
End Sub


Private Sub TxtGridSpacingY_Validate(Cancel As Boolean)
    ' Check if a valid number got entered
    Cancel = True
    If IsNumeric(TxtGridSpacingY.Text) Then
        If CDbl(TxtGridSpacingY.Text) > 0 Then
            Cancel = False
        End If
    End If
End Sub


Private Sub TxtGridSpacingY_KeyPress(KeyAscii As Integer)
    ' The (valid) value needs to be set by pressing Return
    If KeyAscii = vbKeyReturn Then
        Dim InputInvalid As Boolean
        Call TxtGridSpacingY_Validate(InputInvalid)
        If Not InputInvalid Then
            If ChkGridSpacingLock.Value = 1 Then
                TxtGridSpacingX.Text = TxtGridSpacingY.Text
            End If
            
            Call GridSpacing
            TxtGridSpacingY.Tag = ""
        End If
    End If
End Sub


Private Sub TxtOffsetCoordSystemX_GotFocus()
    TxtOffsetCoordSystemX.SelStart = 0
    TxtOffsetCoordSystemX.SelLength = Len(TxtOffsetCoordSystemX.Text)
End Sub

Private Sub TxtOffsetCoordSystemY_GotFocus()
    TxtOffsetCoordSystemY.SelStart = 0
    TxtOffsetCoordSystemY.SelLength = Len(TxtOffsetCoordSystemY.Text)
End Sub

Private Sub TmrCoefIncrDecr_Timer()
    If TmrCoefIncrDecr.Interval = 250 Then TmrCoefIncrDecr.Interval = 50
    
    ' Only change slider value is coefficient not outside the sliders value range
    If CoefIncr = True Then
        If SldCoef.Value > -1000 Then
            SldCoef.Value = SldCoef.Value - 1
        End If
    Else
        If SldCoef.Value < 1000 Then
            SldCoef.Value = SldCoef.Value + 1
        End If
    End If
    
    Call SldCoef_Scroll
End Sub

Private Sub TmrMouseCoordinates_Timer()
    LblMouseCoordsX.Caption = Int((X - FrmMain.ScaleWidth / 2 + KSX) * 100) / 100
    LblMouseCoordsY.Caption = -Int((Y - FrmMain.ScaleHeight / 2 + KSY) * 100) / 100
End Sub
