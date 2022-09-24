VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
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
      Left            =   720
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Left            =   360
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox FrmControl 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   8415
      Left            =   600
      ScaleHeight     =   561
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   777
      TabIndex        =   0
      Top             =   2280
      Width           =   11655
      Begin VB.ListBox List16 
         Height          =   2400
         Left            =   10680
         TabIndex        =   122
         Top             =   5760
         Width           =   615
      End
      Begin VB.ListBox List15 
         Height          =   2400
         Left            =   10080
         TabIndex        =   121
         Top             =   5760
         Width           =   615
      End
      Begin VB.ListBox List14 
         Height          =   2400
         Left            =   9480
         TabIndex        =   120
         Top             =   5760
         Width           =   615
      End
      Begin VB.ListBox List13 
         Height          =   2400
         Left            =   8880
         TabIndex        =   119
         Top             =   5760
         Width           =   615
      End
      Begin VB.ListBox List8 
         Height          =   2400
         Left            =   8280
         TabIndex        =   118
         TabStop         =   0   'False
         Top             =   5760
         Width           =   615
      End
      Begin VB.ListBox List6 
         Height          =   2400
         Left            =   7680
         TabIndex        =   117
         TabStop         =   0   'False
         Top             =   5760
         Width           =   615
      End
      Begin VB.ListBox List5 
         Height          =   2400
         Left            =   7080
         TabIndex        =   116
         TabStop         =   0   'False
         Top             =   5760
         Width           =   615
      End
      Begin VB.ListBox List4 
         Height          =   2400
         Left            =   6480
         TabIndex        =   115
         TabStop         =   0   'False
         Top             =   5760
         Width           =   615
      End
      Begin VB.ListBox List3 
         Height          =   2400
         Left            =   5880
         TabIndex        =   114
         TabStop         =   0   'False
         Top             =   5760
         Width           =   615
      End
      Begin VB.ListBox List2 
         Height          =   2400
         Left            =   5280
         TabIndex        =   113
         TabStop         =   0   'False
         Top             =   5760
         Width           =   615
      End
      Begin VB.Frame FrmCalcFuncValue 
         BackColor       =   &H0080C0FF&
         Caption         =   "Funktionswert"
         Height          =   1095
         Left            =   2040
         TabIndex        =   107
         Top             =   5880
         Width           =   2175
         Begin VB.TextBox TxtCalcFuncValueY 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1140
            Locked          =   -1  'True
            TabIndex        =   112
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox TxtCalcFuncValueX 
            Height          =   285
            Left            =   360
            TabIndex        =   110
            Top             =   720
            Width           =   495
         End
         Begin VB.CommandButton BtnCalcFuncValue 
            BackColor       =   &H0080C0FF&
            Caption         =   "Funktionswert errechnen"
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   108
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   ") ="
            Height          =   255
            Left            =   900
            TabIndex        =   111
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "f ("
            Height          =   255
            Left            =   180
            TabIndex        =   109
            Top             =   720
            Width           =   255
         End
      End
      Begin VB.Frame FrmPlotSettings 
         BackColor       =   &H0080C0FF&
         Caption         =   "Plot-Settings"
         Height          =   1455
         Left            =   2040
         TabIndex        =   102
         Top             =   4320
         Width           =   2175
         Begin VB.CheckBox ChkAlwaysInForeground 
            BackColor       =   &H0080C0FF&
            Caption         =   "Immer im Vordergrund"
            Height          =   255
            Left            =   120
            TabIndex        =   106
            Top             =   1080
            Width           =   1935
         End
         Begin VB.CheckBox ChkAxisLabels 
            BackColor       =   &H0080C0FF&
            Caption         =   "Koordinaten"
            Height          =   255
            Left            =   120
            TabIndex        =   105
            Top             =   840
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CheckBox ChkAxes 
            BackColor       =   &H0080C0FF&
            Caption         =   "Achsenkreuz"
            Height          =   195
            Left            =   120
            TabIndex        =   104
            Top             =   600
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox ChkGrid 
            BackColor       =   &H0080C0FF&
            Caption         =   "Raster"
            Height          =   195
            Left            =   120
            TabIndex        =   103
            Top             =   360
            Value           =   1  'Checked
            Width           =   855
         End
      End
      Begin VB.Frame FmCalcIntegral 
         BackColor       =   &H0080C0FF&
         Caption         =   "Integral"
         Height          =   1935
         Left            =   9600
         TabIndex        =   91
         Top             =   360
         Width           =   1935
         Begin VB.TextBox TxtIntAbs 
            Height          =   375
            Left            =   1100
            TabIndex        =   100
            Top             =   1440
            Width           =   735
         End
         Begin VB.TextBox TxtIntSum 
            Height          =   375
            Left            =   120
            TabIndex        =   98
            Top             =   1440
            Width           =   735
         End
         Begin VB.TextBox TxtIntUpperBound 
            Height          =   285
            Left            =   1200
            TabIndex        =   96
            Text            =   "0"
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox TxtIntLowerBound 
            Height          =   285
            Left            =   1200
            TabIndex        =   94
            Text            =   "0"
            Top             =   360
            Width           =   615
         End
         Begin VB.CommandButton BtnCalcIntegral 
            BackColor       =   &H0080C0FF&
            Caption         =   "Integral errechnen"
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   92
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Betrag"
            Height          =   195
            Left            =   1080
            TabIndex        =   99
            Top             =   1200
            Width           =   465
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Summe"
            Height          =   195
            Left            =   120
            TabIndex        =   97
            Top             =   1200
            Width           =   525
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "b"
            Height          =   195
            Left            =   1080
            TabIndex        =   95
            Top             =   720
            Width           =   90
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "a"
            Height          =   195
            Left            =   1080
            TabIndex        =   93
            Top             =   375
            Width           =   90
         End
      End
      Begin VB.CheckBox ChkGridSpacingLock 
         BackColor       =   &H0080C0FF&
         Caption         =   "Lock"
         Height          =   375
         Left            =   2160
         TabIndex        =   90
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Frame FrmColor 
         BackColor       =   &H0080C0FF&
         Caption         =   "Color"
         Height          =   1095
         Left            =   600
         TabIndex        =   74
         Top             =   5280
         Width           =   675
         Begin VB.PictureBox PicColorMain 
            BackColor       =   &H00FF0000&
            Height          =   778
            Left            =   80
            ScaleHeight     =   0.5
            ScaleLeft       =   1
            ScaleMode       =   0  'User
            ScaleWidth      =   0.344
            TabIndex        =   89
            TabStop         =   0   'False
            Top             =   240
            Width           =   550
         End
         Begin VB.PictureBox PicColorSelArea 
            Height          =   778
            Left            =   670
            ScaleHeight     =   48
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   96
            TabIndex        =   75
            Top             =   240
            Width           =   1500
            Begin VB.PictureBox PicColorPalette 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   1
               Left            =   240
               ScaleHeight     =   16
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   16
               TabIndex        =   87
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
               TabIndex        =   86
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
               TabIndex        =   85
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
               TabIndex        =   84
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
               TabIndex        =   83
               TabStop         =   0   'False
               Top             =   0
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
               TabIndex        =   82
               TabStop         =   0   'False
               Top             =   0
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
               TabIndex        =   81
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
               TabIndex        =   80
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
               TabIndex        =   79
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
               TabIndex        =   78
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
               TabIndex        =   77
               TabStop         =   0   'False
               Top             =   240
               Width           =   240
            End
            Begin VB.PictureBox PicColorPalette 
               BackColor       =   &H00404040&
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   9
               Left            =   720
               ScaleHeight     =   16
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   16
               TabIndex        =   76
               TabStop         =   0   'False
               Top             =   240
               Width           =   240
            End
            Begin VB.Label LblColorSelCustom 
               Alignment       =   2  'Center
               BackColor       =   &H000735BC&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "SELECT  "
               Height          =   240
               Left            =   0
               TabIndex        =   88
               Top             =   480
               Width           =   1440
            End
         End
      End
      Begin VB.CommandButton Command23 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Extrema"
         Height          =   375
         Left            =   9600
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   4440
         Width           =   1815
      End
      Begin VB.CheckBox Check7 
         BackColor       =   &H0080C0FF&
         Caption         =   "Graph erklicken"
         Height          =   375
         Left            =   9600
         TabIndex        =   44
         Top             =   4800
         Width           =   1695
      End
      Begin VB.ListBox List1 
         Height          =   2400
         Left            =   4680
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   5760
         Width           =   615
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H0080C0FF&
         Caption         =   "Eigenschaften"
         Height          =   4935
         Left            =   4680
         TabIndex        =   66
         Top             =   240
         Width           =   4815
         Begin VB.TextBox Text17 
            Height          =   375
            Left            =   120
            TabIndex        =   36
            Top             =   960
            Width           =   4575
         End
         Begin VB.TextBox Text18 
            Height          =   375
            Left            =   120
            TabIndex        =   37
            Top             =   1560
            Width           =   4575
         End
         Begin VB.CommandButton BtnHornerSchema 
            BackColor       =   &H0080C0FF&
            Caption         =   "Horner Schema (L�cken, Pole und Nullstellen) + Linearfaktorzerlegung"
            Enabled         =   0   'False
            Height          =   615
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   240
            Width           =   2895
         End
         Begin VB.ListBox List12 
            Height          =   2400
            Left            =   3360
            TabIndex        =   42
            Top             =   2400
            Width           =   615
         End
         Begin VB.ListBox List11 
            Height          =   2400
            Left            =   2640
            TabIndex        =   41
            Top             =   2400
            Width           =   615
         End
         Begin VB.ListBox List10 
            Height          =   2400
            Left            =   1800
            TabIndex        =   40
            Top             =   2400
            Width           =   615
         End
         Begin VB.ListBox List9 
            Height          =   2400
            Left            =   1080
            TabIndex        =   39
            Top             =   2400
            Width           =   615
         End
         Begin VB.ListBox List7 
            Height          =   2400
            Left            =   120
            TabIndex        =   38
            Top             =   2400
            Width           =   615
         End
         Begin VB.CommandButton Command21 
            BackColor       =   &H0080C0FF&
            Caption         =   "Anzeigen"
            Enabled         =   0   'False
            Height          =   615
            Left            =   3120
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "NS       Vielfachheit   Pole          Ordnung"
            Height          =   255
            Left            =   1080
            TabIndex        =   68
            Top             =   2160
            Width           =   2895
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "Def.-l�cken"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   2160
            Width           =   975
         End
         Begin VB.Line Line1 
            X1              =   240
            X2              =   4440
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
         Top             =   3840
         Width           =   1815
      End
      Begin VB.CommandButton BtnIntegrate 
         BackColor       =   &H0080C0FF&
         Caption         =   "Integrieren"
         Height          =   375
         Left            =   9600
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   2640
         Width           =   1815
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
         TabIndex        =   64
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
         Top             =   3240
         Width           =   1815
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
         Caption         =   "Hauptmen�"
         Enabled         =   0   'False
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
      Begin VB.Frame Frame4 
         BackColor       =   &H0080C0FF&
         Caption         =   "Grad    Koeffizient"
         Height          =   975
         Left            =   600
         TabIndex        =   56
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
         Begin VB.TextBox TxtGradToSetCoefficient 
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
            TabIndex        =   57
            Top             =   280
            Width           =   255
         End
      End
      Begin VB.CommandButton Command16 
         Caption         =   "-"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   8040
         Width           =   375
      End
      Begin VB.CommandButton Command15 
         Caption         =   "+"
         Height          =   255
         Left            =   120
         TabIndex        =   53
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
            Caption         =   "H�heneinheiten"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Breiteneinheiten"
            Height          =   255
            Left            =   120
            TabIndex        =   46
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
         TabIndex        =   45
         Top             =   6480
         Width           =   1335
         Begin VB.CommandButton BtnCalcCodomain 
            BackColor       =   &H0080C0FF&
            Caption         =   "Wertebereich errechnen"
            Enabled         =   0   'False
            Height          =   495
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   101
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox TxtCodomainUpperBound 
            Enabled         =   0   'False
            Height          =   285
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   1440
            Width           =   495
         End
         Begin VB.TextBox TxtCodomainLowerBound 
            Enabled         =   0   'False
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
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Wertebereich:"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   55
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
      Begin ComctlLib.Slider Slider1 
         Height          =   7695
         Left            =   120
         TabIndex        =   52
         Top             =   360
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   13573
         _Version        =   327682
         BorderStyle     =   1
         Orientation     =   1
         Min             =   -10000
         Max             =   10000
         SelStart        =   -1000
         TickStyle       =   3
         Value           =   -1000
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Y:"
         Height          =   255
         Left            =   2120
         TabIndex        =   73
         Top             =   1120
         Width           =   135
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "Y:"
         Height          =   255
         Left            =   840
         TabIndex        =   72
         Top             =   4200
         Width           =   135
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tip"
         Height          =   195
         Left            =   10800
         TabIndex        =   71
         Top             =   5520
         Width           =   225
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hop"
         Height          =   195
         Left            =   10200
         TabIndex        =   70
         Top             =   5520
         Width           =   300
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Dicke"
         Height          =   255
         Left            =   1320
         TabIndex        =   65
         Top             =   5280
         Width           =   495
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Grad(N)"
         Enabled         =   0   'False
         Height          =   255
         Left            =   720
         TabIndex        =   63
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label LblMouseCoordsX 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   62
         Top             =   3960
         Width           =   615
      End
      Begin VB.Label LblMouseCoordsY 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   61
         Top             =   4200
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Grad(Z)"
         Height          =   255
         Left            =   720
         TabIndex        =   60
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "X:"
         Height          =   255
         Left            =   840
         TabIndex        =   59
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
         TabIndex        =   58
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label LblMoveMenu 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Men� verschieben"
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
         TabIndex        =   51
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "X:"
         Height          =   255
         Left            =   2120
         TabIndex        =   50
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
         TabIndex        =   49
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
         TabIndex        =   48
         Top             =   1800
         Width           =   1455
      End
   End
   Begin MSComDlg.CommonDialog Cdg1 
      Left            =   1200
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
Dim X, Y, X1, Y1, Y2, X2, I, V, G, B As Boolean, W, Faktor, KSX, KSY, SFX, SFY, STPX, STPY, MNS As Boolean, MENX, MENY, MCX, MCY, Plus As Boolean, C(), GradDiff, DragX, DragY, DiffZ, DiffN, DiffZA, DiffNA, E, DIFFNR, ASYM, Z, Grad1, DegDen, DegAsymptote As Integer, A1, A2, DefiL, WXK As Boolean, Element
Dim CoefAsymptote() As Double

Option Explicit


Dim aX, aY, aY2, dx, dy

Private Sub ChkGridSpacingLock_Click()
    If ChkGridSpacingLock.Value = 1 Then
        TxtGridSpacingY.Text = TxtGridSpacingX.Text
        Call GridSpacing
    End If
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
        'Form in den Normalzustand zur�cksetzen
        Call SetWindowPos(Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, 3)
    Else
        'Form dauerhaft in den Vordergrund setzen
        Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, 3)
    End If
End Sub

Private Sub ChkRationalFunction_Click()
    GRF = ChkRationalFunction.Value
    
    If ChkRationalFunction.Value = 1 Then
        IsNotRationalFunction = False
        OptDenominator.Enabled = True
        Label21.Enabled = True
        TxtDegreeDenominator.Enabled = True
        BtnAsymptote.Enabled = True
    Else
        IsNotRationalFunction = True
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
    MENX = FrmControl.Left
    MENY = FrmControl.Top
    FrmControl.Visible = False
End Sub

Private Sub BtnDifferentiate_Click()
    If IsNotRationalFunction = True Then
        DegNum = TxtDegreeNumerator.Text
        For I = 1 To DegNum
            CoefNum(I - 1) = CoefNum(I) * (I)
        Next I
        CoefNum(DegNum) = 0
    Else
        Call Graph2
    End If
End Sub

Private Sub BtnCalcCodomain_Click()
    If IsNotRationalFunction = True Then
        DegNum = TxtDegreeNumerator.Text
        For G = 0 To DegNum
            Y1 = Y1 + CoefNum(G) * TxtIntvLowerBound.Text ^ G
        Next G
        TxtCodomainLowerBound.Text = Y1
        Y1 = 0
        
        For G = 0 To DegNum
            Y1 = Y1 + CoefNum(G) * TxtIntvUpperBound.Text ^ G
        Next G
        TxtCodomainUpperBound.Text = Y1
        Y1 = 0
    Else
        DegNum = TxtDegreeNumerator.Text
        For G = 0 To DegNum
            Y1 = Y1 + CoefNum(G) * TxtIntvLowerBound.Text ^ G
        Next G
        
        DegNum = TxtDegreeDenominator.Text
        For G = 0 To DegNum
            Y2 = Y2 + CoefDen(G) * TxtIntvLowerBound.Text ^ G
        Next G
        TxtCodomainLowerBound.Text = Y1 / Y2
        
        Y1 = 0
        Y2 = 0
        
        DegNum = TxtDegreeNumerator.Text
        For G = 0 To DegNum
            Y1 = Y1 + CoefNum(G) * TxtIntvUpperBound.Text ^ G
        Next G
        
        DegNum = TxtDegreeDenominator.Text
        For G = 0 To DegNum
            Y2 = Y2 + CoefDen(G) * TxtIntvUpperBound.Text ^ G
        Next G
        TxtCodomainUpperBound.Text = Y1 / Y2
        
        Y1 = 0
        Y2 = 0
    End If
End Sub

Private Sub BtnCalcFuncValue_Click()
    DegNum = TxtDegreeNumerator.Text
    
    For G = 0 To DegNum
        Y1 = Y1 + CoefNum(G) * TxtCalcFuncValueX.Text ^ G
    Next G
    
    If IsNotRationalFunction = False Then
        DegNum = TxtDegreeDenominator.Text
        
        For G = 0 To DegNum
            Y2 = Y2 + CoefDen(G) * TxtCalcFuncValueX.Text ^ G
        Next G
        
        Y1 = Y1 / Y2
    End If
    
    TxtCalcFuncValueY.Text = Y1
    Y1 = 0
    Y2 = 0
End Sub


Private Sub Command15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Plus = True
    Timer2.Interval = 250
End Sub

Private Sub Command15_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer2.Interval = 0
End Sub

Private Sub Command16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Plus = False
Timer2.Interval = 250
End Sub

Private Sub Command16_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer2.Interval = 0
End Sub
'
'Private Sub BtnAsymptote_Click()
'If IsNotRationalFunction = False Then
'If DegNum > DegDen Then
'GradDiff = DegNum - DegDen
'ReDim C(GradDiff)
'For i = 1 To GradDiff 'DegDen
'CoefNum(i + DegDen) = CoefNum(i + DegDen) - CoefDen(i)
'Next i
'
'For i = 1 To DegDen
'CoefNum(i + GradDiff) = 0
'Next i
'End If
'End If
'End Sub


Private Sub BtnAsymptote_Click()
    DegNum = CDbl(TxtDegreeNumerator.Text)
    DegDen = CDbl(TxtDegreeDenominator.Text)
    
    If IsNotRationalFunction = False Then
        DegAsymptote = DegNum - DegDen
        
        ' If there is as asymptote
        If DegAsymptote >= 0 Then
            Dim CoefNumAsymptote() As Double, CoefDenAsymptote() As Double
            ReDim CoefNumAsymptote(DegNum + 1)
            ReDim CoefDenAsymptote(DegDen + 1)
            ReDim CoefAsymptote(DegAsymptote + 1)
            
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
            
            ' Draw the graph
            Call Graph3
        End If
    End If
End Sub

Private Sub BtnHornerSchema_Click()
    Dim Start, Ende, VZ
    'On Error Resume Next
    Grad1 = TxtDegreeNumerator.Text
    DegDen = TxtDegreeDenominator.Text
    'Call HornerSchema
    
    If IsNotRationalFunction = True Then
        Newton CoefNum, Grad1, True
    Else
        Newton CoefNum, Grad1, True
        Newton CoefDen, DegDen, False
    End If

    If IsNotRationalFunction = False Then
        '''''If Matrix2(0) < 0 Then
        '''''Text17.Text = "Y= -"
        '''''Else
        '''''Text17.Text = "Y="
        '''''End If
        '''''If Factor1 <> 1 Then Text17.Text = Text17.Text & Factor1 & "*("
        
        For I = 1 To Grad1
            If Newton1(I) < 0 Then
                VZ = "+"
            Else
                VZ = "-"
            End If
            Text17.Text = Text17.Text & " (x " + VZ + Str(Abs(Newton1(I))) + ")"
        Next I
        
        If Factor1 <> 1 Then Text17.Text = Text17.Text & " )"
        
        '''''If Matrix3(0) < 0 Then
        '''''Text18.Text = "Y= -"
        '''''Else
        '''''Text18.Text = "Y="
        '''''End If
        '''''If Factor2 <> 1 Then Text18.Text = Text18.Text & Factor2 & "*("
        
        For I = 1 To DegDen
            If Newton2(I) < 0 Then
                VZ = "+"
            Else
                VZ = "-"
            End If
            Text18.Text = Text18.Text & " (x " + VZ + Str(Abs(Newton2(I))) + ")"
        Next I
        '''''If Factor2 <> 1 Then Text18.Text = Text18.Text & " )"
        
        List1.Clear
        List2.Clear
        For I = 1 To Grad1
            If Newton1(I - 1) <> "" Then List1.AddItem (Newton1(I - 1))
        Next I
        For I = 1 To DegDen
            If Newton2(I - 1) <> "" Then List2.AddItem (Newton2(I - 1))
        Next I
        
        List3.Clear
        List4.Clear
        List3.List(0) = List1.List(0)
        List4.List(0) = 1
        For I = 1 To List1.ListCount - 1
            Element = False
            For U = 0 To List3.ListCount - 1
                If List3.List(U) = List1.List(I) Then
                List4.List(U) = List4.List(U) + 1
                Element = True
                Exit For
                End If
            Next U
            If Element = False Then List3.AddItem (List1.List(I)): List4.AddItem (1)
        Next I
            
        '''    For I = 1 To Grad1 - 1
        '''        If List1.List(I) <> List1.List(I - 1) Then
        '''        List3.AddItem (List1.List(I))
        '''        List4.AddItem (1)
        '''        Else
        '''        List4.List(List4.ListCount - 1) = List4.List(List4.ListCount - 1) + 1
        '''        End If
        '''    Next I
        List5.Clear
        List6.Clear
        List5.List(0) = List2.List(0)
        List6.List(0) = 1
        For I = 1 To DegDen - 1
            If List2.List(I) <> List2.List(I - 1) Then
                List5.AddItem (List2.List(I))
                List6.AddItem (1)
            Else
                List6.List(List6.ListCount - 1) = List6.List(List6.ListCount - 1) + 1
            End If
        Next I
        
        If List1.ListCount > List2.ListCount Then
            Ende = List1.ListCount
        Else
            Ende = List2.ListCount
        End If
        
        For N = 0 To List3.ListCount
            For I = 0 To Ende
                If List3.List(N) = List5.List(I) Then
                List7.AddItem (List3.List(N))
                    If List4.List(N) = List6.List(I) Then
                        List4.List(N) = "-"
                        List6.List(I) = "-"
                    ElseIf List4.List(N) > List6.List(I) Then
                        List4.List(N) = List4.List(N) - List6.List(I)
                        List6.List(I) = "-"
                    Else
                        List6.List(I) = List6.List(I) - List4.List(N)
                        List4.List(N) = "-"
                    End If
                End If
            Next I
        Next N
        
        For I = 0 To List3.ListCount - 1
            If List4.List(I) <> "-" Then List9.AddItem (List3.List(I)): List10.AddItem (List4.List(I))
        Next I
        
        For I = 0 To List5.ListCount - 1
            If List6.List(I) <> "-" Then List11.AddItem (List5.List(I)): List12.AddItem (List6.List(I))
        Next I
    Else
        '''''If Matrix2(0) < 0 Then
        '''''Text17.Text = "Y= -"
        '''''Else
        '''''Text17.Text = "Y="
        '''''End If
        '''''If Factor1 <> 1 Then Text17.Text = Text17.Text & Factor1 & "*("
    
        For I = 1 To Grad1
        
            If Newton1(I) < 0 Then
                VZ = "+"
            Else
                VZ = "-"
            End If
            Text17.Text = Text17.Text & " (x " + VZ + Str(Abs(Newton1(I))) + ")"
        Next I
    '''''If Factor1 <> 1 Then Text17.Text = Text17.Text & " )"
    
        List1.Clear
        Newton1(0) = Newton1(0)
        Newton1(1) = Newton1(1)
        Newton1(2) = Newton1(2)
        Newton1(3) = Newton1(3)
        For I = 1 To Grad1
            If Newton1(I - 1) <> "" Then List1.AddItem (Newton1(I - 1))
        Next I
        
        List3.Clear
        List4.Clear
        List3.List(0) = List1.List(0)
        List4.List(0) = 1
        
        For I = 1 To List1.ListCount - 1
            Element = False
            For U = 0 To List3.ListCount - 1
                If List3.List(U) = List1.List(I) Then
                    List4.List(U) = List4.List(U) + 1
                    Element = True
                    Exit For
                End If
            Next U
            If Element = False Then List3.AddItem (List1.List(I)): List4.AddItem (1)
        Next I
        
        '    For I = 1 To Grad1 - 1
        '    If List1.List(I) <> List1.List(I - 1) Then
        '    List3.AddItem (List1.List(I))
        '    List4.AddItem (1)
        '    Else
        '    List4.List(List4.ListCount - 1) = List4.List(List4.ListCount - 1) + 1
        '    End If
        '    Next I
    
        For I = 0 To List3.ListCount
            List9.AddItem (List3.List(I))
        Next I
        
        For I = 0 To List3.ListCount
            List10.AddItem (List4.List(I))
        Next I
    End If
End Sub

Private Sub BtnIntegrate_Click()
    'DegNum = TxtDegreeNumerator.Text
    'For i = 1 To DegNum
    '   CoefNum(i - 1) = CoefNum(i) * (i)
    'Next i
    'CoefNum(DegNum) = 0
    DegNum = TxtDegreeNumerator.Text
    ReDim C(DegNum + 1)
    For I = 0 To DegNum + 1
        C(I) = CoefNum(I)
    Next I
    ReDim CoefNum(DegNum + 2)
    CoefNum(0) = 0
    For I = 0 To DegNum + 1
        CoefNum(I + 1) = C(I) / (I + 1)
    Next I
    TxtDegreeNumerator.Text = TxtDegreeNumerator.Text + 1
    DegNum = TxtDegreeNumerator.Text

    X = -100
    FrmMain.DrawWidth = TxtLineWidth.Text
    FrmMain.ForeColor = PicColorMain.BackColor
    For X1 = 1 To STPX
        V = (X1 / STPX * FrmMain.ScaleWidth - FrmMain.ScaleWidth / 2)
        
        I = X1 / STPX * FrmMain.ScaleWidth
        
        'DegNum = TxtDegreeNumerator.Text
        
        For G = 0 To DegNum
            Y1 = Y1 + CoefNum(G) * V ^ G
        Next G
        
        If IsNotRationalFunction = False Then
            DegNum = TxtDegreeDenominator.Text
            
            For G = 0 To DegNum
                Y2 = Y2 + CoefDen(G) * V ^ G
            Next G
            
            Y1 = Y1 / Y2
        End If
        
        Y1 = FrmMain.ScaleHeight / 2 - Y1
        
        If Y < FrmMain.ScaleHeight + KSY + 100 Then ' +1
            If Y > -FrmMain.ScaleHeight / 2 + KSY + 1 Then
                If FrmMain.ScaleWidth / 2 + TxtIntvLowerBound.Text < I Then
                    If FrmMain.ScaleWidth / 2 + TxtIntvUpperBound.Text > I Then
                        FrmMain.Line (X - KSX, Y - KSY)-(I - KSX, Y1 - KSY)
                    End If
                End If
            End If
        End If
        
        Y = Y1
        X = (X1 - 0) / STPX * FrmMain.ScaleWidth
        Y1 = 0
        Y2 = 0
    Next X1
    FrmMain.DrawWidth = 1
End Sub


Private Sub Command21_Click()
    ' *** Das ganze '0.0001' kann wahrscheinlich weggelassen werden, da Definitionsl�cken ja jetzt �bersprungen werden
    FrmMain.DrawWidth = 3
    
    If IsNotRationalFunction = False Then
        For I = 0 To List7.ListCount - 1
            If List7.List(I) <> "" Then
                DegNum = TxtDegreeNumerator.Text
                For G = 0 To DegNum
                    Y1 = Y1 + CoefNum(G) * (Int(List7.List(I)) + 0.0001) ^ G
                Next G
                
                FrmMain.Circle (List7.List(I) + FrmMain.ScaleWidth / 2, FrmMain.ScaleHeight / 2 - Y1), 0.1, RGB(255, 0, 0)
                Y1 = 0
            Else
                DegNum = TxtDegreeNumerator.Text
                For G = 0 To DegNum
                    Y1 = Y1 + CoefNum(G) * (Int(List7.List(I)) + 0.0001) ^ G
                Next G
                
                DegNum = TxtDegreeDenominator.Text
                For G = 0 To DegNum
                    Y 2 = Y2 + CoefDen(G) * (Int(List7.List(I)) + 0.0001) ^ G
                Next G
                
                FrmMain.Circle (List7.List(I) + FrmMain.ScaleWidth / 2, FrmMain.ScaleHeight / 2 - Y1 / Y2), 0.1, RGB(255, 0, 0)
                Y1 = 0
                Y2 = 0
            End If
        Next I
        
        For I = 0 To List9.ListCount - 1
            If List9.List(I) <> "" Then
                DegNum = TxtDegreeNumerator.Text
                For G = 0 To DegNum
                    Y1 = Y1 + CoefNum(G) * (Int(List9.List(I)) + 0.0001) ^ G
                Next G
                
                FrmMain.Circle (List9.List(I) + FrmMain.ScaleWidth / 2, FrmMain.ScaleHeight / 2 - Y1), 0.1, RGB(255, 0, 0)
                Y1 = 0
            Else
                DegNum = TxtDegreeNumerator.Text
                For G = 0 To DegNum
                    Y1 = Y1 + CoefNum(G) * (Int(List9.List(I)) + 0.0001) ^ G
                Next G
                
                DegNum = TxtDegreeDenominator.Text
                For G = 0 To DegNum
                    Y2 = Y2 + CoefDen(G) * (Int(List9.List(I)) + 0.0001) ^ G
                Next G
                
                FrmMain.Circle (List9.List(I) + FrmMain.ScaleWidth / 2, FrmMain.ScaleHeight / 2 - Y1 / Y2), 0.1, RGB(255, 0, 0)
                Y1 = 0
                Y2 = 0
            End If
        Next I
        
        FrmMain.DrawStyle = 2
        
        For I = 0 To List11.ListCount - 1
            If List11.List(I) <> "" Then
                FrmMain.Line (List11.List(I) + FrmMain.ScaleWidth / 2, 0)-(List11.List(I) + FrmMain.ScaleWidth / 2, FrmMain.ScaleHeight), RGB(255, 0, 0)
                FrmMain.DrawStyle = 0
            End If
        Next I
    Else '****************************************
        For I = 0 To List3.ListCount - 1
            If List3.List(I) <> "" Then
                DegNum = TxtDegreeNumerator.Text
                For G = 0 To DegNum
                    Y1 = Y1 + CoefNum(G) * (Int(List3.List(I)) + 0.0001) ^ G
                Next G
                
                FrmMain.Circle (List3.List(I) + FrmMain.ScaleWidth / 2, FrmMain.ScaleHeight / 2 - Y1), 0.1, RGB(255, 0, 0)
                Y1 = 0
            End If
            
            Y1 = 0
            Y2 = 0
        Next I
    End If '****************************************
End Sub

Private Sub BtnCalcIntegral_Click()
    If IsNotRationalFunction = True Then  ' Integration gebrochen rationaler Funktionen ist viel komplizierter. Siehe z.B.: https://www.youtube.com/watch?v=AOaRHMoYaRw
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
        MsgBox "Not yet implemented for rational functions!"
    End If
End Sub

Private Sub Command23_Click()
    ReDim ZAbl1(0 To DegNum - 1)
    For I = 1 To UBound(CoefNum)
        ZAbl1(I - 1) = CoefNum(I) * I
    Next I
    
    ZAbl1(0) = ZAbl1(0)
    ZAbl1(1) = ZAbl1(1)
    ZAbl1(2) = ZAbl1(2)
    ZAbl1(3) = ZAbl1(3)
    
    ZAbl2(0) = ZAbl2(0)
    ZAbl2(1) = ZAbl2(1)
    ZAbl2(2) = ZAbl2(2)
    ZAbl2(3) = ZAbl2(3)
    '''For I = 1 To UBound(CoefDen)
    '''NAbl1(I - 1) = CoefDen(I) * I
    '''Next I
    
    
    '''For I = 1 To UBound(CoefDen)
    '''NAbl2(I - 1) = CoefDen(I) * I
    '''Next I
    
    Newton ZAbl1, Grad1 - 1, True
    
    For I = 0 To UBound(Newton1) - 1
        List13.AddItem (Newton1(I))
    Next I
    
    ReDim ZAbl2(0 To DegNum - 2)
    For I = 1 To UBound(ZAbl1)
        ZAbl2(I - 1) = ZAbl1(I) * I
    Next I
    
    Newton ZAbl2, Grad1 - 2, True
    
    For I = 0 To UBound(Newton1) - 1
        List14.AddItem (Newton1(I))
    Next I
    
    
    For I = 0 To List13.ListCount - 1
        If fv(List13.List(I) + 10 ^ -5, CoefNum, Grad1) < fv(List13.List(I), CoefNum, Grad1) Then
        List15.AddItem (List13.List(I))
        Else
        List16.AddItem (List13.List(I))
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
    Grad1 = TxtDegreeNumerator.Text
    DegDen = TxtDegreeDenominator.Text
    Frame4.Enabled = True
    FrmMainMenu.Enabled = True
    Label17.Enabled = True
    TxtCodomainLowerBound.Enabled = True
    TxtCodomainUpperBound.Enabled = True
    BtnSaveCoefficients.Enabled = True
    BtnDifferentiate.Enabled = True
    BtnCalcCodomain.Enabled = True
    Command15.Enabled = True
    Command16.Enabled = True
    BtnHornerSchema.Enabled = True
    Command21.Enabled = True
    Slider1.Enabled = True

    If TxtDegreeNumerator.Text < 0 Then TxtDegreeNumerator.Text = 0
    If TxtDegreeDenominator.Text < 0 Then TxtDegreeNumerator.Text = 0
    TxtDegreeNumerator.Text = Int(TxtDegreeNumerator.Text)
    TxtDegreeDenominator.Text = Int(TxtDegreeDenominator.Text)
    DegNum = TxtDegreeNumerator.Text
    If ChkRationalFunction.Value = 1 Then
        IsNotRationalFunction = False
    Else
        IsNotRationalFunction = True
    End If
    
    If ChkRationalFunction.Value = 1 Then DegDen = TxtDegreeDenominator.Text
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
    End
End Sub

Private Sub BtnSaveCoefficients_Click()
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
    If IsNotRationalFunction = False Then
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
    
    Dim FileNum, L�nge, GRP1 As GRP
    FileNum = FreeFile
    L�nge = Len(GRP1)
    Open Filename For Random As FileNum Len = L�nge
    Get #FileNum, , GRP1

    ChkRationalFunction.Value = -CInt(GRP1.GRF)
    IsNotRationalFunction = Not GRP1.GRF
    TxtDegreeNumerator.Text = GRP1.ZG
    DegNum = GRP1.ZG
    Grad1 = GRP1.ZG
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
'IsNotRationalFunction = True ' *** F�r definitionsl�cken�berpr�fung
'FrmMain.Cls
'Call Raster
'Call Koordinaten
'Call Nullpunkt
'Call Graph
'End If
End Sub


Private Sub Form_Load()
    Me.WindowState = vbMaximized
    ChkAlwaysInForeground_Click
    
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
  
    Faktor = Slider1.Value

    Call Raster
    Call Koordinaten

    Plus = True

    If Command() <> "" Then
        Dim FileNum, L�nge, GRP1 As GRP
        FileNum = FreeFile
        L�nge = Len(GRP1)
        Open Mid(Command(), 2, Len(Command()) - 2) For Random As FileNum Len = L�nge
        Get #FileNum, 1, GRP1
    
        ChkRationalFunction.Value = Trim(GRP1.GRF)
        IsNotRationalFunction = 1 - Int(Trim(GRP1.GRF))
        TxtDegreeNumerator.Text = Trim(GRP1.ZG)
        DegNum = Trim(GRP1.ZG)
        Grad1 = Trim(GRP1.ZG)
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
        If IsNotRationalFunction = False Then
            DegNum = TxtDegreeNumerator.Text
            For G = 0 To DegNum
                aY = aY + CoefNum(G) * W ^ G
            Next G
            
            DegNum = TxtDegreeDenominator.Text
            For G = 0 To DegNum
                aY2 = aY2 + CoefDen(G) * W ^ G
            Next G
            aY = aY / aY2
        Else
            'If IsNotRationalFunction = True Then
            DegNum = TxtDegreeNumerator.Text
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
    If Button = vbRightButton Then
        If FrmControl.Visible = False Then 'XXX
            FrmControl.Visible = True
        End If
    End If

    If Button = vbLeftButton Then
        If Check7.Value = 1 Then
            WXK = False 'Wiederholte X-Koordinate
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
                'List13.AddItem (LblMouseCoordsX.Caption)
                'List14.AddItem (LblMouseCoordsY.Caption)
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
                
                IsNotRationalFunction = True ' *** F�r definitionsl�cken�berpr�fung
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

Private Sub OptNumerator_Click()
    If OptNumerator.Value = True Then
        DegNum = TxtDegreeNumerator.Text
    Else
        DegNum = TxtDegreeDenominator.Text
    End If

    If TxtGradToSetCoefficient.Text <> "" Then
        If TxtGradToSetCoefficient.Text < 0 Then TxtGradToSetCoefficient.Text = 0
    End If

    If TxtGradToSetCoefficient.Text > DegNum Then TxtGradToSetCoefficient.Text = DegNum
    TxtGradToSetCoefficient.Text = Int(TxtGradToSetCoefficient.Text)
    Slider1.Value = CoefNum(TxtGradToSetCoefficient.Text) * -100
    TxtSetCoefficient.Text = CoefNum(TxtGradToSetCoefficient.Text)
End Sub

Private Sub PicColorPalette_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    For I = 0 To PicColorPalette.Count - 1
        PicColorPalette(I).BorderStyle = 0
    Next I
    PicColorPalette(Index).BorderStyle = 1
End Sub

Private Sub PicColorPalette_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    PicColorMain.BackColor = PicColorPalette(Index).BackColor
    FrmColor.Width = FrmColor.Tag
End Sub

Private Sub PicColorMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FrmColor.Width = 148
End Sub

Private Sub LblColorSelCustom_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Cdg1.ShowColor
    PicColorMain.BackColor = Cdg1.Color
    FrmColor.Width = FrmColor.Tag
    
    For I = 0 To PicColorPalette.Count - 1
        PicColorPalette(I).BorderStyle = 0
    Next I
End Sub

Private Sub Slider1_Scroll()
    If Faktor <> Slider1.Value Then
        If TxtGradToSetCoefficient.Text <> "" Then
            TxtSetCoefficient.Text = -Slider1.Value / 100
            FrmMain.Cls
            If OptNumerator.Value = True Then
                DegNum = TxtDegreeNumerator.Text
                CoefNum(TxtGradToSetCoefficient.Text) = -Slider1.Value / 100
                For G = 0 To DegNum
                    Y = Y + CoefNum(G) * (-FrmMain.ScaleWidth / 2) ^ G
                Next G
            Else
                DegNum = TxtDegreeDenominator.Text
                CoefDen(TxtGradToSetCoefficient.Text) = -Slider1.Value / 100
                For G = 0 To DegNum
                    Y = Y + CoefNum(G) * (-FrmMain.ScaleWidth / 2) ^ G
                Next G
            End If
            X = 0
            
            Draw
        End If
    End If
End Sub

Private Sub TxtDegreeNumerator_GotFocus()
    TxtDegreeNumerator.SelStart = 0
    TxtDegreeNumerator.SelLength = Len(TxtDegreeNumerator.Text)
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

Private Function Graph()
    If DegNum = 0 And DegDen = 0 Then Exit Function
    
    X = -100
    FrmMain.DrawWidth = TxtLineWidth.Text
    FrmMain.ForeColor = PicColorMain.BackColor
    
    For X1 = 1 + KSX * STPX / FrmMain.ScaleWidth To STPX + KSX * STPX / FrmMain.ScaleWidth
        If IsNotRationalFunction = False Then '***
            V = (X1 / STPX * FrmMain.ScaleWidth - FrmMain.ScaleWidth / 2) '***
            DefiL = 0 '***
            For G = 0 To DegDen '***
                DefiL = DefiL + CoefDen(G) * V ^ G '***
            Next G '*** �berpr�fung aud Definitionsl�cke
        Else '***
            DefiL = 1 '***
        End If '***
        
        If DefiL <> 0 Then '***
            V = (X1 / STPX * FrmMain.ScaleWidth - FrmMain.ScaleWidth / 2)
            
            I = X1 / STPX * FrmMain.ScaleWidth
            
            DegNum = TxtDegreeNumerator.Text
            
            For G = 0 To DegNum
                Y1 = Y1 + CoefNum(G) * V ^ G
            Next G
            
            If IsNotRationalFunction = False Then
                DegDen = TxtDegreeDenominator.Text
                
                For G = 0 To DegDen
                    Y2 = Y2 + CoefDen(G) * V ^ G
                Next G
                
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
            Y1 = 0
            Y2 = 0
        End If '***
    Next X1
    
    FrmMain.DrawWidth = 1
End Function

Private Sub TxtDegreeDenominator_GotFocus()
    TxtDegreeDenominator.SelStart = 0
    TxtDegreeDenominator.SelLength = Len(TxtDegreeDenominator.Text)
End Sub

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
    If TxtLineWidth.Text < 1 Then TxtLineWidth.Text = 1
    If TxtLineWidth.Text > 10 Then TxtLineWidth.Text = 10
    TxtLineWidth.Text = Int(TxtLineWidth.Text)
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



Private Sub TxtGradToSetCoefficient_GotFocus()
    TxtGradToSetCoefficient.SelStart = 0
    TxtGradToSetCoefficient.SelLength = Len(TxtGradToSetCoefficient.Text)

    TxtGradToSetCoefficient.Tag = TxtGradToSetCoefficient.Text
End Sub

Private Sub TxtGradToSetCoefficient_LostFocus()
    ' Reset to previous value if no valid number got entered
    If TxtGradToSetCoefficient.Tag <> "" Then
        TxtGradToSetCoefficient.Text = TxtGradToSetCoefficient.Tag
    End If
End Sub


Private Sub TxtGradToSetCoefficient_Validate(Cancel As Boolean)
    Cancel = True
    
    If OptNumerator.Value = True Then
        DegNum = TxtDegreeNumerator.Text
    Else
        DegNum = TxtDegreeDenominator.Text
    End If
    
    
    If IsNumeric(TxtGradToSetCoefficient.Text) Then
        If CInt(TxtGradToSetCoefficient.Text) = TxtGradToSetCoefficient.Text Then
            If CInt(TxtGradToSetCoefficient.Text) >= 0 And CInt(TxtGradToSetCoefficient.Text) <= DegNum Then
                Cancel = False
            End If
        End If
    End If
End Sub


Private Sub TxtGradToSetCoefficient_KeyPress(KeyAscii As Integer)
    ' The (valid) needs to be set by pressing Return
    If KeyAscii = vbKeyReturn Then
        Dim InputInvalid As Boolean
        Call TxtGradToSetCoefficient_Validate(InputInvalid)
        If Not InputInvalid Then
            DegNum = CInt(TxtGradToSetCoefficient.Text)
            TxtGradToSetCoefficient.Text = DegNum
            If OptNumerator.Value = True Then
                TxtSetCoefficient.Text = CoefNum(DegNum)
                Slider1.Value = CoefNum(TxtGradToSetCoefficient.Text) * -100 ' Invert because of the inverted direction of the slider
            Else
                TxtSetCoefficient.Text = CoefDen(DegNum)
                Slider1.Value = CoefDen(TxtGradToSetCoefficient.Text) * -100
            End If
            TxtGradToSetCoefficient.Tag = ""
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
    
    If OptNumerator.Value = True Then
        DegNum = TxtDegreeNumerator.Text
    Else
        DegNum = TxtDegreeDenominator.Text
    End If
    
    
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
            Slider1.Value = TxtSetCoefficient.Text * -100 ' evtl. mit Int()
        
            If OptNumerator.Value = True Then
                CoefNum(TxtGradToSetCoefficient.Text) = TxtSetCoefficient.Text
            Else
                CoefDen(TxtGradToSetCoefficient.Text) = TxtSetCoefficient.Text
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


' *** Bildschirmaufl�sung nur einmal am Anfang erechnen und als Konstante �bergeben --> schnelleres Zeichnen des Graphen
Private Sub Timer1_Timer()
    If Faktor <> Slider1.Value Then
        If TxtGradToSetCoefficient.Text <> "" Then
            FrmMain.Cls
            CoefNum(TxtGradToSetCoefficient.Text) = -Slider1.Value / 100
            For G = 0 To DegNum
                Y = Y + CoefNum(G) * (-FrmMain.ScaleWidth / 2) ^ G
            Next G
            
            X = 0
            
            Call Raster
            Call Nullpunkt
            Call Koordinaten
            Call Graph
        End If
    End If
End Sub

Private Sub Timer2_Timer()
    If Timer2.Interval = 250 Then Timer2.Interval = 50
    
    ' Only change slider value is coefficient not outside the sliders value range
    If Plus = True Then
        If Slider1.Value > -1000 Then
            Slider1.Value = Slider1.Value - 1
        End If
    Else
        If Slider1.Value < 1000 Then
            Slider1.Value = Slider1.Value + 1
        End If
    End If
    
    If Faktor <> Slider1.Value Then
        TxtSetCoefficient.Text = -Slider1.Value / 100
        FrmMain.Cls
        If OptNumerator.Value = True Then
            DegNum = TxtDegreeNumerator.Text
            CoefNum(TxtGradToSetCoefficient.Text) = -Slider1.Value / 100
            For G = 0 To DegNum
                Y = Y + CoefNum(G) * (-FrmMain.ScaleWidth / 2) ^ G
            Next G
        Else
            DegNum = TxtDegreeDenominator.Text
            CoefDen(TxtGradToSetCoefficient.Text) = -Slider1.Value / 100
            For G = 0 To DegNum
                Y = Y + CoefNum(G) * (-FrmMain.ScaleWidth / 2) ^ G
            Next G
        End If
        X = 0
        
        Draw
    
    End If
End Sub

Private Sub TmrMouseCoordinates_Timer()
    LblMouseCoordsX.Caption = Int((X - FrmMain.ScaleWidth / 2 + KSX) * 100) / 100
    LblMouseCoordsY.Caption = -Int((Y - FrmMain.ScaleHeight / 2 + KSY) * 100) / 100
End Sub

'Sub EnableControlsOn(Frm As Form, State As Integer)
'    Dim I   ' Variable deklarieren.
'    For I = 0 To Frm.Controls.Count - 1
'        If Not TypeOf Frm.Controls(I) Is Menu Then
'            Frm.Controls(I).Enabled = State
'        End If
'    Next I
'End Sub


'TxtDegreeNumerator.Text = CmdDraw.Width / Screen.TwipsPerPixelX / 1280 * 13.32


Private Function Graph2()
    X = -100
    FrmMain.DrawWidth = TxtLineWidth.Text ' mit Screen-Faktoren multiplizieren!
    FrmMain.ForeColor = PicColorMain.BackColor
    For X1 = 1 + KSX * STPX / FrmMain.ScaleWidth To STPX + KSX * STPX / FrmMain.ScaleWidth '
        V = (X1 / STPX * FrmMain.ScaleWidth - FrmMain.ScaleWidth / 2)
        'V = X1 / STPX * FrmMain.ScaleWidth
        E = X1 / STPX * FrmMain.ScaleWidth
        
        DegNum = TxtDegreeNumerator.Text
        DegDen = TxtDegreeDenominator.Text
        
        ReDim C(DegNum - DIFFNR)
        For G = 1 To DegNum - DIFFNR '+1
            C(G - 1) = CoefNum(G) * G
        Next G
        
        For G = 0 To DegNum - 1 - DIFFNR
            DiffNA = DiffNA + C(G) * V ^ G
        Next G
        For I = 0 To DegNum - DIFFNR
            C(DegNum - DIFFNR) = 0
        Next I
        
        ReDim C(DegDen)
        For G = 1 To DegDen - DIFFNR '+1
            C(G - 1) = CoefDen(G) * G
        Next G
        
        For G = 0 To DegDen - 1 - DIFFNR
            DiffZA = DiffZA + C(G) * V ^ G
        Next G
        For I = 0 To DegDen - DIFFNR
            C(DegDen - DIFFNR) = 0
        Next I
        
        For G = 0 To DegNum - DIFFNR
            DiffN = DiffN + CoefNum(G) * V ^ G
        Next G
        
        For G = 0 To DegDen - DIFFNR
            DiffZ = DiffZ + CoefDen(G) * V ^ G
        Next G
        
        Y1 = (DiffNA * DiffZ - DiffN * DiffZA) / (DiffZ ^ 2)
        
        Y1 = FrmMain.ScaleHeight / 2 - Y1
        
        If Y < FrmMain.ScaleHeight + KSY + 100 Then ' +1
            If Y > -FrmMain.ScaleHeight / 2 + KSY + 1 Then
                If FrmMain.ScaleWidth / 2 + TxtIntvLowerBound.Text < I Then
                    If FrmMain.ScaleWidth / 2 + TxtIntvUpperBound.Text > I Then
                        FrmMain.Line (X - KSX, Y - KSY)-(E - KSX, Y1 - KSY)
                    End If
                End If
            End If
        End If
        
        Y = Y1
        X = (X1 - 0) / STPX * FrmMain.ScaleWidth
        'X = (X1 / STPX * FrmMain.ScaleWidth - FrmMain.ScaleWidth / 2)
        Y1 = 0
        Y2 = 0
        DiffN = 0
        DiffNA = 0
        DiffZ = 0
        DiffZA = 0
    Next X1
    
    FrmMain.DrawWidth = 1
    
    DIFFNR = DIFFNR + 1
End Function


Private Function Graph3()
    X = -100
    FrmMain.DrawWidth = TxtLineWidth.Text ' mit Screen-Faktoren multiplizieren!
    FrmMain.ForeColor = PicColorMain.BackColor
    
    For X1 = 1 + KSX * STPX / FrmMain.ScaleWidth To STPX + KSX * STPX / FrmMain.ScaleWidth
        V = (X1 / STPX * FrmMain.ScaleWidth - FrmMain.ScaleWidth / 2)
        
        I = X1 / STPX * FrmMain.ScaleWidth
        
        For G = 0 To DegAsymptote
            Y1 = Y1 + CoefAsymptote(G) * V ^ G
        Next G
        
        Y1 = FrmMain.ScaleHeight / 2 - Y1
        
        If Y < FrmMain.ScaleHeight + KSY + 100 Then ' +1
            If Y > -FrmMain.ScaleHeight / 2 + KSY + 1 Then
                If FrmMain.ScaleWidth / 2 + TxtIntvLowerBound.Text < I Then
                    If FrmMain.ScaleWidth / 2 + TxtIntvUpperBound.Text > I Then
                        FrmMain.Line (X - KSX, Y - KSY)-(I - KSX, Y1 - KSY)
                    End If
                End If
            End If
        End If
        
        Y = Y1
        X = (X1 - 0) / STPX * FrmMain.ScaleWidth
        Y1 = 0
        Y2 = 0
    Next X1
    
    FrmMain.DrawWidth = 1
End Function
