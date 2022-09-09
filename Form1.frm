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
      Left            =   4320
      Top             =   2280
   End
   Begin VB.Timer Timer2 
      Left            =   3960
      Top             =   2280
   End
   Begin VB.Timer Timer1 
      Left            =   3600
      Top             =   2280
   End
   Begin VB.Frame FrmControl 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   9135
      Left            =   360
      TabIndex        =   0
      Top             =   3120
      Width           =   12135
      Begin VB.CheckBox ChkGridSpacingLock 
         BackColor       =   &H0080C0FF&
         Caption         =   "Lock"
         Height          =   375
         Left            =   2160
         TabIndex        =   118
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Frame FrmColor 
         BackColor       =   &H0080C0FF&
         Caption         =   "Color"
         Height          =   1109
         Left            =   600
         TabIndex        =   102
         Top             =   5280
         Width           =   681
         Begin VB.PictureBox PicColorMain 
            BackColor       =   &H00FF0000&
            Height          =   778
            Left            =   80
            ScaleHeight     =   0.5
            ScaleLeft       =   1
            ScaleMode       =   0  'User
            ScaleWidth      =   0.344
            TabIndex        =   117
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
            TabIndex        =   103
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
               TabIndex        =   115
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
               TabIndex        =   114
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
               TabIndex        =   113
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
               TabIndex        =   112
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
               TabIndex        =   111
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
               TabIndex        =   110
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
               TabIndex        =   109
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
               TabIndex        =   108
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
               TabIndex        =   107
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
               TabIndex        =   106
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
               TabIndex        =   105
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
               TabIndex        =   104
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
               TabIndex        =   116
               Top             =   480
               Width           =   1440
            End
         End
      End
      Begin VB.ListBox List16 
         Height          =   2400
         Left            =   10680
         TabIndex        =   97
         Top             =   6480
         Width           =   615
      End
      Begin VB.ListBox List15 
         Height          =   2400
         Left            =   10080
         TabIndex        =   96
         Top             =   6480
         Width           =   615
      End
      Begin VB.CommandButton Command23 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Extrema"
         Height          =   615
         Left            =   9840
         Style           =   1  'Graphical
         TabIndex        =   95
         Top             =   600
         Width           =   1695
      End
      Begin VB.ListBox List14 
         Height          =   2400
         Left            =   9480
         TabIndex        =   94
         Top             =   6480
         Width           =   615
      End
      Begin VB.ListBox List13 
         Height          =   2400
         Left            =   8880
         TabIndex        =   93
         Top             =   6480
         Width           =   615
      End
      Begin VB.TextBox Text23 
         Height          =   375
         Left            =   8640
         TabIndex        =   88
         Top             =   6120
         Width           =   735
      End
      Begin VB.TextBox Text22 
         Height          =   375
         Left            =   8640
         TabIndex        =   87
         Top             =   5400
         Width           =   735
      End
      Begin VB.CommandButton Command22 
         BackColor       =   &H0080C0FF&
         Caption         =   "Integral errechnen"
         Height          =   375
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   5280
         Width           =   1815
      End
      Begin VB.TextBox Text21 
         Height          =   285
         Left            =   7680
         TabIndex        =   85
         Text            =   "0"
         Top             =   5760
         Width           =   615
      End
      Begin VB.TextBox Text19 
         Height          =   285
         Left            =   6720
         TabIndex        =   84
         Text            =   "0"
         Top             =   5760
         Width           =   615
      End
      Begin VB.CheckBox Check7 
         BackColor       =   &H0080C0FF&
         Caption         =   "Graph erklicken"
         Height          =   375
         Left            =   4800
         TabIndex        =   58
         Top             =   5400
         Width           =   2055
      End
      Begin VB.ListBox List1 
         Height          =   2400
         Left            =   4680
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   6480
         Width           =   615
      End
      Begin VB.ListBox List2 
         Height          =   2400
         Left            =   5280
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   6480
         Width           =   615
      End
      Begin VB.ListBox List3 
         Height          =   2400
         Left            =   5880
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   6480
         Width           =   615
      End
      Begin VB.ListBox List4 
         Height          =   2400
         Left            =   6480
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   6480
         Width           =   615
      End
      Begin VB.ListBox List5 
         Height          =   2400
         Left            =   7080
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   6480
         Width           =   615
      End
      Begin VB.ListBox List6 
         Height          =   2400
         Left            =   7680
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   6480
         Width           =   615
      End
      Begin VB.ListBox List8 
         Height          =   2400
         Left            =   8280
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   6480
         Width           =   615
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H0080C0FF&
         Caption         =   "Eigenschaften"
         Height          =   4935
         Left            =   4680
         TabIndex        =   81
         Top             =   240
         Width           =   4815
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   480
            Top             =   1680
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            DefaultExt      =   "gps"
         End
         Begin VB.TextBox Text17 
            Height          =   375
            Left            =   120
            TabIndex        =   44
            Top             =   960
            Width           =   4575
         End
         Begin VB.TextBox Text18 
            Height          =   375
            Left            =   120
            TabIndex        =   45
            Top             =   1560
            Width           =   4575
         End
         Begin VB.CommandButton BtnHornerSchema 
            BackColor       =   &H0080C0FF&
            Caption         =   "Horner Schema (Lücken, Pole und Nullstellen) + Linearfaktorzerlegung"
            Enabled         =   0   'False
            Height          =   615
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   240
            Width           =   2895
         End
         Begin VB.ListBox List12 
            Height          =   2400
            Left            =   3360
            TabIndex        =   50
            Top             =   2400
            Width           =   615
         End
         Begin VB.ListBox List11 
            Height          =   2400
            Left            =   2640
            TabIndex        =   49
            Top             =   2400
            Width           =   615
         End
         Begin VB.ListBox List10 
            Height          =   2400
            Left            =   1800
            TabIndex        =   48
            Top             =   2400
            Width           =   615
         End
         Begin VB.ListBox List9 
            Height          =   2400
            Left            =   1080
            TabIndex        =   47
            Top             =   2400
            Width           =   615
         End
         Begin VB.ListBox List7 
            Height          =   2400
            Left            =   120
            TabIndex        =   46
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
            TabIndex        =   43
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "NS       Vielfachheit   Pole          Ordnung"
            Height          =   255
            Left            =   1080
            TabIndex        =   83
            Top             =   2160
            Width           =   2895
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "Def.-lücken"
            Height          =   255
            Left            =   120
            TabIndex        =   82
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
         Height          =   495
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   4680
         Width           =   2175
      End
      Begin VB.CommandButton BtnIntegrate 
         BackColor       =   &H0080C0FF&
         Caption         =   "Integrieren"
         Height          =   255
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   8400
         Width           =   1815
      End
      Begin VB.TextBox TxtCalcFuncValueX 
         Height          =   285
         Left            =   1920
         TabIndex        =   38
         Top             =   8400
         Width           =   495
      End
      Begin VB.CommandButton BtnCalcFuncValue 
         BackColor       =   &H0080C0FF&
         Caption         =   "Funktionswert errechnen"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   8400
         Width           =   1335
      End
      Begin VB.TextBox TxtCalcFuncValueY 
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   8640
         Width           =   495
      End
      Begin VB.TextBox TxtLineWidth 
         Height          =   285
         Left            =   600
         TabIndex        =   29
         Text            =   "1"
         Top             =   6840
         Width           =   615
      End
      Begin MSComCtl2.FlatScrollBar ScrLineWidth 
         Height          =   495
         Left            =   600
         TabIndex        =   78
         Top             =   7200
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
         Height          =   255
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   8760
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
      Begin VB.Frame Frame5 
         BackColor       =   &H0080C0FF&
         Caption         =   "Hauptmenü"
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
         TabIndex        =   70
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
            TabIndex        =   71
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.CommandButton Command16 
         Caption         =   "-"
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   8040
         Width           =   375
      End
      Begin VB.CommandButton Command15 
         Caption         =   "+"
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H0080C0FF&
         Caption         =   "Wertebereich errechnen"
         Enabled         =   0   'False
         Height          =   495
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   6480
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
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
            TabIndex        =   61
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Breiteneinheiten"
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.CommandButton BtnExit 
         BackColor       =   &H000080FF&
         Caption         =   "Beenden"
         Height          =   615
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   7080
         Width           =   1455
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
      Begin VB.CheckBox ChkGrid 
         BackColor       =   &H0080C0FF&
         Caption         =   "Raster"
         Height          =   195
         Left            =   1320
         TabIndex        =   25
         Top             =   5280
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.TextBox TxtGridSpacingY 
         Height          =   285
         Left            =   3720
         TabIndex        =   19
         Text            =   "1"
         Top             =   1800
         Width           =   735
      End
      Begin VB.CheckBox ChkAxisLabels 
         BackColor       =   &H0080C0FF&
         Caption         =   "Koordinaten"
         Height          =   495
         Left            =   2880
         TabIndex        =   27
         Top             =   5280
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox ChkAxes 
         BackColor       =   &H0080C0FF&
         Caption         =   "Achsenkreuz"
         Height          =   195
         Left            =   1320
         TabIndex        =   26
         Top             =   5640
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox ChkAlwaysInForeground 
         BackColor       =   &H0080C0FF&
         Caption         =   "Immer im Vordergrund"
         Height          =   495
         Left            =   2880
         TabIndex        =   28
         Top             =   5880
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Intervall"
         Height          =   1335
         Left            =   1200
         TabIndex        =   59
         Top             =   6480
         Width           =   1455
         Begin VB.TextBox Text13 
            Enabled         =   0   'False
            Height          =   285
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox Text6 
            Enabled         =   0   'False
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox Text11 
            Height          =   285
            Left            =   240
            TabIndex        =   30
            Text            =   "-1000"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox Text12 
            Height          =   285
            Left            =   840
            TabIndex        =   31
            Text            =   "1000"
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Wertebereich:"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   600
            Width           =   1215
         End
      End
      Begin VB.CommandButton BtnHide 
         BackColor       =   &H0000C0C0&
         Caption         =   "Ausblenden"
         Height          =   375
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   7920
         Width           =   3735
      End
      Begin ComctlLib.Slider Slider1 
         Height          =   7695
         Left            =   120
         TabIndex        =   66
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Y:"
         Height          =   255
         Left            =   2120
         TabIndex        =   101
         Top             =   1120
         Width           =   135
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "Y:"
         Height          =   255
         Left            =   840
         TabIndex        =   100
         Top             =   4200
         Width           =   135
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tip"
         Height          =   195
         Left            =   10680
         TabIndex        =   99
         Top             =   6240
         Width           =   225
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hop"
         Height          =   195
         Left            =   10080
         TabIndex        =   98
         Top             =   6240
         Width           =   300
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Betrag"
         Height          =   195
         Left            =   8640
         TabIndex        =   92
         Top             =   5880
         Width           =   465
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Summe"
         Height          =   195
         Left            =   8640
         TabIndex        =   91
         Top             =   5160
         Width           =   525
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "b"
         Height          =   195
         Left            =   7440
         TabIndex        =   90
         Top             =   5760
         Width           =   90
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "a"
         Height          =   195
         Left            =   6480
         TabIndex        =   89
         Top             =   5760
         Width           =   90
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "X= Y="
         Height          =   375
         Left            =   1560
         TabIndex        =   80
         Top             =   8520
         Width           =   255
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Dicke"
         Height          =   255
         Left            =   600
         TabIndex        =   79
         Top             =   6600
         Width           =   495
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Grad(N)"
         Enabled         =   0   'False
         Height          =   255
         Left            =   720
         TabIndex        =   77
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label LblMouseCoordsX 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   76
         Top             =   3960
         Width           =   615
      End
      Begin VB.Label LblMouseCoordsY 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   75
         Top             =   4200
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Grad(Z)"
         Height          =   255
         Left            =   720
         TabIndex        =   74
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "X:"
         Height          =   255
         Left            =   840
         TabIndex        =   73
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
         TabIndex        =   72
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
         Left            =   2400
         TabIndex        =   65
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "X:"
         Height          =   255
         Left            =   2120
         TabIndex        =   64
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
         TabIndex        =   63
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
         TabIndex        =   62
         Top             =   1800
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X, Y, X1, Y1, Y2, X2, I, V, G, B As Boolean, W, Faktor, KSX, KSY, SFX, SFY, STPX, STPY, MNS As Boolean, MENX, MENY, MCX, MCY, Plus As Boolean, C(), GradDiff, DragX, DragY, DiffZ, DiffN, DiffZA, DiffNA, E, DIFFNR, ASYM, Z, J, H, K(), L(), Faktor2, Grad1, Grad2, Grad3, A1, A2, DefiL, IntVal, IntVal1, IntVal2, KoefChange As Boolean, AuswahlNummer As Integer, CoefficientsZ As String, CoefficientsN As String, KoeffizientenNummer As Integer, EinlesePosition As Integer, WXK As Boolean, Element

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
        NV = True
        OptDenominator.Enabled = False
        Label21.Enabled = True
        TxtDegreeDenominator.Enabled = True
    Else
        NV = False
        OptDenominator.Enabled = False
        OptNumerator.Value = True
        Label21.Enabled = False
        TxtDegreeDenominator.Enabled = False
    End If
End Sub

Private Sub Check7_Click()
    Frame5.Enabled = True
    'If Check7.Value = 0 Then Grad = -1
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
    If NV = True Then
        Grad = TxtDegreeNumerator.Text
        For I = 1 To Grad
            A(I - 1) = A(I) * (I)
        Next I
        A(Grad) = 0
    Else
        Call Graph2
    End If
End Sub

Private Sub Command13_Click()
    If NV = True Then
        Grad = TxtDegreeNumerator.Text
        For G = 0 To Grad
            Y1 = Y1 + A(G) * Text11.Text ^ G
        Next G
        Text6.Text = Y1
        Y1 = 0
        
        For G = 0 To Grad
            Y1 = Y1 + A(G) * Text12.Text ^ G
        Next G
        Text13.Text = Y1
        Y1 = 0
    Else
        Grad = TxtDegreeNumerator.Text
        For G = 0 To Grad
            Y1 = Y1 + A(G) * Text11.Text ^ G
        Next G
        
        Grad = TxtDegreeDenominator.Text
        For G = 0 To Grad
            Y2 = Y2 + D(G) * Text11.Text ^ G
        Next G
        Text6.Text = Y1 / Y2
        
        Y1 = 0
        Y2 = 0
        
        Grad = TxtDegreeNumerator.Text
        For G = 0 To Grad
            Y1 = Y1 + A(G) * Text12.Text ^ G
        Next G
        
        Grad = TxtDegreeDenominator.Text
        For G = 0 To Grad
            Y2 = Y2 + D(G) * Text12.Text ^ G
        Next G
        Text13.Text = Y1 / Y2
        
        Y1 = 0
        Y2 = 0
    End If
End Sub

Private Sub BtnCalcFuncValue_Click()
    Grad = TxtDegreeNumerator.Text
    
    For G = 0 To Grad
        Y1 = Y1 + A(G) * TxtCalcFuncValueX.Text ^ G
    Next G
    
    If NV = False Then
        Grad = TxtDegreeDenominator.Text
        
        For G = 0 To Grad
            Y2 = Y2 + D(G) * TxtCalcFuncValueX.Text ^ G
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
On Error Resume Next
Plus = False
Timer2.Interval = 250
End Sub

Private Sub Command16_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer2.Interval = 0
End Sub
'
'Private Sub BtnAsymptote_Click()
'If NV = False Then
'If Grad > Grad2 Then
'GradDiff = Grad - Grad2
'ReDim C(GradDiff)
'For i = 1 To GradDiff 'Grad2
'A(i + Grad2) = A(i + Grad2) - D(i)
'Next i
'
'For i = 1 To Grad2
'A(i + GradDiff) = 0
'Next i
'End If
'End If
'End Sub


Private Sub BtnAsymptote_Click()
    Grad = TxtDegreeNumerator.Text
    Grad2 = TxtDegreeDenominator.Text
    
    If NV = False Then
        If Grad >= Grad2 Then
            Grad3 = Grad - Grad2
            ReDim K(Grad + 1)
            ReDim L(Grad2 + 1)
            
            For I = 0 To Grad
                K(I) = A(I)
            Next I
            
            For I = 0 To Grad2
                L(I) = D(I)
            Next I
            
            ReDim Z(Grad3 + 1) '1
            For J = 0 To Grad
                If Grad - J >= Grad2 + 0 Then
                    Faktor2 = K(Grad - J) / L(Grad2)
                    Z(Grad3 - J + 0) = Faktor2
                    For H = 0 To Grad2
                        K(Grad - J - H) = K(Grad - J - H) - Faktor2 * L(Grad2 - H)
                    Next H
                End If
            Next J
            Call Graph3
        End If
    End If
End Sub

Private Sub BtnHornerSchema_Click()
    Dim Start, Ende, VZ
    'On Error Resume Next
    Grad1 = TxtDegreeNumerator.Text
    Grad2 = TxtDegreeDenominator.Text
    'Call HornerSchema
    
    If NV = True Then
        Newton A, Grad1, True
    Else
        Newton A, Grad1, True
        Newton D, Grad2, False
    End If

    If NV = False Then
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
        
        For I = 1 To Grad2
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
        For I = 1 To Grad2
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
        For I = 1 To Grad2 - 1
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
    'Grad = TxtDegreeNumerator.Text
    'For i = 1 To Grad
    '   A(i - 1) = A(i) * (i)
    'Next i
    'A(Grad) = 0
    Grad = TxtDegreeNumerator.Text
    ReDim C(Grad + 1)
    For I = 0 To Grad + 1
        C(I) = A(I)
    Next I
    ReDim A(Grad + 2)
    A(0) = 0
    For I = 0 To Grad + 1
        A(I + 1) = C(I) / (I + 1)
    Next I
    TxtDegreeNumerator.Text = TxtDegreeNumerator.Text + 1
    Grad = TxtDegreeNumerator.Text

    X = -100
    FrmMain.DrawWidth = TxtLineWidth.Text
    FrmMain.ForeColor = PicColorMain.BackColor
    For X1 = 1 To STPX
        V = (X1 / STPX * FrmMain.ScaleWidth - FrmMain.ScaleWidth / 2)
        
        I = X1 / STPX * FrmMain.ScaleWidth
        
        'Grad = TxtDegreeNumerator.Text
        
        For G = 0 To Grad
            Y1 = Y1 + A(G) * V ^ G
        Next G
        
        If NV = False Then
            Grad = TxtDegreeDenominator.Text
            
            For G = 0 To Grad
                Y2 = Y2 + D(G) * V ^ G
            Next G
            
            Y1 = Y1 / Y2
        End If
        
        Y1 = FrmMain.ScaleHeight / 2 - Y1
        
        If Y < FrmMain.ScaleHeight + KSY + 100 Then ' +1
            If Y > -FrmMain.ScaleHeight / 2 + KSY + 1 Then
                If FrmMain.ScaleWidth / 2 + Text11.Text < I Then
                    If FrmMain.ScaleWidth / 2 + Text12.Text > I Then
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
    ' *** Das ganze '0.0001' kann wahrscheinlich weggelassen werden, da Definitionslücken ja jetzt übersprungen werden
    FrmMain.DrawWidth = 3
    
    If NV = False Then
        For I = 0 To List7.ListCount - 1
            If List7.List(I) <> "" Then
                Grad = TxtDegreeNumerator.Text
                For G = 0 To Grad
                    Y1 = Y1 + A(G) * (Int(List7.List(I)) + 0.0001) ^ G
                Next G
                
                FrmMain.Circle (List7.List(I) + FrmMain.ScaleWidth / 2, FrmMain.ScaleHeight / 2 - Y1), 0.1, RGB(255, 0, 0)
                Y1 = 0
            Else
                Grad = TxtDegreeNumerator.Text
                For G = 0 To Grad
                    Y1 = Y1 + A(G) * (Int(List7.List(I)) + 0.0001) ^ G
                Next G
                
                Grad = TxtDegreeDenominator.Text
                For G = 0 To Grad
                    Y 2 = Y2 + D(G) * (Int(List7.List(I)) + 0.0001) ^ G
                Next G
                
                FrmMain.Circle (List7.List(I) + FrmMain.ScaleWidth / 2, FrmMain.ScaleHeight / 2 - Y1 / Y2), 0.1, RGB(255, 0, 0)
                Y1 = 0
                Y2 = 0
            End If
        Next I
        
        For I = 0 To List9.ListCount - 1
            If List9.List(I) <> "" Then
                Grad = TxtDegreeNumerator.Text
                For G = 0 To Grad
                    Y1 = Y1 + A(G) * (Int(List9.List(I)) + 0.0001) ^ G
                Next G
                
                FrmMain.Circle (List9.List(I) + FrmMain.ScaleWidth / 2, FrmMain.ScaleHeight / 2 - Y1), 0.1, RGB(255, 0, 0)
                Y1 = 0
            Else
                Grad = TxtDegreeNumerator.Text
                For G = 0 To Grad
                    Y1 = Y1 + A(G) * (Int(List9.List(I)) + 0.0001) ^ G
                Next G
                
                Grad = TxtDegreeDenominator.Text
                For G = 0 To Grad
                    Y2 = Y2 + D(G) * (Int(List9.List(I)) + 0.0001) ^ G
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
                Grad = TxtDegreeNumerator.Text
                For G = 0 To Grad
                    Y1 = Y1 + A(G) * (Int(List3.List(I)) + 0.0001) ^ G
                Next G
                
                FrmMain.Circle (List3.List(I) + FrmMain.ScaleWidth / 2, FrmMain.ScaleHeight / 2 - Y1), 0.1, RGB(255, 0, 0)
                Y1 = 0
            End If
            
            Y1 = 0
            Y2 = 0
        Next I
    End If '****************************************
End Sub

Private Sub Command22_Click()
    If NV = True Then
        IntVal = 0
        IntVal1 = 0
        IntVal2 = 0
        For I = 0 To Grad
            'IntVal1 = IntVal1 + Text19.Text / (I + 1) * A(I) ^ (I + 1)
            'IntVal2 = IntVal2 + Text21.Text / (I + 1) * A(I) ^ (I + 1)
            IntVal1 = IntVal1 + A(I) / (I + 1) * Text19.Text ^ (I + 1)
            IntVal2 = IntVal2 + A(I) / (I + 1) * Text21.Text ^ (I + 1)
        Next I
        IntVal = IntVal2 - IntVal1
        Text22.Text = IntVal
        Text23.Text = Abs(IntVal1) + Abs(IntVal2)
        
        For I = 0 To Abs(Abs(Text19.Text) + Abs(Text21.Text)) * 10 + 1
            If I < Abs(Abs(Text19.Text) + Abs(Text21.Text)) * 10 Then
                Y1 = 0
                For U = 0 To Grad
                    Y1 = Y1 + A(U) * (Text19.Text + I * 0.1) ^ U
                Next U
                FrmMain.Line (Text19.Text + Me.ScaleWidth / 2 + I * 0.1, Me.ScaleHeight / 2 - Y1)-(Text19.Text + Me.ScaleWidth / 2 + I * 0.1, Me.ScaleHeight / 2), RGB(255, 0, 255)
            Else
                Y2 = 0
                For U = 0 To Grad
                    Y2 = Y2 + A(U) * (Text21.Text) ^ U
                Next U
                FrmMain.Line (Text21.Text + Me.ScaleWidth / 2, Me.ScaleHeight / 2 - Y2)-(Text21.Text + Me.ScaleWidth / 2, Me.ScaleHeight / 2), RGB(255, 0, 255)
            End If
        Next I
    End If
End Sub

Private Sub Command23_Click()
    ReDim ZAbl1(0 To Grad - 1)
    For I = 1 To UBound(A)
        ZAbl1(I - 1) = A(I) * I
    Next I
    
    ZAbl1(0) = ZAbl1(0)
    ZAbl1(1) = ZAbl1(1)
    ZAbl1(2) = ZAbl1(2)
    ZAbl1(3) = ZAbl1(3)
    
    ZAbl2(0) = ZAbl2(0)
    ZAbl2(1) = ZAbl2(1)
    ZAbl2(2) = ZAbl2(2)
    ZAbl2(3) = ZAbl2(3)
    '''For I = 1 To UBound(D)
    '''NAbl1(I - 1) = D(I) * I
    '''Next I
    
    
    '''For I = 1 To UBound(D)
    '''NAbl2(I - 1) = D(I) * I
    '''Next I
    
    Newton ZAbl1, Grad1 - 1, True
    
    For I = 0 To UBound(Newton1) - 1
        List13.AddItem (Newton1(I))
    Next I
    
    ReDim ZAbl2(0 To Grad - 2)
    For I = 1 To UBound(ZAbl1)
        ZAbl2(I - 1) = ZAbl1(I) * I
    Next I
    
    Newton ZAbl2, Grad1 - 2, True
    
    For I = 0 To UBound(Newton1) - 1
        List14.AddItem (Newton1(I))
    Next I
    
    
    For I = 0 To List13.ListCount - 1
        If fv(List13.List(I) + 10 ^ -5, A, Grad1) < fv(List13.List(I), A, Grad1) Then
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
    Draw
End Sub

Private Sub CmdClear_Click()
    FrmMain.Cls
    
    Call Raster
    Call Nullpunkt
    Call Koordinaten
End Sub

Private Sub Command3_Click()
    Grad1 = TxtDegreeNumerator.Text
    Grad2 = TxtDegreeDenominator.Text
    Frame4.Enabled = True
    Frame5.Enabled = True
    Label17.Enabled = True
    Text6.Enabled = True
    Text13.Enabled = True
    BtnSaveCoefficients.Enabled = True
    BtnDifferentiate.Enabled = True
    Command13.Enabled = True
    Command15.Enabled = True
    Command16.Enabled = True
    BtnHornerSchema.Enabled = True
    Command21.Enabled = True
    Slider1.Enabled = True

    If TxtDegreeNumerator.Text < 0 Then TxtDegreeNumerator.Text = 0
    If TxtDegreeDenominator.Text < 0 Then TxtDegreeNumerator.Text = 0
    TxtDegreeNumerator.Text = Int(TxtDegreeNumerator.Text)
    TxtDegreeDenominator.Text = Int(TxtDegreeDenominator.Text)
    Grad = TxtDegreeNumerator.Text
    If ChkRationalFunction.Value = 1 Then
        NV = False
    Else
        NV = True
    End If
    
    If ChkRationalFunction.Value = 1 Then Grad2 = TxtDegreeDenominator.Text
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
    ' cdlOFNOverwritePrompt
    ' cdlOFNPathMustExist
    '
    CommonDialog1.ShowSave
    
    Dim FileNum
    Dim GRP1 As GRP
    GRP1.ZCoefficients = ""
    GRP1.NCoefficients = ""
    CoefficientsZ = ""
    CoefficientsN = ""
    FileNum = FreeFile
    
    GRP1.GRF = ChkRationalFunction.Value
    GRP1.ZG = TxtDegreeNumerator.Text
    GRP1.NG = TxtDegreeDenominator.Text
    GRP1.DefL = Text11.Text
    GRP1.DefR = Text12.Text
    GRP1.IntL = Text19.Text
    GRP1.IntR = Text21.Text
    GRP1.Width = TxtLineWidth.Text
    GRP1.Color = PicColorMain.BackColor
    For I = 0 To TxtDegreeNumerator.Text
        CoefficientsZ = CoefficientsZ & ";" & Str(A(I))
    Next I
    If NV = False Then
        For I = 0 To TxtDegreeDenominator.Text
            CoefficientsN = CoefficientsN & ";" & Str(D(I))
        Next I
    End If
    GRP1.ZCoefficients = CoefficientsZ
    GRP1.NCoefficients = CoefficientsN
    
    Open CommonDialog1.Filename For Append As FileNum
    'Print #FileNum, GRP1.GRF & ";" & GRP1.ZG & ";" & GRP1.NG & ";" & GRP1.DefL & ";" & GRP1.DefR & ";" & GRP1.IntL & ";" & GRP1.IntR & ";" & GRP1.Width & ";" & GRP1.Color & GRP1.ZCoefficients & GRP1.NCoefficients
    Print #FileNum, GRP1.GRF & GRP1.ZG & GRP1.NG & GRP1.DefL & GRP1.DefR & GRP1.IntL & GRP1.IntR & GRP1.Width & GRP1.Color & GRP1.ZCoefficients & GRP1.NCoefficients
    Close FileNum
    SetAttr CommonDialog1.Filename, vbReadOnly
    
    GRP1.ZCoefficients = ""
    GRP1.NCoefficients = ""
    CoefficientsZ = ""
    CoefficientsN = ""
End Sub

'Private Sub BtnSaveCoefficients_Click()
'   SaveStringArray App.Path & "\Test.dat", A()
'End Sub
'
'Private Sub BtnLoadCoefficients_Click()
' ReadStringArray App.Path & "\Test.dat", A
'End Sub

Private Sub BtnLoadCoefficients_Click()
    CommonDialog1.Filter = "Graphen (*.gps)|*.gps"
    CommonDialog1.ShowOpen
    
    Dim FileNum, Länge, GRP1 As GRP
    FileNum = FreeFile
    Länge = Len(GRP1)
    Open CommonDialog1.Filename For Random As FileNum Len = Länge
    Get #FileNum, 1, GRP1

    ChkRationalFunction.Value = Trim(GRP1.GRF)
    NV = 1 - Int(Trim(GRP1.GRF))
    TxtDegreeNumerator.Text = Trim(GRP1.ZG)
    Grad = Trim(GRP1.ZG)
    Grad1 = Trim(GRP1.ZG)
    Grad2 = Trim(GRP1.NG)
    TxtDegreeDenominator.Text = Trim(GRP1.NG)
    Text11.Text = Trim(GRP1.DefL)
    Text12.Text = Trim(GRP1.DefR)
    Text19.Text = Trim(GRP1.IntL)
    TxtLineWidth.Text = Trim(GRP1.Width)
    Text21.Text = Trim(GRP1.IntR)
    PicColorMain.BackColor = Trim(GRP1.Color)
    
    ReDim A(0 To GRP1.ZG) ' KoeffizientenNummer, Einleseposition
    KoeffizientenNummer = 0
    For I = 2 To Len(Trim(GRP1.ZCoefficients))
        If Mid(Trim(GRP1.ZCoefficients), I, 1) = ";" Then
            A(KoeffizientenNummer) = Trim(A(KoeffizientenNummer))
            KoeffizientenNummer = KoeffizientenNummer + 1
        Else
            A(KoeffizientenNummer) = A(KoeffizientenNummer) & Mid(Trim(GRP1.ZCoefficients), I, 1)
            A(KoeffizientenNummer) = Trim(A(KoeffizientenNummer))
        End If
    Next I
    
    If GRP1.GRF = 1 Then
        ReDim D(0 To GRP1.NG)
        KoeffizientenNummer = 0
        For I = 2 To Len(Trim(GRP1.NCoefficients))
            If Mid(Trim(GRP1.NCoefficients), I, 1) = ";" Then
                KoeffizientenNummer = KoeffizientenNummer + 1
                D(KoeffizientenNummer) = Trim(D(KoeffizientenNummer))
            Else
                D(KoeffizientenNummer) = D(KoeffizientenNummer) & Mid(Trim(GRP1.NCoefficients), I, 1)
                D(KoeffizientenNummer) = Trim(D(KoeffizientenNummer))
            End If
        Next I
    End If
    
    Close FileNum
    Frame5.Enabled = True
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
'Factor = -(M(I, U) / M(I, I)) '  -(A(U, I)
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
'ReDim A(0 To Dimension - 1)
'For I = 1 To Dimension '+ 1
''A(I - 1) = M(Dimension + 1, Dimension + 1 - I)
'A(I - 1) = M(Dimension + 1, I)
'Next I
'
'NV = True ' *** Für definitionslückenüberprüfung
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
        Dim FileNum, Länge, GRP1 As GRP
        FileNum = FreeFile
        Länge = Len(GRP1)
        Open Mid(Command(), 2, Len(Command()) - 2) For Random As FileNum Len = Länge
        Get #FileNum, 1, GRP1
    
        ChkRationalFunction.Value = Trim(GRP1.GRF)
        NV = 1 - Int(Trim(GRP1.GRF))
        TxtDegreeNumerator.Text = Trim(GRP1.ZG)
        Grad = Trim(GRP1.ZG)
        Grad1 = Trim(GRP1.ZG)
        Grad2 = Trim(GRP1.NG)
        TxtDegreeDenominator.Text = Trim(GRP1.NG)
        Text11.Text = Trim(GRP1.DefL)
        Text12.Text = Trim(GRP1.DefR)
        Text19.Text = Trim(GRP1.IntL)
        TxtLineWidth.Text = Trim(GRP1.Width)
        Text21.Text = Trim(GRP1.IntR)
        PicColorMain.BackColor = Trim(GRP1.Color)
    
        ReDim A(0 To GRP1.ZG) ' KoeffizientenNummer, Einleseposition
        KoeffizientenNummer = 0
        For I = 2 To Len(Trim(GRP1.ZCoefficients))
            If Mid(Trim(GRP1.ZCoefficients), I, 1) = ";" Then
                A(KoeffizientenNummer) = Trim(A(KoeffizientenNummer))
                KoeffizientenNummer = KoeffizientenNummer + 1
            Else
                A(KoeffizientenNummer) = A(KoeffizientenNummer) & Mid(Trim(GRP1.ZCoefficients), I, 1)
                A(KoeffizientenNummer) = Trim(A(KoeffizientenNummer))
            End If
        Next I
    
    
        If GRP1.GRF = 1 Then
            ReDim D(0 To GRP1.NG)
            KoeffizientenNummer = 0
            For I = 2 To Len(Trim(GRP1.NCoefficients))
                If Mid(Trim(GRP1.NCoefficients), I, 1) = ";" Then
                    KoeffizientenNummer = KoeffizientenNummer + 1
                    D(KoeffizientenNummer) = Trim(D(KoeffizientenNummer))
                Else
                    D(KoeffizientenNummer) = D(KoeffizientenNummer) & Mid(Trim(GRP1.NCoefficients), I, 1)
                    D(KoeffizientenNummer) = Trim(D(KoeffizientenNummer))
                End If
            Next I
        End If
    
        Frame5.Enabled = True
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LblMouseCoordsX.Caption = Int((X - FrmMain.ScaleWidth / 2 + KSX) * 100) / 100
    LblMouseCoordsY.Caption = -Int((Y - FrmMain.ScaleHeight / 2 + KSY) * 100) / 100
    Dim Pt As POINTAPI

    Call GetCursorPos(Pt)
    'aX = Pt.X

    W = X - FrmMain.ScaleWidth / 2 + KSX

    If B = True Then
        If NV = False Then
            Grad = TxtDegreeNumerator.Text
            For G = 0 To Grad
                aY = aY + A(G) * W ^ G
            Next G
            
            Grad = TxtDegreeDenominator.Text
            For G = 0 To Grad
                aY2 = aY2 + D(G) * W ^ G
            Next G
            aY = aY / aY2
        Else
            'If NV = True Then
            Grad = TxtDegreeNumerator.Text
            For G = 0 To Grad
                aY = aY + A(G) * W ^ G
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
                        Factor = -(M(I, U) / M(I, I)) '  -(A(U, I)
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
                ReDim A(0 To Dimension - 1)
                For I = 1 To Dimension '+ 1
                    'A(I - 1) = M(Dimension + 1, Dimension + 1 - I)
                    A(I - 1) = M(Dimension + 1, I)
                Next I
                
                NV = True ' *** Für definitionslückenüberprüfung
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

Private Sub LblColorSelCustom_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CommonDialog1.ShowColor
    PicColorMain.BackColor = CommonDialog1.Color
    FrmColor.Width = (PicColorMain.ScaleWidth) / FrmMain.ScaleWidth * Screen.TwipsPerPixelX * 1280
End Sub

'Private Sub LblMoveMenu_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) XXX
'FrmControl.Drag 2
''FrmControl.Left = FrmControl.Left + (X - DragX)
''FrmControl.Top = FrmControl.Top + (Y - DragY)
'FrmControl.Move X - DragX, Y - DragY
'End Sub

Private Sub OptNumerator_Click()
    If OptNumerator.Value = True Then
        Grad = TxtDegreeNumerator.Text
    Else
        Grad = TxtDegreeDenominator.Text
    End If

    If TxtGradToSetCoefficient.Text <> "" Then
        If TxtGradToSetCoefficient.Text < 0 Then TxtGradToSetCoefficient.Text = 0
    End If

    If TxtGradToSetCoefficient.Text > Grad Then TxtGradToSetCoefficient.Text = Grad
    TxtGradToSetCoefficient.Text = Int(TxtGradToSetCoefficient.Text)
    Slider1.Value = A(TxtGradToSetCoefficient.Text) * -100
    TxtSetCoefficient.Text = A(TxtGradToSetCoefficient.Text)
End Sub

Private Sub PicColorPalette_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    For I = 0 To PicColorPalette.Count - 1
        PicColorPalette(I).BorderStyle = 0
    Next I
    PicColorPalette(Index).BorderStyle = 1
End Sub

Private Sub PicColorPalette_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'For i = 0 To PicColorPalette.Count - 1
'PicColorPalette(i).BorderStyle = 0
'Next i
'PicColorPalette(Index).BorderStyle = 1
End Sub

Private Sub PicColorPalette_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    PicColorMain.BackColor = PicColorPalette(Index).BackColor
    FrmColor.Width = 681 ' (PicColorMain.ScaleWidth) / 13.32 * Screen.TwipsPerPixelX * 1280
    'For i = 0 To PicColorPalette.Count - 1
    'PicColorPalette(i).BorderStyle = 0
    'Next i
    'PicColorPalette(Index).BorderStyle = 1
End Sub

Private Sub PicColorMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FrmColor.Width = 1.56 * 1000
End Sub

Private Sub Slider1_Scroll()
    If KoefChange = False Then
        If TxtGradToSetCoefficient.Text > Grad Then TxtGradToSetCoefficient.Text = Grad '###
        
        If Faktor <> Slider1.Value Then
            If TxtGradToSetCoefficient.Text <> "" Then
                TxtSetCoefficient.Text = -Slider1.Value / 100
                FrmMain.Cls
                If OptNumerator.Value = True Then
                    Grad = TxtDegreeNumerator.Text
                    A(TxtGradToSetCoefficient.Text) = -Slider1.Value / 100
                    For G = 0 To Grad
                        Y = Y + A(G) * (-FrmMain.ScaleWidth / 2) ^ G
                    Next G
                Else
                    Grad = TxtDegreeDenominator.Text
                    D(TxtGradToSetCoefficient.Text) = -Slider1.Value / 100
                    For G = 0 To Grad
                        Y = Y + A(G) * (-FrmMain.ScaleWidth / 2) ^ G
                    Next G
                End If
                X = 0
                
                Draw
            End If
        End If
    End If
End Sub

Private Sub TxtDegreeNumerator_GotFocus()
    TxtDegreeNumerator.SelStart = 0
    TxtDegreeNumerator.SelLength = Len(TxtDegreeNumerator.Text)
End Sub

Private Sub Text11_GotFocus()
    Text11.SelStart = 0
    Text11.SelLength = Len(Text11.Text)
End Sub

Private Sub Text12_GotFocus()
    Text12.SelStart = 0
    Text12.SelLength = Len(Text12.Text)
End Sub

Private Sub Text12_LostFocus()
    If Text12.Text <> "" Then
        If Text12.Text <= Text11.Text Then Text12.Text = Text11.Text + 1
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
    If Grad = 0 And Grad2 = 0 Then Exit Function
    
    X = -100
    FrmMain.DrawWidth = TxtLineWidth.Text
    FrmMain.ForeColor = PicColorMain.BackColor
    
    For X1 = 1 + KSX * STPX / FrmMain.ScaleWidth To STPX + KSX * STPX / FrmMain.ScaleWidth
        If NV = False Then '***
            V = (X1 / STPX * FrmMain.ScaleWidth - FrmMain.ScaleWidth / 2) '***
            DefiL = 0 '***
            For G = 0 To Grad2 '***
                DefiL = DefiL + D(G) * V ^ G '***
            Next G '*** Überprüfung aud Definitionslücke
        Else '***
            DefiL = 1 '***
        End If '***
        
        If DefiL <> 0 Then '***
            V = (X1 / STPX * FrmMain.ScaleWidth - FrmMain.ScaleWidth / 2)
            
            I = X1 / STPX * FrmMain.ScaleWidth
            
            Grad = TxtDegreeNumerator.Text
            
            For G = 0 To Grad
                Y1 = Y1 + A(G) * V ^ G
            Next G
            
            If NV = False Then
                Grad2 = TxtDegreeDenominator.Text
                
                For G = 0 To Grad2
                    Y2 = Y2 + D(G) * V ^ G
                Next G
                
                Y1 = Y1 / Y2
            End If
            
            Y1 = FrmMain.ScaleHeight / 2 - Y1
            
            If Y < FrmMain.ScaleHeight + KSY + 100 Then ' +1
                If Y > -FrmMain.ScaleHeight / 2 + KSY - 100 Then '+ 1 Then
                    If FrmMain.ScaleWidth / 2 + Text11.Text < I Then
                        If FrmMain.ScaleWidth / 2 + Text12.Text > I Then
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

Private Sub Text19_GotFocus()
    Text19.SelStart = 0
    Text19.SelLength = Len(Text19.Text)
End Sub

Private Sub TxtSetCoefficient_GotFocus()
    TxtSetCoefficient.SelStart = 0
    TxtSetCoefficient.SelLength = Len(TxtSetCoefficient.Text)
End Sub

Private Sub TxtSetCoefficient_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Slider1.Value = TxtSetCoefficient.Text * -100 ' evtl. mit Int()
        
        If OptNumerator.Value = True Then
            A(TxtGradToSetCoefficient.Text) = TxtSetCoefficient.Text
        Else
            D(TxtGradToSetCoefficient.Text) = TxtSetCoefficient.Text
        End If
        
        Draw
    End If
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

Private Sub Text21_GotFocus()
    Text21.SelStart = 0
    Text21.SelLength = Len(Text21.Text)
End Sub

Private Sub Text22_GotFocus()
    Text22.SelStart = 0
    Text22.SelLength = Len(Text22.Text)
End Sub

Private Sub Text23_GotFocus()
    Text23.SelStart = 0
    Text23.SelLength = Len(Text23.Text)
End Sub


Private Sub TxtGradToSetCoefficient_Change()
    On Error Resume Next
    
    If OptNumerator.Value = True Then
            Grad = TxtDegreeNumerator.Text
    Else
        Grad = TxtDegreeDenominator.Text
    End If
    
    If TxtGradToSetCoefficient.Text <> "" Then
        If TxtGradToSetCoefficient.Text < 0 Then TxtGradToSetCoefficient.Text = 0
        
        'If TxtGradToSetCoefficient.Text > Grad Then TxtGradToSetCoefficient.Text = Grad
        
        If OptNumerator.Value = True Then
            TxtSetCoefficient.Text = A(TxtGradToSetCoefficient.Text)
            Slider1.Value = -A(TxtGradToSetCoefficient.Text) * 100
        Else
            TxtSetCoefficient.Text = D(TxtGradToSetCoefficient.Text)
            Slider1.Value = -D(TxtGradToSetCoefficient.Text) * 100
        End If
    End If
End Sub

Private Sub TxtGradToSetCoefficient_GotFocus()
    On Error Resume Next
    TxtGradToSetCoefficient.SelStart = 0
    TxtGradToSetCoefficient.SelLength = Len(TxtGradToSetCoefficient.Text)
    KoefChange = True
End Sub

Private Sub TxtGradToSetCoefficient_LostFocus()
    On Error Resume Next
    If OptNumerator.Value = True Then
        Grad = TxtDegreeNumerator.Text
        If TxtGradToSetCoefficient.Text <> "" Then
            If TxtGradToSetCoefficient.Text < 0 Then TxtGradToSetCoefficient.Text = 0
        End If

        If TxtGradToSetCoefficient.Text > Grad Then TxtGradToSetCoefficient.Text = Grad
        TxtGradToSetCoefficient.Text = Int(TxtGradToSetCoefficient.Text)
        Slider1.Value = A(TxtGradToSetCoefficient.Text) * -100
        TxtSetCoefficient.Text = A(TxtGradToSetCoefficient.Text)
    Else
        Grad = TxtDegreeDenominator.Text
        If TxtGradToSetCoefficient.Text <> "" Then
        If TxtGradToSetCoefficient.Text < 0 Then TxtGradToSetCoefficient.Text = 0
    End If
    
        If TxtGradToSetCoefficient.Text > Grad Then TxtGradToSetCoefficient.Text = Grad
        TxtGradToSetCoefficient.Text = Int(TxtGradToSetCoefficient.Text)
        Slider1.Value = D(TxtGradToSetCoefficient.Text) * -100
        TxtSetCoefficient.Text = D(TxtGradToSetCoefficient.Text)
    End If
    
    KoefChange = False
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
        If TxtUnitsWidth.Text > 0 Then
            Cancel = False
        End If
    End If
End Sub


Private Sub TxtUnitsWidth_KeyPress(KeyAscii As Integer)
    ' The (valid) needs to be set by pressing Return
    If KeyAscii = vbKeyReturn Then
        If ChkProportional.Value = 1 Then
            TxtUnitsHeight.Text = Int(TxtUnitsWidth.Text / STPX * (STPY) * 100) / 100
        End If
        Call ScalePlot
        TxtUnitsWidth.Tag = ""
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
        If TxtUnitsHeight.Text > 0 Then
            Cancel = False
        End If
    End If
End Sub


Private Sub TxtUnitsHeight_KeyPress(KeyAscii As Integer)
    ' The (valid) needs to be set by pressing Return
    If KeyAscii = vbKeyReturn Then
        If ChkProportional.Value = 1 Then
            TxtUnitsWidth.Text = Int(TxtUnitsHeight.Text / STPY * STPX * 100) / 100
        End If
        Call ScalePlot
        TxtUnitsHeight.Tag = ""
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
        If TxtGridSpacingX.Text > 0 Then
            Cancel = False
        End If
    End If
End Sub


Private Sub TxtGridSpacingX_KeyPress(KeyAscii As Integer)
    ' The (valid) needs to be set by pressing Return
    If KeyAscii = vbKeyReturn Then
        If ChkGridSpacingLock.Value = 1 Then
            TxtGridSpacingY.Text = TxtGridSpacingX.Text
        End If
        
        Call GridSpacing
        TxtGridSpacingX.Tag = ""
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
        If TxtGridSpacingY.Text > 0 Then
            Cancel = False
        End If
    End If
End Sub


Private Sub TxtGridSpacingY_KeyPress(KeyAscii As Integer)
    ' The (valid) needs to be set by pressing Return
    If KeyAscii = vbKeyReturn Then
        If ChkGridSpacingLock.Value = 1 Then
            TxtGridSpacingX.Text = TxtGridSpacingY.Text
        End If
        
        Call GridSpacing
        TxtGridSpacingY.Tag = ""
    End If
End Sub


Private Sub Text11_LostFocus()
    If Text11.Text <> "" Then
        If Text11.Text >= Text12.Text Then Text11.Text = Text12.Text - 1
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


' *** Bildschirmauflösung nur einmal am Anfang erechnen und als Konstante übergeben --> schnelleres Zeichnen des Graphen
Private Sub Timer1_Timer()
    If Faktor <> Slider1.Value Then
        If TxtGradToSetCoefficient.Text <> "" Then
            FrmMain.Cls
            A(TxtGradToSetCoefficient.Text) = -Slider1.Value / 100
            For G = 0 To Grad
                Y = Y + A(G) * (-FrmMain.ScaleWidth / 2) ^ G
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
            Grad = TxtDegreeNumerator.Text
            A(TxtGradToSetCoefficient.Text) = -Slider1.Value / 100
            For G = 0 To Grad
                Y = Y + A(G) * (-FrmMain.ScaleWidth / 2) ^ G
            Next G
        Else
            Grad = TxtDegreeDenominator.Text
            D(TxtGradToSetCoefficient.Text) = -Slider1.Value / 100
            For G = 0 To Grad
                Y = Y + A(G) * (-FrmMain.ScaleWidth / 2) ^ G
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
        
        Grad = TxtDegreeNumerator.Text
        Grad2 = TxtDegreeDenominator.Text
        
        ReDim C(Grad - DIFFNR)
        For G = 1 To Grad - DIFFNR '+1
            C(G - 1) = A(G) * G
        Next G
        
        For G = 0 To Grad - 1 - DIFFNR
            DiffNA = DiffNA + C(G) * V ^ G
        Next G
        For I = 0 To Grad - DIFFNR
            C(Grad - DIFFNR) = 0
        Next I
        
        ReDim C(Grad2)
        For G = 1 To Grad2 - DIFFNR '+1
            C(G - 1) = D(G) * G
        Next G
        
        For G = 0 To Grad2 - 1 - DIFFNR
            DiffZA = DiffZA + C(G) * V ^ G
        Next G
        For I = 0 To Grad2 - DIFFNR
            C(Grad2 - DIFFNR) = 0
        Next I
        
        For G = 0 To Grad - DIFFNR
            DiffN = DiffN + A(G) * V ^ G
        Next G
        
        For G = 0 To Grad2 - DIFFNR
            DiffZ = DiffZ + D(G) * V ^ G
        Next G
        
        Y1 = (DiffNA * DiffZ - DiffN * DiffZA) / (DiffZ ^ 2)
        
        Y1 = FrmMain.ScaleHeight / 2 - Y1
        
        If Y < FrmMain.ScaleHeight + KSY + 100 Then ' +1
            If Y > -FrmMain.ScaleHeight / 2 + KSY + 1 Then
                If FrmMain.ScaleWidth / 2 + Text11.Text < I Then
                    If FrmMain.ScaleWidth / 2 + Text12.Text > I Then
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
        
        For G = 0 To Grad3
            Y1 = Y1 + Z(G) * V ^ G
        Next G
        
        Y1 = FrmMain.ScaleHeight / 2 - Y1
        
        If Y < FrmMain.ScaleHeight + KSY + 100 Then ' +1
            If Y > -FrmMain.ScaleHeight / 2 + KSY + 1 Then
                If FrmMain.ScaleWidth / 2 + Text11.Text < I Then
                    If FrmMain.ScaleWidth / 2 + Text12.Text > I Then
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
