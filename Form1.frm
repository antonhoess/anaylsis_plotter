VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   12555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10785
   ControlBox      =   0   'False
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8.719
   ScaleMode       =   5  'Zoll
   ScaleWidth      =   7.49
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text17 
      Height          =   375
      Left            =   5640
      TabIndex        =   94
      Text            =   "1"
      Top             =   3240
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6000
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Ende"
      Height          =   375
      Left            =   480
      TabIndex        =   65
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Timer Timer3 
      Left            =   4320
      Top             =   2280
   End
   Begin VB.Timer Timer2 
      Left            =   3960
      Top             =   2280
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3600
      Top             =   2280
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Steuerung"
      Height          =   11415
      Left            =   480
      TabIndex        =   4
      Top             =   2760
      Width           =   4695
      Begin VB.CommandButton Command17 
         BackColor       =   &H0080C0FF&
         Caption         =   "Asymtote"
         Height          =   495
         Left            =   3240
         Style           =   1  'Grafisch
         TabIndex        =   91
         Top             =   8400
         Width           =   1335
      End
      Begin VB.CommandButton Command20 
         BackColor       =   &H0080C0FF&
         Caption         =   "Zurückdifferenzieren"
         Height          =   495
         Left            =   1560
         Style           =   1  'Grafisch
         TabIndex        =   89
         Top             =   8400
         Width           =   1575
      End
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   840
         TabIndex        =   87
         Top             =   9120
         Width           =   495
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H0080C0FF&
         Caption         =   "Funktionswert errechnen"
         Height          =   495
         Left            =   120
         Style           =   1  'Grafisch
         TabIndex        =   86
         Top             =   8400
         Width           =   1335
      End
      Begin VB.TextBox Text16 
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   85
         Top             =   9360
         Width           =   495
      End
      Begin VB.TextBox Text20 
         Height          =   285
         Left            =   600
         TabIndex        =   84
         Text            =   "1"
         Top             =   6840
         Width           =   615
      End
      Begin MSComCtl2.FlatScrollBar FlatScrollBar1 
         Height          =   495
         Left            =   600
         TabIndex        =   82
         Top             =   7200
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   2
         Min             =   1
         Max             =   10
         Orientation     =   8323072
         Value           =   10
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H0080C0FF&
         Caption         =   "COL"
         Height          =   1095
         Left            =   720
         TabIndex        =   66
         Top             =   5280
         Width           =   495
         Begin VB.PictureBox Picture1 
            Height          =   855
            Left            =   480
            ScaleHeight     =   0.552
            ScaleMode       =   5  'Zoll
            ScaleWidth      =   1.01
            TabIndex        =   68
            Top             =   240
            Width           =   1510
            Begin VB.PictureBox Picture2 
               BackColor       =   &H00404040&
               BorderStyle     =   0  'Kein
               Height          =   255
               Index           =   9
               Left            =   720
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   78
               Top             =   240
               Width           =   255
            End
            Begin VB.PictureBox Picture2 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'Kein
               Height          =   255
               Index           =   11
               Left            =   1200
               ScaleHeight     =   0.177
               ScaleMode       =   5  'Zoll
               ScaleWidth      =   0.177
               TabIndex        =   80
               Top             =   240
               Width           =   255
            End
            Begin VB.PictureBox Picture2 
               BackColor       =   &H00004080&
               BorderStyle     =   0  'Kein
               Height          =   255
               Index           =   10
               Left            =   960
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   79
               Top             =   240
               Width           =   255
            End
            Begin VB.PictureBox Picture2 
               BackColor       =   &H00C0C0FF&
               BorderStyle     =   0  'Kein
               Height          =   255
               Index           =   8
               Left            =   480
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   77
               Top             =   240
               Width           =   255
            End
            Begin VB.PictureBox Picture2 
               BackColor       =   &H000080FF&
               BorderStyle     =   0  'Kein
               Height          =   255
               Index           =   7
               Left            =   240
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   76
               Top             =   240
               Width           =   255
            End
            Begin VB.PictureBox Picture2 
               BackColor       =   &H00FF8080&
               BorderStyle     =   0  'Kein
               Height          =   255
               Index           =   6
               Left            =   0
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   75
               Top             =   240
               Width           =   255
            End
            Begin VB.PictureBox Picture2 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'Kein
               Height          =   255
               Index           =   5
               Left            =   1200
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   74
               Top             =   0
               Width           =   255
            End
            Begin VB.PictureBox Picture2 
               BackColor       =   &H00C000C0&
               BorderStyle     =   0  'Kein
               Height          =   255
               Index           =   4
               Left            =   960
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   73
               Top             =   0
               Width           =   255
            End
            Begin VB.PictureBox Picture2 
               BackColor       =   &H0000FF00&
               BorderStyle     =   0  'Kein
               Height          =   255
               Index           =   3
               Left            =   720
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   72
               Top             =   0
               Width           =   255
            End
            Begin VB.PictureBox Picture2 
               BackColor       =   &H0080FFFF&
               BorderStyle     =   0  'Kein
               Height          =   255
               Index           =   2
               Left            =   480
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   71
               Top             =   0
               Width           =   255
            End
            Begin VB.PictureBox Picture2 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'Kein
               Height          =   255
               Index           =   1
               Left            =   240
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   70
               Top             =   0
               Width           =   255
            End
            Begin VB.PictureBox Picture2 
               BackColor       =   &H000000FF&
               BorderStyle     =   0  'Kein
               Height          =   255
               Index           =   0
               Left            =   0
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   69
               Top             =   0
               Width           =   255
            End
            Begin VB.Label Label27 
               Alignment       =   2  'Zentriert
               BackColor       =   &H000735BC&
               BorderStyle     =   1  'Fest Einfach
               Caption         =   "SELECT  "
               Height          =   375
               Left            =   0
               TabIndex        =   81
               Top             =   480
               Width           =   1575
            End
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00FF0000&
            Height          =   855
            Left            =   0
            ScaleHeight     =   0.552
            ScaleMode       =   5  'Zoll
            ScaleWidth      =   0.344
            TabIndex        =   67
            Top             =   240
            Width           =   550
         End
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H0080C0FF&
         Caption         =   "Differentieren"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         Style           =   1  'Grafisch
         TabIndex        =   34
         Top             =   4680
         Width           =   2055
      End
      Begin VB.TextBox Text14 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   56
         Text            =   "0"
         Top             =   3360
         Width           =   495
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H0080C0FF&
         Caption         =   "Gebr. Rat. Fkt"
         Height          =   435
         Left            =   720
         Style           =   1  'Grafisch
         TabIndex        =   55
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1320
         TabIndex        =   54
         Text            =   "3"
         Top             =   2640
         Width           =   495
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H0080C0FF&
         Caption         =   "Frame5"
         Enabled         =   0   'False
         Height          =   1095
         Left            =   600
         TabIndex        =   50
         Top             =   1560
         Width           =   1335
         Begin VB.CommandButton Command1 
            BackColor       =   &H0080C0FF&
            Caption         =   "GO!"
            Default         =   -1  'True
            Height          =   255
            Left            =   120
            Style           =   1  'Grafisch
            TabIndex        =   53
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H0080C0FF&
            Caption         =   "Clear"
            Height          =   255
            Left            =   120
            Style           =   1  'Grafisch
            TabIndex        =   52
            Top             =   480
            Width           =   1095
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H0080C0FF&
            Caption         =   "Trace"
            Height          =   255
            Left            =   120
            Style           =   1  'Grafisch
            TabIndex        =   51
            Top             =   720
            Width           =   1095
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H0080C0FF&
         Caption         =   "Grad    Koeffizient"
         Enabled         =   0   'False
         Height          =   855
         Left            =   600
         TabIndex        =   44
         Top             =   240
         Width           =   1455
         Begin VB.OptionButton Option2 
            BackColor       =   &H0080C0FF&
            Caption         =   "N"
            Enabled         =   0   'False
            Height          =   255
            Left            =   600
            TabIndex        =   48
            Top             =   480
            Width           =   375
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H0080C0FF&
            Caption         =   "Z"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   480
            Value           =   -1  'True
            Width           =   375
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   720
            TabIndex        =   46
            Text            =   "0"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   120
            TabIndex        =   45
            Text            =   "0"
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Zentriert
            BackStyle       =   0  'Transparent
            Caption         =   "="
            Height          =   255
            Left            =   480
            TabIndex        =   49
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.CommandButton Command16 
         Caption         =   "-"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   8040
         Width           =   375
      End
      Begin VB.CommandButton Command15 
         Caption         =   "+"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H0080C0FF&
         Caption         =   "Wertebereich errechnen"
         Enabled         =   0   'False
         Height          =   495
         Left            =   2880
         Style           =   1  'Grafisch
         TabIndex        =   35
         Top             =   6480
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080C0FF&
         Caption         =   "Koeffizienten"
         Height          =   255
         Left            =   720
         Style           =   1  'Grafisch
         TabIndex        =   29
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Proportionalität"
         Height          =   1935
         Left            =   2040
         TabIndex        =   22
         Top             =   2640
         Width           =   2175
         Begin VB.CheckBox Check1 
            BackColor       =   &H0080C0FF&
            Caption         =   "Proportional"
            Height          =   375
            Left            =   120
            Style           =   1  'Grafisch
            TabIndex        =   26
            Top             =   1440
            Value           =   1  'Aktiviert
            Width           =   1935
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H0080C0FF&
            Caption         =   "Proportionieren"
            Height          =   495
            Left            =   120
            Style           =   1  'Grafisch
            TabIndex        =   25
            Top             =   960
            Width           =   1935
         End
         Begin VB.TextBox Text5 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1440
            TabIndex        =   24
            Text            =   "10"
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   1440
            TabIndex        =   23
            Text            =   "13"
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Längeneinheiten"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Breiteneinheiten"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H000080FF&
         Caption         =   "Beenden"
         Height          =   615
         Left            =   2880
         Style           =   1  'Grafisch
         TabIndex        =   21
         Top             =   7080
         Width           =   1455
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H0080C0FF&
         Caption         =   "Koeffizienten speichern"
         Enabled         =   0   'False
         Height          =   435
         Left            =   720
         Style           =   1  'Grafisch
         TabIndex        =   20
         Top             =   4320
         Width           =   1095
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H0080C0FF&
         Caption         =   "Koeffizienten laden"
         Height          =   375
         Left            =   720
         Style           =   1  'Grafisch
         TabIndex        =   19
         Top             =   4800
         Width           =   1095
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H0080C0FF&
         Caption         =   "Verschieben (Koordinatensystem)"
         Height          =   615
         Left            =   2880
         Style           =   1  'Grafisch
         TabIndex        =   18
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   2280
         TabIndex        =   17
         Text            =   "0"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   2280
         TabIndex        =   16
         Text            =   "0"
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   3720
         TabIndex        =   15
         Text            =   "1"
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H0080C0FF&
         Caption         =   "strecken"
         Height          =   375
         Left            =   2040
         Style           =   1  'Grafisch
         TabIndex        =   14
         Top             =   2160
         Width           =   2415
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Raster"
         Height          =   495
         Left            =   1320
         Style           =   1  'Grafisch
         TabIndex        =   13
         Top             =   5280
         Value           =   1  'Aktiviert
         Width           =   1335
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   3720
         TabIndex        =   12
         Text            =   "1"
         Top             =   1800
         Width           =   735
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H0080C0FF&
         Caption         =   "Koordinaten"
         Height          =   495
         Left            =   2880
         Style           =   1  'Grafisch
         TabIndex        =   11
         Top             =   5280
         Value           =   1  'Aktiviert
         Width           =   1455
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H0080C0FF&
         Caption         =   "Achsenkreuz"
         Height          =   495
         Left            =   1320
         Style           =   1  'Grafisch
         TabIndex        =   10
         Top             =   5880
         Value           =   1  'Aktiviert
         Width           =   1335
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H0080C0FF&
         Caption         =   "Immer im Vordergrund"
         Height          =   495
         Left            =   2880
         Style           =   1  'Grafisch
         TabIndex        =   9
         Top             =   5880
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Intervall"
         Height          =   1335
         Left            =   1200
         TabIndex        =   6
         Top             =   6480
         Width           =   1455
         Begin VB.TextBox Text13 
            Enabled         =   0   'False
            Height          =   285
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   42
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox Text6 
            Enabled         =   0   'False
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   41
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox Text11 
            Height          =   285
            Left            =   240
            TabIndex        =   8
            Text            =   "-1000"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox Text12 
            Height          =   285
            Left            =   840
            TabIndex        =   7
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
            TabIndex        =   43
            Top             =   600
            Width           =   1215
         End
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H0000C0C0&
         Caption         =   "Ausblenden"
         Height          =   375
         Left            =   1320
         Style           =   1  'Grafisch
         TabIndex        =   5
         Top             =   7920
         Width           =   3015
      End
      Begin ComctlLib.Slider Slider1 
         Height          =   7695
         Left            =   120
         TabIndex        =   36
         Top             =   360
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   13573
         _Version        =   327682
         BorderStyle     =   1
         Enabled         =   0   'False
         Orientation     =   1
         Min             =   -1000
         Max             =   1000
         TickStyle       =   3
      End
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2370
         Left            =   1440
         TabIndex        =   90
         Top             =   9000
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   12648447
         BackColor       =   16512
         Appearance      =   1
         MonthBackColor  =   128
         StartOfWeek     =   662831106
         TitleBackColor  =   8388608
         TitleForeColor  =   12632064
         TrailingForeColor=   8454016
         CurrentDate     =   37576
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "X= Y="
         Height          =   375
         Left            =   480
         TabIndex        =   88
         Top             =   9240
         Width           =   255
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Dicke"
         Height          =   255
         Left            =   720
         TabIndex        =   83
         Top             =   6600
         Width           =   495
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Grad(N)"
         Enabled         =   0   'False
         Height          =   255
         Left            =   720
         TabIndex        =   62
         Top             =   3360
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   61
         Top             =   3840
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   60
         Top             =   4080
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Grad(Z)"
         Height          =   255
         Left            =   720
         TabIndex        =   59
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "X Y"
         Height          =   375
         Left            =   840
         TabIndex        =   58
         Top             =   3840
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
         TabIndex        =   57
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fest Einfach
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
         TabIndex        =   33
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "X  Y"
         Height          =   615
         Left            =   2160
         TabIndex        =   32
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Strecken (Faktor)-X"
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
         TabIndex        =   31
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Strecken (Faktor)-Y"
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
         TabIndex        =   30
         Top             =   1800
         Width           =   1455
      End
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Polynom-Anzahl"
      Height          =   255
      Left            =   5640
      TabIndex        =   95
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label18 
      Caption         =   "Über Ableitung Wendepunkte, Extremwerte und über Horner-Schema Nullpunkte, Definitionslücken, Pole, hebbare Lücken errechnen.Nicht nur mir ganzzahligen Exponenten arbeiten, sondern über ein zweites Array einmal Koeffizienten und über das zweite die Exponenten speichern und ganz normal in der Schleife durchrechnen lassen - das gleiche glit dann auch für das Differenzieren"
      Height          =   1455
      Left            =   2640
      TabIndex        =   93
      Top             =   840
      Width           =   4335
   End
   Begin VB.Label Label20 
      Caption         =   "Beim Proportionieren Graph weiterhin zeichnen lassen z.B., wenn Faktor 1 ist!"
      Height          =   615
      Left            =   480
      TabIndex        =   92
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   120
      Left            =   0
      Top             =   0
      Width           =   120
   End
   Begin VB.Label Label26 
      Caption         =   "bei form2 mit links und rechts-Pfeiltasten bewegen"
      Height          =   255
      Left            =   2640
      TabIndex        =   64
      Top             =   600
      Width           =   3855
   End
   Begin VB.Label Label9 
      Caption         =   "Bei Trace Cursorposition aktualisieren (vielleicht über getcursor pos!)"
      Height          =   255
      Left            =   2640
      TabIndex        =   63
      Top             =   360
      Width           =   4935
   End
   Begin VB.Label Label23 
      Caption         =   "Bei Bruch mehr als nur einaml differenzieren"
      Height          =   375
      Left            =   480
      TabIndex        =   40
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label24 
      Caption         =   "Form1.BorderStyle=0 (evtl.)"
      Height          =   255
      Left            =   2640
      TabIndex        =   37
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label11 
      Caption         =   "Proportional bei Raster nach Breite von Form1 angleichen, denn ansonsten stimmt es nicht überein"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label16 
      Caption         =   "evtl abschnittsweise Definition"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label10 
      Caption         =   "Arrays in Datei speichern und aus Datei laden"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Einstellungen speichern"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X, Y, X1, Y1, Y2, X2, i, V, G, B As Boolean, W, Faktor, KSX, KSY, SFX, SFY, STPX, STPY, MNS As Boolean, MENX, MENY, MCX, MCY, SliderValue, Plus As Boolean, C(), GradDiff, DragX, DragY, DiffZ, DiffN, DiffZA, DiffNA, E, DIFFNR, ASYM, Z, J, H, K(), L(), Faktor2, Grad3, A1, A2

Option Explicit

Private Declare Function SetCursorPos Lib "user32" (ByVal _
        X As Long, ByVal Y As Long) As Long

Private Declare Function GetCursorPos Lib "user32" _
        (lpPoint As POINTAPI) As Long

Private Type POINTAPI
  X As Long
  Y As Long
End Type

Dim aX, aY, aY2, dx, dy

Private Sub Check1_Click()
If Check1.Value = 0 Then
Text5.Enabled = True
Else
Text5.Enabled = False
Text5.Text = Int(Text4.Text / STPX * (STPY) * 100) / 100
End If
If Check1.Value = 1 Then
Form1.ScaleHeight = Text5.Text
Else
Form1.ScaleHeight = Int(Text4.Text / STPX * (STPY) * 100) / 100
End If
Form1.Cls
Call Raster
Call Nullpunkt
Call Koordinaten
Call Graph
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
Call Raster
Call Nullpunkt
Call Koordinaten
Call Graph
Else
Form1.Cls
Call Nullpunkt
Call Koordinaten
Call Graph
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
Call Raster
Call Nullpunkt
Call Koordinaten
Call Graph
Else
Form1.Cls
Call Raster
Call Nullpunkt
Call Graph
End If
End Sub

Private Sub Check4_Click()
If Check4.Value = 1 Then
Call Raster
Call Nullpunkt
Call Koordinaten
Call Graph
Else
Form1.Cls
Call Raster
Call Koordinaten
Call Graph
End If
End Sub

Private Sub Check5_Click()
If Check5.Value = 0 Then
'Form in den Normalzustand zurücksetzen
Call SetWindowPos(Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, 3)
Else
'Form dauerhaft in den Vordergrund setzen
Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, 3)
End If
End Sub

Private Sub Check6_Click()
GRF = Check6.Value
If Check6.Value = 1 Then
NV = True
Option2.Enabled = False
Label21.Enabled = True
Text14.Enabled = True
Else
NV = False
Option2.Enabled = False
Option1.Value = True
Label21.Enabled = False
Text14.Enabled = False
End If

End Sub

Private Sub Command10_Click()
Form1.Cls
SFX = Text9.Text
SFY = Text10.Text
Call Raster
Call Nullpunkt
Call Koordinaten
Call Graph
End Sub

Private Sub Command11_Click()
MENX = Frame3.Left
MENY = Frame3.Top
Frame3.Visible = False
End Sub

Private Sub Command12_Click()
If NV = True Then
Grad = Text1.Text
For i = 1 To Grad
A(i - 1) = A(i) * (i)
Next i
A(Grad) = 0
Else
Call Graph2
End If
End Sub

Private Sub Command13_Click()
If NV = True Then

Grad = Text1.Text
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

Grad = Text1.Text
For G = 0 To Grad
Y1 = Y1 + A(G) * Text11.Text ^ G
Next G

Grad = Text14.Text
For G = 0 To Grad
Y2 = Y2 + D(G) * Text11.Text ^ G
Next G
Text6.Text = Y1 / Y2

Y1 = 0
Y2 = 0

Grad = Text1.Text
For G = 0 To Grad
Y1 = Y1 + A(G) * Text12.Text ^ G
Next G

Grad = Text14.Text
For G = 0 To Grad
Y2 = Y2 + D(G) * Text12.Text ^ G
Next G
Text13.Text = Y1 / Y2

Y1 = 0
Y2 = 0

End If
End Sub

Private Sub Command14_Click()
On Error Resume Next
Grad = Text1.Text

For G = 0 To Grad
Y1 = Y1 + A(G) * Text15.Text ^ G
Next G

If NV = False Then
Grad = Text14.Text

For G = 0 To Grad
Y2 = Y2 + D(G) * Text15.Text ^ G
Next G

Y1 = Y1 / Y2
End If

Text16.Text = Y1
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
'Private Sub Command17_Click()
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


Private Sub Command17_Click()
On Error Resume Next
Grad = Text1.Text
Grad2 = Text14.Text
If NV = False Then
If Grad >= Grad2 Then
Grad3 = Grad - Grad2
ReDim K(Grad + 1)
ReDim L(Grad2 + 1)

For i = 0 To Grad
K(i) = A(i)
Next i

For i = 0 To Grad2
L(i) = D(i)
Next i

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

Private Sub Command18_Click()
End
End Sub


Private Sub Command20_Click()
'Grad = Text1.Text
'For i = 1 To Grad
'A(i - 1) = A(i) * (i)
'Next i
'A(Grad) = 0
Grad = Text1.Text
ReDim C(Grad + 1)
For i = 0 To Grad + 1
C(i) = A(i)
Next i
ReDim A(Grad + 2)
A(0) = 0
For i = 0 To Grad + 1
A(i + 1) = C(i) / (i + 1)
Next i
Text1.Text = Text1.Text + 1
Grad = Text1.Text


On Error Resume Next
X = -100
Form1.DrawWidth = Text20.Text
Form1.ForeColor = Picture3.BackColor
For X1 = 1 To STPX
V = (X1 / STPX * Form1.ScaleWidth - Form1.ScaleWidth / 2)

i = X1 / STPX * Form1.ScaleWidth

'Grad = Text1.Text

For G = 0 To Grad
Y1 = Y1 + A(G) * V ^ G
Next G

If NV = False Then
Grad = Text14.Text

For G = 0 To Grad
Y2 = Y2 + D(G) * V ^ G
Next G

Y1 = Y1 / Y2
End If

Y1 = Form1.ScaleHeight / 2 - Y1

If Y < Form1.ScaleHeight + KSY + 100 Then ' +1
If Y > -Form1.ScaleHeight / 2 + KSY + 1 Then
If Form1.ScaleWidth / 2 + Text11.Text < i Then
If Form1.ScaleWidth / 2 + Text12.Text > i Then
Form1.Line (X - KSX, Y - KSY)-(i - KSX, Y1 - KSY)
End If
End If
End If
End If

Y = Y1
X = (X1 - 0) / STPX * Form1.ScaleWidth
Y1 = 0
Y2 = 0
Next X1
Form1.DrawWidth = 1
End Sub


Private Sub Command4_Click()
If B = True Then
B = False
'Timer3.Interval = 0
Else
B = True
'Timer3.Interval = 1
End If

End Sub

Private Sub Command1_Click()
X = -100 '-1

Call Raster
Call Nullpunkt
Call Koordinaten
Call Graph

End Sub

Private Sub Command2_Click()
Form1.Cls

Call Raster
Call Nullpunkt
Call Koordinaten

End Sub

Private Sub Command3_Click()
If Text17.Text < 1 Then Text17.Text = 1
Poly = Text17.Text
Frame4.Enabled = True
Frame5.Enabled = True
Label17.Enabled = True
Text6.Enabled = True
Text13.Enabled = True
Command7.Enabled = True
Command12.Enabled = True
Command13.Enabled = True
Command15.Enabled = True
Command16.Enabled = True
Slider1.Enabled = True
On Error Resume Next
If Text1.Text < 0 Then Text1.Text = 0
If Text14.Text < 0 Then Text1.Text = 0
Text1.Text = Int(Text1.Text)
Text14.Text = Int(Text14.Text)
Grad = Text1.Text
If Check6.Value = 1 Then
NV = False
Else
NV = True
End If
If Check6.Value = 1 Then Grad2 = Text14.Text
Form2.Show (1)
End Sub

Private Sub Command5_Click()
On Error Resume Next
Form1.Cls
Form1.ScaleMode = 0
Picture1.ScaleMode = 0
For i = 0 To Picture2.Count
Picture2(i).ScaleMode = 0
Next i
Picture3.ScaleMode = 0
Form1.ScaleWidth = Text4.Text
If Check1.Value = 0 Then
Form1.ScaleHeight = Text5.Text
Else
Form1.ScaleHeight = Int(Text4.Text / STPX * (STPY) * 100) / 100
End If

Call Raster
Call Nullpunkt
Call Koordinaten
Call Graph
End Sub

Private Sub Command6_Click()
End
End Sub

'Private Sub Command7_Click()
'   SaveStringArray App.Path & "\Test.dat", A()
'End Sub
'
'Private Sub Command8_Click()
' ReadStringArray App.Path & "\Test.dat", A
'End Sub

Private Sub Command9_Click()
KSX = -Text7.Text
KSY = Text8.Text
Form1.Cls
Call Raster
Call Nullpunkt
Call Koordinaten
Call Graph
End Sub


Private Sub FlatScrollBar1_Change()
Text20.Text = -FlatScrollBar1.Value + 11
End Sub

Private Sub Form_Activate()
Text4.Text = Int(Me.ScaleWidth * 100) / 100
Text5.Text = Int(Me.ScaleHeight * 100) / 100
End Sub

'
'Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
'    For I = 0 To Form1.Controls.Count - 1
''        If Not TypeOf Frm.Controls(I) Is Menu Then
''            Frm.Controls(I).Enabled = State
''        End If
'Frame3.Move X, Y 'X - DragX, Y - DragY
'    Next I
'
''Frame3.Move X, Y 'X - DragX, Y - DragY
'End Sub

Private Sub Form_DragDrop(source As Control, X As Single, Y As Single)
On Error Resume Next
'Frame3.Move X - DragX, Y - DragY

End Sub

Private Sub Form_Load()
Me.WindowState = 2
If Check5.Value = 1 Then
'Form dauerhaft in den Vordergrund setzen
Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, 3)
Else
'Form dauerhaft in den Vordergrund setzen
Call SetWindowPos(Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, 3)
End If
 
 KSX = 0
 KSY = 0
 SFX = 1
 SFY = 1
 
 STPX = Screen.Width / Screen.TwipsPerPixelX
 STPY = Screen.Height / Screen.TwipsPerPixelY - 22
 
Call Nullpunkt

Y = Form1.ScaleHeight / 2
X = 0

 
  B = False
  
  Faktor = Slider1.Value

Call Raster
Call Koordinaten

Plus = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Label1.Caption = Int((X - Form1.ScaleWidth / 2 + KSX) * 100) / 100
Label2.Caption = -Int((Y - Form1.ScaleHeight / 2 + KSY) * 100) / 100
 Dim Pt As POINTAPI
  
    Call GetCursorPos(Pt)
      'aX = Pt.X

W = X - Form1.ScaleWidth / 2 + KSX


If B = True Then
If NV = False Then
Grad = Text1.Text
For G = 0 To Grad
aY = aY + A(G) * W ^ G
Next G

Grad = Text14.Text
For G = 0 To Grad
aY2 = aY2 + D(G) * W ^ G
Next G
aY = aY / aY2
Else
'If NV = True Then
Grad = Text1.Text
For G = 0 To Grad
aY = aY + A(G) * W ^ G
Next G
'End If
End If

aY = -aY + Form1.ScaleHeight / 2
End If

If B = True Then
Call SetCursorPos(X / Form1.ScaleWidth * STPX, (aY - KSY) / Form1.ScaleHeight * (STPY) + 20)
Call GetCursorPos(Pt)
Label1.Caption = Pt.X
Label2.Caption = Pt.Y
aY = 0
aY2 = 0
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
If Frame3.Visible = False Then
Frame3.Visible = True
End If
End If
End Sub

Private Sub Form_Resize()
Form1.Cls
Call Raster
Call Nullpunkt
Call Koordinaten
Call Graph
End Sub



Private Sub Frame3_DragDrop(source As Control, X As Single, Y As Single)
Frame3.Left = Frame3.Left + (X - DragX)
Frame3.Top = Frame3.Top + (Y - DragY)
End Sub

Private Sub Image1_DragDrop(source As Control, X As Single, Y As Single)
Frame3.Move X / Screen.TwipsPerPixelX / 1280 * Form1.ScaleWidth - DragX, Y / Screen.TwipsPerPixelY / 1024 * Form1.ScaleHeight - DragY       'X, Y
Image1.Visible = False
Frame3.Visible = True
End Sub

Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragX = (Label19.Left + X) / Screen.TwipsPerPixelX / 1280 * Form1.ScaleWidth 'X - Frame3.Left 'X - Frame3.Left
DragY = (Label19.Top + Y) / Screen.TwipsPerPixelY / 1024 * Form1.ScaleHeight 'Y - Frame3.Top 'Y - Frame3.Top
Image1.Left = 0
Image1.Top = 0
Image1.Width = Form1.Width
Image1.Height = Form1.Height
Image1.Visible = True
Frame3.Visible = False
Frame3.Drag 1
'Text1.Text = Command1.Width / Screen.TwipsPerPixelX / 1280 * 13.32
End Sub

Private Sub Label27_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
CommonDialog1.ShowColor
Picture3.BackColor = CommonDialog1.Color
Frame6.Width = (Picture3.ScaleWidth) / Form1.ScaleWidth * Screen.TwipsPerPixelX * 1280
End Sub

'Private Sub Label19_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Frame3.Drag 2
''Frame3.Left = Frame3.Left + (X - DragX)
''Frame3.Top = Frame3.Top + (Y - DragY)
'Frame3.Move X - DragX, Y - DragY
'End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
Grad = Text1.Text
Else
Grad = Text14.Text
End If

On Error Resume Next
If Text3.Text <> "" Then
If Text3.Text < 0 Then Text3.Text = 0
End If

If Text3.Text > Grad Then Text3.Text = Grad
Text3.Text = Int(Text3.Text)
Slider1.Value = A(Text3.Text) * -100
Text2.Text = A(Text3.Text)
End Sub

Private Sub Picture2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
For i = 0 To Picture2.Count - 1
Picture2(i).BorderStyle = 0
Next i
Picture2(Index).BorderStyle = 1
End Sub

Private Sub Picture2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'For i = 0 To Picture2.Count - 1
'Picture2(i).BorderStyle = 0
'Next i
'Picture2(Index).BorderStyle = 1
End Sub

Private Sub Picture2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture3.BackColor = Picture2(Index).BackColor
Frame6.Width = (Picture3.ScaleWidth) / 13.32 * Screen.TwipsPerPixelX * 1280
'For i = 0 To Picture2.Count - 1
'Picture2(i).BorderStyle = 0
'Next i
'Picture2(Index).BorderStyle = 1
End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Frame6.Width = (Picture1.ScaleWidth + Picture3.ScaleWidth) / 13.32 * Screen.TwipsPerPixelX * 1280
'Frame6.Width = (Picture3.ScaleWidth) / Form1.ScaleWidth * Screen.TwipsPerPixelX * 1280
End Sub

Private Sub Picture3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Frame6.Width = (Picture3.ScaleWidth) / Form1.ScaleWidth * Screen.TwipsPerPixelX * 1280
End Sub

Private Sub Slider1_Scroll()
If Faktor <> Slider1.Value Then
If Text3.Text <> "" Then

Text2.Text = -Slider1.Value / 100
Form1.Cls
If Option1.Value = True Then
Grad = Text1.Text
A(Text3.Text) = -Slider1.Value / 100
For G = 0 To Grad
Y = Y + A(G) * (-Form1.ScaleWidth / 2) ^ G
Next G
Else
Grad = Text14.Text
D(Text3.Text) = -Slider1.Value / 100
For G = 0 To Grad
Y = Y + A(G) * (-Form1.ScaleWidth / 2) ^ G
Next G
End If
X = 0

Call Raster
Call Nullpunkt
Call Koordinaten
Call Graph

End If
End If
End Sub





Private Sub Text12_LostFocus()
If Text12.Text <> "" Then
If Text12.Text <= Text11.Text Then Text12.Text = Text11.Text + 1
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Slider1.Value = Text2.Text * -100 ' evtl. mit Int()
If Option1.Value = True Then
A(Text3.Text) = Text2.Text
Else
D(Text3.Text) = Text2.Text
End If
Form1.Cls
Call Raster
Call Nullpunkt
Call Koordinaten
Call Graph
End If
End Sub


Private Function Koordinaten()
If Check3.Value = 1 Then

  Form1.ForeColor = RGB(0, 100, 0)
  
For i = -Int(Form1.ScaleWidth / 2 - KSX) To Int(Form1.ScaleWidth / 2 + KSX)
Form1.CurrentY = Form1.ScaleHeight / 2 - KSY
Form1.CurrentX = Form1.ScaleWidth / 2 - KSX + i
Form1.Print Int(i)
Next i

For i = -Int(Form1.ScaleHeight / 2 - KSY) To Int(Form1.ScaleHeight / 2 + KSY)
Form1.CurrentX = Form1.ScaleWidth / 2 - KSX
Form1.CurrentY = Form1.ScaleHeight / 2 - KSY + i
Form1.Print Int(-i)
Next i

Form1.ForeColor = RGB(0, 0, 255)

End If
End Function

Private Function Raster()

If Check2.Value = 1 Then
  
  Form1.DrawStyle = 2
  Form1.ForeColor = RGB(255, 0, 0)
   
  For i = Int(-1 / SFX) To Int(Form1.ScaleWidth / SFX) + 1
  Form1.Line ((Form1.ScaleWidth / 2 / SFX - (Int(Form1.ScaleWidth / 2 / SFX))) * SFX + i * SFX - (KSX - SFX * Int(KSX / SFX)), 0)-((Form1.ScaleWidth / 2 / SFX - (Int(Form1.ScaleWidth / 2 / SFX))) * SFX + i * SFX - (KSX - SFX * Int(KSX / SFX)), Form1.ScaleHeight) '(KSX - Int(KSX) --> beim Strecken anpassen
  Next i
  
  For i = Int(-1 / SFY) To Int(Form1.ScaleHeight / SFY) + 1
  Form1.Line (0, (Form1.ScaleHeight / 2 / SFY - (Int(Form1.ScaleHeight / 2 / SFY))) * SFY + i * SFY - (KSY - SFY * Int(KSY / SFY)))-(Form1.ScaleWidth, (Form1.ScaleHeight / 2 / SFY - (Int(Form1.ScaleHeight / 2 / SFY))) * SFY + i * SFY - (KSY - SFY * Int(KSY / SFY)))
  Next i
  
  Form1.ForeColor = RGB(0, 0, 255)
  Form1.DrawStyle = 0
  
  End If

End Function

Private Function Nullpunkt()
If Check4.Value = 1 Then

Form1.DrawWidth = 3
Form1.ForeColor = 0

Form1.Line (Form1.ScaleWidth / 2 - KSX, 0)-(Form1.ScaleWidth / 2 - KSX, Form1.ScaleHeight)
Form1.Line (0, Form1.ScaleHeight / 2 - KSY)-(Form1.ScaleWidth, Form1.ScaleHeight / 2 - KSY)

Form1.DrawWidth = 1
Form1.ForeColor = RGB(0, 0, 255)

End If
End Function

Private Function Graph()
On Error Resume Next
X = -100
Form1.DrawWidth = Text20.Text
Form1.ForeColor = Picture3.BackColor
For X1 = 1 + KSX * STPX / Form1.ScaleWidth To STPX + KSX * STPX / Form1.ScaleWidth
V = (X1 / STPX * Form1.ScaleWidth - Form1.ScaleWidth / 2)

i = X1 / STPX * Form1.ScaleWidth

Grad = Text1.Text

For G = 0 To Grad
Y1 = Y1 + A(G) * V ^ G
Next G

If NV = False Then
Grad2 = Text14.Text

For G = 0 To Grad2
Y2 = Y2 + D(G) * V ^ G
Next G

Y1 = Y1 / Y2
End If

Y1 = Form1.ScaleHeight / 2 - Y1

If Y < Form1.ScaleHeight + KSY + 100 Then ' +1
If Y > -Form1.ScaleHeight / 2 + KSY + 1 Then
If Form1.ScaleWidth / 2 + Text11.Text < i Then
If Form1.ScaleWidth / 2 + Text12.Text > i Then
Form1.Line (X - KSX, Y - KSY)-(i - KSX, Y1 - KSY)
End If
End If
End If
End If

Y = Y1
X = (X1 - 0) / STPX * Form1.ScaleWidth
Y1 = 0
Y2 = 0
Next X1
Form1.DrawWidth = 1
End Function

Private Sub Text20_LostFocus()
If Text20.Text < 1 Then Text20.Text = 1
If Text20.Text > 10 Then Text20.Text = 10
Text20.Text = Int(Text20.Text)
End Sub

Private Sub Text3_Change()
On Error Resume Next
If Option1.Value = True Then
Grad = Text1.Text
Else
Grad = Text14.Text
End If
If Text3.Text <> "" Then
If Text3.Text < 0 Then Text3.Text = 0
If Text3.Text > Grad Then Text3.Text = Grad
If Option1.Value = True Then
Text2.Text = A(Text3.Text)
Slider1.Value = -A(Text3.Text) * 100
Else
Text2.Text = D(Text3.Text)
Slider1.Value = -D(Text3.Text) * 100
End If
End If
End Sub

Private Sub Text3_LostFocus()
On Error Resume Next
If Option1.Value = True Then
Grad = Text1.Text
If Text3.Text <> "" Then
If Text3.Text < 0 Then Text3.Text = 0
End If

If Text3.Text > Grad Then Text3.Text = Grad
Text3.Text = Int(Text3.Text)
Slider1.Value = A(Text3.Text) * -100
Text2.Text = A(Text3.Text)
Else
Grad = Text14.Text
If Text3.Text <> "" Then
If Text3.Text < 0 Then Text3.Text = 0
End If

If Text3.Text > Grad Then Text3.Text = Grad
Text3.Text = Int(Text3.Text)
Slider1.Value = D(Text3.Text) * -100
Text2.Text = D(Text3.Text)
End If
End Sub

Private Sub Text4_Change()
If Text4.Text <> "" Then
If Text4.Text < 0 Then Text4.Text = 0
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
If Check1.Value = 1 Then
Text5.Text = Int(Text4.Text / STPX * (STPY) * 100) / 100
End If
End If
End Sub

Private Sub Text4_LostFocus()
If Check1.Value = 0 Then
Else
Text5.Text = Int(Text4.Text / STPX * (STPY) * 100) / 100 '*** vielleicht als Konstante definieren
End If
End Sub

Private Sub Text11_LostFocus()
If Text11.Text <> "" Then
If Text11.Text >= Text12.Text Then Text11.Text = Text12.Text - 1
End If
End Sub

' *** Bildschirmauflösung nur einmal am Anfang erechnen und als Konstante übergeben --> schnelleres Zeichnen des Graphen
Private Sub Timer1_Timer()
'If SliderValue = Slider1.Value Then
If Faktor <> Slider1.Value Then
If Text3.Text <> "" Then
Form1.Cls
A(Text3.Text) = -Slider1.Value / 100
For G = 0 To Grad
Y = Y + A(G) * (-Form1.ScaleWidth / 2) ^ G
Next G

X = 0

Call Raster
Call Nullpunkt
Call Koordinaten
Call Graph

End If
End If
'End If
'SliderValue = Slider1.Value
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

Text2.Text = -Slider1.Value / 100
Form1.Cls
If Option1.Value = True Then
Grad = Text1.Text
A(Text3.Text) = -Slider1.Value / 100
For G = 0 To Grad
Y = Y + A(G) * (-Form1.ScaleWidth / 2) ^ G
Next G
Else
Grad = Text14.Text
D(Text3.Text) = -Slider1.Value / 100
For G = 0 To Grad
Y = Y + A(G) * (-Form1.ScaleWidth / 2) ^ G
Next G
End If
X = 0

Call Raster
Call Nullpunkt
Call Koordinaten
Call Graph

End If
End Sub

Private Sub Timer3_Timer()
Label1.Caption = Int((X - Form1.ScaleWidth / 2 + KSX) * 100) / 100
Label2.Caption = -Int((Y - Form1.ScaleHeight / 2 + KSY) * 100) / 100
End Sub

'Sub EnableControlsOn(Frm As Form, State As Integer)
'    Dim I   ' Variable deklarieren.
'    For I = 0 To Frm.Controls.Count - 1
'        If Not TypeOf Frm.Controls(I) Is Menu Then
'            Frm.Controls(I).Enabled = State
'        End If
'    Next I
'End Sub


'Text1.Text = Command1.Width / Screen.TwipsPerPixelX / 1280 * 13.32


Private Function Graph2()
On Error Resume Next
X = -100
Form1.DrawWidth = Text20.Text ' mit Screen-Faktoren multiplizieren!
Form1.ForeColor = Picture3.BackColor
For X1 = 1 + KSX * STPX / Form1.ScaleWidth To STPX + KSX * STPX / Form1.ScaleWidth '
V = (X1 / STPX * Form1.ScaleWidth - Form1.ScaleWidth / 2)
'V = X1 / STPX * Form1.ScaleWidth
E = X1 / STPX * Form1.ScaleWidth

Grad = Text1.Text
Grad2 = Text14.Text

ReDim C(Grad - DIFFNR)
For G = 1 To Grad - DIFFNR '+1
C(G - 1) = A(G) * G
Next G

For G = 0 To Grad - 1 - DIFFNR
DiffNA = DiffNA + C(G) * V ^ G
Next G
For i = 0 To Grad - DIFFNR
C(Grad - DIFFNR) = 0
Next i

ReDim C(Grad2)
For G = 1 To Grad2 - DIFFNR '+1
C(G - 1) = D(G) * G
Next G

For G = 0 To Grad2 - 1 - DIFFNR
DiffZA = DiffZA + C(G) * V ^ G
Next G
For i = 0 To Grad2 - DIFFNR
C(Grad2 - DIFFNR) = 0
Next i

For G = 0 To Grad - DIFFNR
DiffN = DiffN + A(G) * V ^ G
Next G

For G = 0 To Grad2 - DIFFNR
DiffZ = DiffZ + D(G) * V ^ G
Next G



Y1 = (DiffNA * DiffZ - DiffN * DiffZA) / (DiffZ ^ 2)


Y1 = Form1.ScaleHeight / 2 - Y1

If Y < Form1.ScaleHeight + KSY + 100 Then ' +1
If Y > -Form1.ScaleHeight / 2 + KSY + 1 Then
If Form1.ScaleWidth / 2 + Text11.Text < i Then
If Form1.ScaleWidth / 2 + Text12.Text > i Then
Form1.Line (X - KSX, Y - KSY)-(E - KSX, Y1 - KSY)
End If
End If
End If
End If

Y = Y1
X = (X1 - 0) / STPX * Form1.ScaleWidth
'X = (X1 / STPX * Form1.ScaleWidth - Form1.ScaleWidth / 2)
Y1 = 0
Y2 = 0
DiffN = 0
DiffNA = 0
DiffZ = 0
DiffZA = 0
Next X1
Form1.DrawWidth = 1

DIFFNR = DIFFNR + 1
End Function


Private Function Graph3()
On Error Resume Next
X = -100
Form1.DrawWidth = Text20.Text ' mit Screen-Faktoren multiplizieren!
Form1.ForeColor = Picture3.BackColor
For X1 = 1 + KSX * STPX / Form1.ScaleWidth To STPX + KSX * STPX / Form1.ScaleWidth
V = (X1 / STPX * Form1.ScaleWidth - Form1.ScaleWidth / 2)

i = X1 / STPX * Form1.ScaleWidth

For G = 0 To Grad3
Y1 = Y1 + Z(G) * V ^ G
Next G

Y1 = Form1.ScaleHeight / 2 - Y1

If Y < Form1.ScaleHeight + KSY + 100 Then ' +1
If Y > -Form1.ScaleHeight / 2 + KSY + 1 Then
If Form1.ScaleWidth / 2 + Text11.Text < i Then
If Form1.ScaleWidth / 2 + Text12.Text > i Then
Form1.Line (X - KSX, Y - KSY)-(i - KSX, Y1 - KSY)
End If
End If
End If
End If

Y = Y1
X = (X1 - 0) / STPX * Form1.ScaleWidth
Y1 = 0
Y2 = 0
Next X1
Form1.DrawWidth = 1
End Function
