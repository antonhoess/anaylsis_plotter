VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   11460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10785
   ControlBox      =   0   'False
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7.958
   ScaleMode       =   5  'Zoll
   ScaleWidth      =   7.49
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text14 
      Height          =   375
      Left            =   6240
      TabIndex        =   56
      Text            =   "0"
      Top             =   5160
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6000
      Top             =   3960
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Steuerung"
      Height          =   8535
      Left            =   840
      TabIndex        =   4
      Top             =   2520
      Width           =   4575
      Begin VB.CheckBox Check6 
         Caption         =   "Gebr. Rat. Fkt"
         Height          =   315
         Left            =   600
         TabIndex        =   55
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   5160
         Width           =   735
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   5160
         Width           =   735
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H0080C0FF&
         Caption         =   "Wertebereich errechnen"
         Height          =   495
         Left            =   1080
         Style           =   1  'Grafisch
         TabIndex        =   47
         Top             =   5520
         Width           =   2775
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H0080C0FF&
         Caption         =   "Differentieren"
         Height          =   375
         Left            =   2040
         Style           =   1  'Grafisch
         TabIndex        =   46
         Top             =   4680
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080C0FF&
         Caption         =   "GO!"
         Height          =   375
         Left            =   720
         Style           =   1  'Grafisch
         TabIndex        =   35
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1320
         TabIndex        =   34
         Text            =   "3"
         Top             =   2400
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Clear"
         Height          =   375
         Left            =   720
         Style           =   1  'Grafisch
         TabIndex        =   33
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080C0FF&
         Caption         =   "Koeffizienten"
         Height          =   375
         Left            =   720
         Style           =   1  'Grafisch
         TabIndex        =   32
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0080C0FF&
         Caption         =   "Trace"
         Height          =   375
         Left            =   720
         Style           =   1  'Grafisch
         TabIndex        =   31
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1320
         TabIndex        =   30
         Text            =   "0"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   720
         TabIndex        =   29
         Text            =   "0"
         Top             =   480
         Width           =   375
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
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Breiteneinheiten"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H000080FF&
         Caption         =   "Beenden"
         Height          =   615
         Left            =   2520
         Style           =   1  'Grafisch
         TabIndex        =   21
         Top             =   7320
         Width           =   1455
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H0080C0FF&
         Caption         =   "Koeffizienten speichern"
         Height          =   195
         Left            =   720
         Style           =   1  'Grafisch
         TabIndex        =   20
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H0080C0FF&
         Caption         =   "Koeffizienten laden"
         Height          =   255
         Left            =   720
         Style           =   1  'Grafisch
         TabIndex        =   19
         Top             =   4920
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
         Left            =   960
         Style           =   1  'Grafisch
         TabIndex        =   13
         Top             =   6120
         Value           =   1  'Aktiviert
         Width           =   1455
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
         Left            =   2520
         Style           =   1  'Grafisch
         TabIndex        =   11
         Top             =   6120
         Value           =   1  'Aktiviert
         Width           =   1455
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H0080C0FF&
         Caption         =   "Achsenkreuz"
         Height          =   495
         Left            =   960
         Style           =   1  'Grafisch
         TabIndex        =   10
         Top             =   6720
         Value           =   1  'Aktiviert
         Width           =   1455
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H0080C0FF&
         Caption         =   "Immer im Vordergrund"
         Height          =   495
         Left            =   2520
         Style           =   1  'Grafisch
         TabIndex        =   9
         Top             =   6720
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Intervall"
         Height          =   615
         Left            =   960
         TabIndex        =   6
         Top             =   7320
         Width           =   1455
         Begin VB.TextBox Text11 
            Height          =   285
            Left            =   240
            TabIndex        =   8
            Text            =   "-10"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox Text12 
            Height          =   285
            Left            =   840
            TabIndex        =   7
            Text            =   "10"
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H0000C0C0&
         Caption         =   "Ausblenden"
         Height          =   375
         Left            =   960
         Style           =   1  'Grafisch
         TabIndex        =   5
         Top             =   8040
         Width           =   3015
      End
      Begin ComctlLib.Slider Slider1 
         Height          =   8175
         Left            =   240
         TabIndex        =   51
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   14420
         _Version        =   327682
         BorderStyle     =   1
         Orientation     =   1
         Min             =   -100
         Max             =   100
         TickStyle       =   3
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Grad      Koeffizient"
         Height          =   255
         Left            =   720
         TabIndex        =   52
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Wertebereich:"
         Height          =   255
         Left            =   1080
         TabIndex        =   50
         Top             =   5160
         Width           =   1215
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Verschieben möglich bei gedrückter Maus"
         Height          =   375
         Left            =   2280
         TabIndex        =   45
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   255
         Left            =   1080
         TabIndex        =   44
         Top             =   4080
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         Height          =   255
         Left            =   1080
         TabIndex        =   43
         Top             =   4440
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Grad"
         Height          =   255
         Left            =   840
         TabIndex        =   42
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "X     Y"
         Height          =   615
         Left            =   840
         TabIndex        =   41
         Top             =   4080
         Width           =   135
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "="
         Height          =   255
         Left            =   1080
         TabIndex        =   40
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "X  Y"
         Height          =   615
         Left            =   2040
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
         Top             =   1800
         Width           =   1455
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
         TabIndex        =   36
         Top             =   3840
         Width           =   1095
      End
   End
   Begin VB.Label Label22 
      Caption         =   "Zuerst Raster, dann Nullpunkt"
      Height          =   255
      Left            =   2760
      TabIndex        =   58
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label21 
      Caption         =   "Grad"
      Height          =   255
      Left            =   5760
      TabIndex        =   57
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label Label20 
      Caption         =   "Value von Slider1 ändern"
      Height          =   255
      Left            =   2760
      TabIndex        =   54
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label18 
      Caption         =   "Bei Trace Verschiebung von Koordinatensystem einberechnen und auch bei Cursorpositionsanzeige"
      Height          =   855
      Left            =   2760
      TabIndex        =   53
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label11 
      Caption         =   "Proportional bei Raster nach Breite von Form1 angleichen, denn ansonsten stimmt es nicht überein"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label16 
      Caption         =   "gebr.rat Fkt.; evtl abschnittsweise Definition"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label10 
      Caption         =   "Arrays in Datei speichern und aus Datei laden"
      Height          =   855
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Einstellungen speichern"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   1320
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X, Y, X1, Y1, Y2, X2, i, V, G, B As Boolean, W, Faktor, KSX, KSY, SFX, SFY, STPX, STPY, MNS As Boolean, MENX, MENY, MCX, MCY, SliderValue

Option Explicit

Private Declare Function SetCursorPos Lib "user32" (ByVal _
        X As Long, ByVal Y As Long) As Long

Private Declare Function GetCursorPos Lib "user32" _
        (lpPoint As POINTAPI) As Long

Private Type POINTAPI
  X As Long
  Y As Long
End Type

Dim aX, aY, dx, dy

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
Call Nullpunkt
Call Raster
Call Koordinaten
Call Graph
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
Call Nullpunkt
Call Raster
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
Call Nullpunkt
Call Raster
Call Koordinaten
Call Graph
Else
Form1.Cls
Call Nullpunkt
Call Raster
Call Graph
End If
End Sub

Private Sub Check4_Click()
If Check4.Value = 1 Then
Call Nullpunkt
Call Raster
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
Else
NV = False
End If

End Sub

Private Sub Command10_Click()
Form1.Cls
SFX = Text9.Text
SFY = Text10.Text
Call Nullpunkt
Call Raster
Call Koordinaten
Call Graph
End Sub

Private Sub Command11_Click()
MENX = Frame3.Left
MENY = Frame3.Top
Frame3.Visible = False
End Sub

Private Sub Command12_Click()
For i = 1 To Grad
A(i - 1) = A(i) * (i)
Next i
A(Grad) = 0
End Sub

Private Sub Command13_Click()
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
End Sub

Private Sub Command4_Click()
If B = True Then
B = False
Else
B = True
End If
End Sub

Private Sub Command1_Click()
'For G = 0 To Grad
'y = y + A(G) * (-Form1.ScaleWidth / 2) ^ G
'Next G
'y = -y + Form1.ScaleHeight / 2

'Y = Form1.ScaleHeight / 2

X = -1 '0

Call Raster
Call Koordinaten
Call Graph

End Sub

Private Sub Command2_Click()
Form1.Cls

Call Nullpunkt
Call Raster
Call Koordinaten

End Sub

Private Sub Command3_Click()
'Frame1.Visible = True
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
Form2.Show (1)
End Sub

Private Sub Command5_Click()
On Error Resume Next
Form1.Cls
Form1.ScaleMode = 0
Form1.ScaleWidth = Text4.Text
If Check1.Value = 0 Then
Form1.ScaleHeight = Text5.Text
Else
Form1.ScaleHeight = Int(Text4.Text / STPX * (STPY) * 100) / 100
End If
' *** Rastergröße je näch größe um das 10-fache vergrößern oder verkleinern
Call Nullpunkt
Call Raster
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
Call Nullpunkt
Call Raster
Call Koordinaten
Call Graph
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
'ReDim A(0)
'ReDim D(1)
'  A(0) = 1
'  D(0) = 0
'  D(1) = 2
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = Int((X - Form1.ScaleWidth / 2) * 100) / 100 + KSX
Label2.Caption = -Int((Y - Form1.ScaleHeight / 2) * 100) / 100 + KSY

 Dim Pt As POINTAPI
  
    Call GetCursorPos(Pt)
      'aX = Pt.X

W = X - Form1.ScaleWidth / 2


If B = True Then
For G = 0 To Grad
aY = aY + A(G) * W ^ G
Next G
aY = -aY + Form1.ScaleHeight / 2
End If

'If aY > 0 Then aY = 0
'If aY < Form1.ScaleHeight Then aY = Form1.ScaleHeight
      'aY = Pt.Y

If B = True Then Call SetCursorPos(X / Form1.ScaleWidth * STPX - KSX, aY / Form1.ScaleHeight * (STPY) + 20 - KSY)
Label1.Caption = Int((X - Form1.ScaleWidth / 2) * 100) / 100
Label2.Caption = -Int((Y - Form1.ScaleHeight / 2) * 100) / 100
aY = 0
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
Call Nullpunkt
Call Raster
Call Koordinaten
Call Graph
End Sub
'
'Private Sub Slider1_Change()
'Form1.Cls
''Call Raster
'Call Graph
'End Sub

Private Sub Slider1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Faktor <> Slider1.Value Then
'Slider1.Value = -Slider1.Value

Text2.Text = Slider1.Value / 10
Form1.Cls
A(Text3.Text) = Slider1.Value / 10
For G = 0 To Grad
Y = Y + A(G) * (-Form1.ScaleWidth / 2) ^ G
Next G

X = 0

Call Nullpunkt
Call Raster
Call Koordinaten
Call Graph

End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Slider1.Value = Text2.Text * 10 ' evtl. mit Int()
A(Text3.Text) = Text2.Text
Form1.Cls
Call Nullpunkt
Call Raster
Call Koordinaten
Call Graph
End If
End Sub

Private Sub Text3_Change()
On Error Resume Next
If Text3.Text <> "" Then
If Text3.Text < 0 Then Text3.Text = 0
End If
If Text3.Text > Grad Then Text3.Text = Grad
Text3.Text = Int(Text3.Text)
Slider1.Value = A(Text3.Text) * 10
Text2.Text = A(Text3.Text)
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

'Private Function Graph()
'On Error Resume Next
'X = -100
'Form1.DrawWidth = 1 '3
'
'For X1 = 1 To STPX
'V = (X1 / STPX * Form1.ScaleWidth - Form1.ScaleWidth / 2)
'
'i = X1 / STPX * Form1.ScaleWidth
'
'For G = 0 To Grad
'Y1 = Y1 + A(G) * V ^ G
'Next G
'
'Y1 = Form1.ScaleHeight / 2 - Y1
'
'If Y < Form1.ScaleHeight + KSY + 1 Then
'If Y > -Form1.ScaleHeight / 2 + KSY + 1 Then
'If Form1.ScaleWidth / 2 + Text11.Text < i Then
'If Form1.ScaleWidth / 2 + Text12.Text > i Then
'Form1.Line (X - KSX, Y - KSY)-(i - KSX, Y1 - KSY)
'End If
'End If
'End If
'End If
'
'Y = Y1
'X = (X1 - 0) / STPX * Form1.ScaleWidth
'Y1 = 0
'Next X1
'Form1.DrawWidth = 1
'End Function

Private Function Graph()
On Error Resume Next
X = -100
Form1.DrawWidth = 1 '3

For X1 = 1 To STPX
V = (X1 / STPX * Form1.ScaleWidth - Form1.ScaleWidth / 2)

i = X1 / STPX * Form1.ScaleWidth

Grad = Text1.Text

For G = 0 To Grad
Y1 = Y1 + A(G) * V ^ G
Next G

Grad = Text14.Text

For G = 0 To Grad
Y2 = Y2 + D(G) * V ^ G
Next G

Y1 = Y1 / Y2

Y1 = Form1.ScaleHeight / 2 - Y1

If Y < Form1.ScaleHeight + KSY + 1 Then
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

Private Sub Text4_Change()
If Text4.Text < 0 Then Text4.Text = 0
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

' *** Bildschirmauflösung nur einmal am Anfang erechnen und als Konstante übergeben --> schnelleres Zeichnen des Graphen
Private Sub Timer1_Timer()
If SliderValue = Slider1.Value Then
If Faktor <> Slider1.Value Then

'Text2.Text = Slider1.Value / 10
Form1.Cls
A(Text3.Text) = Slider1.Value / 10
For G = 0 To Grad
Y = Y + A(G) * (-Form1.ScaleWidth / 2) ^ G
Next G

X = 0

Call Nullpunkt
Call Raster
Call Koordinaten
Call Graph

End If
End If
SliderValue = Slider1.Value
End Sub
