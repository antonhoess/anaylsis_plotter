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
   Begin VB.Frame Frame3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Steuerung"
      Height          =   7095
      Left            =   600
      TabIndex        =   7
      Top             =   3840
      Width           =   4215
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080C0FF&
         Caption         =   "GO!"
         Height          =   375
         Left            =   240
         Style           =   1  'Grafisch
         TabIndex        =   38
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   840
         TabIndex        =   37
         Text            =   "3"
         Top             =   2400
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Clear"
         Height          =   375
         Left            =   240
         Style           =   1  'Grafisch
         TabIndex        =   36
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080C0FF&
         Caption         =   "Koeffizienten"
         Height          =   375
         Left            =   240
         Style           =   1  'Grafisch
         TabIndex        =   35
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0080C0FF&
         Caption         =   "Trace"
         Height          =   375
         Left            =   240
         Style           =   1  'Grafisch
         TabIndex        =   34
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   840
         TabIndex        =   33
         Text            =   "0"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   240
         TabIndex        =   32
         Text            =   "0"
         Top             =   240
         Width           =   375
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Proportionalität"
         Height          =   1815
         Left            =   1560
         TabIndex        =   25
         Top             =   2640
         Visible         =   0   'False
         Width           =   2175
         Begin VB.CheckBox Check1 
            BackColor       =   &H0080C0FF&
            Caption         =   "Proportional"
            Height          =   255
            Left            =   120
            Style           =   1  'Grafisch
            TabIndex        =   29
            Top             =   1440
            Value           =   1  'Aktiviert
            Width           =   1935
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H0080C0FF&
            Caption         =   "Command5"
            Height          =   495
            Left            =   120
            Style           =   1  'Grafisch
            TabIndex        =   28
            Top             =   960
            Width           =   1935
         End
         Begin VB.TextBox Text5 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1440
            TabIndex        =   27
            Text            =   "10"
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   1440
            TabIndex        =   26
            Text            =   "13"
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Längeneinheiten"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Breiteneinheiten"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H000080FF&
         Caption         =   "Beenden"
         Height          =   615
         Left            =   2040
         Style           =   1  'Grafisch
         TabIndex        =   24
         Top             =   5760
         Width           =   1455
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H0080C0FF&
         Caption         =   "Koeffizienten speichern"
         Height          =   375
         Left            =   240
         Style           =   1  'Grafisch
         TabIndex        =   23
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H0080C0FF&
         Caption         =   "Koeffizienten laden"
         Height          =   375
         Left            =   240
         Style           =   1  'Grafisch
         TabIndex        =   22
         Top             =   4080
         Width           =   1095
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H0080C0FF&
         Caption         =   "Verschieben (Koordinatensystem)"
         Height          =   615
         Left            =   2400
         Style           =   1  'Grafisch
         TabIndex        =   21
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   1800
         TabIndex        =   20
         Text            =   "0"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   1800
         TabIndex        =   19
         Text            =   "0"
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   3240
         TabIndex        =   18
         Text            =   "1"
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H0080C0FF&
         Caption         =   "strecken"
         Height          =   375
         Left            =   1560
         Style           =   1  'Grafisch
         TabIndex        =   17
         Top             =   2160
         Width           =   2415
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Raster"
         Height          =   495
         Left            =   480
         Style           =   1  'Grafisch
         TabIndex        =   16
         Top             =   4560
         Value           =   1  'Aktiviert
         Width           =   1455
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   3240
         TabIndex        =   15
         Text            =   "1"
         Top             =   1800
         Width           =   735
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H0080C0FF&
         Caption         =   "Koordinaten"
         Height          =   495
         Left            =   2040
         Style           =   1  'Grafisch
         TabIndex        =   14
         Top             =   4560
         Value           =   1  'Aktiviert
         Width           =   1455
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H0080C0FF&
         Caption         =   "Achsenkreuz"
         Height          =   495
         Left            =   480
         Style           =   1  'Grafisch
         TabIndex        =   13
         Top             =   5160
         Value           =   1  'Aktiviert
         Width           =   1455
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H0080C0FF&
         Caption         =   "Immer im Vordergrund"
         Height          =   495
         Left            =   2040
         Style           =   1  'Grafisch
         TabIndex        =   12
         Top             =   5160
         Value           =   1  'Aktiviert
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Intervall"
         Height          =   615
         Left            =   480
         TabIndex        =   9
         Top             =   5760
         Width           =   1455
         Begin VB.TextBox Text11 
            Height          =   285
            Left            =   240
            TabIndex        =   11
            Text            =   "-3"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox Text12 
            Height          =   285
            Left            =   840
            TabIndex        =   10
            Text            =   "4"
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H0000C0C0&
         Caption         =   "Ausblenden"
         Height          =   375
         Left            =   480
         Style           =   1  'Grafisch
         TabIndex        =   8
         Top             =   6480
         Width           =   3015
      End
      Begin VB.Label Label19 
         Caption         =   "Verschieben möglich bei gedrückter Maus"
         Height          =   375
         Left            =   1800
         TabIndex        =   48
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   255
         Left            =   600
         TabIndex        =   47
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         Height          =   255
         Left            =   600
         TabIndex        =   46
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Grad"
         Height          =   255
         Left            =   360
         TabIndex        =   45
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "X     Y"
         Height          =   615
         Left            =   360
         TabIndex        =   44
         Top             =   3120
         Width           =   135
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "="
         Height          =   255
         Left            =   600
         TabIndex        =   43
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "X  Y"
         Height          =   615
         Left            =   1560
         TabIndex        =   42
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
         Left            =   1560
         TabIndex        =   41
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
         Left            =   1560
         TabIndex        =   40
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
         Left            =   360
         TabIndex        =   39
         Top             =   2880
         Width           =   1095
      End
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   11295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   19923
      _Version        =   327682
      BorderStyle     =   1
      Orientation     =   1
      Min             =   -100
      Max             =   100
   End
   Begin VB.Label Label18 
      Caption         =   "Differenzieren (Ableitfunktion), Wertebereich errechnen"
      Height          =   495
      Left            =   2760
      TabIndex        =   6
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label11 
      Caption         =   "Proportional bei Raster nach Breite von Form1 angleichen, denn ansonsten stimmt es nicht überein"
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label16 
      Caption         =   "gebr.rat Fkt.; evtl abschnittsweise Definition"
      Height          =   615
      Left            =   600
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Arrays in Datei speichern und aus Datei laden"
      Height          =   855
      Left            =   600
      TabIndex        =   3
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label9 
      Caption         =   "Bei Change Grapg neu zeichnen"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Einstellungen speichern"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   2400
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X, Y, X1, Y1, X2, i, V, G, B As Boolean, W, Faktor, KSX, KSY, SFX, SFY, STPX, STPY, MNS As Boolean, MENX, MENY, MCX, MCY

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
Frame1.Visible = True
On Error Resume Next
If Text1.Text < 0 Then Text1.Text = 0
Text1.Text = Int(Text1.Text)
Grad = Text1.Text
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
  
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = Int((X - Form1.ScaleWidth / 2) * 100) / 100
Label2.Caption = -Int((Y - Form1.ScaleHeight / 2) * 100) / 100

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

If B = True Then Call SetCursorPos(X / Form1.ScaleWidth * STPX, aY / Form1.ScaleHeight * (STPY) + 20)
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

Private Sub Frame3_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Label19_Click()

End Sub

Private Sub Slider1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Faktor <> Slider1.Value Then

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
Form1.Cls
Call Raster
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

Private Function Graph()
On Error Resume Next
X = -100
Form1.DrawWidth = 1 '3

For X1 = 1 To STPX
V = (X1 / STPX * Form1.ScaleWidth - Form1.ScaleWidth / 2)

i = X1 / STPX * Form1.ScaleWidth

For G = 0 To Grad
Y1 = Y1 + A(G) * V ^ G
Next G

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
