VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   11460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9240
   ControlBox      =   0   'False
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7.958
   ScaleMode       =   5  'Zoll
   ScaleWidth      =   6.417
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command6 
      Caption         =   "Beenden"
      Height          =   375
      Left            =   1200
      TabIndex        =   22
      Top             =   10320
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Frame1"
      Height          =   1815
      Left            =   720
      TabIndex        =   14
      Top             =   8040
      Visible         =   0   'False
      Width           =   2175
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1080
         TabIndex        =   18
         Text            =   "13"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   17
         Text            =   "10"
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H0080C0FF&
         Caption         =   "Command5"
         Height          =   495
         Left            =   120
         Style           =   1  'Grafisch
         TabIndex        =   16
         Top             =   960
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Proportional"
         Height          =   255
         Left            =   120
         Style           =   1  'Grafisch
         TabIndex        =   15
         Top             =   1440
         Value           =   1  'Aktiviert
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "scalewidth"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "scaleheight"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   855
      End
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   11295
      Left            =   0
      TabIndex        =   9
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
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   600
      TabIndex        =   12
      Text            =   "0"
      Top             =   4320
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1200
      TabIndex        =   10
      Text            =   "0"
      Top             =   4320
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080C0FF&
      Caption         =   "Trace"
      Height          =   375
      Left            =   600
      Style           =   1  'Grafisch
      TabIndex        =   8
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Koeffizienten"
      Height          =   375
      Left            =   600
      Style           =   1  'Grafisch
      TabIndex        =   5
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Clear"
      Height          =   375
      Left            =   600
      Style           =   1  'Grafisch
      TabIndex        =   2
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Text            =   "2"
      Top             =   6360
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "GO!"
      Height          =   375
      Left            =   600
      Style           =   1  'Grafisch
      TabIndex        =   0
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "Alles in Abhängigkeit der Bildschirmauflösung stellen"
      Height          =   735
      Left            =   2760
      TabIndex        =   23
      Top             =   6360
      Width           =   3615
   End
   Begin VB.Label Label9 
      Caption         =   "Bei Change Grapg neu zeichnen"
      Height          =   495
      Left            =   1920
      TabIndex        =   21
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "="
      Height          =   255
      Left            =   960
      TabIndex        =   13
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "Einstellungen speichern"
      Height          =   495
      Left            =   840
      TabIndex        =   11
      Top             =   7560
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "X     Y"
      Height          =   615
      Left            =   720
      TabIndex        =   7
      Top             =   6840
      Width           =   135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Grad"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   7200
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   6840
      Width           =   615
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   0.417
      X2              =   1.083
      Y1              =   2.417
      Y2              =   2.417
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0.75
      X2              =   0.75
      Y1              =   2.083
      Y2              =   2.75
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x, y, X1, Y1, X2, I, V, G, B As Boolean, W, Faktor

Option Explicit

Private Declare Function SetCursorPos Lib "user32" (ByVal _
        x As Long, ByVal y As Long) As Long

Private Declare Function GetCursorPos Lib "user32" _
        (lpPoint As POINTAPI) As Long

Private Type POINTAPI
  x As Long
  y As Long
End Type

Dim aX, aY, dx, dy 'Dim aX%, aY%, dx%, dy%

Private Sub Check1_Click()
If Check1.Value = 0 Then
Text5.Enabled = True
Else
Text5.Enabled = False
Text5.Text = Int(Text4.Text / 1280 * 1002 * 100) / 100
End If
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

x = -1 '0

Call Graph

End Sub

Private Sub Command2_Click()
Form1.Cls

  Form1.DrawStyle = 2
  Form1.ForeColor = RGB(255, 0, 0)
  
  For I = 0 To Int(Form1.ScaleWidth / 2) + 2
  Form1.Line (Form1.ScaleWidth / 2 - Int(Form1.ScaleWidth / 2) + I, 0)-(Form1.ScaleWidth / 2 - Int(Form1.ScaleWidth / 2) + I, Form1.ScaleHeight)
  Next I
    
  For I = 0 To Int(Form1.ScaleHeight)
  Form1.Line (0, Form1.ScaleHeight / 2 - Int(Form1.ScaleHeight / 2) + I)-(Form1.ScaleWidth, Form1.ScaleHeight / 2 - Int(Form1.ScaleHeight / 2) + I)
  Next I
  
  Form1.ForeColor = RGB(0, 0, 255)
  Form1.DrawStyle = 0
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
Form1.ScaleHeight = Int(Text4.Text / 1280 * 1002 * 100) / 100
End If
' *** Rastergröße je näch größe um das 10-fache vergrößern oder verkleinern
Call Graph
Call Raster
End Sub

Private Sub Command6_Click()
End
End Sub

Private Sub Form_Load()
 Me.WindowState = 2
 
Call Nullpunkt

y = Form1.ScaleHeight / 2
x = 0

 
  B = False
  
  Faktor = Slider1.Value

Call Raster
  
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label1.Caption = Int((x - Form1.ScaleWidth / 2) * 100) / 100
Label2.Caption = -Int((y - Form1.ScaleHeight / 2) * 100) / 100

 Dim Pt As POINTAPI
  
    Call GetCursorPos(Pt)
      'aX = Pt.X

W = x - Form1.ScaleWidth / 2


If B = True Then
For G = 0 To Grad
aY = aY + A(G) * W ^ G
Next G
aY = -aY + Form1.ScaleHeight / 2
End If

'If aY > 0 Then aY = 0
'If aY < Form1.ScaleHeight Then aY = Form1.ScaleHeight
      'aY = Pt.Y

If B = True Then Call SetCursorPos(x / Form1.ScaleWidth * 1280, aY / Form1.ScaleHeight * 1002 + 20)
Label1.Caption = Int((x - Form1.ScaleWidth / 2) * 100) / 100
Label2.Caption = -Int((y - Form1.ScaleHeight / 2) * 100) / 100
aY = 0
End Sub

Private Sub Form_Resize()
Form1.Cls
Call Nullpunkt
Call Raster
End Sub

Private Sub Slider1_Change()
Form1.Cls
Call Raster
Call Graph
End Sub

Private Sub Slider1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Faktor <> Slider1.Value Then

Text2.Text = Slider1.Value / 10
Form1.Cls
A(Text3.Text) = Slider1.Value / 10
For G = 0 To Grad
y = y + A(G) * (-Form1.ScaleWidth / 2) ^ G
Next G

x = 0

Call Raster
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

Private Function Raster()
  
  Form1.DrawStyle = 2
  Form1.ForeColor = RGB(255, 0, 0)
   
  For I = 0 To Int(Form1.ScaleWidth)
  Form1.Line (Form1.ScaleWidth / 2 - Int(Form1.ScaleWidth / 2) + I, 0)-(Form1.ScaleWidth / 2 - Int(Form1.ScaleWidth / 2) + I, Form1.ScaleHeight)
  Next I
  
  For I = 0 To Int(Form1.ScaleHeight)
  Form1.Line (0, Form1.ScaleHeight / 2 - Int(Form1.ScaleHeight / 2) + I)-(Form1.ScaleWidth, Form1.ScaleHeight / 2 - Int(Form1.ScaleHeight / 2) + I)
  Next I
  
  Form1.ForeColor = RGB(0, 0, 255)
  Form1.DrawStyle = 0

End Function

Private Function Nullpunkt()
Line1.X1 = Form1.ScaleWidth / 2
Line1.Y1 = 0
Line1.X2 = Form1.ScaleWidth / 2
Line1.Y2 = Form1.ScaleHeight

Line2.X1 = 0
Line2.Y1 = Form1.ScaleHeight / 2
Line2.X2 = Form1.ScaleWidth
Line2.Y2 = Form1.ScaleHeight / 2
End Function

Private Function Graph()
On Error Resume Next
x = -100
Form1.DrawWidth = 3

For X1 = 1 To 1280
V = (X1 / 1280 * Form1.ScaleWidth - Form1.ScaleWidth / 2)

I = X1 / 1280 * Form1.ScaleWidth

For G = 0 To Grad
Y1 = Y1 + A(G) * V ^ G
Next G

Y1 = Form1.ScaleHeight / 2 - Y1

If y < Form1.ScaleHeight / 1 Then
If y > -Form1.ScaleHeight / 2 Then
Form1.Line (x, y)-(I, Y1)
End If
End If

y = Y1
x = (X1 - 0) / 1280 * Form1.ScaleWidth
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
Text5.Text = Int(Text4.Text / 1280 * 1002 * 100) / 100
End If
End If
End Sub

Private Sub Text4_LostFocus()
If Check1.Value = 0 Then
Else
Text5.Text = Int(Text4.Text / 1280 * 1002 * 100) / 100 '*** vielleicht als Konstante definieren
End If
End Sub
