VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   9270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11760
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6.438
   ScaleMode       =   5  'Zoll
   ScaleWidth      =   8.167
   StartUpPosition =   3  'Windows-Standard
   WindowState     =   2  'Maximiert
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Koeffizienten"
      Height          =   375
      Left            =   600
      Style           =   1  'Grafisch
      TabIndex        =   5
      Top             =   5760
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
      Top             =   5400
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
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Zoom-Funktion und Trace-Funktion hinzufügen über setzen der Cursor-Position ; Koeffizienten übert Slider verändern und dann Graph sofortz wiedert neu zeichnen ; evtl. Raster auch noch hinzufügen"
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   6960
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "X     Y"
      Height          =   615
      Left            =   720
      TabIndex        =   7
      Top             =   6120
      Width           =   135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Grad"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   6480
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   6120
      Width           =   615
   End
   Begin VB.Line Line2 
      X1              =   1.5
      X2              =   3
      Y1              =   4.333
      Y2              =   4.833
   End
   Begin VB.Line Line1 
      X1              =   2.25
      X2              =   2.5
      Y1              =   1.167
      Y2              =   3.083
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X, Y, X1, Y1, X2, I, V, G


'Dim I As Long
'Dim Anzahl As Long

Private Sub Command1_Click()
Y = Form1.ScaleHeight / 2
X = 0
For X1 = 1 To 1280
V = (X1 / 1280 * Form1.ScaleWidth - Form1.ScaleWidth / 2)
'Y1 = Form1.ScaleHeight / 2 - (Text1.Text * V ^ 0 + Text2.Text * V ^ 1 + Text3.Text * V ^ 2 + Text4.Text * V ^ 3)

I = X1 / 1280 * Form1.ScaleWidth ' - Form1.ScaleWidth / 2

For G = 0 To Grad
Y1 = Y1 + A(G) * V ^ G
Next G

Y1 = Form1.ScaleHeight / 2 - Y1
'Y1 = -Y1

If Y < Form1.ScaleHeight / 1 Then
If Y > -Form1.ScaleHeight / 2 Then
Form1.Line (X, Y)-(I, Y1) '*** eigentlich anfangs nur dise Zeile
End If
End If

Y = Y1
X = (X1 - 0) / 1280 * Form1.ScaleWidth
Y1 = 0
Next X1
End Sub

Private Sub Command2_Click()
Form1.Cls
End Sub

Private Sub Command3_Click()
If Text1.Text < 0 Then Text1.Text = 0
Grad = Text1.Text
Form2.Show
End Sub

Private Sub Form_Load()
Line1.X1 = Form1.ScaleWidth / 2
Line1.Y1 = 0
Line1.X2 = Form1.ScaleWidth / 2
Line1.Y2 = Form1.ScaleHeight

Line2.X1 = 0
Line2.Y1 = Form1.ScaleHeight / 2
Line2.X2 = Form1.ScaleWidth
Line2.Y2 = Form1.ScaleHeight / 2

Y = Form1.ScaleHeight / 2
X = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = Int((X - Form1.ScaleWidth / 2) * 100) / 100
Label2.Caption = -Int((Y - Form1.ScaleHeight / 2) * 100) / 100
End Sub

Private Sub Form_Resize()
Line1.X1 = Form1.ScaleWidth / 2
Line1.Y1 = 0
Line1.X2 = Form1.ScaleWidth / 2
Line1.Y2 = Form1.ScaleHeight

Line2.X1 = 0
Line2.Y1 = Form1.ScaleHeight / 2
Line2.X2 = Form1.ScaleWidth
Line2.Y2 = Form1.ScaleHeight / 2
End Sub


'Dim A() As Long
'Dim I As Long
'Dim Anzahl As Long
'
'Anzahl = 1000
'ReDim a(Anzahl)
'
'For i = 0 To Anzahl
'a(i) = i
'Next
'
