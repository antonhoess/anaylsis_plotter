VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Koeffizienten"
   ClientHeight    =   1965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2295
   ControlBox      =   0   'False
   FillStyle       =   0  'Ausgefüllt
   LinkTopic       =   "Form2"
   ScaleHeight     =   1965
   ScaleWidth      =   2295
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command3 
      Caption         =   "Ende"
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Weiter"
      Height          =   375
      Left            =   840
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Zurück"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Eingeben"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      Style           =   1  'Grafisch
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      Caption         =   "Geben Sie den Koeffizienten für den 0-ten Grad ein!"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public GGN, KZ, KZ2
Dim T, Sprung As Boolean

Private Sub Command1_Click()
If KZ2 = KZ Then KZ = KZ + 1
KZ2 = KZ2 + 1
Sprung = False
If GGN = -1 Then GGN = 0

If Form2.Visible = True Then
Text1.SetFocus
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End If

If NV = True Then
Label1.Caption = "Geben Sie den Koeffizienten für den " & GGN + 1 & "-ten Grad ein!"
A(GGN) = Text1.Text
GGN = GGN + 1
Else
If GRF = True Then
If Grad <> 1 Then
A(GGN) = Text1.Text
Else
GGN = 1
Text1.Text = A(GGN - 1)
End If '***
GGN = GGN + 1
If GGN = Grad + 1 Then
Label1.Caption = "Geben Sie den Koeffizienten für den " & 0 & "-ten Nenner-Grad ein!"
Beep
Else
Label1.Caption = "Geben Sie den Koeffizienten für den " & GGN + 0 & "-ten Zähler-Grad ein!"
End If

Else
If GGN = Grad Then
Label1.Caption = "Geben Sie den Koeffizienten für den " & 0 & "-ten Nenner-Grad ein!"
Else
Label1.Caption = "Geben Sie den Koeffizienten für den " & GGN + 1 & "-ten Nenner-Grad ein!"
End If
D(GGN) = Text1.Text
GGN = GGN + 1
End If
End If

If Grad - GGN = -1 Then
If NV = True Then
Unload Me
Else
If GRF = True Then
GRF = False
GGN = 0
Grad = Form1.Text14.Text
Else
Unload Me
End If
End If
End If

End Sub

Private Sub Command2_Click()
'If GGN <> 0 Then KZ2 = KZ2 - 1
If NV = True Then
If GGN = 0 Then MsgBox "Jetzt geht es nicht mehr weiter", vbOKOnly, "Hinweis"
If GGN <> 0 Then
KZ2 = KZ2 - 1
GGN = GGN - 1
Label1.Caption = "Geben Sie den Koeffizienten für den " & GGN & "-ten Grad ein!"
 Text1.Text = A(GGN)
End If
Else
If GRF = True Then
If GGN = 0 Then MsgBox "Jetzt geht es nicht mehr weiter", vbOKOnly, "Hinweis"
If Sprung = True Then Sprung = False: GGN = Grad + 1: Grad = Grad + 1 '***
If GGN <> 0 Then
KZ2 = KZ2 - 1
GGN = GGN - 1
Text1.Text = A(GGN)
End If
If GGN = 0 Then
Label1.Caption = "Geben Sie den Koeffizienten für den " & 0 & "-ten Nenner-Grad ein!"
Else
Label1.Caption = "Geben Sie den Koeffizienten für den " & GGN + 0 & "-ten Zähler-Grad ein!"
End If
Else
If GGN <> 0 Then
KZ2 = KZ2 - 1
GGN = GGN - 1
Text1.Text = D(GGN)
Label1.Caption = "Geben Sie den Koeffizienten für den " & GGN + 0 & "-ten Nenner-Grad ein!"
Else 'If GGN = 0 Then
'Label1.Caption = "Geben Sie den Koeffizienten für den " & 0 & "-ten Nenner-Grad ein!"
If Grad <> 1 Then
GGN = Grad + 1
GRF = True
Label1.Caption = "Geben Sie den Koeffizienten für den " & GGN + 0 & "-ten Zähler-Grad ein!"
Text1.Text = A(GGN + 0) '+1
Sprung = True
Else
GGN = Grad + 1
GRF = True
Label1.Caption = "Geben Sie den Koeffizienten für den 0-ten Zähler-Grad ein!"
Text1.Text = A(0)
Sprung = True
End If
End If
End If
End If

End Sub

Private Sub Command3_Click()
If NV = True Then
For T = GGN To Grad + 1
A(GGN) = 0
Next T
Else
If GRF = True Then
For T = GGN To Grad + 1
A(GGN) = True
Next T
Grad = Form1.Text14.Text
For T = GGN To Grad + 1
D(GGN) = True
Next T
Else
Grad = Form1.Text14.Text
For T = GGN To Grad + 1
D(GGN) = 0
Next T
End If
End If

Unload Me
End Sub

Private Sub Command4_Click()

If NV = True Then
Label1.Caption = "Geben Sie den Koeffizienten für den " & GGN + 1 & "-ten Grad ein!"
If KZ2 < KZ Then
Text1.Text = A(GGN)
Else
A(GGN) = 0
Text1.Text = 0
End If
GGN = GGN + 1
Else
If GRF = True Then
A(GGN) = 0
Text1.Text = 0
GGN = GGN + 1
If GGN = Grad + 1 Then
Label1.Caption = "Geben Sie den Koeffizienten für den " & 0 & "-ten Nenner-Grad ein!"
Else
Label1.Caption = "Geben Sie den Koeffizienten für den " & GGN + 0 & "-ten Zähler-Grad ein!"
End If

Else
If GGN = Grad Then
Label1.Caption = "Geben Sie den Koeffizienten für den " & 0 & "-ten Nenner-Grad ein!"
Else
Label1.Caption = "Geben Sie den Koeffizienten für den " & GGN + 1 & "-ten Nenner-Grad ein!"
End If
If KZ2 < KZ Then
Text1.Text = D(GGN)
Else
D(GGN) = 0
Text1.Text = 0
End If
GGN = GGN + 1
End If
End If

Text1.SetFocus
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)

If Grad - GGN = -1 Then
If NV = True Then
Unload Me
Else
If GRF = True Then
GRF = False
GGN = 0
Grad = Form1.Text14.Text
Else
Unload Me
End If
End If
End If

End Sub

Private Sub Form_Activate()

Text1.SetFocus
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)

End Sub
'
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyF1 Then Unload Me
'End Sub

Private Sub Form_Load()
Form2.KeyPreview = True
If Form1.Check6.Value = 1 Then
GRF = True
Else
GRF = False
End If

If Form1.Check5.Value = 1 Then
'Form dauerhaft in den Vordergrund setzen
Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, 3)
Else
'Form dauerhaft in den Vordergrund setzen
Call SetWindowPos(Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, 3)
End If

GGN = 0
ReDim A(Grad + 1)
ReDim D(Form1.Text14.Text + 14)

If NV = False Then
Label1.Caption = "Geben Sie den Koeffizienten für den " & 0 & "-ten Zähler-Grad ein!"
Else
Label1.Caption = "Geben Sie den Koeffizienten für den " & 0 & "-ten Grad ein!"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If NV = False Then
Form1.Option2.Enabled = True
Else
Form1.Option2.Enabled = False
Form1.Option1.Value = True
End If
'Form1.Command1.SetFocus
Grad = Form1.Text1.Text
KZ = 0
End Sub

'Private Sub Text1_KeyPress(KeyAscii As Integer)
'If KeyAscii = &H25 Then Unload Me
'End Sub
Private Sub Text1_Change()

End Sub
