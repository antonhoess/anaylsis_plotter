VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Koeffizienten"
   ClientHeight    =   1965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2295
   ControlBox      =   0   'False
   FillStyle       =   0  'Ausgef�llt
   LinkTopic       =   "Form2"
   ScaleHeight     =   1965
   ScaleWidth      =   2295
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command3 
      Caption         =   "Ende"
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Weiter"
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Zur�ck"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Eingeben"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      Caption         =   "Geben Sie den Koeffizienten f�r den 0-ten Grad ein!"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public GGN   '=GradGegenNull
Dim T

'*** Bei Form2.Load und bei Eingeben Textfeld aktivieren, Text aktualisieren

'Private Sub Command1_Click()
'If Grad - GGN = 0 Then
'If GRF = False Then
''Unload Me
'Else
'GRF = False
'Grad = Form1.Text14.Text
'GGN = 0
'End If
'End If
'
''If Grad - GGN = -1 Then
''If GRF = False Then
''Unload Me
''End If
''End If
'
''Label1.Caption = "Geben Sie den Koeffizienten f�r den " & GGN + 1 & "-ten Grad ein!"
'If NV = False Then
'If GRF = True Then
'Label1.Caption = "Geben Sie den Koeffizienten f�r den " & GGN + 1 & "-ten Grad ein!"
'A(GGN) = Text1.Text
'GGN = GGN + 1
'Else
'Label1.Caption = "Geben Sie den Koeffizienten f�r den " & GGN + 0 & "-ten Grad ein!"
'D(GGN) = Form1.Text14.Text
'GGN = GGN + 1
'End If
'Else
'Label1.Caption = "Geben Sie den Koeffizienten f�r den " & GGN + 1 & "-ten Grad ein!"
'A(GGN) = Text1.Text
'GGN = GGN + 1
'End If
'
''If NV = False Then
''If Grad - GGN = 0 Then
''If GRF = False Then
''GGN = 0
''End If
''End If
''End If
'
''GGN = GGN + 1
''If Grad - GGN = -1 Then Unload Me
'
''''If Grad - GGN = -1 Then
''''If GRF = False Then
''''Unload Me
''''Else
''''GRF = False
''''Grad = Form1.Text14.Text
''''GGN = 0
''''End If
''''End If
'
'''If NV = False Then
'''If Grad - GGN = 0 Then
'''If GRF = False Then
'''GGN = 0
'''End If
'''End If
'''End If
'End Sub

Private Sub Command1_Click()
'Label1.Caption = "Geben Sie den Koeffizienten f�r den " & GGN + 1 & "-ten Grad ein!"
'A(GGN) = Text1.Text
'GGN = GGN + 1
'If Grad - GGN = -1 Then Unload Me


If NV = True Then
Label1.Caption = "Geben Sie den Koeffizienten f�r den " & GGN + 1 & "-ten Grad ein!"
A(GGN) = Text1.Text
GGN = GGN + 1
Else
If GRF = True Then
Label1.Caption = "Geben Sie den Koeffizienten f�r den " & GGN + 1 & "-ten Grad ein!"
A(GGN) = Text1.Text
GGN = GGN + 1
Else
If GGN = 0 Then
Label1.Caption = "Geben Sie den Koeffizienten f�r den " & GGN + 1 & "-ten Grad ein!"
Else
Label1.Caption = "Geben Sie den Koeffizienten f�r den " & GGN + 1 & "-ten Grad ein!"
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
If GGN = 0 Then MsgBox "Jetzt geht es nicht mehr weiter", vbOKOnly, "Hinweis"
If GGN <> 0 Then
GGN = GGN - 1
Label1.Caption = "Geben Sie den Koeffizienten f�r den " & GGN & "-ten Grad ein!"
 Text1.Text = A(GGN)
End If
End Sub

Private Sub Command3_Click()
For T = GGN To Grad + 1
If GRF = True Then
A(GGN) = 0
Else
D(GGN) = 0
End If
Next T
Unload Me
End Sub

Private Sub Command4_Click()
Text1.Text = 0
Label1.Caption = "Geben Sie den Koeffizienten f�r den " & GGN + 1 & "-ten Grad ein!"
A(GGN) = 0
GGN = GGN + 1
If Grad - GGN = -1 Then Unload Me
End Sub

Private Sub Form_Load()
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
End Sub
