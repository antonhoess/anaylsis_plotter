VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Koeffizienten"
   ClientHeight    =   1530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2295
   ControlBox      =   0   'False
   FillStyle       =   0  'Ausgefüllt
   LinkTopic       =   "Form2"
   ScaleHeight     =   1530
   ScaleWidth      =   2295
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command3 
      Caption         =   "Ende"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Zurück"
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Eingeben"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   855
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
      Caption         =   "Geben Sie den Koeffizienten für den 0-ten Grad ein!"
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

Private Sub Command1_Click()
Label1.Caption = "Geben Sie den Koeffizienten für den " & GGN + 1 & "-ten Grad ein!"
A(GGN) = Text1.Text
GGN = GGN + 1
If Grad - GGN = -1 Then Unload Me
End Sub

Private Sub Command2_Click()
If GGN = 0 Then MsgBox "Jetzt geht es nicht mehr weiter", vbOKOnly, "Hinweis"
If GGN <> 0 Then
GGN = GGN - 1
Label1.Caption = "Geben Sie den Koeffizienten für den " & GGN & "-ten Grad ein!"
 Text1.Text = A(GGN)
End If
End Sub

Private Sub Command3_Click()
For T = GGN To Grad + 1
A(GGN) = 0
Next T
Unload Me
End Sub

Private Sub Form_Load()
GGN = 0
ReDim A(Grad + 1)
End Sub
