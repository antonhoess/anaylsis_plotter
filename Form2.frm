VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Koeffizienten"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2295
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   2295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "Eingeben"
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
Public B, GGN  '=GradGegenNull
'Dim A()

Private Sub Command1_Click()
Label1.Caption = "Geben Sie den Koeffizienten für den " & GGN + 1 & "-ten Grad ein!"
A(GGN) = Text1.Text
GGN = GGN + 1
If Grad - GGN = -1 Then Unload Me
End Sub

Private Sub Form_Load()
GGN = 0
ReDim A(Grad + 1)
End Sub
