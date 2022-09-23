VERSION 5.00
Begin VB.Form FrmCoefficients 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Koeffizienten"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2295
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   2295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      Caption         =   "Ende"
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080FF80&
      Caption         =   "Weiter"
      Height          =   375
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000C0C0&
      Caption         =   "Zurück"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Eingeben"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
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
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Geben Sie den Koeffizienten für den 0-ten Grad ein!"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "FrmCoefficients"
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

If FrmCoefficients.Visible = True Then
Text1.SetFocus
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End If

If IsNotRationalFunction = True Then
Label1.Caption = "Geben Sie den Koeffizienten für den " & GGN + 1 & "-ten DegNum ein!"
CoefNum(GGN) = Text1.Text
GGN = GGN + 1
Else
If GRF = True Then
If DegNum <> 1 Then
CoefNum(GGN) = Text1.Text
Else
GGN = 1
Text1.Text = CoefNum(GGN - 1)
End If '***
GGN = GGN + 1
If GGN = DegNum + 1 Then
Label1.Caption = "Geben Sie den Koeffizienten für den " & 0 & "-ten Nenner-Grad ein!"
Beep
Else
Label1.Caption = "Geben Sie den Koeffizienten für den " & GGN + 0 & "-ten Zähler-Grad ein!"
End If

Else
If GGN = DegNum Then
Label1.Caption = "Geben Sie den Koeffizienten für den " & 0 & "-ten Nenner-Grad ein!"
Else
Label1.Caption = "Geben Sie den Koeffizienten für den " & GGN + 1 & "-ten Nenner-Grad ein!"
End If
CoefDen(GGN) = Text1.Text
GGN = GGN + 1
End If
End If

If DegNum - GGN = -1 Then
If IsNotRationalFunction = True Then
Unload Me
Else
If GRF = True Then
GRF = False
GGN = 0
DegNum = FrmMain.TxtDegreeDenominator.Text
Else
Unload Me
End If
End If
End If

End Sub

Private Sub Command2_Click()
'If GGN <> 0 Then KZ2 = KZ2 - 1
If IsNotRationalFunction = True Then
If GGN = 0 Then MsgBox "Jetzt geht es nicht mehr weiter", vbOKOnly, "Hinweis"
If GGN <> 0 Then
KZ2 = KZ2 - 1
GGN = GGN - 1
Label1.Caption = "Geben Sie den Koeffizienten für den " & GGN & "-ten DegNum ein!"
 Text1.Text = CoefNum(GGN)
End If
Else
If GRF = True Then
If GGN = 0 Then MsgBox "Jetzt geht es nicht mehr weiter", vbOKOnly, "Hinweis"
If Sprung = True Then Sprung = False: GGN = DegNum + 1: DegNum = DegNum + 1 '***
If GGN <> 0 Then
KZ2 = KZ2 - 1
GGN = GGN - 1
Text1.Text = CoefNum(GGN)
End If
If GGN = 0 Then
Label1.Caption = "Geben Sie den Koeffizienten für den " & 0 & "-ten Nenner-DegNum ein!"
Else
Label1.Caption = "Geben Sie den Koeffizienten für den " & GGN + 0 & "-ten Zähler-DegNum ein!"
End If
Else
If GGN <> 0 Then
KZ2 = KZ2 - 1
GGN = GGN - 1
Text1.Text = CoefDen(GGN)
Label1.Caption = "Geben Sie den Koeffizienten für den " & GGN + 0 & "-ten Nenner-DegNum ein!"
Else 'If GGN = 0 Then
'Label1.Caption = "Geben Sie den Koeffizienten für den " & 0 & "-ten Nenner-DegNum ein!"
If DegNum <> 1 Then
GGN = DegNum + 1
GRF = True
Label1.Caption = "Geben Sie den Koeffizienten für den " & GGN + 0 & "-ten Zähler-DegNum ein!"
Text1.Text = CoefNum(GGN + 0) '+1
Sprung = True
Else
GGN = DegNum + 1
GRF = True
Label1.Caption = "Geben Sie den Koeffizienten für den 0-ten Zähler-DegNum ein!"
Text1.Text = CoefNum(0)
Sprung = True
End If
End If
End If
End If

End Sub

Private Sub BtnCoefficients_Click()
    If IsNotRationalFunction = True Then
        For T = GGN To DegNum + 1
            CoefNum(GGN) = 0
        Next T
    Else
        If GRF = True Then
            For T = GGN To DegNum + 1
                CoefNum(GGN) = True
            Next T
            DegNum = FrmMain.TxtDegreeDenominator.Text
            For T = GGN To DegNum + 1
                CoefDen(GGN) = True
            Next T
        Else
            DegNum = FrmMain.TxtDegreeDenominator.Text
            For T = GGN To DegNum + 1
                CoefDen(GGN) = 0
            Next T
        End If
    End If
    
    Unload Me
End Sub

Private Sub BtnTrace_Click()

If IsNotRationalFunction = True Then
Label1.Caption = "Geben Sie den Koeffizienten für den " & GGN + 1 & "-ten DegNum ein!"
If KZ2 < KZ Then
Text1.Text = CoefNum(GGN)
Else
CoefNum(GGN) = 0
Text1.Text = 0
End If
GGN = GGN + 1
Else
If GRF = True Then
CoefNum(GGN) = 0
Text1.Text = 0
GGN = GGN + 1
If GGN = DegNum + 1 Then
Label1.Caption = "Geben Sie den Koeffizienten für den " & 0 & "-ten Nenner-DegNum ein!"
Else
Label1.Caption = "Geben Sie den Koeffizienten für den " & GGN + 0 & "-ten Zähler-DegNum ein!"
End If

Else
If GGN = DegNum Then
Label1.Caption = "Geben Sie den Koeffizienten für den " & 0 & "-ten Nenner-DegNum ein!"
Else
Label1.Caption = "Geben Sie den Koeffizienten für den " & GGN + 1 & "-ten Nenner-DegNum ein!"
End If
If KZ2 < KZ Then
Text1.Text = CoefDen(GGN)
Else
CoefDen(GGN) = 0
Text1.Text = 0
End If
GGN = GGN + 1
End If
End If

Text1.SetFocus
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)

If DegNum - GGN = -1 Then
If IsNotRationalFunction = True Then
Unload Me
Else
If GRF = True Then
GRF = False
GGN = 0
DegNum = FrmMain.TxtDegreeDenominator.Text
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

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 1 Then
Select Case KeyCode

Case vbKeyLeft

'If GGN <> 0 Then KZ2 = KZ2 - 1
If IsNotRationalFunction = True Then
If GGN = 0 Then MsgBox "Jetzt geht es nicht mehr weiter", vbOKOnly, "Hinweis"
If GGN <> 0 Then
KZ2 = KZ2 - 1
GGN = GGN - 1
Label1.Caption = "Geben Sie den Koeffizienten für den " & GGN & "-ten DegNum ein!"
 Text1.Text = CoefNum(GGN)
End If
Else
If GRF = True Then
If GGN = 0 Then MsgBox "Jetzt geht es nicht mehr weiter", vbOKOnly, "Hinweis"
If Sprung = True Then Sprung = False: GGN = DegNum + 1: DegNum = DegNum + 1 '***
If GGN <> 0 Then
KZ2 = KZ2 - 1
GGN = GGN - 1
Text1.Text = CoefNum(GGN)
End If
If GGN = 0 Then
Label1.Caption = "Geben Sie den Koeffizienten für den " & 0 & "-ten Nenner-DegNum ein!"
Else
Label1.Caption = "Geben Sie den Koeffizienten für den " & GGN + 0 & "-ten Zähler-DegNum ein!"
End If
Else
If GGN <> 0 Then
KZ2 = KZ2 - 1
GGN = GGN - 1
Text1.Text = CoefDen(GGN)
Label1.Caption = "Geben Sie den Koeffizienten für den " & GGN + 0 & "-ten Nenner-DegNum ein!"
Else 'If GGN = 0 Then
'Label1.Caption = "Geben Sie den Koeffizienten für den " & 0 & "-ten Nenner-DegNum ein!"
If DegNum <> 1 Then
GGN = DegNum + 1
GRF = True
Label1.Caption = "Geben Sie den Koeffizienten für den " & GGN + 0 & "-ten Zähler-DegNum ein!"
Text1.Text = CoefNum(GGN + 0) '+1
Sprung = True
Else
GGN = DegNum + 1
GRF = True
Label1.Caption = "Geben Sie den Koeffizienten für den 0-ten Zähler-DegNum ein!"
Text1.Text = CoefNum(0)
Sprung = True
End If
End If
End If
End If


Case vbKeyRight

If IsNotRationalFunction = True Then
Label1.Caption = "Geben Sie den Koeffizienten für den " & GGN + 1 & "-ten DegNum ein!"
If KZ2 < KZ Then
Text1.Text = CoefNum(GGN)
Else
CoefNum(GGN) = 0
Text1.Text = 0
End If
GGN = GGN + 1
Else
If GRF = True Then
CoefNum(GGN) = 0
Text1.Text = 0
GGN = GGN + 1
If GGN = DegNum + 1 Then
Label1.Caption = "Geben Sie den Koeffizienten für den " & 0 & "-ten Nenner-DegNum ein!"
Else
Label1.Caption = "Geben Sie den Koeffizienten für den " & GGN + 0 & "-ten Zähler-DegNum ein!"
End If

Else
If GGN = DegNum Then
Label1.Caption = "Geben Sie den Koeffizienten für den " & 0 & "-ten Nenner-DegNum ein!"
Else
Label1.Caption = "Geben Sie den Koeffizienten für den " & GGN + 1 & "-ten Nenner-DegNum ein!"
End If
If KZ2 < KZ Then
Text1.Text = CoefDen(GGN)
Else
CoefDen(GGN) = 0
Text1.Text = 0
End If
GGN = GGN + 1
End If
End If

Text1.SetFocus
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)

If DegNum - GGN = -1 Then
If IsNotRationalFunction = True Then
Unload Me
Else
If GRF = True Then
GRF = False
GGN = 0
DegNum = FrmMain.TxtDegreeDenominator.Text
Else
Unload Me
End If
End If
End If


End Select
End If
End Sub

Private Sub Form_Load()
    FrmCoefficients.KeyPreview = True
    
    If FrmMain.ChkRationalFunction.Value = 1 Then
        GRF = True
    Else
        GRF = False
    End If
    
    GGN = 0
    ReDim CoefNum(DegNum + 1)
    ReDim CoefDen(FrmMain.TxtDegreeDenominator.Text + 14)
    
    If IsNotRationalFunction = False Then
        Label1.Caption = "Geben Sie den Koeffizienten für den " & 0 & "-ten Zähler-DegNum ein!"
    Else
        Label1.Caption = "Geben Sie den Koeffizienten für den " & 0 & "-ten DegNum ein!"
    End If
End Sub

Private Sub FrmCoefficients_Unload(Cancel As Integer)
    If IsNotRationalFunction = False Then
        FrmMain.OptDenominator.Enabled = True
    Else
        FrmMain.OptDenominator.Enabled = False
        FrmMain.OptNumerator.Value = True
    End If
    
    'FrmMain.Command1.SetFocus
    DegNum = FrmMain.Text1.Text
    KZ = 0
End Sub

'Private Sub Text1_KeyPress(KeyAscii As Integer)
'If KeyAscii = &H25 Then Unload Me
'End Sub

