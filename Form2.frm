VERSION 5.00
Begin VB.Form FrmCoefficients 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Koeffizienten"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2295
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   2295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BtnExit 
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
   Begin VB.CommandButton BtnNext 
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
   Begin VB.CommandButton BtnPrev 
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
   Begin VB.CommandButton BtnEnter 
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
   Begin VB.TextBox TxtCoef 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label LblEntryInfo 
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
Option Explicit

Dim NumeratorActive As Boolean
Dim DegCur As Integer

' XXX
'Private Function IsInteger(Number As String) As Boolean
'    IsInteger = False
'
'    If IsNumeric(Number) Then
'        If CInt(Number) = CDbl(Number) Then
'            IsInteger = True
'        End If
'    End If
'End Function


Private Sub LoadCoefficients()
    ' Load current coefficient into text box and select it for faster entering
    If NumeratorActive Then
        TxtCoef.Text = CoefNum(DegCur)
    Else
        TxtCoef.Text = CoefDen(DegCur)
    End If
    
    If FrmCoefficients.Visible = True Then
        TxtCoef.SetFocus
        TxtCoef.SelStart = 0
        TxtCoef.SelLength = Len(TxtCoef.Text)
    End If

    ' Update label with degree information
    If Not IsRationalFunction Then
        LblEntryInfo.Caption = "Geben Sie den Koeffizienten für den " & DegCur & "-ten Grad ein!"
    Else
        If NumeratorActive Then
            LblEntryInfo.Caption = "Geben Sie den Koeffizienten für den " & DegCur & "-ten Nenner-Grad ein!"
        Else
            LblEntryInfo.Caption = "Geben Sie den Koeffizienten für den " & DegCur & "-ten Zähler-Grad ein!"
        End If
    End If
End Sub

Private Sub NextCoef()
    ' Check is entry is valid
    If Not IsNumeric(TxtCoef.Text) Then
        MsgBox "Please enter a numeric value!"
        Exit Sub
    End If
    
    ' Store entered value
    If NumeratorActive Then
        CoefNum(DegCur) = CDbl(TxtCoef.Text)
    Else
        CoefDen(DegCur) = CDbl(TxtCoef.Text)
    End If
    
    ' Switch to next degree, switch from numerator to denominator (is function is rational) and end value entry when entered last value
    If Not IsRationalFunction Then
        If DegCur < DegNum Then
            DegCur = DegCur + 1
        Else
            Unload Me
            Exit Sub
        End If
    Else
        If NumeratorActive Then
            If DegCur < DegNum Then
                DegCur = DegCur + 1
            Else
                DegCur = 0
                NumeratorActive = False
                Beep
            End If
        Else
            If DegCur < DegDen Then
                DegCur = DegCur + 1
            Else
                Unload Me
                Exit Sub
            End If
        End If
    End If

    Call LoadCoefficients
End Sub


Private Sub PrevCoef()
    ' Check is entry is valid
    If Not IsNumeric(TxtCoef.Text) Then
        MsgBox "Please enter a numeric value!"
        Exit Sub
    End If
    
    ' Store entered value
    If NumeratorActive Then
        CoefNum(DegCur) = CDbl(TxtCoef.Text)
    Else
        CoefDen(DegCur) = CDbl(TxtCoef.Text)
    End If
    
    ' Switch to previous degree, switch from denominator to numerator (is function is rational) and show message when arrived at the beginning
    If Not IsRationalFunction Then
        If DegCur > 0 Then
            DegCur = DegCur - 1
        Else
            MsgBox "Jetzt geht es nicht mehr weiter", vbOKOnly, "Hinweis"
            Exit Sub
        End If
    Else
        If NumeratorActive Then
            If DegCur > 0 Then
                DegCur = DegCur - 1
            Else
                MsgBox "Jetzt geht es nicht mehr weiter", vbOKOnly, "Hinweis"
            End If
        Else
            If DegCur > 0 Then
                DegCur = DegCur - 1
            Else
                DegCur = DegNum
                NumeratorActive = True
                Beep
            End If
        End If
    End If

    Call LoadCoefficients
End Sub


Private Sub Form_Load()
    FrmCoefficients.KeyPreview = True
    NumeratorActive = True
    DegCur = 0
    
    Call LoadCoefficients
End Sub


Private Sub Form_Activate()
    TxtCoef.SetFocus
    TxtCoef.SelStart = 0
    TxtCoef.SelLength = Len(TxtCoef.Text)
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If Shift = 1 Then
        Select Case KeyCode
        
        Case vbKeyLeft
            Call PrevCoef
        
        Case vbKeyRight
            Call NextCoef
        End Select
    End If
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub


Private Sub BtnEnter_Click()
    Call NextCoef
End Sub


Private Sub BtnPrev_Click()
    Call PrevCoef
End Sub


Private Sub BtnNext_Click()
    Call NextCoef
End Sub


Private Sub BtnExit_Click()
    Unload Me
End Sub

'XXX
'Private Sub TxtCoef_KeyPress(KeyAscii As Integer)
'If KeyAscii = &H25 Then Unload Me
'End Sub

