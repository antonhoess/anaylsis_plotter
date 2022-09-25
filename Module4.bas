Attribute VB_Name = "Module4"
Public Xneu, Newton1(), Newton2(), N2 As Boolean, ZFC, NFC, Grad5, Koefficients2, ZAbl1(), ZAbl2(), NAbl1(), NAbl2()

Public Sub Newton(Koefficients, Grad4, N2)
    Dim I As Integer
    'mit ZFC arbeiten (Zählerfaktorcounter)
    Dim Checkvalue, Untergrad, Fertig ', Factor2
    
    ReDim Newton1(0)
    ReDim Newton2(0)
    ReDim Koefficients2(0)
    Grad5 = Grad4
    'Fertig = False
    Koefficients2 = Koefficients
    ''Koefficients2(0) = Koefficients2(0)
    ''Koefficients2(1) = Koefficients2(1)
    ''Koefficients2(2) = Koefficients2(2)
    ''Newton1(0) = Newton1(0)
    ''Newton1(1) = Newton1(1)
    ''Newton1(2) = Newton1(2)
    ''Newton1(3) = Newton1(3)
    'ZFC = 1
    'NFC = 1
    ZFC = 0
    NFC = 0
    
    If N2 = True Then
        If Grad5 > 0 Then
            ReDim Newton1(0 To Grad5)
        Else
            ReDim Newton1(0)
        End If
    Else
        If Grad5 > 0 Then
            ReDim Newton2(0 To Grad5)
        Else
            ReDim Newton2(0)
        End If
    End If
    
    Xneu = 10 ^ 10
    Untergrad = 0
    
    For I = 0 To Grad5
        If Koefficients2(I) = 0 Then
        Untergrad = Untergrad + 1
        
            If N2 = True Then
                Newton1(ZFC) = 0
                ZFC = ZFC + 1
            Else
                Newton2(NFC) = 0
                NFC = NFC + 1
            End If
        Else
            Exit For
        End If
    Next I
    
    For I = 0 To Grad5 - Untergrad
        Koefficients2(I) = Koefficients2(I + Untergrad)
    Next I
    
    Grad5 = Grad5 - Untergrad
    If Grad5 = 0 Then Exit Sub
    ReDim Preserve Koefficients2(0 To Grad5)
    
    Do While Fertig = False
        For I = 1 To 1000
            Checkvalue = Xneu
            
            Xneu = Xneu - fv(Xneu, Koefficients2, Grad5) / fav(Xneu, Koefficients2, Grad5)
            If Xneu - fv(Xneu, Koefficients2, Grad5) / fav(Xneu, Koefficients2, Grad5) = Checkvalue Then Exit For
            If I = 999 Then Exit Do
        Next I
            If N2 = True Then
                Newton1(ZFC) = (Int(Xneu * 10 ^ 6 + 0.1) / 10 ^ 6)
                ZFC = ZFC + 1
            Else
                Newton2(NFC) = (Int(Xneu * 10 ^ 6 + 0.1) / 10 ^ 6)
                NFC = NFC + 1
            End If
    Call Polynomdivision(Koefficients2, Grad5)
    If Grad5 = 0 Then Exit Do
    Xneu = 10 ^ 10 'Text1.Text
    
    DoEvents
    Loop
End Sub

Public Sub Polynomdivision(Koefficients2, Grad5)
    Dim I As Integer
    ''Factor1 = 1 / Koefficients2(Grad5)
    ''For I = 0 To Grad5
    ''    Koefficients2(I) = Koefficients2(I) * Factor1
    ''Next I
    
    For I = 1 To Grad5
        Koefficients2(Grad5 - I) = Koefficients2(Grad5 - I + 1) * (Xneu) + Koefficients2(Grad5 - I)
    Next I
    
    For I = 0 To Grad5 - 1
        Koefficients2(I) = Koefficients2(I + 1)
    Next I
    Grad5 = Grad5 - 1
    ReDim Preserve Koefficients2(0 To Grad5)
    
    Fertig = True
    For I = 0 To Grad5
        If Koefficients2(I) > 10 ^ -5 And Koefficients2(I) > -10 ^ -5 Then Fertig = False
    Next I
End Sub

Public Function fv(Z1, Koefficients2, Grad5)
    fv = 0
    For U = 0 To Grad5
        fv = fv + Koefficients2(U) * Z1 ^ U
    Next U
End Function

Public Function fav(Z1, Koefficients2, Grad5)
    fav = 0
    For U = 1 To Grad5
        fav = fav + (Koefficients2(U) * U) * Z1 ^ (U - 1)
    Next U
End Function


