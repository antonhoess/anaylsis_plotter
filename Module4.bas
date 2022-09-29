Attribute VB_Name = "NewtonMethod"
Public Koefficients2, Grad5

' N2 = Store Result to array Newton2. There is Newton1 and Newton2, which are global and get processed alter by the calling function
Public Function Newton(Koefficients, Grad4, N2 As Boolean)
    Dim I As Integer
    Dim Xneu As Double
    Dim ZFC As Integer, NFC As Integer ' ZählerFaktorCounter, NennerFaktorCounter
    Dim Newton1() As Double, Newton2() As Double
    'mit ZFC arbeiten
    Dim Checkvalue, Untergrad, Fertig ', Factor2
    
    ReDim Newton1(0)
    ReDim Newton2(0)
    ReDim Koefficients2(0)
    Grad5 = Grad4
    Koefficients2 = Koefficients
    ZFC = 0
    NFC = 0

    If N2 = True Then
        ReDim Newton1(0 To Grad5)
    Else
        ReDim Newton2(0 To Grad5)
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
    If Grad5 <> 0 Then
        ReDim Preserve Koefficients2(0 To Grad5)
        
        Do While Fertig = False
            For I = 1 To 1000
                Checkvalue = Xneu
                
                Xneu = Xneu - GetFuncvalByXvalFromNonFractionalFunction(Xneu, Koefficients2) / GetFuncvalByXvalFromNonFractionalFunction(Xneu, DifferentiateFunctionCoefs(Koefficients2))
                If Xneu - GetFuncvalByXvalFromNonFractionalFunction(Xneu, Koefficients2) / GetFuncvalByXvalFromNonFractionalFunction(Xneu, DifferentiateFunctionCoefs(Koefficients2)) = Checkvalue Then Exit For
                If I = 999 Then Exit Do
            Next I
            
            If N2 = True Then
                Newton1(ZFC) = (Int(Xneu * 10 ^ 6 + 0.1) / 10 ^ 6)
                ZFC = ZFC + 1
            Else
                Newton2(NFC) = (Int(Xneu * 10 ^ 6 + 0.1) / 10 ^ 6)
                NFC = NFC + 1
            End If
            
            Call Polynomdivision(Koefficients2, Xneu)
            If Grad5 = 0 Then Exit Do
            Xneu = 10 ^ 10 'Text1.Text
            
            DoEvents
        Loop
    End If
    
    If N2 = True Then
        Newton = Newton1
    Else
        Newton = Newton2
    End If
End Function

Public Sub Polynomdivision(Koefficients2, X)
    Dim I As Integer
    ''Factor1 = 1 / Koefficients2(Grad5)
    ''For I = 0 To Grad5
    ''    Koefficients2(I) = Koefficients2(I) * Factor1
    ''Next I
    
    ' Coef(n-1) = Coef(n-1) + Coef(n) * x
    Dim Degree As Integer
    Degree = UBound(Koefficients2)
    
    For I = 1 To Degree
        Koefficients2(Degree - I) = Koefficients2(Degree - I + 1) * X + Koefficients2(Degree - I)
    Next I
    
    For I = 0 To Degree - 1
        Koefficients2(I) = Koefficients2(I + 1)
    Next I
    Degree = Degree - 1
    Grad5 = Grad5 - 1
    ReDim Preserve Koefficients2(0 To Degree)
    
    Fertig = True
    For I = 0 To Degree
        If Koefficients2(I) > 10 ^ -5 And Koefficients2(I) > -10 ^ -5 Then Fertig = False
    Next I
End Sub


Public Function GetFuncvalByXvalFromNonFractionalFunction(ByVal X As Double, ByRef Coefficients) As Double
    Dim Value As Double
    Dim Deg As Integer
    Dim D As Integer
    
    Deg = UBound(Coefficients)
    For D = 0 To Deg
        Value = Value + Coefficients(D) * X ^ D
    Next D
    
    GetFuncvalByXvalFromNonFractionalFunction = Value
End Function


Public Function DifferentiateFunctionCoefs(ByRef Coefficients) As Double()
    Dim Degree As Integer
    Degree = UBound(Coefficients)
    Dim Result() As Double
    ReDim Result(0 To Degree - 1)

    For I = 1 To Degree
        Result(I - 1) = Coefficients(I) * I
    Next I
    
    DifferentiateFunctionCoefs = Result
End Function
