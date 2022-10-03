Attribute VB_Name = "NewtonMethod"
Option Explicit


Public Type NewtonResult
    NullCnt As Integer
    Nulls() As Double
End Type


' N2 = Store Result to array Newton2. There is Result and Newton2, which are global and get processed alter by the calling function
Public Function Newton(Coef, N2 As Boolean) As NewtonResult
    Dim I As Integer
    Dim Xneu As Double
    Dim Degree As Integer
    Dim Result As NewtonResult
    Dim DegZero
    Dim X As Double ', Factor2
    Dim NullsPrecision As Integer
    Dim Fertig As Boolean
    Dim IterCnt As Integer
    
    If UBound(Coef) > 0 Then
        Xneu = 10 ^ 10
        Result.NullCnt = 0
        Degree = UBound(Coef)
        ReDim Result.Nulls(0 To Degree - 1)
        DegZero = 0
        NullsPrecision = 4
        
        ' XXX ? Remove all first 0-coefficients beginning from degree 0 upwards
        For I = 0 To Degree
            If Coef(I) = 0 Then
                DegZero = DegZero + 1
                Result.Nulls(Result.NullCnt) = 0
                Result.NullCnt = Result.NullCnt + 1
            Else
                Exit For
            End If
        Next I
        
        ' XXX? Reduce degrees by number of 0-coefficients and shrink array of coeffcients
        If DegZero > 0 Then
            For I = 0 To Degree - DegZero
                Coef(I) = Coef(I + DegZero)
            Next I
        
            Degree = Degree - DegZero
            ReDim Preserve Coef(0 To Degree)
        End If
        
        If Degree > 0 Then  ' XXX Was passiert sonst, z.B. wenn alle Koeffizienten 0 sind? Testen.
            Do While True
                ' Newton method: find nulls by repeatedly calculating x_n+1 = x_n + f(x_n) / f'(x_n)
                ' If no null was found after many iteration, supposedly there is no more null
                Xneu = 10 ^ 10
                IterCnt = 0
                Do While True
                    X = Xneu
                    Xneu = Xneu - GetFuncValByX(Xneu, Coef) / GetFuncValByX(Xneu, GetDiffFuncCoefs(Coef))
                    
                    If Round(Xneu, NullsPrecision) = Round(X, NullsPrecision) Then
                        Exit Do
                    End If
                    IterCnt = IterCnt + 1
                    
                    If IterCnt = 1000 Then
                        Fertig = True
                        Exit Do
                    End If
                Loop
                
                If Fertig Then
                    Exit Do
                End If
                
                ' Add null to list, divide polynomial by found null and decrease degree
                Result.Nulls(Result.NullCnt) = Round(Xneu, 6)
                Result.NullCnt = Result.NullCnt + 1
                
                Call Nullstellendivision(Coef, Xneu)
                            
                Degree = Degree - 1
                If Degree = 0 Then Exit Do
                
                ' Update GUI
                DoEvents
            Loop
        End If
        
        ReDim Preserve Result.Nulls(0 To Result.NullCnt - 1)
    Else
        Result.NullCnt = 0
    End If
    
    Newton = Result
End Function

Public Sub Nullstellendivision(Coef, X)
    Dim I As Integer
    Dim Degree As Integer
    Degree = UBound(Coef)
    ''Factor1 = 1 / Coef(Degree)
    ''For I = 0 To Degree
    ''    Coef(I) = Coef(I) * Factor1
    ''Next I
    
    ' Coef(n-1) = Coef(n-1) + Coef(n) * x
    For I = 1 To Degree
        Coef(Degree - I) = Coef(Degree - I) + Coef(Degree - I + 1) * X
    Next I
    
    ' Koreffizienten nachrücken
    For I = 0 To Degree - 1
        Coef(I) = Coef(I + 1)
    Next I
    
    ' Grad um 1 verringern und Array entsprechend verkleinern
    Degree = Degree - 1
    ReDim Preserve Coef(0 To Degree)
End Sub


Public Function GetFuncValByX(ByVal X As Double, ByRef Coefficients) As Double
    Dim Value As Double
    Dim Deg As Integer
    Dim D As Integer
    
    Deg = UBound(Coefficients)
    For D = 0 To Deg
        Value = Value + Coefficients(D) * X ^ D
    Next D
    
    GetFuncValByX = Value
End Function


Public Function GetDiffFuncCoefs(ByRef Coefficients) As Double()
    Dim Degree As Integer
    Degree = UBound(Coefficients)
    Dim Result() As Double
    
    If Degree > 0 Then
        ReDim Result(0 To Degree - 1)
    
        Dim I As Integer
        For I = 1 To Degree
            Result(I - 1) = Coefficients(I) * I
        Next I
    Else
        ReDim Result(0 To 0)
        Result(0) = 0
    End If
    
    GetDiffFuncCoefs = Result
End Function
