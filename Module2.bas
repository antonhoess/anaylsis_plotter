Attribute VB_Name = "Module2"
Public N, R, C, E, F, X, Matrix() As Single, Matrix2() As Single, Matrix3() As Single, Horner1(), Horner2(), Ergebnis, Factor1, Factor2

Public Function HornerSchema()
    Grad1 = Form1.Text1.Text
    Grad2 = Form1.Text14.Text
    
    If NV = False Then
    
        If Grad1 > 0 Then
            ReDim Horner1(1 To Grad1)
        Else
            ReDim Horner1(1)
        End If
        
        If Grad2 > 0 Then
            ReDim Horner2(1 To Grad2)
        Else
            ReDim Horner1(1)
        End If
        
        ReDim Matrix2(0 To Grad1)
        ReDim Matrix3(0 To Grad2)
        
        For I = -Grad1 To 0
            Matrix2(I + Grad1) = A(-I)
        Next I
        
        For I = -Grad2 To 0
            Matrix3(I + Grad2) = D(-I)
        Next I
        
        
        Factor1 = Matrix2(0)
        For I = 0 To Grad1
            Matrix2(I) = Matrix2(I) / Factor1
        Next I
        
        Factor2 = Matrix3(0)
        For I = 0 To Grad2
            Matrix3(I) = Matrix3(I) / Factor2
        Next I
        
        For N = 1 To Grad1
            For I = -100 To 100
                B = Matrix2(0)
                For X = 1 To Grad1 - N + 1
                    B = B * I + Matrix2(X)
                Next X
                If B = 0 Then
                    ReDim Matrix(0 To Grad1 - N + 1)
                    For E = 0 To Grad1 - N + 1
                        Matrix(E) = Matrix2(E)
                    Next E
                    ReDim Matrix2(0 To Grad1)
                    ReDim Matrix2(0 To Grad1 - N)
                    F = 0
                    For E = 0 To Grad1 - N
                        F = F * I + Matrix(E)
                        Matrix2(E) = F
                    Next E
                    Horner1(N) = I
                    Exit For
                End If
            Next I
        Next N
        
        For N = 1 To Grad2
            For I = -100 To 100
                B = Matrix3(0)
                For X = 1 To Grad2 - N + 1
                    B = B * I + Matrix3(X)
                Next X
                If B = 0 Then
                    ReDim Matrix(0 To Grad2 - N + 1)
                    For E = 0 To Grad2 - N + 1
                        Matrix(E) = Matrix3(E)
                    Next E
                    ReDim Matrix3(0 To Grad2)
                    ReDim Matrix3(0 To Grad2 - N)
                    F = 0
                    For E = 0 To Grad2 - N
                        F = F * I + Matrix(E)
                        Matrix3(E) = F
                    Next E
                    Horner2(N) = I
                    Exit For
                End If
            Next I
        Next N
    Else
        If Grad1 > 0 Then
            ReDim Horner1(1 To Grad1)
        Else
            ReDim Horner1(1)
        End If
        
        ReDim Matrix2(0 To Grad1)
        
        For I = -Grad1 To 0
            Matrix2(I + Grad1) = A(-I)
        Next I
        
        Factor1 = Matrix2(0)
        For I = 0 To Grad1
            Matrix2(I) = Matrix2(I) / Factor1
        Next I
        
        For N = 1 To Grad1
            For I = -100 To 100
                B = Matrix2(0)
                For X = 1 To Grad1 - N + 1
                    B = B * I + Matrix2(X)
                Next X
                If B = 0 Then
                    ReDim Matrix(0 To Grad1 - N + 1)
                    For E = 0 To Grad1 - N + 1
                        Matrix(E) = Matrix2(E)
                    Next E
                    ReDim Matrix2(0 To Grad1)
                    ReDim Matrix2(0 To Grad1 - N)
                    F = 0
                    For E = 0 To Grad1 - N
                        F = F * I + Matrix(E)
                        Matrix2(E) = F
                    Next E
                    Horner1(N) = I
                    Exit For
                End If
            Next I
        Next N
    End If
End Function

