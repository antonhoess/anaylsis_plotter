Attribute VB_Name = "SaveArray"
'Code von Benjamin Wilger
'Benjamin@ActiveVB.de
'Copyright (C) 2001-2002
Option Explicit

Private Declare Sub RtlMoveMemory Lib "kernel32" (dest As Any, source As Any, ByVal bytes As Long)

Public Sub ReadStringArray(ByVal Filename As String, StringArray() As String)
    Dim FNum As Integer
    Dim LBounds() As Long, DimCount As Integer
    Dim UBounds() As Long
    
    'Freie Dateinummer herausfinden
    FNum = FreeFile
    'Datei mit binärem Zugriff öffnen
    Open Filename For Binary As FNum
    'Anzahl der Dimensionen auslesen
    Get FNum, , DimCount
    'Array für die LBounds-Daten
    ReDim LBounds(0 To DimCount - 1)
    'Array für die UBounds-Daten
    ReDim UBounds(0 To DimCount - 1)
    '... einlesen
    Get FNum, , LBounds 'Erst die LBounds und dann die UBounds
    Get FNum, , UBounds
    'nun das Zielarray den Informationen entsprechend dimensionieren
    Select Case DimCount
        Case 1
            ReDim StringArray(LBounds(0) To UBounds(0))
        Case 2
            ReDim StringArray(0 To UBounds(0), _
                    LBounds(1) To UBounds(1))
        Case 3
            ReDim StringArray(LBounds(0) To UBounds(0), _
                    LBounds(1) To UBounds(1), _
                    LBounds(2) To UBounds(2))
        Case 4
            ReDim StringArray(LBounds(0) To UBounds(0), _
                    LBounds(1) To UBounds(1), _
                    LBounds(2) To UBounds(2), _
                    LBounds(3) To UBounds(3))
        Case 5
            ReDim StringArray(LBounds(0) To UBounds(0), _
                    LBounds(1) To UBounds(1), _
                    LBounds(2) To UBounds(2), _
                    LBounds(3) To UBounds(3), _
                    LBounds(4) To UBounds(4))
        Case 6
            ReDim StringArray(LBounds(0) To UBounds(0), _
                    LBounds(1) To UBounds(1), _
                    LBounds(2) To UBounds(2), _
                    LBounds(3) To UBounds(3), _
                    LBounds(4) To UBounds(4), _
                    LBounds(5) To UBounds(5))
        Case 7
            ReDim StringArray(LBounds(0) To UBounds(0), _
                    LBounds(1) To UBounds(1), _
                    LBounds(2) To UBounds(2), _
                    LBounds(3) To UBounds(3), _
                    LBounds(4) To UBounds(4), _
                    LBounds(5) To UBounds(5), _
                    LBounds(6) To UBounds(6))
        Case 8
            ReDim StringArray(LBounds(0) To UBounds(0), _
                    LBounds(1) To UBounds(1), _
                    LBounds(2) To UBounds(2), _
                    LBounds(3) To UBounds(3), _
                    LBounds(4) To UBounds(4), _
                    LBounds(5) To UBounds(5), _
                    LBounds(6) To UBounds(6), _
                    LBounds(7) To UBounds(7))
        Case 9
            ReDim StringArray(LBounds(0) To UBounds(0), _
                    LBounds(1) To UBounds(1), _
                    LBounds(2) To UBounds(2), _
                    LBounds(3) To UBounds(3), _
                    LBounds(4) To UBounds(4), _
                    LBounds(5) To UBounds(5), _
                    LBounds(6) To UBounds(6), _
                    LBounds(7) To UBounds(7), _
                    LBounds(8) To UBounds(8))
        Case 10
            ReDim StringArray(LBounds(0) To UBounds(0), _
                    LBounds(1) To UBounds(1), _
                    LBounds(2) To UBounds(2), _
                    LBounds(3) To UBounds(3), _
                    LBounds(4) To UBounds(4), _
                    LBounds(5) To UBounds(5), _
                    LBounds(6) To UBounds(6), _
                    LBounds(7) To UBounds(7), _
                    LBounds(8) To UBounds(8), _
                    LBounds(9) To UBounds(9))
        Case Else
            Err.Raise vbObjectError + 100, , "Zu viele Dimensionen!"
    End Select
    'das Array einlesen lassen.
    Get FNum, , StringArray
    'Datei schließen
    Close FNum
    
End Sub

Public Sub SaveStringArray(ByVal Filename As String, StringArray() As String)
    Dim FNum As Integer
    Dim Dimensions As Long
    
    Dimensions = Dimension(StringArray)
    
    If Dimensions > 10 Then
        Err.Raise vbObjectError + 100, , "Zu viele Dimensionen!"
    End If
    'Freie Dateinummer herausfinden
    FNum = FreeFile
    'Falls bereits vorhanden, die Datei löschen, da sonst ältere Werte
    'vielleicht bestehen bleiben
    If Dir(Filename) <> "" Then Kill Filename
    'Datei zum binären Zugriff öffnen, ggf. erstellen
    Open Filename For Binary As FNum
    Put FNum, , CInt(Dimensions) 'Anzahl der Dimensionen schreiben
    
    Put FNum, , CLng(LBound(StringArray, 1)) 'Untergrenzen speichern
    If Dimensions >= 2 Then Put FNum, , CLng(LBound(StringArray, 2))
    If Dimensions >= 3 Then Put FNum, , CLng(LBound(StringArray, 3))
    If Dimensions >= 4 Then Put FNum, , CLng(LBound(StringArray, 4))
    If Dimensions >= 5 Then Put FNum, , CLng(LBound(StringArray, 5))
    If Dimensions >= 6 Then Put FNum, , CLng(LBound(StringArray, 6))
    If Dimensions >= 7 Then Put FNum, , CLng(LBound(StringArray, 7))
    If Dimensions >= 8 Then Put FNum, , CLng(LBound(StringArray, 8))
    If Dimensions >= 9 Then Put FNum, , CLng(LBound(StringArray, 9))
    If Dimensions = 10 Then Put FNum, , CLng(LBound(StringArray, 10))
    
    Put FNum, , CLng(UBound(StringArray, 1)) 'Obergrenzen speichern
    If Dimensions >= 2 Then Put FNum, , CLng(UBound(StringArray, 2))
    If Dimensions >= 3 Then Put FNum, , CLng(UBound(StringArray, 3))
    If Dimensions >= 4 Then Put FNum, , CLng(UBound(StringArray, 4))
    If Dimensions >= 5 Then Put FNum, , CLng(UBound(StringArray, 5))
    If Dimensions >= 6 Then Put FNum, , CLng(UBound(StringArray, 6))
    If Dimensions >= 7 Then Put FNum, , CLng(UBound(StringArray, 7))
    If Dimensions >= 8 Then Put FNum, , CLng(UBound(StringArray, 8))
    If Dimensions >= 9 Then Put FNum, , CLng(UBound(StringArray, 9))
    If Dimensions = 10 Then Put FNum, , CLng(UBound(StringArray, 10))
    'Den kompletten Array speichern
    Put FNum, , StringArray
    'Datei wieder schließen
    Close FNum
    
End Sub

'Normale String, Integer/Long, Byte und Double/Single Arrays können
'in Variant konvertiert werden, sodass wir mit dieser Methode
'die Dimensionen herausfinden können.
'Vielen Dank an Jost Schwider(www.vb-tec.de) für diesen Code!
' -> http://vb-tec.de/arrdim.htm
Private Function Dimension(ByRef avarArray As Variant) As Integer
    Dim Ptr As Long
    
    If IsArray(avarArray) Then
        Ptr = VarPtr(avarArray) + 8     'VB-Array
        RtlMoveMemory Ptr, ByVal Ptr, 4 'SafeArrayDescriptor
        RtlMoveMemory Ptr, ByVal Ptr, 4 'SafeArray-Struktur
        If Ptr Then RtlMoveMemory Dimension, ByVal Ptr, 2
    Else
        Err.Raise 13 'Type mismatch
    End If
End Function

'Falls Du den Code abwandelst, das das Zielarray nicht mehr
'in ein Variant konvertiert werden kann(beispielsweise bei UDTs),
'benutze diese Funktion. Sie ermittelt die Dimensionen mit der Brechstangen-
'Methode :-)
Private Function DimensionBrechstange(ByRef avarArray As Variant) As Long
    Dim I As Long
    Dim tmpBound As Integer
    
    On Error Resume Next
    Do
        I = I + 1
        tmpBound = UBound(avarArray, I)
    Loop Until Err
    Err.Clear
    
    DimensionBrechstange = I - 1
End Function


