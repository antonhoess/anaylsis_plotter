Attribute VB_Name = "Module1"
Public CoefNum() As Double, CoefDen() As Double, DegNum As Integer, DegDen As Integer, GRF As Boolean, IsNotRationalFunction As Boolean, Dimension, Factor, C1(), D1(), U, S, M(), O()
' XXX
' IsNotRationalFunction = Inficates if currently a rational function is defined

Public Type GRP
    GRF As Boolean 'String * 1 ' Gebrochen Rationale Funktion
    ZG As Integer 'String * 2 ' Zähler Grad
    NG As Integer ' String * 2 ' Nenner Grad
    DefL As Double 'String * 30 ' Interval Untergrenze (links)
    DefR As Double ' String * 30 ' Interval Obergrenze (rechts)
    IntL As Double ' String * 30 ' Integral Untergrenze (links)
    IntR As Double ' String * 30 ' Integral Obergrenze (rechts)
    Width As Integer ' String * 2 ' Linenstärke
    Color As Long ' String * 8 ' Linienfarbe
    ZCoefficients As String * 1000 ' Koeffizienten Zähler
    NCoefficients As String * 1000 ' Koeffizienten Nenner
End Type

