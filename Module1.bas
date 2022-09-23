Attribute VB_Name = "Module1"
Public CoefNum() As Double, CoefDen() As Double, DegNum As Integer, DegDen As Integer, GRF As Boolean, IsNotRationalFunction As Boolean, Dimension, Factor, C1(), D1(), U, S, M(), O()
' XXX
' IsNotRationalFunction = Inficates if currently a rational function is defined

Public Type GRP
    GRF As String * 1 ' Gebrochen Rationale Funktion
    ZG As String * 2 ' Zähler Grad
    NG As String * 2 ' Nenner Grad
    DefL As String * 30 ' Interval Untergrenze (links)
    DefR As String * 30 ' Interval Obergrenze (rechts)
    IntL As String * 30 ' Integral Untergrenze (links)
    IntR As String * 30 ' Integral Obergrenze (rechts)
    Width As String * 2 ' Linenstärke
    Color As String * 8 ' Linienfarbe
    ZCoefficients As String * 1000 ' Koeffizienten Zähler
    NCoefficients As String * 1000 ' Koeffizienten Nenner
End Type

