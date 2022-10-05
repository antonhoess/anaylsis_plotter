Attribute VB_Name = "Module1"
Public CoefNum() As Double, CoefDen() As Double, DegNum As Integer, DegDen As Integer, GRF As Boolean, IsRationalFunction As Boolean, Dimension, Factor, C1(), D1(), U, S, M(), O()

Public Type GRP
    GRF As Boolean ' Gebrochen Rationale Funktion
    ZG As Integer ' Zähler Grad
    NG As Integer ' Nenner Grad
    DefL As Double ' Interval Untergrenze (links)
    DefR As Double ' Interval Obergrenze (rechts)
    IntL As Double ' Integral Untergrenze (links)
    IntR As Double ' Integral Obergrenze (rechts)
    Width As Integer ' Linenstärke
    Color As Long ' Linienfarbe
    ZCoefficients As String * 1000 ' Koeffizienten Zähler
    NCoefficients As String * 1000 ' Koeffizienten Nenner
End Type


Public Type RationalFunction
    IsRational As Boolean
    DegNum As Integer
    DegDen As Integer
    CoefNum() As Double
    CoefDen() As Double
End Type


Public Type DrawSettings
    DrawMode As Integer
    DrawStyle As Integer
    DrawWidth As Integer
    FillColor As Long
    ForeColor As Long
End Type


Public Function GetDrawSettings(ByRef Frm As Form) As DrawSettings
    Dim Settings As DrawSettings
    
    Settings.DrawMode = Frm.DrawMode
    Settings.DrawStyle = Frm.DrawStyle
    Settings.DrawWidth = Frm.DrawWidth
    Settings.FillColor = Frm.FillColor
    Settings.ForeColor = Frm.ForeColor
    
    GetDrawSettings = Settings
End Function


Public Function SetDrawSettings(ByRef Frm As Form, ByRef Settings As DrawSettings)
    Frm.DrawMode = Settings.DrawMode
    Frm.DrawStyle = Settings.DrawStyle
    Frm.DrawWidth = Settings.DrawWidth
End Function
