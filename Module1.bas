Attribute VB_Name = "Module1"
Public A(), Grad, D(), GRF As Boolean, NV As Boolean, Grad2, Dimension, Factor, C1(), D1(), U, S, M(), O()

Public Type GRP
    GRF As String * 1
    ZG As String * 2
    NG As String * 2
    DefL As String * 30
    DefR As String * 30
    IntL As String * 30
    IntR As String * 30
    Width As String * 2
    Color As String * 8
    ZCoefficients As String * 1000
    NCoefficients As String * 1000
End Type

