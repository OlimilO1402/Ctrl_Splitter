Attribute VB_Name = "MNew"
Option Explicit

Public Function Splitter(BolMDI As Boolean, Owner As Object, Container As Object, Name As String, LeftTop As Control, RghtBot As Control) As Splitter
    Set Splitter = New Splitter: Splitter.New_ BolMDI, Owner, Container, Name, LeftTop, RghtBot
End Function
