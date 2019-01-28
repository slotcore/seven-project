VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FuncionesFchHor"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function RestarHoras(HoraInicio As Date, HoraFinal As Date) As Date
    Dim Num1, Num2, total As Double
    Dim Hora, minuto As Integer
    Dim xCad As String
    
    total = HoraInicio - HoraFinal
    RestarHoras = CDate(Format(total, "hh:mm"))
End Function

Function DeHoraNumeros(Hora As Date) As Double
    Dim Num1, Num2, total As Double
    Dim Horas, minuto As Integer
    Dim xCad As String
    
    Num1 = (Hour(Hora) * 3600)
    Num1 = Num1 + (Minute(Hora) * 60)

    DeHoraNumeros = (Num1 / 3600)
End Function

Function HallaDiasMes(Fecha As Date) As Integer
    HallaDiasMes = (DateSerial(Year(Fecha), Month(Fecha) + 1, Day(Fecha)) - DateSerial(Year(Fecha), Month(Fecha), Day(Fecha)))
End Function
