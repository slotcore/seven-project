Attribute VB_Name = "Funciones"
Option Explicit

Function F_NulosN(valor_nulo As Variant) As Double
    If Trim(valor_nulo) = "" Or IsNull(valor_nulo) Then
        F_NulosN = 0
    Else
        'Si el valor no es nulo retorna el valor original
        If IsNumeric(valor_nulo) Then
            F_NulosN = valor_nulo
        Else
            F_NulosN = 0
        End If
    End If
End Function

Function F_NulosC(valor_nulo As Variant) As String
    If Trim(valor_nulo) = "" Or IsNull(valor_nulo) Then
        F_NulosC = ""
    Else
        'Si el valor no es nulo retorna el valor original
        If IsNumeric(valor_nulo) Then
            F_NulosC = Trim(valor_nulo)
        Else
            F_NulosC = Trim(valor_nulo)
        End If
    End If
End Function

Sub F_RST_Busq(rstBusq As ADODB.Recordset, TxtSQLoTabla As String, xConeccion As Connection)
On Error GoTo LaCague:
    If rstBusq.State = adStateOpen Then
        rstBusq.Close
    End If
    
    rstBusq.CursorLocation = adUseClient
    rstBusq.CursorType = adOpenForwardOnly
    rstBusq.LockType = adLockOptimistic
    
    rstBusq.ActiveConnection = xConeccion
    rstBusq.Open F_NulosC(TxtSQLoTabla), , , , adAsyncFetch
    'adAsyncFetch
    Exit Sub
LaCague:
    MsgBox "No se pudo guardar el recorset por el siguiente motivo :" + Trim(Err.Description)
    Set rstBusq = Nothing
End Sub

Sub F_RST_Mant(rstMant As ADODB.Recordset, TxtSQLoTabla As String, xConeccion As Connection)
    
On Error GoTo LaCague:

    If rstMant.State = adStateOpen Then
        rstMant.Close
    End If
    rstMant.CursorLocation = adUseClient
    rstMant.CursorType = adOpenKeyset
    rstMant.LockType = adLockOptimistic
    If F_NulosC(TxtSQLoTabla) = "" Then
        'Crea un recordset desconectado de una tabla
        rstMant.Open
        Exit Sub
    End If
    rstMant.ActiveConnection = xConeccion
    rstMant.Open TxtSQLoTabla, , , , adAsyncFetch
    Exit Sub
LaCague:
    MsgBox "No se pudo guardar el recorset por el siguiente motivo :" + Trim(Err.Description)
    Set rstMant = Nothing
End Sub


