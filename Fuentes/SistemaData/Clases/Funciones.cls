VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Funciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function GetCodeTable(Tabla As String, CampoId As String, Optional mConexion As ADODB.Connection) As Long
    Dim xRstTabla As New ADODB.Recordset
    Dim dataBase As New SistemaData.EDataBase
    
    If Not mConexion Is Nothing Then Set dataBase.Connection = mConexion
    
    dataBase.CommandText = "SELECT MAX(" & CampoId & ") AS " & CampoId & " " _
                    + vbCr + "FROM " & Tabla

    'dataBase.CommandText = "SELECT TOP 1 " & CampoId & " FROM " & Tabla & " ORDER BY  " & CampoId & " DESC "
    Set xRstTabla = dataBase.GetRecordset()
    If xRstTabla.RecordCount = 0 Then
        GetCodeTable = 1
    Else
        xRstTabla.MoveLast
        
        If IsNumeric(xRstTabla(CampoId)) Then
            GetCodeTable = xRstTabla(CampoId) + 1
        Else
            GetCodeTable = 1
        End If
    End If
    
    Set xRstTabla = Nothing
    Set dataBase = Nothing
End Function

Public Function LeerLineaINI(RutaArchivoIni As String, TextoLinea As String, Posicion As String) As String
    LeerLineaINI = mLeerLineaINI(RutaArchivoIni, TextoLinea, Posicion)
End Function

Function Encriptar(DataValue As Variant) As Variant
    Dim x As Long
    Dim Temp As String
    Dim TempNum As Long
    Dim TempChar As String
    Dim TempChar2 As String
      
    For x = 1 To Len(DataValue)
        TempChar2 = Mid(DataValue, x, 1)
        TempNum = Int(Asc(TempChar2) / 16)
          
        If ((TempNum * 16) < Asc(TempChar2)) Then
                 
            TempChar = ConvToHex(Asc(TempChar2) - (TempNum * 16))
            Temp = Temp & ConvToHex(TempNum) & TempChar
        Else
            Temp = Temp & ConvToHex(TempNum) & "0"
          
        End If
    Next x
      
    Encriptar = Temp
End Function

Private Function ConvToHex(x As Long) As String
    If x > 9 Then
        ConvToHex = Chr(x + 55)
    Else
        ConvToHex = CStr(x)
    End If
End Function

Private Function ConvToInt(x As String) As Long
    Dim x1 As String
    Dim x2 As String
    Dim Temp As Long
      
    x1 = Mid(x, 1, 1)
    x2 = Mid(x, 2, 1)
    If IsNumeric(x1) Then
        Temp = 16 * Int(x1)
    Else
        Temp = (Asc(x1) - 55) * 16
    End If
    If IsNumeric(x2) Then
        Temp = Temp + Int(x2)
    Else
        Temp = Temp + (Asc(x2) - 55)
    End If
    ' retorno
    ConvToInt = Temp
End Function

Function Desencriptar(DataValue As Variant) As Variant
    Dim x As Long
    Dim Temp As String
    Dim HexByte As String
      
    For x = 1 To Len(DataValue) Step 2
        HexByte = Mid(DataValue, x, 2)
        Temp = Temp & Chr(ConvToInt(HexByte))
    Next x
    ' retorno
    Desencriptar = Temp
End Function

