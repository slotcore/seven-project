Attribute VB_Name = "Funciones"
Option Explicit

Function CargarTMPResultado(xFchIni As String, xFchFin As String, xCodigoBalance As Integer, Rst As ADODB.Recordset) As ADODB.Recordset

End Function

Function CargarTMPBalance(xFchIni As String, xFchFin As String, xCodigoBalance As Integer, Rst As ADODB.Recordset) As ADODB.Recordset
    Dim RstBal As New ADODB.Recordset
    Dim RstSal As New ADODB.Recordset
    Dim A, B As Integer
    
    Dim nSQL As String
    
''    'cargamos los movimientos del periodo
''    RST_Busq RstBal, "TRANSFORM Sum(con_diario.impdebsol) AS SumaDeimpdebsol SELECT con_balancedet.idbal, con_balancecab.descripcion, " _
''        & " Sum(IIf([impdebdol]=0,[impdebsol],[impdebdol]*[con_tc].[impven])) AS totdebsol, Sum(IIf([imphabdol]=0,[imphabsol],[imphabdol]*[con_tc].[impven])) AS tothabsol, " _
''        & " Sum(IIf([impdebdol]=0,[impdebsol],[impdebdol]*[con_tc].[impven])-IIf([imphabdol]=0,[imphabsol],[imphabdol]*[con_tc].[impven])) AS saldo " _
''        & " FROM con_balance LEFT JOIN (con_balancecab LEFT JOIN (con_planctas RIGHT JOIN ((con_diario RIGHT JOIN con_balancedet ON con_diario.idcue = con_balancedet.idcuenta) " _
''        & " LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON con_planctas.id = con_balancedet.idcuenta) ON (con_balancecab.id = con_balancedet.idbal) " _
''        & " AND (con_balancecab.idcab = con_balancedet.idcab)) ON con_balance.id = con_balancecab.idcab WHERE (((con_diario.fchasi)>=CDate('" & xFchIni & "') " _
''        & " And (con_diario.fchasi)<=CDate('" & xFchFin & "')) AND ((con_balancecab.idcab)=" & xCodigoBalance & ")) GROUP BY con_balancedet.idbal, con_balancecab.descripcion" _
''        & " PIVOT con_balancedet.idcuenta", xCon
''
''
''    'cargamos los saldos
''    RST_Busq RstSal, "TRANSFORM Sum(con_diario.impdebsol) AS SumaDeimpdebsol SELECT con_balancedet.idbal, con_balancecab.descripcion, " _
''        & " Sum(IIf([impdebdol]=0,[impdebsol],[impdebdol]*[con_tc].[impven])) AS totdebsol, Sum(IIf([imphabdol]=0,[imphabsol],[imphabdol]*[con_tc].[impven])) AS tothabsol, " _
''        & " Sum(IIf([impdebdol]=0,[impdebsol],[impdebdol]*[con_tc].[impven])-IIf([imphabdol]=0,[imphabsol],[imphabdol]*[con_tc].[impven])) AS saldo " _
''        & " FROM con_balance LEFT JOIN (con_balancecab LEFT JOIN (con_planctas RIGHT JOIN ((con_diario RIGHT JOIN con_balancedet ON con_diario.idcue = con_balancedet.idcuenta) " _
''        & " LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON con_planctas.id = con_balancedet.idcuenta) ON (con_balancecab.id = con_balancedet.idbal) " _
''        & " AND (con_balancecab.idcab = con_balancedet.idcab)) ON con_balance.id = con_balancecab.idcab WHERE (((con_balancedet.idbal) Is Not Null) " _
''        & " AND ((con_diario.fchasi)<CDate('" & xFchIni & "') Or (con_diario.fchasi) Is Null) AND ((con_balancecab.idcab)=" & xCodigoBalance & ")) " _
''        & " GROUP BY con_balancedet.idbal, con_balancecab.descripcion PIVOT con_balancedet.idcuenta", xCon
    
    'cargamos los movimientos del periodo
    nSQL = "SELECT con_balancedet.idbal, con_balancecab.descripcion, Sum(IIf([impdebdol]=0,[impdebsol],[impdebdol]*[con_tc].[impven])) AS totdebsol, Sum(IIf([imphabdol]=0,[imphabsol],[imphabdol]*[con_tc].[impven])) AS tothabsol, Sum(IIf([impdebdol]=0,[impdebsol],[impdebdol]*[con_tc].[impven])-IIf([imphabdol]=0,[imphabsol],[imphabdol]*[con_tc].[impven])) AS saldo " _
        + vbCr + " FROM con_balance LEFT JOIN (con_balancecab LEFT JOIN (con_planctas RIGHT JOIN ((con_diario RIGHT JOIN con_balancedet ON con_diario.idcue = con_balancedet.idcuenta) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON con_planctas.id = con_balancedet.idcuenta) ON (con_balancecab.id = con_balancedet.idbal) AND (con_balancecab.idcab = con_balancedet.idcab)) ON con_balance.id = con_balancecab.idcab " _
        + vbCr + " WHERE (((con_diario.fchasi)>=CDate('" & xFchIni & "') And (con_diario.fchasi)<=CDate('" & xFchFin & "')) AND ((con_balancecab.idcab)=" & xCodigoBalance & ")) " _
        + vbCr + " GROUP BY con_balancedet.idbal, con_balancecab.descripcion;"
    
    RST_Busq RstBal, nSQL, xCon
    
    
    'cargamos los saldos
    nSQL = "SELECT con_balancedet.idbal, con_balancecab.descripcion, Sum(IIf([impdebdol]=0,[impdebsol],[impdebdol]*[con_tc].[impven])) AS totdebsol, Sum(IIf([imphabdol]=0,[imphabsol],[imphabdol]*[con_tc].[impven])) AS tothabsol, Sum(IIf([impdebdol]=0,[impdebsol],[impdebdol]*[con_tc].[impven])-IIf([imphabdol]=0,[imphabsol],[imphabdol]*[con_tc].[impven])) AS saldo " _
        + vbCr + " FROM con_balance LEFT JOIN (con_balancecab LEFT JOIN (con_planctas RIGHT JOIN ((con_diario RIGHT JOIN con_balancedet ON con_diario.idcue = con_balancedet.idcuenta) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON con_planctas.id = con_balancedet.idcuenta) ON (con_balancecab.id = con_balancedet.idbal) AND (con_balancecab.idcab = con_balancedet.idcab)) ON con_balance.id = con_balancecab.idcab " _
        + vbCr + " WHERE (((con_balancedet.idbal) Is Not Null) AND ((con_diario.fchasi)<CDate('" & xFchIni & "') Or (con_diario.fchasi) Is Null) AND ((con_balancecab.idcab)=" & xCodigoBalance & ")) " _
        + vbCr + " GROUP BY con_balancedet.idbal, con_balancecab.descripcion; "

    RST_Busq RstSal, nSQL, xCon

    If RstBal.RecordCount = 0 Then
        MsgBox "No se ha encontrado movimientos en el periodo especificado para efectura un Balance General", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Set RstBal = Nothing
        Exit Function
    End If
    
    RstBal.MoveFirst
    
    'agregamos los movimientos del periodo al rst temporal
    For A = 1 To RstBal.RecordCount
        Rst.AddNew
        Rst("idbal") = RstBal("idbal")
        Rst("descri") = NulosC(RstBal("descripcion"))
        Rst("movperdeb") = NulosN(RstBal("totdebsol"))
        Rst("movperhab") = NulosN(RstBal("tothabsol"))
        
        RstBal.MoveNext
        If RstBal.EOF = True Then Exit For
    Next A
    
    'agregamos los movimientos iniciales al rst temporal
    If RstSal.RecordCount <> 0 Then
        RstSal.MoveFirst
        For A = 1 To RstSal.RecordCount
            Rst.MoveFirst
            Rst.Find "idbal = " & NulosN(RstSal("idbal")) & ""
            If Rst.EOF = True Then
                Rst.AddNew
                Rst("idbal") = NulosN(RstSal("idbal"))
                Rst("descri") = NulosC(RstSal("descripcion"))
            End If
            Rst("salinideb") = NulosN(RstSal("totdebsol"))
            Rst("salinihab") = NulosN(RstSal("tothabsol"))
            Rst.Update
            
            RstSal.MoveNext
            If RstSal.EOF = True Then Exit For
        Next A
    End If
    
    'hacemos los calculos para determinar el saldo
    Rst.MoveFirst
    
    Dim RstDis As New ADODB.Recordset
    Dim RstCueBal As New ADODB.Recordset
    Dim xIdBalance As Integer
    
    xIdBalance = 1
    Rst.MoveFirst
    
    For A = 1 To Rst.RecordCount
        Rst("saldeb") = NulosN(Rst("movperdeb")) + NulosN(Rst("salinideb"))
        Rst("salhab") = NulosN(Rst("movperhab")) + NulosN(Rst("salinihab"))
        Rst("resultado") = NulosN(Rst("saldeb")) - NulosN(Rst("salhab"))
        
        RST_Busq RstCueBal, "SELECT DISTINCT con_balancedet.idcab, con_balancecab.id, con_planctas.dissegsal, con_balancecab.tipo, con_balancedet.idgru " _
            & " FROM con_planctas RIGHT JOIN (con_balancecab LEFT JOIN con_balancedet ON (con_balancecab.id = con_balancedet.idbal) AND (con_balancecab.idcab = con_balancedet.idcab)) " _
            & " ON con_planctas.id = con_balancedet.idcuenta Where (((con_balancedet.idcab) = " & xIdBalance & ") And ((con_balancecab.id) = " & Rst("idbal") & ") And ((con_planctas.dissegsal) = -1)) " _
            & " ORDER BY con_balancedet.idgru", xCon
        
        Dim xDebe, xHaber As Double
        xDebe = 0
        xHaber = 0
        If RstCueBal.RecordCount <> 0 Then
            If RstCueBal("tipo") = 2 Then Rst("resultado") = 0
            
            For B = 1 To RstCueBal.RecordCount
                RST_Busq RstDis, "SELECT Sum(IIf([impdebdol]<>0,[impdebdol]*[con_tc].[impven],[impdebsol])) AS debe, " _
                    & " Sum(IIf([imphabdol]<>0,[imphabdol]*[con_tc].[impven],[imphabsol])) AS haber FROM (con_planctas LEFT JOIN (con_diario LEFT JOIN con_tc " _
                    & " ON con_diario.fchdoc = con_tc.fecha) ON con_planctas.id = con_diario.idcue) RIGHT JOIN (con_balancecab LEFT JOIN con_balancedet " _
                    & " ON (con_balancecab.id = con_balancedet.idbal) AND (con_balancecab.idcab = con_balancedet.idcab)) ON con_planctas.id = con_balancedet.idcuenta " _
                    & " WHERE (((con_balancecab.idcab)=" & xIdBalance & ") AND ((con_balancecab.id)=" & NulosN(Rst("idbal")) & ") AND ((con_planctas.dissegsal)=-1) " _
                    & " AND ((con_balancedet.idgru)=" & NulosN(RstCueBal("idgru")) & "))", xCon

                If RstCueBal("tipo") = 1 Then
                    If (NulosN(RstDis("debe")) - NulosN(RstDis("haber"))) < 0 Then
                        'si el saldo es menor a 0 le sumamos al saldo principal el saldo negativo de la cuenta
                        'restamos al resultado el saldo negativos
                        If NulosN(RstCueBal("tipo")) = 1 Then
                            Rst("saldeb") = Rst("saldeb") - RstDis("debe")
                            Rst("salhab") = Rst("salhab") - RstDis("haber")
                            Rst("resultado") = (Rst("saldeb") - Rst("salhab"))
                        End If
                    End If
                Else
                    If (NulosN(RstDis("debe")) - NulosN(RstDis("haber"))) < 0 Then
                        Rst("resultado") = Rst("resultado") + Abs(NulosN(RstDis("debe")) - NulosN(RstDis("haber")))
                    End If
                End If
                    
                xDebe = xDebe + NulosN(RstDis("debe"))
                xHaber = xHaber + NulosN(RstDis("haber"))
                
                RstCueBal.MoveNext
                If RstCueBal.EOF = True Then Exit For
            Next B
        End If
        
        Rst.MoveNext
        If Rst.EOF = True Then Exit For
    Next A
    Set CargarTMPBalance = Rst
End Function

Function EjecutarFormula(FORMULA As String, Rst As ADODB.Recordset) As Double
    Dim xCad, xIdBal, xSigno As String
    Dim RstTmp As New ADODB.Recordset
    
    Set RstTmp = Rst
    
    Dim A As Integer
    Dim xTot, xTot1 As Double
    xIdBal = ""
    xTot = 0
    Dim EsPrimero As Boolean
    
    EsPrimero = True
    'If RstBal.RecordCount = 0 Then
    '    EjecutarFormula = 0
    '    Exit Function
    'End If
    
    For A = 1 To Len(Trim(FORMULA))
        xCad = ""
        xCad = xCad + Mid(Trim(FORMULA), A, 1)
        If xCad = "+" Or xCad = "-" Then
            xSigno = xCad
            If RstTmp.State <> 0 Then
                RstTmp.MoveFirst
                RstTmp.Find "idbal = " & NulosN(xIdBal) & ""
                If RstTmp.EOF = False Then
                    If RstTmp.EOF = False Then
                        If xCad = "+" Then xTot = xTot + Abs(NulosN(RstTmp("resultado")))
                        If xCad = "-" Then
                            If EsPrimero = True Then
                                xTot = Abs(NulosN(RstTmp("resultado"))) - xTot
                            Else
                                xTot = xTot - Abs(NulosN(RstTmp("resultado")))
                            End If
                        End If
                    End If
                End If
            End If
            EsPrimero = False
            xIdBal = ""
        Else
            xIdBal = xIdBal + xCad
        End If
    Next A
    
    If NulosC(xIdBal) <> "" Then
        If RstTmp.State <> 0 Then
            RstTmp.MoveFirst
            RstTmp.Find "idbal = " & NulosN(xIdBal) & ""
            If RstTmp.EOF = False Then
                If RstTmp.EOF = False Then
                    If xSigno = "+" Then xTot = xTot + Abs(NulosN(RstTmp("resultado")))
                    If xSigno = "-" Then xTot = xTot - Abs(NulosN(RstTmp("resultado")))
                End If
            End If
        End If
    End If
    EjecutarFormula = xTot
End Function

