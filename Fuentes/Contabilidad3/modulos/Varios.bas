Attribute VB_Name = "Varios"

Public Enum e_CategoriaConcepto
    e_Remuneracion = 1
    e_Aportacion = 2
    e_Descuento = 3
End Enum

'*****************************************************************************************************
'*****************************************************************************************************
'*****************************************************************************************************
'--para editor de formula

Public Sub eliminar_nodo(arb_obj As TreeView, Optional nd As Node, Optional key As String, Optional texto As String)
  
    If Not Nothing Is nd Then
        If MsgBox("Seguro desea eliminar " & nd.Text, vbQuestion + vbYesNo) = vbYes Then
            arb_obj.Nodes.Remove nd.key
        End If
        Exit Sub
    End If
    
    If Not IsMissing(key) Then
       If Trim(texto) <> "" Then
          If MsgBox(Trim("Seguro desea eliminar " & texto), vbQuestion + vbYesNo) = vbNo Then Exit Sub
       End If
       arb_obj.Nodes.Remove key
       Exit Sub
    End If

End Sub
'----------Agregar un node al treeview
Public Sub agregar_nodo(arb_obj As TreeView, _
                        relative As String, _
                        key As String, _
                        Text As String, _
                        Optional imagenA As String, _
                        Optional imagenB As String)
                        
    On Error Resume Next
    Dim nds As Node
    'Set nds = arb_obj.Nodes.Add()
    If imagenA <> "" And imagenB = "" Then
       Set nds = arb_obj.Nodes.Add(relative, tvwChild, key, Text, imagenA)
    ElseIf imagenA <> "" And imagenB <> "" Then
       Set nds = arb_obj.Nodes.Add(relative, tvwChild, key, Text, imagenA, imagenB)
    ElseIf imagenA = "" And imagenB <> "" Then
       imagenA = imagenB
       Set nds = arb_obj.Nodes.Add(relative, tvwChild, key, Text, imagenA)
    Else
       Set nds = arb_obj.Nodes.Add(relative, tvwChild, key, Text)
    End If
    Err.Clear
End Sub



Public Sub SacarRepetidasArray(X() As String, cant As Long)
    Dim i As Integer
    Dim J As Integer
    Dim Y As Variant
    Dim z As Variant
    ReDim Y(cant)
    ReDim z(cant)
    Y = X
    For i = 0 To cant
        For J = i + 1 To cant - 1
            If Y(i) = Y(J) Then
               Y(J) = -1
            End If
        Next
    Next
    
    
    J = 0
    
    For i = 0 To cant - 1
        If Y(i) <> "-1" Then
           ReDim Preserve X(J)
           X(J) = Y(i)
           J = J + 1
        End If
    Next
    cant = J
End Sub



Function MostrarFormulaEquivalente(IdCpto As Long) As String
    '===================================================================================================
    'Creado : 06/10/08 Por: Johan Castro
    'Propósito: procedimiento para obtener la formula expresado en la descripcion del conceptos involucrados
    '
    'Entradas:  IdCpto codigo del concepto que viene a ser la formula
    '
    'Resultados: Formula equivalente a las variables
    '
    'Otros: Para obtener esta formula se tendra que consultar los conceptos involucrados en la formula
    '       luego se reemplazara la variable por la descripcion del concepto
    '
    'Modificado :
    
    '===================================================================================================

    Dim rst As New ADODB.Recordset
    Dim nSQl As String
    Dim nFormula As String
    
    nSQl = "SELECT con_concepto.formula,con_concepto_1.id, con_concepto_1.descripcion,con_concepto_1.variable " _
        + vbCr + " FROM (con_concepto INNER JOIN con_conceptodet ON con_concepto.id = con_conceptodet.idcpto) INNER JOIN con_concepto AS con_concepto_1 ON con_conceptodet.idref = con_concepto_1.id " _
        + vbCr + " WHERE con_concepto.origen=-1 and con_concepto.id=" & IdCpto & ""

    RST_Busq rst, nSQl, xCon
    
    If rst.State = 0 Then Exit Function
    
    If rst.RecordCount <> 0 Then
        rst.MoveFirst
        nFormula = NulosC(rst("formula"))
        Do While Not rst.EOF
            '--reemplazar la variable por la descripcion
            nFormula = Replace(nFormula, rst("variable"), rst("descripcion"))
            
            rst.MoveNext
        Loop
            
    End If
        
    Set rst = Nothing
    
    MostrarFormulaEquivalente = nFormula

End Function

