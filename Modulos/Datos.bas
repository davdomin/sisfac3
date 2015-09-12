Attribute VB_Name = "Datos"
Public Conexion As New ADODB.Connection
Public pArrValues() As String
Sub Respaldo()
    On Error Resume Next
    Dim Nombre As String
    Dim fso As New FileSystemObject
    Nombre = Format(Now, "ddmmyy-hhmm") & ".resp"
    fso.CopyFile App.Path & "\Base de datos\datos.MDB", "d:\" & Nombre, True
    fso.CopyFile App.Path & "\Base de datos\datos.MDB", "e:\" & Nombre, True
    fso.CopyFile App.Path & "\Base de datos\datos.MDB", "f:\" & Nombre, True
    fso.CopyFile App.Path & "\Base de datos\datos.MDB", "g:\" & Nombre, True
    fso.CopyFile App.Path & "\Base de datos\datos.MDB", "h:\" & Nombre, True
    fso.CopyFile App.Path & "\Base de datos\datos.MDB", "i:\" & Nombre, True
    fso.CopyFile App.Path & "\Base de datos\datos.MDB", "o:\" & Nombre, True
End Sub
Sub Borrado_General()
If MsgBox("Esta Seguro de que desea comenzar el sistema desde 0", vbYesNo) = vbYes Then
    Respaldo
    Conexion.Execute "delete from ordeneje"
    Conexion.Execute "delete from ordenempe"
    Conexion.Execute "delete from ordenTDet"
    Conexion.Execute "delete from Pagos"
    Conexion.Execute "delete from pedidoEnt"
    Conexion.Execute "delete from pedidoDet"
    Conexion.Execute "delete from pedidoEnc"
    Conexion.Execute "delete from Reponer"
    Conexion.Execute "delete from ReponerPro"
    Conexion.Execute "delete from Guias"
    Conexion.Execute "delete from salidaCaja"
    Conexion.Execute "delete from cotizacionEnc"
    Conexion.Execute "delete from cotizacionDet"
    
    
    Conexion.Execute "delete from Factura"
    
    Conexion.Execute "delete from PresupuestoEnc"
    Conexion.Execute "delete from PresupuestoDet"
    Conexion.Execute "delete from ClienteVehiculos"
    Conexion.Execute "delete from Vehiculos"
    Conexion.Execute "delete from Modelos"
    Conexion.Execute "delete from Marcas"
        
    Conexion.Execute "delete from Gastos"
    Conexion.Execute "delete from Cuentas"
    Conexion.Execute "delete from Proveedores"
    Conexion.Execute "delete from Bancos"
    Conexion.Execute "delete from Entregas"
    Conexion.Execute "delete from PagosEmpDet"
    Conexion.Execute "delete from PagosEmp"
    Conexion.Execute "delete from PagDeuda"

    
    Conexion.Execute "delete from Clientes"
    Conexion.Execute "delete from Empleados"
    Conexion.Execute "delete from Productos_Almacen"
    Conexion.Execute "delete from ReponerPro"
    
    Conexion.Execute "delete from Productos"
    Conexion.Execute "delete from TipoP"
    Conexion.Execute "delete from cargos"
    Conexion.Execute "delete from categorias"
    Conexion.Execute "delete from tipoAl"
    Conexion.Execute "delete from movCuenta"
    Conexion.Execute "insert into proveedores(codProveedor,rif,RazonSocial)   values(0,'0','Empleados')"
    MsgBox "Operacion Completada"
End If

End Sub
Function inicializar()
If MsgBox("Esta Seguro de que desea inicializar el sistema", vbYesNo) = vbYes Then
    Respaldo
    Conexion.Execute "delete from ordeneje"
    Conexion.Execute "delete from ordenempe"
    Conexion.Execute "delete from ordenTDet"
    Conexion.Execute "delete from Pagos"
    Conexion.Execute "delete from pedidoEnt"
    Conexion.Execute "delete from pedidoDet"
    Conexion.Execute "delete from pedidoEnc"
    Conexion.Execute "delete from Reponer"
    Conexion.Execute "delete from gastos"
    Conexion.Execute "delete from Proveedores"
    Conexion.Execute "delete from ReponerPro"
    
    Conexion.Execute "delete from Guias"
    Conexion.Execute "delete from salidaCaja"
    Conexion.Execute "delete from cotizacionEnc"
    Conexion.Execute "delete from cotizacionDet"
    
    Conexion.Execute "delete from factura"
    Conexion.Execute "delete from boleta"
    
    Conexion.Execute "INSERT INTO proveedores(codProveedor,rif,RazonSocial)   VALUES (0,'0','Empleados')"
    
    MsgBox "Operacion Completada"
    inicializar = True
Else
    inicializar = False
End If
End Function
Sub CargarReporte(filtro As String, Archivo As String, Formulas() As String, Optional printer As Boolean)
Dim sisFacReport As New SReport
Dim cont As Integer
Dim fo As New FileSystemObject
Dim f As Variant
Dim data As String
Dim value As String
    sisFacReport.LoadExternal Archivo
    If printer Then
        sisFacReport.accion = "printer"
    End If
    If Not fo.FileExists(Archivo) Then Exit Sub
On Error Resume Next
    sisFacReport.setParam "nombreEmpresa", NombreEmpresa
    sisFacReport.setParam "rifEmpresa", RifEmpresa
    sisFacReport.setParam "SloganEmpresa", SloganEmpresa
    sisFacReport.setParam "direccionEmpresa", DireccionEmpresa
    On Error GoTo 0
    cont = 4
    For Each f In Formulas
        If Len(f) = 0 Then GoTo sig
        paraName = Split(f, "=")(0)
        paraValue = Replace(Split(f, "=")(1), "'", "")
        sisFacReport.setParam paraName, paraValue
sig:
        cont = cont + 1
    Next
imprimir:
    sisFacReport.SelectionFormula = filtro
    sisFacReport.PrintReport
End Sub
Public Function getConexionString(ByVal datasource As String)
    getConexionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & datasource & ";Persist Security Info=False"
End Function
Public Function getRutaDb()
Dim conf As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim ruta As String
    ChDir App.Path
    ChDrive App.Path
    conf.Open getConexionString("configuracion.mdb")
    rs.Open "SELECT * FROM conf", conf
    If Not rs.EOF Then
        ruta = rs("ruta")
    Else
        ruta = App.Path
    End If
    rutaBD = ruta
    getRutaDb = ruta
    Proyecto.cargarConfiguracion (rs)
    conf.Close
End Function
Sub Conectar()
     Conexion.Open getConexionString(getRutaDb())
End Sub
Function generarCodigo(Tabla As String, campoClave As String)
    Dim rs As New ADODB.Recordset
    rs.Open "SELECT MAX([" & campoClave & "]) FROM [" & Tabla & "]", Conexion
    If Not IsNumeric(rs(0)) Then
        generarCodigo = 1
    Else
        generarCodigo = rs(0) + 1
    End If
End Function
Sub MostrarCatalogo(sql As String)
    Catalogo.sql = sql
    Catalogo.Show vbModal
End Sub

Function MostrarCampo(Tabla As String, Campo As String, filtro As String)
On Error Resume Next
    Dim rs As New ADODB.Recordset
    rs.Open "select *from [" & Tabla & "] where " & filtro, Conexion
    If rs.EOF Then
        MostrarCampo = ""
    Else
        If Not IsNull(rs(Campo)) Then
            MostrarCampo = rs(Campo)
        Else
            MostrarCampo = ""
        End If
    
    End If
End Function
Function Existe(Tabla As String, filtro As String)
On Error Resume Next
    Dim rs As New ADODB.Recordset
    rs.Open "select *from [" & Tabla & "] where " & filtro, Conexion
    Existe = Not rs.EOF
End Function

Function llenarCombo(iSql As String, combo As ComboBox)
    Dim rs As New ADODB.Recordset
    rs.Open iSql, Conexion
    combo.Clear
    While Not rs.EOF
        If Not IsNull(rs(0)) Then combo.AddItem rs(0)
        rs.MoveNext
    Wend
    rs.Close

End Function

Sub CargarValores(sql As String)
    Dim lIndex As Long
    Dim sValue As String
    Dim t As New ADODB.Recordset
     t.Open sql, Conexion
    lIndex = 0
    While Not t.EOF
        sValue = t(1) & ""
        ReDim Preserve pArrValues(lIndex)
        If sValue <> "" Then
            pArrValues(lIndex) = sValue
        End If
        lIndex = lIndex + 1
        t.MoveNext
    Wend
       
    Close
End Sub

Public Function AutoCompletar_TextBox(textBox As textBox)
Dim i As Integer
On Error Resume Next
Dim posSelect As Integer
    Select Case (bKeyBack Or Len(textBox.Text) = 0)
        Case True
            bKeyBack = False
            Exit Function
    End Select
    With textBox
        For i = 0 To UBound(pArrValues)
            If InStr(1, pArrValues(i), .Text, vbTextCompare) = 1 Then
                posSelect = .SelStart
                .Text = pArrValues(i)
                .SelStart = posSelect
                .SelLength = Len(.Text) - posSelect
                Exit For
            End If
        Next i
    End With
End Function

Function saldoActual() As Double
    Dim rs As New ADODB.Recordset
    Dim ultimoMov As Long
    rs.Open "select max(codMovimiento)from movCuenta", Conexion
    If IsNull(rs(0)) Then
        saldoActual = 0
        Exit Function
    Else
        ultimoMov = rs(0)
    End If
    rs.Close
    rs.Open "select saldo from movCuenta where codMovimiento=" & ultimoMov, Conexion
    saldoActual = rs(0)
End Function
