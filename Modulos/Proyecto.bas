Attribute VB_Name = "Proyecto"
Public rutaBD As String
Public UsuarioSession As String
Public codUsuario As Integer
Public PIVa As Double
Public NombreEmpresa As String
Public SloganEmpresa As String
Public DireccionEmpresa As String
Public RifEmpresa As String
Public SerieImpresora As String
Public Produccion As Boolean
Public Facturacion As Boolean
Public ColorLetras As ColorConstants
Public ColorTextoLetras As Variant
Public ColorFondoLetras As Variant
Public ColorFondo As Variant
Public NivelEntro As String
Public bKeyBack As Boolean
Public inicialGuia As Boolean
Public inicialIvaIncluido As Boolean
Public inventario As Boolean
Public cambioSunat As Boolean
Public urlSunat As String



Private Sub setRutas()
    ChDir App.Path
    ChDrive App.Path
End Sub
Sub BorrarEspacios()
    Dim t As New ADODB.Recordset
    t.Open "select *from productos", Conexion
    While Not t.EOF
        Desc = Trim(t("descripcion"))
        nd = ""
        For i = 1 To Len(Desc)
            L = Mid(Desc, i, 1)
            ls = Mid(Desc, i + 1, 1)
                If L = " " And ls = " " Then
                
                Else
                    nd = nd & L
                
                End If
        Next
        If nd <> Desc Then
            sql = "update [productos] set descripcion='" & nd & "' where descripcion='" & Desc & "'"
            Conexion.Execute sql
        End If
            
        
        t.MoveNext
    Wend
    MsgBox "Listo"

End Sub
Public Sub cargarConfiguracion(rs)
On Error Resume Next
    inicialGuia = rs("guia")
    inicialIvaIncluido = rs("iva_incluido")
    cambioSunat = rs("cambio_sunat")
    urlSunat = rs("url_sunat")
    PIVa = (rs("iva") / 100)
    NombreEmpresa = rs("nombre_empresa")
    SloganEmpresa = rs("slogan_empresa")
    RifEmpresa = rs("rif_empresa")
    DireccionEmpresa = rs("direccion_empresa")
    inventario = rs("inventario")
End Sub
Sub Incializar()
    ColorLetras = vbBlack
    ColorFondo = RGB(144, 207, 214)
    Proyecto.ColorTextoLetras = ColorLetras
    Proyecto.ColorFondoLetras = RGB(255, 255, 255)
    Produccion = False
    Facturacion = False
    setRutas
    Datos.Conectar
    'BorrarEspacios
End Sub
Sub Main()
    Incializar
    Form1.Show
    Form1.Enabled = False
    Form17.Show
'    Form7.Show
End Sub
Function rellenarCeros(num)
    t = num
    While Len(t) <= 5
        t = "0" & t
    Wend
    rellenarCeros = t
End Function

