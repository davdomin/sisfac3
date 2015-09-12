VERSION 5.00
Begin VB.Form Form12 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ordenes de Trabajo"
   ClientHeight    =   9195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10035
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9195
   ScaleWidth      =   10035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List6 
      BackColor       =   &H00CEFFFC&
      Height          =   1860
      Left            =   7320
      TabIndex        =   40
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "|<"
      Height          =   375
      Left            =   3720
      TabIndex        =   39
      ToolTipText     =   "Haga Click Aqui para ir al primer Registro"
      Top             =   8760
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Caption         =   "<"
      Height          =   375
      Left            =   4440
      TabIndex        =   38
      ToolTipText     =   "Haga Click Aqui para ir al anterior  Registro"
      Top             =   8760
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      Caption         =   ">"
      Height          =   375
      Left            =   5160
      TabIndex        =   37
      ToolTipText     =   "Haga Click Aqui para ir al siguiente Registro"
      Top             =   8760
      Width           =   615
   End
   Begin VB.CommandButton Command11 
      Caption         =   ">|"
      Height          =   375
      Left            =   5880
      TabIndex        =   36
      ToolTipText     =   "Haga Click Aqui para ir al ultimo Registro"
      Top             =   8760
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3600
      TabIndex        =   35
      ToolTipText     =   "Haga Click aqui para deshacer el registro actual"
      Top             =   8160
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   4800
      TabIndex        =   34
      ToolTipText     =   "Haga Clcik Aqui para cambiiar los valores de este registro"
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   8880
      TabIndex        =   33
      ToolTipText     =   "Haga Click aqui para borra definitivamente este registrop"
      Top             =   6840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   6120
      TabIndex        =   32
      ToolTipText     =   "Haga Cliick Aqui para buscar un registro"
      Top             =   8160
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Salir"
      Height          =   375
      Left            =   7320
      TabIndex        =   31
      ToolTipText     =   "Haga click Aqui para volver al menu prinicpal"
      Top             =   8160
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   2280
      TabIndex        =   29
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton QuitarE 
      Caption         =   "-"
      Height          =   255
      Left            =   7020
      TabIndex        =   28
      Top             =   6720
      Width           =   495
   End
   Begin VB.CommandButton AgregarE 
      Caption         =   "+"
      Height          =   255
      Left            =   6420
      TabIndex        =   27
      Top             =   6720
      Width           =   495
   End
   Begin VB.CommandButton QuitarM 
      Caption         =   "-"
      Height          =   255
      Left            =   9240
      TabIndex        =   26
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton AgregarM 
      Caption         =   "+"
      Height          =   255
      Left            =   8640
      TabIndex        =   25
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton Catalogo1 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   3060
      TabIndex        =   24
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   960
      TabIndex        =   23
      Top             =   8160
      Width           =   1215
   End
   Begin VB.ListBox List5 
      BackColor       =   &H00CEFFFC&
      Height          =   960
      Left            =   1620
      TabIndex        =   22
      Top             =   7080
      Width           =   4575
   End
   Begin VB.ComboBox Combo2 
      Height          =   345
      Left            =   1620
      TabIndex        =   20
      Text            =   "Combo1"
      Top             =   6720
      Width           =   4695
   End
   Begin VB.ListBox List4 
      BackColor       =   &H00CEFFFC&
      Height          =   1860
      Left            =   6120
      TabIndex        =   19
      Top             =   4320
      Width           =   1215
   End
   Begin VB.ListBox List3 
      BackColor       =   &H00CEFFFC&
      Height          =   1860
      Left            =   1620
      TabIndex        =   18
      Top             =   4320
      Width           =   4575
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   6240
      TabIndex        =   16
      Text            =   "Text5"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   1620
      TabIndex        =   15
      Text            =   "Combo1"
      Top             =   3960
      Width           =   4695
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00CEFFFC&
      Height          =   1185
      Left            =   2700
      TabIndex        =   13
      Top             =   2280
      Width           =   4695
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00CEFFFC&
      Height          =   1185
      Left            =   1620
      TabIndex        =   11
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   330
      Left            =   1620
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   330
      Left            =   3780
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "09:00:00pm"
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   1620
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "18/01/1986"
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label LblUnidad 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   7320
      TabIndex        =   41
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Label12"
      Height          =   255
      Left            =   7920
      TabIndex        =   30
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Empleado"
      Height          =   255
      Left            =   1620
      TabIndex        =   21
      Top             =   6480
      Width           =   4695
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cantidad"
      Height          =   255
      Left            =   6240
      TabIndex        =   17
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Material"
      Height          =   255
      Left            =   1620
      TabIndex        =   14
      Top             =   3720
      Width           =   4695
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Descripcion"
      Height          =   255
      Left            =   2700
      TabIndex        =   12
      Top             =   2040
      Width           =   4695
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cantidad"
      Height          =   255
      Left            =   1620
      TabIndex        =   10
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "label6"
      Height          =   255
      Left            =   1620
      TabIndex        =   9
      Top             =   1680
      Width           =   5775
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
      Height          =   495
      Left            =   900
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Pedido"
      Height          =   255
      Left            =   900
      TabIndex        =   6
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Hora:"
      Height          =   255
      Left            =   3060
      TabIndex        =   4
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      Height          =   255
      Left            =   900
      TabIndex        =   2
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de Orden:"
      Height          =   255
      Left            =   6240
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label13 
      Caption         =   "Unidad"
      Height          =   495
      Left            =   7320
      TabIndex        =   42
      Top             =   3720
      Width           =   1215
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public campoClave
Public Tabla
Sub buscarAlmacen()
Dim codAlmacen As Integer
    codAlmacen = Val(Datos.MostrarCampo("Almacen", "codigoAl", "descripcion='" & Combo1.Text & "'"))
    LblUnidad.Caption = Datos.MostrarCampo("Almacen", "Unidad", "CodigoAl=" & codAlmacen)
End Sub

Sub MaterialesProductos(codProducto As Integer, Cantidad As Double)
    Dim rs As New ADODB.Recordset
    Dim nombreMaterial As String
    Dim Unidad As String
    rs.Open "select  *from Productos_Almacen where codProducto= " & codProducto, Conexion
    
    While Not rs.EOF
        nombreMaterial = Datos.MostrarCampo("Almacen", "Descripcion", "codigoAl=" & rs("codAlmacen"))
        Unidad = Datos.MostrarCampo("Almacen", "Unidad", "CodigoAl=" & rs("codAlmacen"))
        List3.AddItem nombreMaterial
        List4.AddItem rs("cantidad") * Cantidad
        List6.AddItem Unidad
        rs.MoveNext
    Wend
    
    

End Sub

Sub CargarTablas()
    campoClave = "numeroO"
    Tabla = "ordenTEnc"
End Sub
Sub MostrarMateriales(NumeroO As Integer)
    Dim rs As New ADODB.Recordset
    Dim iSql As String
    Dim nombreMaterial As String
    Dim Unidad As String
    iSql = "select *from ordenTDet where numeroO=" & NumeroO
    rs.Open iSql, Conexion
    List3.Clear
    List4.Clear
    List6.Clear
    While Not rs.EOF
        nombreMaterial = Datos.MostrarCampo("Almacen", "Descripcion", "CodigoAl=" & rs("CodigoAl"))
        Unidad = Datos.MostrarCampo("Almacen", "Unidad", "CodigoAl=" & rs("CodigoAl"))
        List3.AddItem nombreMaterial
        List4.AddItem rs("cantidad")
        List6.AddItem Unidad
        rs.MoveNext
    Wend
End Sub
Sub MostrarEmpleados(NumeroO As Integer)
    Dim rs As New ADODB.Recordset
    Dim iSql As String
    Dim nombreEmpleado As String
    iSql = "select *from ordenEmpe where numeroO=" & NumeroO
    rs.Open iSql, Conexion
    List5.Clear
    While Not rs.EOF
        nombreEmpleado = Datos.MostrarCampo("Empleados", "Nombre", "CodEmpleado=" & rs("CodEmpleado"))
        List5.AddItem nombreEmpleado
        rs.MoveNext
    Wend
End Sub

Sub Mostrar(rs As ADODB.Recordset)
    Text1.Text = rs("numeroO")
    Text2.Text = rs("fecha")
    Text3.Text = rs("hora")
    Text4.Text = rs("codPedido")
    Label12.Caption = rs("Status")
    MostrarMateriales rs("numeroO")
    MostrarEmpleados rs("numeroO")
    
End Sub
Sub guardarEmpleados()
Dim NumeroO As Integer
Dim i As Integer
Dim codEmpleado As Integer
Dim iSql As String
    
    NumeroO = Val(Text1.Text)
    Conexion.Execute "delete from ordenEmpe where numeroO=" & NumeroO
    For i = 0 To List5.ListCount - 1
        codEmpleado = Datos.MostrarCampo("Empleados", "CodEmpleado", "Nombre='" & List5.List(i) & "'")
        iSql = "insert into OrdenEmpe(NumeroO,CodEmpleado) values(" _
        & "" & NumeroO & "," _
        & "" & codEmpleado & ")"
        Conexion.Execute iSql
    Next
End Sub
Sub ActualizarMateriales(codAlmacen As Integer, Cantidad As Double)
    Dim iSql As String
    iSql = "Update Almacen set stock=stock-" & Cantidad & " where codigoAl=" & codAlmacen
    Conexion.Execute iSql
    
End Sub
Sub guardarMateriales()
Dim NumeroO As Integer
Dim i As Integer
Dim codigoAl As Integer
Dim Cantidad As Double
Dim iSql As String
    NumeroO = Val(Text1.Text)
    Conexion.Execute "delete from ordenTDet where numeroO=" & NumeroO
    For i = 0 To List3.ListCount - 1
        codigoAl = Datos.MostrarCampo("almacen", "CodigoAl", "Descripcion='" & List3.List(i) & "'")
        Cantidad = Val(List4.List(i))
        iSql = "insert into OrdenTDet(numeroO,codigoAl,cantidad) values (" _
        & "" & NumeroO & "," _
        & "" & codigoAl & "," _
        & "" & Cantidad & ")"
        Conexion.Execute iSql
        ActualizarMateriales codigoAl, Cantidad
    Next i
    


End Sub
Sub Guardar()
    Dim A As VbMsgBoxResult
    Dim NumeroO As Integer
    Dim iSql As String
    Dim rs As New ADODB.Recordset
    NumeroO = Val(Text1.Text)
    If Val(Text4.Text) = 0 Then
        A = MsgBox("Error en el codigo de pedido", vbCritical)
       Exit Sub
    End If
    
    iSql = "select *from ordenTEnc where numeroO=" & NumeroO
    rs.Open iSql, Conexion
    If rs.EOF Then
        iSql = "insert into ordenTEnc(numeroO,fecha,hora,codPedido,status) values (" _
            & "" & NumeroO & "," _
            & "'" & Date & "'," _
             & "'" & Time & "'," _
            & "" & Val(Text4.Text) & "," _
            & "'" & "Pendiente" & "')"
    Else
        iSql = "update ordenTEnc set " _
        & "codPedido=" & Val(Text4.Text) & "' where numeroO =" & NumeroO
    End If
    Conexion.Execute iSql
    
    iSql = "Update pedidoEnc set status='En Ejecucion' where codPedido=" & Val(Text4.Text)
    Conexion.Execute iSql
    
    guardarMateriales
    guardarEmpleados
    
    
End Sub

'Sub Agregar()
Sub MostrarPedido()
    Dim codPedido As Integer
    Dim iSql As String
    Dim rs As New ADODB.Recordset
    Dim rd As New ADODB.Recordset
    
    codPedido = Val(Text4.Text)
    iSql = "select *from pedidoEnc where CodPedido=" & codPedido
    rs.Open iSql, Conexion
    If rs.EOF Then
        Label6.Caption = ""
        List1.Clear
        List2.Clear
    Else
        Label6.Caption = Datos.MostrarCampo("Clientes", "razonsocial", "codCliente=" & rs("codCliente"))
        rd.Open "select *from PedidoDet where CodPedido=" & codPedido, Conexion
        List1.Clear
        List2.Clear
        While Not rd.EOF
            If Datos.MostrarCampo("Productos", "seFabrica", "codProducto=" & rd("codProducto")) = "Si" Then
                List1.AddItem rd("cantidad")
                List2.AddItem Datos.MostrarCampo("productos", "descripcion", "codProducto=" & rd("codProducto"))
                MaterialesProductos rd("codProducto"), rd("cantidad")
            End If
                rd.MoveNext
            
        Wend
        
    End If
    
    
    
End Sub
Sub Limpiar()
    Text1.Text = ""
    Text2.Text = Date
    Text3.Text = Time
    Text4.Text = ""
    Text5.Text = ""
    
    Combo1.Text = ""
    Combo2.Text = ""
    Label6.Caption = ""
    Label12.Caption = ""
    List1.Clear
    List2.Clear
    List3.Clear
    List4.Clear
    List5.Clear
    List6.Clear
    LblUnidad.Caption = ""
    

End Sub
Sub Bloquear(es As Boolean)
    Label1.Enabled = es
    Label2.Enabled = es
    Label3.Enabled = es
    Label4.Enabled = es
    Label5.Enabled = es
    Label6.Enabled = es
    Label7.Enabled = es
    Label8.Enabled = es
    Label9.Enabled = es
    Label10.Enabled = es
    Label11.Enabled = es
    Label12.Enabled = es
    Label13.Enabled = es
    LblUnidad.Enabled = es
    List1.Enabled = es
    List2.Enabled = es
    List3.Enabled = es
    List4.Enabled = es
    List5.Enabled = es
    List6.Enabled = es
    Text1.Enabled = es
    Text2.Enabled = es
    Text3.Enabled = es
    Text4.Enabled = es
    Text5.Enabled = es
    Combo1.Enabled = es
    Combo2.Enabled = es
    Catalogo1.Enabled = es
    AgregarE.Enabled = es
    AgregarM.Enabled = es
    QuitarE.Enabled = es
    QuitarM.Enabled = es
End Sub
Sub Agregar_material()
    List3.AddItem Combo1.Text
    List4.AddItem Val(Text5.Text)
    Text5.Text = ""
    Combo1.Text = ""
    Combo1.SetFocus
    
End Sub
Sub agregar_Empleado()
    List5.AddItem Combo2.Text
    Combo2.Text = ""
    Combo2.SetFocus
End Sub

Private Sub AgregarE_Click()
    agregar_Empleado
End Sub

Private Sub AgregarM_Click()
    Agregar_material
End Sub


Private Sub Catalogo1_Click()
    Datos.MostrarCatalogo "select *from PedidosPendientes"
    Text4.Text = Catalogo.Resultado
End Sub

Private Sub Combo1_Click()
    buscarAlmacen
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        agregar_Empleado
    End If
End Sub

Private Sub Command1_Click()
    'Limpiar
    'Bloquear True
    'Text1.Text = Datos.generarCodigo("OrdenTEnc", "numeroO")
    Formularios.Nuevo Me
    Text4.SetFocus
End Sub


Private Sub Command10_Click()
    Formularios.Siguiente Me

End Sub

Private Sub Command11_Click()
    Formularios.Ultimo Me
        

End Sub

Private Sub Command2_Click()
    Guardar
    Formularios.Botones Me, 1
    Bloquear False
End Sub

Private Sub Command3_Click()
    Formularios.Cancelar Me
End Sub

Private Sub Command4_Click()
    Formularios.Modificar Me
End Sub

Private Sub Command6_Click()
    MostrarCatalogo "select *from [OrdenesClientes]"
    Text1.Text = Catalogo.Resultado
    Formularios.Buscar Me
End Sub

Private Sub Command7_Click()
    Unload Me
End Sub

Private Sub Command8_Click()
    Formularios.Primero Me
End Sub

Private Sub Command9_Click()
    Formularios.Anterior Me

End Sub

Private Sub Form_Load()
    Formularios.ColorLabels ColorLetras, Me
    Datos.llenarCombo "select descripcion from almacen", Combo1
    Datos.llenarCombo "select nombre from empleados", Combo2
    Bloquear False
    Limpiar
    CargarTablas
    Formularios.Cancelar Me
End Sub

Private Sub Text4_Change()
    MostrarPedido
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Agregar_material
    End If

End Sub
