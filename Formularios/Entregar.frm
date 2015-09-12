VERSION 5.00
Begin VB.Form Form30 
   Caption         =   "Entregar"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   9840
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form30"
   ScaleHeight     =   7650
   ScaleWidth      =   9840
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command11 
      Caption         =   ">|"
      Height          =   375
      Left            =   5520
      TabIndex        =   31
      ToolTipText     =   "Haga Click Aqui para ir al ultimo Registro"
      Top             =   7275
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      Caption         =   ">"
      Height          =   375
      Left            =   4800
      TabIndex        =   30
      ToolTipText     =   "Haga Click Aqui para ir al siguiente Registro"
      Top             =   7275
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Caption         =   "<"
      Height          =   375
      Left            =   4080
      TabIndex        =   29
      ToolTipText     =   "Haga Click Aqui para ir al anterior  Registro"
      Top             =   7275
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "|<"
      Height          =   375
      Left            =   3360
      TabIndex        =   28
      ToolTipText     =   "Haga Click Aqui para ir al primer Registro"
      Top             =   7275
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Salir"
      Height          =   375
      Left            =   7800
      TabIndex        =   27
      ToolTipText     =   "Haga click Aqui para volver al menu prinicpal"
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   6600
      TabIndex        =   26
      ToolTipText     =   "Haga Cliick Aqui para buscar un registro"
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   5400
      TabIndex        =   25
      ToolTipText     =   "Haga Click aqui para borra definitivamente este registrop"
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   4080
      TabIndex        =   24
      ToolTipText     =   "Haga Clcik Aqui para cambiiar los valores de este registro"
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2880
      TabIndex        =   23
      ToolTipText     =   "Haga Click aqui para deshacer el registro actual"
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   1680
      TabIndex        =   22
      ToolTipText     =   "Haga Click Aqui para guardar los cambios en este registro"
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   480
      TabIndex        =   21
      ToolTipText     =   "Haga Click Aqui para Agregar un Nuevo Registro"
      Top             =   6720
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   5520
      TabIndex        =   20
      Text            =   "Text2"
      Top             =   5880
      Width           =   1695
   End
   Begin VB.ComboBox Combo2 
      Height          =   360
      Left            =   1440
      TabIndex        =   17
      Top             =   5880
      Width           =   3975
   End
   Begin VB.ListBox List1 
      Height          =   1980
      Left            =   1440
      TabIndex        =   16
      Top             =   3360
      Width           =   6735
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   1560
      TabIndex        =   3
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label16 
      Caption         =   "Monto a pagar"
      Height          =   495
      Left            =   5520
      TabIndex        =   19
      Top             =   5520
      Width           =   2415
   End
   Begin VB.Label Label11 
      Caption         =   "Nombre del Empleado"
      Height          =   495
      Left            =   1440
      TabIndex        =   18
      Top             =   5520
      Width           =   3375
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Marca"
      Height          =   375
      Left            =   1440
      TabIndex        =   14
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Marca"
      Height          =   375
      Left            =   3840
      TabIndex        =   12
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Marca"
      Height          =   375
      Left            =   6120
      TabIndex        =   11
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label9 
      Caption         =   "Año"
      Height          =   495
      Left            =   6120
      TabIndex        =   10
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Marca"
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label7 
      Caption         =   "Color"
      Height          =   495
      Left            =   3720
      TabIndex        =   8
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Marca"
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Modelo"
      Height          =   495
      Left            =   1440
      TabIndex        =   6
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Marca"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Marca"
      Height          =   495
      Left            =   3480
      TabIndex        =   4
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Placa del vehiculo"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Numero de Entrega"
      Height          =   495
      Left            =   7440
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label15 
      Caption         =   "Fecha"
      Height          =   495
      Left            =   1440
      TabIndex        =   15
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label13 
      Caption         =   "Hora"
      Height          =   495
      Left            =   3840
      TabIndex        =   13
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "Form30"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Tabla As String
Public campoClave As String
Sub MostrarDetalles()
    Dim codPedido As Integer
    Dim codVehiculo As Integer
    Dim Acum As Double
    Dim rs As New ADODB.Recordset
    codVehiculo = Val(Datos.MostrarCampo("Vehiculos", "codVehiculo", "Placa='" & Combo1.Text & "'"))
    codPedido = Val(Datos.MostrarCampo("PedidoEnc", "codPedido", "Status='Pendiente' and codVehiculo=" & codVehiculo))
    rs.Open "select *from PedidoDet where codPedido=" & codPedido, Conexion
    List1.Clear
    Acum = 0
    While Not rs.EOF
        List1.AddItem Datos.MostrarCampo("Productos", "Descripcion", "codProducto=" & rs("codProducto"))
        Acum = Acum + Val(Datos.MostrarCampo("Productos", "PrecioPaga", "codProducto=" & rs("codProducto")))
        rs.MoveNext
    Wend
    Text2.Text = Acum
    
End Sub
Sub MostrarVehiculo()
    Dim Filtro As String
    Dim codModelo As Integer
    Dim codMarca As Integer
    Filtro = "Placa='" & Combo1.Text & "'"
    codModelo = Val(Datos.MostrarCampo("Vehiculos", "codModelo", Filtro))
    codMarca = Datos.MostrarCampo("Modelos", "codMarca", "CodModelo=" & codModelo)
    
    Label4.Caption = Datos.MostrarCampo("Marcas", "Nombre", "codMarca=" & codMarca & "")
    Label6.Caption = Datos.MostrarCampo("Modelos", "Nombre", "codModelo=" & codModelo & "")
    
    Label8.Caption = Datos.MostrarCampo("Vehiculos", "Color", Filtro)
    Label10.Caption = Datos.MostrarCampo("Vehiculos", "Año", Filtro)
    MostrarDetalles
    
End Sub

Sub CargarTablas()
    Tabla = "Entregas"
    campoClave = "codEntrega"
End Sub

Function SqlActualizacion()
    Dim iSql As String
    Dim rs As New ADODB.Recordset
    Dim codEmpleado As Integer
    Dim codPedido As Integer
    Dim codVehiculo As Integer
    
    codVehiculo = Val(Datos.MostrarCampo("Vehiculos", "codVehiculo", "Placa='" & Combo1.Text & "'"))
    codPedido = Val(Datos.MostrarCampo("PedidoEnc", "codPedido", "Status='Pendiente' and codVehiculo=" & codVehiculo))
    codEmpleado = Val(Datos.MostrarCampo("Empleados", "codEmpleado", "Nombre='" & Combo2.Text & "'"))
    
    
    rs.Open "Select *from [" & Tabla & "] where [" & campoClave & "] =" & Val(Text1.Text), Conexion
    codModelo = Datos.MostrarCampo("Modelos", "codModelo", "nombre='" & Combo2.Text & "'")
    If rs.EOF Then
        
        iSql = "insert into [" & Tabla & "] (" _
        & "codEntrega," _
        & "Fecha," _
        & "Hora," _
        & "codPedido," _
        & "codEmpleado," _
        & "Monto) values(" _
        & "" & Text1.Text & "," _
        & "'" & Date & "'," _
        & "'" & Time & "'," _
        & "" & codPedido & "," _
        & "" & codEmpleado & "," _
        & "" & Val(Text2.Text) & ")"
    Else
        iSql = "UPDATE [" & Tabla & "] set " _
            & "Fecha='" & Date & "'," _
            & "Hora='" & Time & "'," _
            & "codPedido=" & codPedido & "," _
            & "codEmpleado=" & codEmpleado & " " _
            & "Monto=" & Val(Text2.Text) & " " _
            & "Where codEntrega=" & Text1.Text & ""
    End If
    Conexion.Execute "Update PedidoEnc set status='Terminado' where codPedido=" & codPedido
    SqlActualizacion = iSql
End Function

Sub Bloquear(es As Boolean)
    Text1.Enabled = es
    Text2.Enabled = es
    Combo1.Enabled = es
    Combo2.Enabled = es
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
    Label14.Enabled = es
    Label15.Enabled = es
    Label16.Enabled = es
    List1.Enabled = es
End Sub
Sub Limpiar()
    Text1.Text = ""
    Text2.Text = ""
    Combo1.Text = ""
    Combo2.Text = ""
    Label14.Caption = Date
    Label12.Caption = Time
    Label4.Caption = ""
    Label6.Caption = ""
    Label8.Caption = ""
    Label10.Caption = ""
    List1.Clear
End Sub
Sub Mostrar(rs As ADODB.Recordset)
On Error Resume Next
Dim codVehiculo As Integer
    Text1.Text = rs(0)
    Label14.Caption = rs(1)
    Label12.Caption = rs(2)
    codVehiculo = Datos.MostrarCampo("pedidoEnc", "codVehiculo", "codPedido=" & rs(3))
    Combo1.Text = Datos.MostrarCampo("Vehiculos", "Placa", "codVehiculo=" & codVehiculo)
    MostrarVehiculo
    'Pendiente el Detalle
    Combo2.Text = Datos.MostrarCampo("Empleados", "Nombre", "codEmpleado=" & rs(4))
    Text2.Text = rs(5)
End Sub

Private Sub Combo1_Click()
    MostrarVehiculo
End Sub

Private Sub Command1_Click()
    Formularios.Nuevo Me
End Sub

Private Sub Command13_Click()
    SendKeys "{TAB}"
End Sub

Private Sub Command2_Click()
    Formularios.Guardar Me
    Datos.llenarCombo "Select placa from porReparar", Combo1
    Datos.llenarCombo "Select nombre from empleados", Combo2

End Sub

Private Sub Command3_Click()
    Formularios.Cancelar Me
End Sub

Private Sub Command4_Click()
    Formularios.Modificar Me
End Sub

Private Sub Command5_Click()
    Formularios.Eliminar Me
End Sub

Private Sub Command6_Click()
'    MostrarCatalogo "select codEntrega, razonsocial,direccion,telefono ,cedrif from [" & Tabla & "]"
    'Text1.Text = Catalogo.Resultado
    'Formularios.Buscar Me
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
Private Sub Command10_Click()
    Formularios.Siguiente Me
End Sub

Private Sub Command11_Click()
    Formularios.Ultimo Me
End Sub


Private Sub Form_Load()
    Formularios.ColorLabels ColorLetras, Me
    Limpiar
    Formularios.Botones Me, 1
    Bloquear False
    CargarTablas
    Datos.llenarCombo "Select distinct placa from porReparar", Combo1
    Datos.llenarCombo "Select nombre from empleados", Combo2
    
End Sub









