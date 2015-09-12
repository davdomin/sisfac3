VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form50 
   Caption         =   "Cotizacion"
   ClientHeight    =   8685
   ClientLeft      =   1485
   ClientTop       =   1875
   ClientWidth     =   11145
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form50"
   ScaleHeight     =   8685
   ScaleWidth      =   11145
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   4800
      Width           =   1335
   End
   Begin VB.ListBox List4 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   8520
      TabIndex        =   29
      Top             =   5160
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   330
      ItemData        =   "Cotizacion.frx":0000
      Left            =   8520
      List            =   "Cotizacion.frx":000A
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   4800
      Width           =   1335
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   7200
      TabIndex        =   27
      Top             =   5160
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   8880
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   435
      Left            =   1080
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   3960
      TabIndex        =   2
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   405
      Left            =   480
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1680
      Width           =   5655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   6240
      TabIndex        =   16
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   7440
      TabIndex        =   4
      Text            =   "Text5"
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Text            =   "Text5"
      Top             =   2280
      Width           =   4695
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      Text            =   "Text5"
      Top             =   2280
      Width           =   4695
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   4800
      Width           =   5415
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   5880
      TabIndex        =   9
      Top             =   4800
      Width           =   1335
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   480
      TabIndex        =   11
      Top             =   5160
      Width           =   5415
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   5880
      TabIndex        =   12
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Nuevo"
      Height          =   495
      Left            =   2760
      TabIndex        =   15
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   4080
      TabIndex        =   14
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   5400
      TabIndex        =   13
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Salir"
      Height          =   495
      Left            =   6720
      TabIndex        =   5
      Top             =   8040
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1815
      Left            =   480
      TabIndex        =   17
      Top             =   2760
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   3201
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Catalogo de Productos"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   120
      Top             =   -120
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   1
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Base Imp."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   30
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Moneda"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9000
      TabIndex        =   28
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de Cotizacion:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   26
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   25
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Hora:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   24
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "DNI/RUC del Cliente"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   23
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre o Razon Social del Cliente:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   22
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   21
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   20
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Producto"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   19
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Precio"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   18
      Top             =   4560
      Width           =   1815
   End
End
Attribute VB_Name = "Form50"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Total As Double
Sub MostrarCotizacion(codCliente As Integer)
    Dim rs As New ADODB.Recordset
    Dim filtro As String
    rs.Open "select *from cotizacionEnc where  codcliente=" & codCliente, Conexion
    If Not rs.EOF Then
    
        Text1.Text = rs("codCotizacion")
        filtro = "CodCliente=" & codCliente & ""
        Text4.Text = Datos.MostrarCampo("Clientes", "razonSocial", filtro)
        BuscarCliente
        Dim rsDet As New ADODB.Recordset
        rsDet.Open "select *from cotizacionDet where codCotizacion=" & Val(Text1.Text), Conexion
        List1.Clear
        List2.Clear
        List3.Clear
        List4.Clear
        While Not rsDet.EOF
            List1.AddItem Datos.MostrarCampo("productos", "descripcion", "codProducto=" & rsDet("codProducto") & "")
            List2.AddItem rsDet("precio")
            List3.AddItem conIva(rsDet("precio"))
            List4.AddItem (rsDet("Moneda"))
            rsDet.MoveNext
        Wend
        rsDet.Close
    End If
        
        
        
        
        
        
    

End Sub
Sub BuscarCliente()
    Dim filtro As String
    filtro = "RazonSocial='" & Text4.Text & "'"
    If Existe("Clientes", filtro) Then
        Text5.Text = Datos.MostrarCampo("Clientes", "rif", filtro)
        Text6.Text = Datos.MostrarCampo("Clientes", "direccion", filtro)
        Text7.Text = Datos.MostrarCampo("Clientes", "telefono", filtro)
    Else
        Text5.Text = ""
        Text6.Text = ""
        Text7.Text = ""
    End If
End Sub
Sub calcularBaseImponible()
    Text10.Text = Formularios.conIva(Val(Text9.Text))
End Sub
Sub mostrarGrilla()
    sqlbase = "select codigo,[descripcion articulo],[stock actual] from productosOrdenados2"
    If Text8.Text <> "" Then
        sqlbase = sqlbase + " where [descripcion articulo] like '%" & Text8.Text & "%'"
    End If
    
    
    
    
    Adodc1.ConnectionString = Conexion.ConnectionString
    Adodc1.RecordSource = sqlbase
    Adodc1.Refresh
    
    Set DataGrid1.DataSource = Adodc1
    DataGrid1.Columns(0).Visible = False
    DataGrid1.Columns(1).Width = 7000

End Sub
Sub bloquear(st As Boolean)
    Text1.Enabled = st
    Text2.Enabled = st
    Text3.Enabled = st
    Text4.Enabled = st
    Text5.Enabled = st
    Text6.Enabled = st
    Text7.Enabled = st
    Text8.Enabled = st
    Text9.Enabled = st
    Combo1.Enabled = st
    Text10.Enabled = st
    'Text11.Enabled = st
    'Text12.Enabled = st
    List1.Enabled = st
    List2.Enabled = st
    List3.Enabled = st
    'List4.Enabled = st
    Command1.Enabled = st
    DataGrid1.Enabled = st
End Sub
Sub Limpiar()
    Total = 0
    Text1.Text = ""
    Text2.Text = Date
    Text3.Text = Time
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Combo1.ListIndex = 0
    'Text10.Text = ""
    'Text11.Text = ""
    'Text12.Text = ""
    List1.Clear
    List2.Clear
    List3.Clear
    List4.Clear
    mostrarGrilla
End Sub
Sub cancelar()
    Limpiar
    bloquear False
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Agregar
    End If
End Sub

Private Sub Command1_Click()
        Catalogo.sql = "SELECT *FROM Clientes"
        Catalogo.Show vbModal

End Sub

Private Sub Command2_Click()
    cancelar
    bloquear True
    Text1.Text = Datos.generarCodigo("cotizacionEnc", "codCotizacion")
    Text4.SetFocus
End Sub
Sub GuardarCliente()
    Dim filtro As String
    filtro = "RazonSocial='" & Text4.Text & "'"
    If Not Existe("Clientes", filtro) Then
        Dim codCliente As Integer
        codCliente = Val(Datos.generarCodigo("Clientes", "CodCliente"))
        
        sql = "Insert Into Clientes (CodCliente,RazonSocial,cedrif,Direccion,Telefono) values(" _
        & "" & codCliente & "," _
        & "'" & Text4.Text & "'," _
        & "'" & Text5.Text & "'," _
        & "'" & Text6.Text & "'," _
        & "'" & Text7.Text & "')"
        Conexion.Execute sql
    End If
End Sub
Sub Guardar()
    Dim sql As String
    If Text4.Text = "" Then
        MsgBox "Debe seleccionar un  Cliente"
        Exit Sub
    End If
    Conexion.Execute "delete from CotizacionEnc where codCotizacion=" & Val(Text1.Text)
    Conexion.Execute "delete from CotizacionDet where codCotizacion=" & Val(Text1.Text)
    GuardarCliente
    Dim codCliente As Integer
    codCliente = Val(Datos.MostrarCampo("Clientes", "codCliente", "razonSocial='" & Text4.Text & "'"))
    
    sql = "insert into CotizacionEnc (codCotizacion,Fecha,hora,codCliente) values(" _
    & "" & Val(Text1.Text) & "," _
    & "'" & Date & "'," _
    & "'" & Time & "'," _
    & "" & codCliente & ")"
    Conexion.Execute sql
    
    For i = 0 To List1.ListCount - 1
        Dim CodDetalle As Integer
        Dim codProducto As Integer
        CodDetalle = Datos.generarCodigo("CotizacionDet", "CodDetalle")
        codProducto = Datos.MostrarCampo("productos", "codProducto", "descripcion='" & List1.List(i) & "'")
        sql = "insert into cotizacionDet(CodDetalle,CodCotizacion,CodProducto,Precio,Moneda) values(" _
        & "" & CodDetalle & "," _
        & "'" & Val(Text1.Text) & "'," _
        & "" & codProducto & "," _
        & "" & Val(List2.List(i)) & "," _
        & "'" & List4.List(i) & "')"
        
        Conexion.Execute sql
        
    Next
    
    
End Sub

Private Sub Command3_Click()
    Guardar
    cancelar
End Sub

Private Sub Command4_Click()
    cancelar
End Sub

Private Sub Command5_Click()
    Unload Me
End Sub

Private Sub DataGrid1_DblClick()
    Text8.Text = Adodc1.Recordset(1)
    Text9.SetFocus
    
End Sub

Private Sub Form_Load()
    Formularios.ColorLabels ColorLetras, Me
    Datos.llenarCombo "select Monedas from Monedas", Combo1
    cancelar
End Sub

Private Sub List1_Click()
Seleccionar List1.ListIndex
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        Eliminar List1.ListIndex
    End If
    
End Sub

Private Sub List2_Click()
Seleccionar List2.ListIndex
End Sub

Private Sub List3_Click()
Seleccionar List3.ListIndex
End Sub

Private Sub List4_Click()
Seleccionar List3.ListIndex
End Sub

Private Sub Text10_Change()
'calcularMontoTotal
End Sub
Sub LimpiarProductos()
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Combo1.Text = ""
 '   Text10.Text = ""
  '  Text11.Text = ""
    Text8.SetFocus
End Sub
Sub Agregar()
    If Not Existe("Productos", "descripcion='" & Text8.Text & "'") Then Exit Sub
    If Val(Text9.Text) <> 0 And Len(Combo1.Text) <> 0 Then
        List1.AddItem Text8.Text
        List2.AddItem Val(Text9.Text)
        List3.AddItem Val(Text10.Text)
        List4.AddItem Combo1.Text
'        Total = Total + Val(Text11.Text)
        LimpiarProductos
  '      Text12.Text = Total
    End If
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Agregar
    End If
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
    BuscarCliente
    Dim t As New ADODB.Recordset
    t.Open "select *from cotizacionEnc where codCliente=" & Val(Datos.MostrarCampo("clientes", "codCliente", "razonsocial='" & Text4.Text & "'")), Conexion
    If Not t.EOF Then
        MostrarCotizacion t("codCliente")
    End If
    
    
    
End Sub

Private Sub Text8_Change()
    mostrarGrilla
    Datos.AutoCompletar_TextBox Text8
End Sub

Private Sub Text8_GotFocus()
    Datos.CargarValores "select *from productos order by descripcion"
End Sub

Private Sub Text8_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyBack, vbKeyDelete
            Select Case Len(Text8.Text)
                Case Is <> 0
                    bKeyBack = True
            End Select
    End Select
End Sub


Private Sub Text4_Change()
    Datos.AutoCompletar_TextBox Text4
End Sub

Private Sub Text4_GotFocus()
    Datos.CargarValores "select CODCliente,RAZONSOCIAL from Clientes order by RAZONSOCIAL"
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyBack, vbKeyDelete
            Select Case Len(Text4.Text)
                Case Is <> 0
                    bKeyBack = True
            End Select
    End Select
    If KeyCode = 114 Then
        Command1_Click
    End If

End Sub


Private Sub Text9_Change()
calcularBaseImponible
End Sub

Private Sub Text9_GotFocus()
    
    Text9.Text = Datos.MostrarCampo("Productos", "Precio", "Descripcion='" & Text1.Text & "'")
    
    Text9.SelStart = 0
    Text9.SelLength = Len(Text9.Text)
End Sub
Sub Seleccionar(Index As Integer)
    List1.Selected(Index) = True
    List2.Selected(Index) = True
    List3.Selected(Index) = True
    List4.Selected(Index) = True
    
End Sub
Sub Eliminar(Index As Integer)
    List1.RemoveItem Index
    List2.RemoveItem Index
    List3.RemoveItem Index
    List4.RemoveItem Index
End Sub


Private Sub Text9_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Agregar
    End If
End Sub
