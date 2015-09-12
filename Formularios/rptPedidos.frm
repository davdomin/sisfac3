VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form16 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Cierre de Pedidos"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   10845
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form16"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option3 
      Caption         =   "Boletas"
      Height          =   255
      Left            =   5760
      TabIndex        =   24
      Top             =   3960
      Width           =   1695
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Facturas"
      Height          =   255
      Left            =   3960
      TabIndex        =   23
      Top             =   3960
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Todos"
      Height          =   255
      Left            =   2160
      TabIndex        =   22
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Por Moneda"
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Excluir Notas de ventas"
      Height          =   255
      Left            =   8040
      TabIndex        =   20
      Top             =   1560
      Width           =   2415
   End
   Begin VB.ComboBox Combo3 
      Height          =   345
      Left            =   2520
      TabIndex        =   19
      Text            =   "Combo3"
      Top             =   2880
      Width           =   3615
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Utilidades"
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Libro de Ventas"
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ingresos"
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Detallado"
      Height          =   255
      Left            =   6720
      TabIndex        =   14
      Top             =   1560
      Width           =   1215
   End
   Begin VB.PictureBox cr1 
      Height          =   480
      Left            =   7560
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   25
      Top             =   4920
      Width           =   1200
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      Height          =   375
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Reporte de Ventas"
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Buscar"
      Height          =   375
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   2520
      TabIndex        =   9
      Top             =   2040
      Width           =   3615
   End
   Begin VB.ComboBox Combo2 
      Height          =   345
      Left            =   4800
      TabIndex        =   7
      Top             =   1320
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   2520
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   129761281
      CurrentDate     =   40288
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   129761281
      CurrentDate     =   40288
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuarios"
      Height          =   375
      Left            =   2520
      TabIndex        =   18
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   960
      TabIndex        =   13
      Top             =   2520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
      Height          =   495
      Left            =   2520
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Estatus de Trabajo"
      Height          =   255
      Left            =   4800
      TabIndex        =   6
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Estatus de Pago"
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta:"
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Desde"
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Formulas() As String
Function Filtro2() As String
    Dim f As String
    Dim fechaDesde As String
    Dim fechaHasta As String
    Dim StatusPago As String
    Dim Status As String
    Dim codCliente As Integer
    Dim codUsuario As Integer
    Dim cont As Integer
    
    fechaDesde = "#" & Format(DTPicker1.value, "mm/dd/yyyy") & "#"
    fechaHasta = "#" & Format(DTPicker2.value, "mm/dd/yyyy") & "#"
    StatusPago = "'" & Combo1.Text & "'"
    Status = "'" & Combo2.Text & "'"
    codCliente = Val(Label6.Caption)
    codUsuario = Val(Datos.MostrarCampo("usuarios", "codusuario", "alias='" & Combo3.Text & "'"))
    
    f = "{ReportePedidos.fecha}>=" & fechaDesde & " and {ReportePedidos.fecha}<=" & fechaHasta
    If Option2.value Then
        f = f & " AND {ReportePedidos.tipo} = 'Factura'"
    End If
    If Option3.value Then
        f = f & " AND {ReportePedidos.tipo} = 'Boleta'"
    End If
    cont = 0
    ReDim Preserve Formulas(cont)
    Formulas(cont) = "fechaDesde=" & fechaDesde
    
    cont = cont + 1
    ReDim Preserve Formulas(cont)
    Formulas(cont) = "fechaHasta=" & fechaHasta
    
    Filtro2 = f
    
End Function

Function filtro() As String
    Dim f As String
    Dim fechaDesde As String
    Dim fechaHasta As String
    Dim StatusPago As String
    Dim Status As String
    Dim codCliente As Integer
    Dim codUsuario As Integer
    Dim cont As Integer
    
    fechaDesde = "#" & Format(DTPicker1.value, "mm/dd/yyyy") & "#"
    fechaHasta = "#" & Format(DTPicker2.value, "mm/dd/yyyy") & "#"
    StatusPago = "'" & Combo1.Text & "'"
    Status = "'" & Combo2.Text & "'"
    codCliente = Val(Label6.Caption)
    codUsuario = Val(Datos.MostrarCampo("usuarios", "codusuario", "alias='" & Combo3.Text & "'"))
    
    f = "{pedidoenc.fecha}>=" & fechaDesde & " and {pedidoenc.fecha}<=" & fechaHasta
    cont = 0
    ReDim Preserve Formulas(cont)
    Formulas(cont) = "fechaDesde=" & fechaDesde
    
    cont = cont + 1
    ReDim Preserve Formulas(cont)
    Formulas(cont) = "fechaHasta=" & fechaHasta
    
    If Len(StatusPago) > 2 Then
        f = f + " And {PedidoEnc.statusPago}=" & StatusPago
        
        cont = cont + 1
        ReDim Preserve Formulas(cont)
        Formulas(cont) = "statusPago=" & StatusPago
    Else
        If Combo1.Text = "" Then
            f = f + " And {PedidoEnc.statusPago}<>'Anulado'"
            cont = cont + 1
        End If
    End If
    If Len(Status) > 2 Then
        f = f + " And {PedidoEnc.status}=" & Status
        
        cont = cont + 1
        ReDim Preserve Formulas(cont)
        Formulas(cont) = "status=" & Status
        
        
    End If
    If codCliente <> 0 Then
        f = f + " And {PedidoEnc.codCliente}=" & codCliente
        cont = cont + 1
        ReDim Preserve Formulas(cont)
        Formulas(cont) = "Cliente='" & Text1.Text & "'"
     
    End If
    
    If codUsuario <> 0 Then
        f = f + " And {PedidoEnc.codUsuario}=" & codUsuario
        cont = cont + 1
        ReDim Preserve Formulas(cont)
        Formulas(cont) = "Vend='" & Combo3.Text & "'"
     
    End If
    If Check2.value = 1 Then
        f = f + " And {PedidoEnc.enLetras}<>'Guia de Venta '"
    End If
    
    filtro = f
    
End Function

Private Sub Command1_Click()
    Datos.MostrarCatalogo "Select CodCliente,razonsocial from clientes"
    Label6.Caption = Catalogo.Resultado
    Text1.Text = Datos.MostrarCampo("Clientes", "razonsocial", "codCliente=" & Val(Label6.Caption))
End Sub

Private Sub Command2_Click()
    Dim f As String
    Dim Archivo As String
    Dim Detallado As Boolean
    f = filtro
    Detallado = Check1.value <> 1
    If Detallado Then
        Archivo = App.Path & "\reportes\pedidos.rpt"
    Else
        Archivo = App.Path & "\reportes\detallado.rpt"
    End If
    Datos.CargarReporte filtro, Archivo, Formulas
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Command4_Click()
    Dim f As String
    Dim Archivo As String
    Dim Detallado As Boolean
    f = filtro
    Detallado = Check1.value <> 1
    Archivo = App.Path & "\reportes\pagos.rpt"
    Datos.CargarReporte filtro, Archivo, Formulas

End Sub

Private Sub Command5_Click()
    Dim f As String
    Dim Archivo As String
    Dim Detallado As Boolean
    f = filtro
    Detallado = Check1.value <> 1
    Archivo = App.Path & "\reportes\ventas.rpt"
    
    Datos.CargarReporte filtro, Archivo, Formulas

End Sub

Private Sub Command6_Click()
    Dim f As String
    Dim Archivo As String
    Dim Detallado As Boolean
    f = filtro
    Detallado = Check1.value <> 1
    If Detallado Then
        Archivo = App.Path & "\reportes\ganancias.rpt"
    Else
        Archivo = App.Path & "\reportes\gananciasdet.rpt"
    End If
    Datos.CargarReporte filtro, Archivo, Formulas
End Sub

Private Sub Command7_Click()
    Dim f As String
    Dim Archivo As String
    Archivo = App.Path & "\reportes\rptMoneda.rpt"
    f = Filtro2
    Datos.CargarReporte Filtro2, Archivo, Formulas
End Sub

Private Sub Form_Load()
    Formularios.ColorLabels ColorLetras, Me
    DTPicker1.value = Date
    DTPicker2.value = Date
    Formularios.cargarStatusPago Combo1
    Formularios.cargarStatusTrabajo Combo2
    Check1.BackColor = Me.BackColor
    Check2.BackColor = Me.BackColor
    Option1.BackColor = Me.BackColor
    Option2.BackColor = Me.BackColor
    Option3.BackColor = Me.BackColor
    Datos.llenarCombo "select alias from usuarios", Combo3
End Sub
