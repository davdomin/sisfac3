VERSION 5.00
Begin VB.Form Form31 
   Caption         =   "Cancelar Deudas"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   9885
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form31"
   ScaleHeight     =   7050
   ScaleWidth      =   9885
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo3 
      Height          =   360
      Left            =   3120
      TabIndex        =   31
      Top             =   5040
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      Caption         =   "Forma de Pago"
      Height          =   615
      Left            =   3120
      TabIndex        =   28
      Top             =   3720
      Width           =   3495
      Begin VB.OptionButton Option2 
         Caption         =   "Cheque"
         Height          =   240
         Left            =   1800
         TabIndex        =   30
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Efectivo"
         Height          =   255
         Left            =   360
         TabIndex        =   29
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   600
      TabIndex        =   26
      ToolTipText     =   "Haga Click Aqui para Agregar un Nuevo Registro"
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   1800
      TabIndex        =   25
      ToolTipText     =   "Haga Click Aqui para guardar los cambios en este registro"
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3000
      TabIndex        =   24
      ToolTipText     =   "Haga Click aqui para deshacer el registro actual"
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   4200
      TabIndex        =   23
      ToolTipText     =   "Haga Clcik Aqui para cambiiar los valores de este registro"
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   5520
      TabIndex        =   22
      ToolTipText     =   "Haga Click aqui para borra definitivamente este registrop"
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   6720
      TabIndex        =   21
      ToolTipText     =   "Haga Cliick Aqui para buscar un registro"
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Salir"
      Height          =   375
      Left            =   7920
      TabIndex        =   20
      ToolTipText     =   "Haga click Aqui para volver al menu prinicpal"
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "|<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   19
      ToolTipText     =   "Haga Click Aqui para ir al primer Registro"
      Top             =   6555
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   18
      ToolTipText     =   "Haga Click Aqui para ir al anterior  Registro"
      Top             =   6555
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   17
      ToolTipText     =   "Haga Click Aqui para ir al siguiente Registro"
      Top             =   6555
      Width           =   615
   End
   Begin VB.CommandButton Command11 
      Caption         =   ">|"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   16
      ToolTipText     =   "Haga Click Aqui para ir al ultimo Registro"
      Top             =   6555
      Width           =   615
   End
   Begin VB.TextBox Text5 
      Height          =   360
      Left            =   3120
      TabIndex        =   12
      Top             =   4560
      Width           =   5055
   End
   Begin VB.TextBox Text4 
      Height          =   360
      Left            =   3120
      TabIndex        =   11
      Top             =   3360
      Width           =   1695
   End
   Begin VB.ComboBox Combo2 
      Height          =   360
      Left            =   3120
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   2880
      Width           =   3015
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   3120
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   2400
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      Height          =   360
      Left            =   6960
      TabIndex        =   6
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   360
      Left            =   4080
      TabIndex        =   4
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   7920
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Banco"
      Height          =   255
      Left            =   720
      TabIndex        =   27
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000009&
      Caption         =   "Label9"
      Height          =   375
      Left            =   480
      TabIndex        =   15
      Top             =   840
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de Cuenta "
      Height          =   255
      Left            =   480
      TabIndex        =   14
      Top             =   480
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "&Numero del Cheque:"
      Height          =   255
      Left            =   720
      TabIndex        =   13
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "&Monto:"
      Height          =   255
      Left            =   720
      TabIndex        =   10
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&Concepto:"
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   2880
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Hora:"
      Height          =   255
      Left            =   6120
      TabIndex        =   5
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor:"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   2400
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de Pago"
      Height          =   255
      Left            =   5880
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
End
Attribute VB_Name = "Form31"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Tabla As String
Public campoClave As String
Sub MostrarNumeroCuenta()
    Dim codProveedor As Integer
    Dim codBanco As Integer
    Dim Resta As Double
    Dim Total As Double
    Dim Pagado As Double
    codProveedor = Datos.MostrarCampo("Proveedores", "codProveedor", "razonsocial='" & Combo1.Text & "'")
    
    Label9.Caption = Datos.MostrarCampo("Gastos", "nroCuenta", "codProveedor=" & codProveedor & " and total>pagado and concepto='" & Combo2.Text & "'")
'    codBanco = Datos.MostrarCampo("Cuentas", "codBanco", "NroCuenta='" & Label9.Caption & "'")
 '   Label9.Caption = Label9.Caption & " - " & Datos.MostrarCampo("Bancos", "Nombre", "codBanco=" & codBanco)

    Total = Val(Datos.MostrarCampo("Gastos", "Total", "codProveedor=" & codProveedor & " and total>pagado and concepto='" & Combo2.Text & "'"))
    Pagado = Val(Datos.MostrarCampo("Gastos", "Pagado", "codProveedor=" & codProveedor & " and total>pagado and concepto='" & Combo2.Text & "'"))
    Resta = Total - Pagado
    Text4.Text = Resta
End Sub

Sub CargarTablas()
    Tabla = "pagDeuda"
    campoClave = "codPagDeuda"
End Sub

Function SqlActualizacion()
    Dim iSql As String
    Dim rs As New ADODB.Recordset
    Dim codGasto As Integer
    Dim codBanco As Integer
    
    codProveedor = Datos.MostrarCampo("Proveedores", "codProveedor", "razonsocial='" & Combo1.Text & "'")
    codGasto = Val(Datos.MostrarCampo("Gastos", "codGasto", "codProveedor=" & codProveedor & " and total>pagado and concepto='" & Combo2.Text & "'"))
    
    rs.Open "Select *from [" & Tabla & "] where [" & campoClave & "] =" & Text1.Text, Conexion
    If rs.EOF Then
        If Option1.value Then
            iSql = "insert into [" & Tabla & "] (" _
            & "codPagDeuda," _
            & "Fecha," _
            & "Hora," _
            & "codGasto,FormaPago," _
            & "Monto)" _
            & " values(" _
            & "'" & Text1.Text & "'," _
            & "'" & Text2.Text & "'," _
            & "'" & Text3.Text & "'," _
            & "" & Val(codGasto) & "," _
            & "'Efectivo'," _
            & "" & Text4.Text & ")"
        Else
            codBanco = Val(Datos.MostrarCampo("Bancos", "codBanco", "nombre='" & Combo3.Text & "'"))
            iSql = "insert into [" & Tabla & "] (" _
            & "codPagDeuda," _
            & "Fecha," _
            & "Hora," _
            & "codGasto,FormaPago," _
            & "Cheque,codbanco," _
            & "Monto)" _
            & " values(" _
            & "'" & Text1.Text & "'," _
            & "'" & Text2.Text & "'," _
            & "'" & Text3.Text & "'," _
            & "" & Val(codGasto) & "," _
            & "'Cheque'," _
            & "'" & Text5.Text & "'," _
            & "" & Val(codBanco) & "," _
            & "" & Text4.Text & ")"
           ' MsgBox iSql
        End If
        Conexion.Execute "update Gastos set pagado=pagado +" & Val(Text4.Text) & " where codGasto=" & codGasto
    Else
        iSql = "UPDATE [" & Tabla & "] set " _
            & "Fecha='" & Text2.Text & "'," _
            & "Hora='" & Text3.Text & "'," _
            & "CodGasto=" & codGasto & "," _
            & "Cheque='" & Text5.Text & "'," _
            & "Monto=" & Val(Text4.Text) & "" _
            & "Where codPagDeuda=" & Text1.Text & ""
    End If
    SqlActualizacion = iSql
    
End Function

Sub Bloquear(es As Boolean)
    Text1.Enabled = es
    Text2.Enabled = es
    Text3.Enabled = es
    Text4.Enabled = es
    Combo1.Enabled = es
    Combo2.Enabled = es
    Text5.Enabled = es
    Combo3.Enabled = es
    Frame1.Enabled = es
    Label1.Enabled = es
    Label2.Enabled = es
    Label3.Enabled = es
    Label4.Enabled = es
    Label5.Enabled = es
    Label6.Enabled = es
    Label7.Enabled = es
    Label8.Enabled = es
    Label9.Enabled = es
    Option1_Click
End Sub
Sub Limpiar()
    Text1.Text = ""
    Text2.Text = Date
    Text3.Text = Time
    Text4.Text = ""
    Text5.Text = ""
    Combo1.Text = ""
    Combo2.Text = ""
    Label9.Caption = ""
    Combo3.Text = ""
    Option1.value = True
End Sub
Sub Mostrar(rs As ADODB.Recordset)
On Error Resume Next
Dim codProveedor As Integer
    Text1.Text = rs(0)
    Text2.Text = rs(1)
    Text3.Text = rs(2)
    Text4.Text = rs(4)
    Text5.Text = rs(5)
    codProveedor = Val(Datos.MostrarCampo("Gastos", "codProveedor", "codGasto=" & rs(3)))
    Combo1.Text = Datos.MostrarCampo("Proveedores", "razonsocial", "codProveedor=" & codProveedor)
    Combo2.Text = Datos.MostrarCampo("Gastos", "Concepto", "codGasto=" & rs(3))
End Sub

Sub CargarConceptos()
    Dim codProveedor As Integer
    codProveedor = Datos.MostrarCampo("Proveedores", "codProveedor", "razonsocial='" & Combo1.Text & "'")
    Datos.llenarCombo "select Concepto from Gastos  where codProveedor=" & codProveedor & " and pagado < total", Combo2
    If Combo2.ListCount > 0 Then
        Combo2.Text = Combo2.List(0)
        MostrarNumeroCuenta
    End If
End Sub

Private Sub Combo1_Click()
    CargarConceptos
End Sub

Private Sub Combo2_Click()
    MostrarNumeroCuenta
End Sub

Private Sub Command1_Click()
   Formularios.Nuevo Me
   Combo1.SetFocus
End Sub

Private Sub Command2_Click()
    If Option2.value Then
        GuardarBanco
    End If
    Formularios.Guardar Me
    Datos.llenarCombo "Select razonsocial from deudas", Combo1
    Datos.llenarCombo "Select nombre from bancos", Combo3

End Sub

Private Sub Command3_Click()
    Formularios.Cancelar Me
End Sub

Private Sub Command4_Click()
    Formularios.modificar Me
End Sub

Private Sub Command5_Click()
    Formularios.Eliminar Me
End Sub

Private Sub Command6_Click()
    MostrarCatalogo "select *from [" & Tabla & "]"
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
Private Sub Command10_Click()
    Formularios.Siguiente Me
End Sub

Private Sub Command11_Click()
    Formularios.Ultimo Me
End Sub

Private Sub Form_Load()
    Formularios.ColorLabels ColorLetras, Me
    Frame1.BackColor = Me.BackColor
    Option1.BackColor = Me.BackColor
    Option2.BackColor = Me.BackColor
    Limpiar
    Datos.llenarCombo "Select razonsocial from deudas", Combo1
    
    Formularios.Botones Me, 1
    Bloquear False
    CargarTablas
End Sub
Sub GuardarBanco()
    Dim t As New ADODB.Recordset
    t.Open "select *from bancos where nombre='" & Combo3.Text & "'", Conexion
    If Not t.EOF Then
        Conexion.Execute "insert into bancos(codBanco,Nombre) values(" _
        & "" & Datos.generarCodigo("bancos", "codBanco") & "," _
        & "'" & Combo3.Text & "')"
    End If


        
        
End Sub




Private Sub Option2_Click()

    Combo3.Enabled = True
    Text5.Enabled = True

End Sub

Private Sub Option1_Click()
    Text5.Text = ""
    Combo3.Text = ""
    Combo3.Enabled = False
    Text5.Enabled = False
End Sub
