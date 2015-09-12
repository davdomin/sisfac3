VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Almacen"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   9075
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      TabIndex        =   8
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton Catalogo2 
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   29
      ToolTipText     =   "Permite Buscar en el catalogo"
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Catalogo1 
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   28
      ToolTipText     =   "Permite Buscar en el catalogo"
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton Command11 
      Caption         =   ">|"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5430
      TabIndex        =   27
      ToolTipText     =   "Haga Click Aqui para ir al ultimo Registro"
      Top             =   5235
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4710
      TabIndex        =   26
      ToolTipText     =   "Haga Click Aqui para ir al siguiente Registro"
      Top             =   5235
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3990
      TabIndex        =   25
      ToolTipText     =   "Haga Click Aqui para ir al anterior  Registro"
      Top             =   5235
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "|<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3270
      TabIndex        =   24
      ToolTipText     =   "Haga Click Aqui para ir al primer Registro"
      Top             =   5235
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7710
      TabIndex        =   23
      ToolTipText     =   "Haga click Aqui para volver al menu prinicpal"
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6510
      TabIndex        =   22
      ToolTipText     =   "Haga Cliick Aqui para buscar un registro"
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Eliminar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5310
      TabIndex        =   21
      ToolTipText     =   "Haga Click aqui para borra definitivamente este registrop"
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Modificar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3990
      TabIndex        =   20
      ToolTipText     =   "Haga Clcik Aqui para cambiiar los valores de este registro"
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2790
      TabIndex        =   19
      ToolTipText     =   "Haga Click aqui para deshacer el registro actual"
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1590
      TabIndex        =   18
      ToolTipText     =   "Haga Click Aqui para guardar los cambios en este registro"
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   390
      TabIndex        =   9
      ToolTipText     =   "Haga Click Aqui para Agregar un Nuevo Registro"
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3600
      TabIndex        =   7
      Top             =   2940
      Width           =   500
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3570
      TabIndex        =   6
      Top             =   2580
      Width           =   1935
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3570
      TabIndex        =   5
      Top             =   2220
      Width           =   500
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      TabIndex        =   4
      Top             =   1860
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3570
      TabIndex        =   3
      Top             =   1500
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3570
      TabIndex        =   2
      Top             =   1140
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3570
      TabIndex        =   1
      Top             =   780
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6210
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   420
      Width           =   1455
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Unidad de Medición"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   32
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descripcion"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4080
      TabIndex        =   31
      Top             =   2940
      Width           =   2775
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descripcion"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4080
      TabIndex        =   30
      Top             =   2220
      Width           =   2775
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Categoria"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1410
      TabIndex        =   17
      Top             =   2940
      Width           =   2055
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Color"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1410
      TabIndex        =   16
      Top             =   2580
      Width           =   2055
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Almacen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1410
      TabIndex        =   15
      Top             =   2220
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Maximo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1410
      TabIndex        =   14
      Top             =   1860
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Minimo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1410
      TabIndex        =   13
      Top             =   1500
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Actual"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1410
      TabIndex        =   12
      Top             =   1140
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   11
      Top             =   780
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5250
      TabIndex        =   10
      Top             =   420
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Tabla As String
Public campoClave As String
Sub CargarTablas()
    Tabla = "Almacen"
    campoClave = "codigoAl"
End Sub

Function SqlActualizacion()
    Dim iSql As String
    Dim rs As New ADODB.Recordset
    rs.Open "Select *from [" & Tabla & "] where [" & campoClave & "] =" & Val(Text1.Text), Conexion
    If rs.EOF Then
        iSql = "insert into [" & Tabla & "] (" _
        & "codigoAl," _
        & "Descripcion," _
        & "Stock," _
        & "StockMin," _
        & "StockMax," _
        & "CodTipoAl," _
        & "Color," _
        & "CodCategoria," _
        & "Unidad)" _
        & " values(" _
        & "" & Text1.Text & "," _
        & "'" & Text2.Text & "'," _
        & "" & Text3.Text & "," _
        & "" & Text4.Text & "," _
        & "" & Text5.Text & "," _
        & "" & Text6.Text & "," _
        & "'" & Text7.Text & "'," _
        & "" & Text8.Text & "," _
        & "'" & Unidad & "')"
    Else
        iSql = "UPDATE [" & Tabla & "] set " _
            & "Descripcion='" & Text2.Text & "'," _
            & "Stock=" & Text3.Text & "," _
            & "StockMin=" & Text4.Text & "," _
            & "StockMax=" & Text5.Text & "," _
            & "CodTipoAl=" & Text6.Text & "," _
            & "Color='" & Text7.Text & "'," _
            & "CodCategoria=" & Text8.Text & "," _
            & "Unidad='" & Text9.Text & "'" _
            & "Where codigoAl=" & Text1.Text & ""
    End If
    
    SqlActualizacion = iSql
    
End Function
Sub MostrarTipoAl()
    Dim CodTipoAl As Integer
    CodTipoAl = Val(Text6.Text)
    Label9.Caption = Datos.MostrarCampo("tipoAl", "Descripcion", "CodTipoAl=" & CodTipoAl)
    End Sub
Sub MostrarCategoria()
    Dim codCategoria As Integer
    codCategoria = Val(Text8.Text)
    Label10.Caption = Datos.MostrarCampo("categorias", "Descripcion", "codCategoria=" & codCategoria)
End Sub

Sub bloquear(es As Boolean)
    Text1.Enabled = es
    Text2.Enabled = es
    Text3.Enabled = es
    Text4.Enabled = es
    Text5.Enabled = es
    Text6.Enabled = es
    Text7.Enabled = es
    Text8.Enabled = es
    Text9.Enabled = es
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
    Catalogo1.Enabled = es
    Catalogo2.Enabled = es
End Sub
Sub Limpiar()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Label9.Caption = ""
    Label10.Caption = ""
End Sub
Sub Mostrar(rs As ADODB.Recordset)
On Error Resume Next
    Text1.Text = rs(0)
    Text2.Text = rs(1)
    Text3.Text = rs(2)
    Text4.Text = rs(3)
    Text5.Text = rs(4)
    Text6.Text = rs(5)
    Text7.Text = rs(6)
    Text8.Text = rs(7)
    Text9.Text = rs(8)
End Sub

Private Sub Catalogo1_Click()
    MostrarCatalogo "select *from TipoAl"
    Text6.Text = Catalogo.Resultado
End Sub

Private Sub Catalogo2_Click()
    MostrarCatalogo "select *from Categorias"
    Text8.Text = Catalogo.Resultado
End Sub

Private Sub Command1_Click()
    Formularios.Nuevo Me
End Sub


Private Sub Command2_Click()
    Formularios.Guardar Me
End Sub

Private Sub Command3_Click()
    Formularios.cancelar Me
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
    Formularios.Botones Me, 1
    bloquear False
    CargarTablas
End Sub

Private Sub Text6_Change()
    MostrarTipoAl
End Sub

Private Sub Text8_Change()
    MostrarCategoria
End Sub
