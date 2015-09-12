VERSION 5.00
Begin VB.Form Form13 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alimentar Almacen"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9870
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form13"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   9870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "Salir"
      Height          =   375
      Left            =   6788
      TabIndex        =   15
      ToolTipText     =   "Haga click Aqui para volver al menu prinicpal"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   5588
      TabIndex        =   14
      ToolTipText     =   "Haga Cliick Aqui para buscar un registro"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   6908
      TabIndex        =   18
      ToolTipText     =   "Haga Clcik Aqui para cambiiar los valores de este registro"
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   1748
      TabIndex        =   17
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   3068
      TabIndex        =   16
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   5588
      TabIndex        =   13
      ToolTipText     =   "Haga Clcik Aqui para cambiiar los valores de este registro"
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4388
      TabIndex        =   12
      ToolTipText     =   "Haga Click aqui para deshacer el registro actual"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command11 
      Caption         =   ">|"
      Height          =   375
      Left            =   5708
      TabIndex        =   11
      ToolTipText     =   "Haga Click Aqui para ir al ultimo Registro"
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      Caption         =   ">"
      Height          =   375
      Left            =   4988
      TabIndex        =   10
      ToolTipText     =   "Haga Click Aqui para ir al siguiente Registro"
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Caption         =   "<"
      Height          =   375
      Left            =   4268
      TabIndex        =   9
      ToolTipText     =   "Haga Click Aqui para ir al anterior  Registro"
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "|<"
      Height          =   375
      Left            =   3548
      TabIndex        =   8
      ToolTipText     =   "Haga Click Aqui para ir al primer Registro"
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   330
      Left            =   6840
      TabIndex        =   7
      Top             =   1320
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   2355
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   1320
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   2355
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox Text1 
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
      Height          =   375
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label6"
      Height          =   255
      Left            =   9000
      TabIndex        =   20
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "En Almacen:"
      Height          =   255
      Left            =   7800
      TabIndex        =   19
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Cantidad"
      Height          =   255
      Left            =   5955
      TabIndex        =   6
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Material"
      Height          =   255
      Left            =   1635
      TabIndex        =   4
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha:"
      Height          =   255
      Left            =   1635
      TabIndex        =   2
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Numero de Reposicion"
      Height          =   255
      Left            =   5640
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Tabla As String
Public campoClave As String
Sub MostrarStock()
    Dim rs As New ADODB.Recordset
    Dim iSql  As String
    Dim codAlmacen As Integer
    codAlmacen = Val(Datos.MostrarCampo("Almacen", "codigoAl", "Descripcion='" & Combo1.Text & "'"))
    iSql = "select *from Almacen where CodigoAl=" & codAlmacen
    rs.Open iSql, Conexion
    If rs.EOF Then
        Label6.Caption = ""
    Else
        Label6.Caption = rs("stock")
    End If

End Sub
Sub Reponer_Almacen()
    Dim codAlmacen As Integer
    Dim Cantidad As Double
    codAlmacen = Val(Datos.MostrarCampo("Almacen", "codigoAl", "Descripcion='" & Combo1.Text & "'"))
    Cantidad = Val(Text3.Text)
    Conexion.Execute "Update Almacen set stock=stock+" & Cantidad & " where codigoAl=" & codAlmacen
End Sub

Sub CargarTablas()
    Tabla = "Reponer"
    campoClave = "CodReponer"
End Sub

Function SqlActualizacion()
    Dim iSql As String
    Dim rs As New ADODB.Recordset
    rs.Open "Select *from [" & Tabla & "] where [" & campoClave & "] =" & Text1.Text, Conexion
    If rs.EOF Then
        iSql = "insert into [" & Tabla & "] (" _
        & "CodReponer," _
        & "Fecha," _
        & "CodAlmacen," _
        & "Cantidad)" _
        & " values(" _
        & "" & Text1.Text & "," _
        & "'" & Text2.Text & "'," _
        & "" & Datos.MostrarCampo("Almacen", "CodigoAl", "Descripcion='" & Combo1.Text & "'") & "," _
        & "'" & Text3.Text & "')"
    Else
        iSql = "UPDATE [" & Tabla & "] set " _
            & "fecha='" & Text2.Text & "'" _
            & "Where CodReponer=" & Text1.Text & ""
    End If
    SqlActualizacion = iSql
    
End Function

Sub Bloquear(es As Boolean)
    Text1.Enabled = es
    Text2.Enabled = es
    Text3.Enabled = es
    Combo1.Enabled = es
    Label1.Enabled = es
    Label2.Enabled = es
    Label2.Enabled = es
    Label3.Enabled = es
    Label4.Enabled = es
    Label5.Enabled = es
    Label6.Enabled = es
End Sub
Sub Limpiar()
    Text1.Text = ""
    Text2.Text = Date
    Text3.Text = ""
    Combo1.Text = ""
    Label6.Caption = ""
End Sub
Sub Mostrar(rs As ADODB.Recordset)
On Error Resume Next
    Text1.Text = rs(0)
    Text2.Text = rs(1)
    Combo1.Text = rs(2)
    Text3.Text = rs(3)
End Sub
Private Sub Combo1_Click()
    MostrarStock
End Sub
Private Sub Command1_Click()
    Formularios.Nuevo Me
    Combo1.SetFocus
End Sub
Private Sub Command2_Click()
    Formularios.Guardar Me
    Reponer_Almacen
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
    Bloquear False
    CargarTablas
    Limpiar
    Datos.llenarCombo "select descripcion from almacen", Combo1
End Sub
Private Sub Text3_Validate(Cancel As Boolean)
    Text3.Text = Round(Text3.Text, 2)
End Sub
