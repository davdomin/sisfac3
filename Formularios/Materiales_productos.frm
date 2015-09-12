VERSION 5.00
Begin VB.Form Form11 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Materiales por Producto"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7815
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List3 
      Height          =   2310
      Left            =   5400
      TabIndex        =   15
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   4200
      TabIndex        =   12
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "-"
      Height          =   225
      Left            =   5760
      TabIndex        =   11
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      Default         =   -1  'True
      Height          =   225
      Left            =   5280
      TabIndex        =   10
      Top             =   1200
      Width           =   375
   End
   Begin VB.ListBox List2 
      Height          =   2310
      Left            =   4200
      TabIndex        =   9
      Top             =   2160
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2310
      Left            =   1080
      TabIndex        =   8
      Top             =   2160
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      Height          =   330
      Left            =   4200
      TabIndex        =   6
      Top             =   1800
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   1080
      TabIndex        =   5
      Top             =   1800
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   720
      Width           =   6855
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.Label LblUnidad 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   5400
      TabIndex        =   14
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Unidad"
      Height          =   495
      Left            =   5400
      TabIndex        =   13
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad:"
      Height          =   375
      Left            =   4200
      TabIndex        =   7
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Material de almacen"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion del Producto:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo del Producto"
      Height          =   255
      Left            =   4320
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public codProducto As Integer
Sub buscarAlmacen()
Dim codAlmacen As Integer
    codAlmacen = Val(Datos.MostrarCampo("Almacen", "codigoAl", "descripcion='" & Combo1.Text & "'"))
    LblUnidad.Caption = Datos.MostrarCampo("Almacen", "Unidad", "CodigoAl=" & codAlmacen)
End Sub
Sub BorrarDetalle()
    Dim codAlmacen
    Dim Indice As Integer
    Indice = List1.ListIndex
    If Indice <> -1 Then
        codAlmacen = Datos.MostrarCampo("Almacen", "CodigoAl", "Descripcion='" & List1.List(Indice) & "'")
        Conexion.Execute "delete from [Productos_Almacen] where codProducto=" & codProducto & " and CodAlmacen=" & codAlmacen
        MostrarMateriales
    End If
End Sub
Sub Seleccionar(Indice As Integer)
    List1.Selected(Indice) = True
    List2.Selected(Indice) = True
    List3.Selected(Indice) = True
End Sub
Sub MostrarMateriales()
    Dim rs As New ADODB.Recordset
    Dim iSql As String
    Dim Descripcion As String
    Dim Unidad As String
    rs.Open "select *from [Productos_Almacen] where codProducto=" & Val(codProducto), Conexion
    List1.Clear
    List2.Clear
    While Not rs.EOF
        Descripcion = Datos.MostrarCampo("Almacen", "Descripcion", "CodigoAL=" & rs("codAlmacen"))
        Unidad = Datos.MostrarCampo("Almacen", "Unidad", "CodigoAl=" & rs("codAlmacen"))
        
        List1.AddItem Descripcion
        List2.AddItem rs("cantidad")
        List3.AddItem Unidad
        rs.MoveNext
    Wend
    
    
    

End Sub
Sub Agregar_material()

    Dim iSql As String
    Dim codAlmacen As Integer
    codAlmacen = Datos.MostrarCampo("Almacen", "CodigoAL", "Descripcion='" & Combo1.Text & "'")
    iSql = "insert into [Productos_Almacen]  (codProducto,codAlmacen,cantidad) values(" & codProducto & "," & codAlmacen & "," & Text3.Text & ")"
    Conexion.Execute iSql
    Combo1.Text = ""
    Text3.Text = ""
    Combo1.SetFocus
    
    MostrarMateriales
End Sub
Sub Buscar_Productos()
    Text1.Text = codProducto
    Text2.Text = Datos.MostrarCampo("Productos", "Descripcion", "CodProducto=" & codProducto)
    
End Sub

Private Sub Combo1_Click()
buscarAlmacen
End Sub

Private Sub Command1_Click()
    If Text3.Text <> "" Then
        Agregar_material
    Else
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Command2_Click()
        BorrarDetalle
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Formularios.ColorLabels ColorLetras, Me
    MostrarMateriales
    Buscar_Productos
    Datos.llenarCombo "select descripcion from almacen", Combo1
End Sub

Private Sub List1_Click()
    Seleccionar List1.ListIndex
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        BorrarDetalle
    End If
End Sub

Private Sub List2_Click()
    Seleccionar List2.ListIndex
End Sub
