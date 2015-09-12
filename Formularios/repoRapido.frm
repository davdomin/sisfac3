VERSION 5.00
Begin VB.Form Form37 
   Caption         =   "Inicializado rapido"
   ClientHeight    =   6705
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10815
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form37"
   ScaleHeight     =   6705
   ScaleWidth      =   10815
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "Cambiar Nombre"
      Height          =   615
      Left            =   9360
      TabIndex        =   28
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   6360
      TabIndex        =   3
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Agregar Producto"
      Height          =   375
      Left            =   480
      TabIndex        =   22
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox Text7 
      Height          =   360
      Left            =   3480
      TabIndex        =   21
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   4200
      Width           =   5535
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Default         =   -1  'True
      Height          =   240
      Left            =   -1200
      TabIndex        =   18
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   6360
      TabIndex        =   4
      Top             =   3720
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Codigo de Barras"
      Height          =   255
      Left            =   7320
      TabIndex        =   16
      Top             =   120
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Seguido"
      Height          =   255
      Left            =   6000
      TabIndex        =   15
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   -720
      TabIndex        =   14
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   6000
      TabIndex        =   13
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Omitir"
      Height          =   495
      Left            =   4440
      TabIndex        =   8
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   3000
      TabIndex        =   7
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   6360
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   6360
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   6360
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   480
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   720
      Width           =   8775
   End
   Begin VB.Label Label10 
      Caption         =   "Label9"
      Height          =   375
      Left            =   480
      TabIndex        =   27
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   375
      Left            =   480
      TabIndex        =   26
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label Label8 
      Caption         =   "&Costo:"
      Height          =   255
      Left            =   4200
      TabIndex        =   25
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label lblEspecial 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7800
      TabIndex        =   24
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label lblNormal 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7800
      TabIndex        =   23
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "&Filtro"
      Height          =   375
      Left            =   2880
      TabIndex        =   20
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Codigo de barras"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2160
      TabIndex        =   19
      Top             =   4200
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "Ubicacion Almacen:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4200
      TabIndex        =   17
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Precio &Especial:"
      Height          =   255
      Left            =   4200
      TabIndex        =   12
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Precio &Normal:"
      Height          =   255
      Left            =   4200
      TabIndex        =   11
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Existencia:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4200
      TabIndex        =   10
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Descripcion:"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "Form37"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t As New ADODB.Recordset
Sub Actualizar()
    Conexion.Execute "update productos set " _
    & " stock=" & Val(Text2.Text) & "," _
    & " precio=" & Val(lblNormal.Caption) & "," _
    & " almacen='" & Text5.Text & "'," _
    & " Costo=" & Val(Text8.Text) & "," _
    & " barras='" & Text6.Text & "'," _
    & " precioPaga=" & Val(lblEspecial.Caption) & " " _
    & " where descripcion='" & Text1.Text & "'"
End Sub
Sub Mostrar()
On Error Resume Next
    Text1.Text = t("descripcion")
    Text2.Text = t("stock")
    Text3.Text = t("precio")
    Text4.Text = t("precioPaga")
    Text5.Text = t("almacen")
    Text6.Text = t("barras")
    Text8.Text = t("costo")
    Label9.Caption = t("talla")
    Label10.Caption = t("Color")
    
End Sub

Private Sub Command1_Click()
    Actualizar
    IrSiguiente
End Sub

Private Sub Command2_Click()
    IrSiguiente
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Command4_Click()
On Error Resume Next
    SendKeys "{TAB}"
End Sub

Private Sub Command5_Click()
    If Me.ActiveControl <> Command1 Then
    On Error Resume Next
        SendKeys "{TAB}"
    Else
        Command1_Click
    End If
End Sub

Private Sub Command6_Click()
    Form5.Show vbModal

End Sub

Private Sub Command7_Click()
Dim nombreNuevo As String
    nombreNuevo = InputBox("Int. Nombre", "Nombre Nuevo", Text1.Text)
    Conexion.Execute "update productos set descripcion='" & nombreNuevo & "' where descripcion='" & Text1.Text & "' and talla='" & Label9.Caption & "'"
    Text1.Text = nombreNuevo
End Sub

Private Sub Form_Load()
On Error Resume Next
    Label9.Caption = ""
    Label10.Caption = ""
    t.Close
    t.Open "select *from productos order by descripcion", Conexion
    Mostrar
End Sub
Sub IrSiguiente()
On Error Resume Next
    t.MoveNext
    Mostrar
    If t.EOF Then
        MsgBox "Completado"
    End If
    Text3.SetFocus
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    t.Close
End Sub

Private Sub Text2_GotFocus()
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
End Sub

Private Sub Text3_Change()
Dim sinIGV As Double
sinIGV = sinIva(Val(Text3.Text))
lblNormal.Caption = sinIGV
End Sub

Private Sub Text3_GotFocus()
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)

End Sub

Private Sub Text4_Change()
Dim sinIGV As Double
sinIGV = sinIva(Val(Text4.Text))
lblEspecial.Caption = sinIGV

End Sub

Private Sub Text4_GotFocus()
    Text4.SelStart = 0
    Text4.SelLength = Len(Text4.Text)
End Sub
Private Sub Text5_GotFocus()
    Text5.SelStart = 0
    Text5.SelLength = Len(Text5.Text)
End Sub

Private Sub Text7_Change()
'On Error Resume Next
    t.Close
    t.Open "select *from productos where descripcion like'" & Text7.Text & "%'  order by descripcion", Conexion
    Mostrar
End Sub

Private Sub Text8_GotFocus()
    Text8.SelStart = 0
    Text8.SelLength = Len(Text8.Text)
End Sub
