VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form form41 
   Caption         =   "Cuenta Corriente"
   ClientHeight    =   8565
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16950
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form41"
   ScaleHeight     =   8565
   ScaleWidth      =   16950
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Agregar Movimiento"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   6
      Top             =   360
      Width           =   13455
      Begin VB.CommandButton Command5 
         Caption         =   "Reportes"
         Height          =   375
         Left            =   5760
         TabIndex        =   20
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Salir"
         Height          =   375
         Left            =   8880
         TabIndex        =   19
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Borrar Ultimo"
         Height          =   375
         Left            =   7200
         TabIndex        =   18
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   4320
         TabIndex        =   16
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   2880
         TabIndex        =   15
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   840
         TabIndex        =   14
         Text            =   "Text2"
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   480
         Left            =   8400
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Text            =   "movimientosCuenta.frx":0000
         Top             =   360
         Width           =   4935
      End
      Begin VB.ComboBox Combo1 
         Height          =   360
         ItemData        =   "movimientosCuenta.frx":0006
         Left            =   5160
         List            =   "movimientosCuenta.frx":0010
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   360
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   840
         TabIndex        =   8
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   55050241
         CurrentDate     =   40794
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Monto"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Concepto"
         Height          =   255
         Left            =   7320
         TabIndex        =   11
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Movimiento"
         Height          =   255
         Left            =   3120
         TabIndex        =   9
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   2055
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grilla 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   2280
      Width           =   16215
      _ExtentX        =   28601
      _ExtentY        =   11033
      _Version        =   393216
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      ScrollBars      =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15120
      TabIndex        =   17
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Egresos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12240
      TabIndex        =   5
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Ingresos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10080
      TabIndex        =   4
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Concepto"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   3
      Top             =   2040
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Movimiento"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2040
      Width           =   1095
   End
End
Attribute VB_Name = "form41"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Agregar()
    Dim codMovimiento As Long
    Dim fecha As String
    Dim Ingresos As Double
    Dim Egresos As Double
    Dim Concepto As String
    Dim Monto As Double
    Dim sql As String
    Dim tipMovimiento As String
    Dim nuevoSaldo As Double

    
    
    codMovimiento = Datos.generarCodigo("movCuenta", "codMovimiento")
    fecha = Format(DTPicker1.value, "dd/mm/yyyy")
    Monto = Val(Text2.Text)
    If Monto <= 0 Then
        MsgBox "Monto Invalido"
        Exit Sub
    End If
    If Combo1.ListIndex = -1 Then
        MsgBox "Debe elegir un Tipo de Movimiento"
        Exit Sub
    End If
    If Combo1.ListIndex = 0 Then
        Ingresos = Monto
        Egresos = 0
        nuevoSaldo = saldoActual + Monto
        
    Else
    
        Egresos = Monto
        Ingresos = 0
        nuevoSaldo = saldoActual - Monto
    End If
    Concepto = Text1.Text
    tipMovimiento = Combo1.Text
    sql = "insert into movCuenta(codMovimiento,tipMovimiento,fecha,deposito,retiro,saldo,concepto)values(" _
    & "" & codMovimiento & "," _
    & "'" & tipMovimiento & "'," _
    & "'" & fecha & "'," _
    & "" & Ingresos & "," _
    & "" & Egresos & "," _
    & "" & nuevoSaldo & "," _
    & "'" & Concepto & "')"
    Conexion.Execute sql
    MsgBox "Agregado"
    
    
    
    
    
        
        
    
End Sub

Sub Limpiar()
    DTPicker1.value = Date
    Combo1.Text = ""
    Text1.Text = ""
    Text2.Text = ""
End Sub
Private Sub Command1_Click()
    Agregar
    mostrarGrilla
    Limpiar
    Combo1.SetFocus
End Sub

Private Sub Command2_Click()
    Limpiar
End Sub

Private Sub Command3_Click()
    borrarUltimo
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Command5_Click()
    Form42.Show vbModal
End Sub

Private Sub Form_Load()
    Formularios.ColorLabels ColorLetras, Me
    Frame1.BackColor = Me.BackColor
 inicializar
 mostrarGrilla

End Sub
Sub borrarUltimo()
    Dim rs As New ADODB.Recordset
    Dim sql As String
    Dim ultimoMov As Integer
    If MsgBox("Esta seguro que desea eliminar el ultimo movimiento", vbYesNo) = vbNo Then Exit Sub
    
    
    
    sql = "select top 10 * from movCuenta order by codMovimiento desc"
    rs.Open sql, Conexion
    If Not rs.EOF Then
        ultimoMov = Val(rs("codMovimiento"))
        Conexion.Execute "delete from movCuenta where codMovimiento=" & ultimoMov
        MsgBox "Movimiento Eliminado"
        mostrarGrilla
    Else
        MsgBox "Nada que Eliminar"
    End If
    
End Sub
Sub mostrarGrilla()
    Dim rs As New ADODB.Recordset
    Dim sql As String
    Dim fila As Integer
    Dim saldo As Double
    Dim primerMov As Integer
    
    
    fila = 1
    sql = "select top 10 * from movCuenta order by codMovimiento desc"
    rs.Open sql, Conexion
    If Not rs.EOF Then
        grilla.Rows = fila
        grilla.TextMatrix(fila - 1, 2) = "Saldo Final"
        grilla.TextMatrix(fila - 1, 5) = rs("Saldo")
        fila = fila + 1
    End If
    
    While Not rs.EOF
        saldo = rs("Saldo")
        grilla.Rows = fila
        grilla.TextMatrix(fila - 1, 0) = rs("fecha")
        grilla.TextMatrix(fila - 1, 1) = rs("tipMovimiento")
        grilla.TextMatrix(fila - 1, 2) = rs("concepto")
        primerMov = rs("codMovimiento")
        
        If rs("deposito") <> 0 Then
            grilla.TextMatrix(fila - 1, 3) = rs("deposito")
        Else
            grilla.TextMatrix(fila - 1, 3) = ""
        
        End If
        If rs("retiro") <> 0 Then
            grilla.TextMatrix(fila - 1, 4) = rs("retiro")
        Else
            grilla.TextMatrix(fila - 1, 4) = ""
        End If
        grilla.TextMatrix(fila - 1, 5) = rs("Saldo")
        
            

        
       
        fila = fila + 1
        rs.MoveNext
        
    Wend
    primerMov = primerMov - 1
    If primerMov = 0 Then
        saldo = 0
    Else
        rs.Close
        rs.Open "select * from movCuenta where codMovimiento=" & primerMov, Conexion
        If rs.EOF Then
            saldo = 0
        Else
            saldo = rs("Saldo")
        End If
    End If
    
    
    grilla.Rows = fila
    grilla.TextMatrix(fila - 1, 2) = "Saldo Inicial"
    grilla.TextMatrix(fila - 1, 5) = saldo
        
      
    
    
    
End Sub
Sub inicializar()


    grilla.Cols = 6
    Limpiar
    
    grilla.ColWidth(0) = 1400
    grilla.ColWidth(1) = 2500
    grilla.ColWidth(2) = 4500
    grilla.ColWidth(3) = 2500
    grilla.ColWidth(4) = 2500
    grilla.ColWidth(5) = 2500




End Sub

