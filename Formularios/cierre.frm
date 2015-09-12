VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form44 
   Caption         =   "Cierre de caja"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   8985
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form44"
   ScaleHeight     =   5685
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   4560
      TabIndex        =   7
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   1680
      TabIndex        =   6
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   1335
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "cierre.frx":0000
      Top             =   3240
      Width           =   7455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1680
      Width           =   2655
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2910
      _ExtentX        =   5133
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   51576833
      CurrentDate     =   40824
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Observacion:"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Monto en Caja:"
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha del Cierre"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form44"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function montoCaja() As Double
    Dim rs As New ADODB.Recordset
    rs.Open "select sum(total) from pedidoEnc where fecha=#" & Format(Me.MonthView1.value, "mm/dd/yyyy") & "#", Conexion
    If rs.EOF Then
        montoCaja = -1
    Else
        If Not IsNumeric(rs(0)) Then
            montoCaja = -1
        Else
            montoCaja = Val(rs(0))
        End If
    End If
        
    
End Function
Private Sub Command1_Click()
Dim Fecha As String
Dim Hora As String
Dim SalidasCaja As Double
Dim montoReal As Double

Dim Observacion As String
Dim sql As String
Dim rs As New ADODB.Recordset
Dim codCierre As Integer
Dim codUsuario As Integer
    If MonthView1.value > Date Then
        MsgBox "No se pueden cuadrar cajas futuras"
        Exit Sub
    End If
    Fecha = MonthView1.value
    rs.Open "select *from  cierre where fecha=#" & Format(Fecha, "mm/dd/yyyy") & "#", Conexion
    If Not rs.EOF Then
        MsgBox "No se puede  volver a grabar un cierre de caja"
        Exit Sub
    End If
    rs.Close
    codCierre = Datos.generarCodigo("cierre", "codCierre")
    Hora = Time
    SalidasCaja = 0
    montoReal = Val(Text1.Text)
    Observacion = Text2.Text
    codUsuario = Val(Datos.MostrarCampo("usuarios", "codUsuario", "alias='" & Proyecto.UsuarioSession & "'"))
    sql = "insert into cierre(codCierre,Fecha,Hora,salidasCaja,montoCaja,montoReal,Observacion,codUsuario) values(" _
    & "" & codCierre & "," _
    & "'" & Fecha & "'," _
    & "'" & Hora & "'," _
    & "" & SalidasCaja & "," _
    & "" & montoCaja & "," _
    & "" & montoReal & "," _
    & "'" & Observacion & "'," _
    & "" & codUsuario & ")"
    If MsgBox("esta seguro que los datos sumistrados son correctos, esta accion no se podra deshacer", vbYesNo) = vbYes Then
        Conexion.Execute sql
    End If
    MsgBox "Cierre de Caja realizado"
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Formularios.ColorLabels ColorLetras, Me
    
    Me.MonthView1.value = Date
    Text1.Text = ""
    Text2.Text = ""
    
End Sub
